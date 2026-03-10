import xarray as xr
import numpy as np
import matplotlib.pyplot as plt
import cartopy.crs as ccrs
import cartopy.feature as cfeature
from pathlib import Path
import pandas as pd
from matplotlib import patheffects

# ===================== Configuration =====================
DATA_DIR = Path(__file__).resolve().parent

FORCINGS = {
    "GHG": "TraCE-21K-ghg-only.monthly.TS.nc",
    "ORB": "TraCE-21K-orb-only.monthly.TS.nc",
    "ICE": "TraCE-21K-ice-only.monthly.TS.nc",
    "FWF": "TraCE-21K-fwf-only.monthly.TS.nc",
}

PERIODS = [
    (11.0, 10.0),   # period 0: oldest, 11-10 ka
    (7.0, 6.0),     # period 1: 7-6 ka
    (4.5, 3.5),     # period 2: 4.5-3.5 ka
    (1.5, 0.5),     # period 3: youngest, 1.5-0.5 ka
]

# Difference-pair definition: (younger index, older index)
DIFF_PAIRS = [
    (3, 2),  # diffs[0]: 1.5-0.5 - 4.5-3.5 → 1-4 ka minus 4-6.5 ka
    (2, 1),  # diffs[1]: 4.5-3.5 - 7-6 → 4-6.5 ka minus 6.5-10.5 ka
    (1, 0),  # diffs[2]: 7-6 - 11-10 → 6.5-10.5 ka minus 10.5-11 ka
]

CMAP = "RdBu_r"
VMIN, VMAX = -2, 2   # Adjust according to the actual difference range

OUTPUT_FOLDER = Path(__file__).parent / "Trace_Diff_Maps_Three_Slices"
# ===============================================

def get_age_ka(time_values):
    return -time_values

def monthly_to_annual_mean(da):
    age_ka = get_age_ka(da["time"].values)
    year_labels = np.floor(age_ka).astype(int)
    da = da.assign_coords(year=("time", year_labels))
    annual = da.groupby("year").mean("time")
    annual_time = - (annual["year"] + 0.5)
    annual = annual.assign_coords(time=("year", annual_time.values))
    annual = annual.swap_dims({"year": "time"}).sortby("time")
    return annual.drop_vars("year")

def mean_in_period(annual_da, ka_start, ka_end):
    age = get_age_ka(annual_da["time"].values)
    mask = (age >= ka_end) & (age <= ka_start)
    if mask.sum() == 0:
        print(f"Warning: {ka_start}–{ka_end} ka  contains no data")
        return None
    return annual_da.isel(time=mask).mean("time")

def main():
    OUTPUT_FOLDER.mkdir(exist_ok=True)
    print("Output folder: ", OUTPUT_FOLDER)

    all_diffs = {}

    for name, fname in FORCINGS.items():
        path = DATA_DIR / fname
        if not path.is_file():
            print(f"File not found: {fname}")
            continue

        print(f"\nProcessing {name} ...")
        ds = xr.open_dataset(path, decode_times=False)
        ts = ds["TS"]

        annual = monthly_to_annual_mean(ts)

        period_means = []
        for start, end in PERIODS:
            m = mean_in_period(annual, start, end)
            period_means.append(m)

        diffs_for_name = []
        period_labels = ['1-4 ka', '4-6.5 ka', '6.5-10.5 ka']

        file_labels = ['1-4ka_minus_4-6.5ka', '4-6.5ka_minus_6.5-10.5ka', '6.5-10.5ka_minus_10.5-11ka']

        for i, (young_idx, old_idx) in enumerate(DIFF_PAIRS):
            if period_means[young_idx] is None or period_means[old_idx] is None:
                continue

            diff = period_means[young_idx] - period_means[old_idx]
            diffs_for_name.append(diff)

            file_label = file_labels[i]

            nc_out = OUTPUT_FOLDER / f"{name}_diff_{file_label}.nc"
            diff.to_dataset(name="TS_diff").to_netcdf(nc_out)
            print(f"  Saved nc: {nc_out.name}")

            df = pd.DataFrame({
                "lat": np.repeat(diff.lat.values, len(diff.lon)),
                "lon": np.tile(diff.lon.values, len(diff.lat)),
                "diff_K": diff.values.ravel()
            })
            xlsx_out = OUTPUT_FOLDER / f"{name}_diff_{file_label}.xlsx"
            df.to_excel(xlsx_out, index=False)
            print(f"  Saved xlsx: {xlsx_out.name}")

            fig = plt.figure(figsize=(10, 5))
            ax = fig.add_subplot(111, projection=ccrs.PlateCarree())
            im = ax.pcolormesh(diff.lon, diff.lat, diff,
                               cmap=CMAP, vmin=VMIN, vmax=VMAX,
                               transform=ccrs.PlateCarree(), shading='auto')
            ax.add_feature(cfeature.COASTLINE, lw=0.6)
            ax.set_global()
            ax.set_title(f"{name} {period_labels[i]} minus older period", fontsize=13, pad=12)

            cbar = fig.colorbar(im, ax=ax, orientation='vertical', shrink=0.6, pad=0.04,
                                label="ΔT (K)", extend='both')
            plt.tight_layout(rect=[0, 0.03, 0.92, 0.97])

            png_out = OUTPUT_FOLDER / f"{name}_diff_{file_label}.png"
            plt.savefig(png_out, dpi=300, bbox_inches='tight')
            plt.close()
            print(f"  Saved png: {png_out.name}")

        all_diffs[name] = diffs_for_name

    print("\nAll done. Three difference maps have been generated.")

    # ================== Combined stacked figure ==================
    forcing_order = list(FORCINGS.keys())
    col_labels = ['1-4 ka', '4-6.5 ka', '6.5-10.5 ka']

    fig = plt.figure(figsize=(18, 12), constrained_layout=True)
    gs = fig.add_gridspec(nrows=4, ncols=3, hspace=0.1, wspace=0.05)

    for row, name in enumerate(forcing_order):
        diffs = all_diffs.get(name, [])
        if len(diffs) != 3:
            print(f"Warning: {name} is missing difference fields; skipping the combined row")
            continue

        # Corrected order: left 1-4 (diffs[0]), middle 4-6.5 (diffs[1]), right 6.5-10.5 (diffs[2])
        col_diffs = [diffs[0], diffs[1], diffs[2]]

        for col in range(3):
            ax = fig.add_subplot(gs[row, col], projection=ccrs.PlateCarree())

            diff = col_diffs[col]
            im = ax.pcolormesh(diff.lon, diff.lat, diff,
                               cmap=CMAP, vmin=VMIN, vmax=VMAX,
                               transform=ccrs.PlateCarree(), shading='auto')
            ax.add_feature(cfeature.COASTLINE, lw=0.6)
            ax.set_global()

            ax.set_title(f"{name} ({col_labels[col]})", fontsize=11, pad=10)

            # Upper-left panel label with white outline
            label = chr(65 + row * 3 + col)  # A, B, C ... L
            ax.text(0.02, 0.98, label, transform=ax.transAxes,
                    fontsize=14, fontweight='bold', va='top', ha='left', color='white',
                    path_effects=[patheffects.withStroke(linewidth=4, foreground='black')])

            cbar = fig.colorbar(im, ax=ax, orientation='vertical', shrink=0.6, pad=0.04,
                                label="ΔT (K)", extend='both')
            cbar.ax.tick_params(labelsize=8)

    combined_png = OUTPUT_FOLDER / "combined_diff_maps_corrected.png"
    plt.savefig(combined_png, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Saved final combined figure: {combined_png}")

if __name__ == "__main__":
    main()