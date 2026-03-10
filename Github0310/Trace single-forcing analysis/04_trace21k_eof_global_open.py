from pathlib import Path
import numpy as np
import pandas as pd
import xarray as xr

CFG = {
    "DATA_DIR": Path(__file__).resolve().parent,
    "OUT_DIR": Path(__file__).resolve().parent,
    "OUT_DIRNAME": "EOF_space_outputs",
    "FILES": {
        "GHG": "TraCE-21K-ghg-only.monthly.TS.nc",
        "ORB": "TraCE-21K-orb-only.monthly.TS.nc",
        "ICE": "TraCE-21K-ice-only.monthly.TS.nc",
        "FWF": "TraCE-21K-fwf-only.monthly.TS.nc",
    },
    "VAR": "TS",
    "AGE_RANGE_KA": (0.0, 12.0),
    "N_MODES": 5,
    "DOMAINS": {
        "GLOBAL": {"lat_min": -90.0, "lat_max": 90.0},
        "NH":     {"lat_min": 0.0,   "lat_max": 90.0},
    },
    "OUT_PREFIX": "TraCE21K_TS_SpatialEOF_annual_anom_wsqrtcos",
    "TIME_SCALES": {
        "annual": 1,
        "centennial": 100,
        "millennial": 1000,
    },
}

def ensure_dir(p: Path) -> Path:
    p.mkdir(parents=True, exist_ok=True)
    return p

def time_to_age_kaBP(time_kaBP: xr.DataArray) -> xr.DataArray:
    return -time_kaBP

def select_time_window(ds: xr.Dataset, age_min: float, age_max: float) -> xr.Dataset:
    t = ds["time"]
    tmin = -age_max
    tmax = -age_min
    m = (t >= tmin) & (t <= tmax) & (t <= 0.0)
    return ds.sel(time=m)

def add_year_coord(ds: xr.Dataset) -> xr.Dataset:
    age_ka = time_to_age_kaBP(ds["time"])
    age_year = age_ka * 1000.0
    year_bp = np.floor(age_year.values + 1e-9).astype(np.int32)
    return ds.assign_coords(
        age_kaBP=("time", age_ka.values),
        year_bp=("time", year_bp)
    )

def monthly_to_annual_mean(da: xr.DataArray) -> xr.DataArray:
    annual = da.groupby("year_bp").mean("time")
    annual = annual.assign_coords(age_kaBP=("year_bp", annual["year_bp"].values.astype(np.float64) / 1000.0))
    annual = annual.swap_dims({"year_bp": "age_kaBP"})
    annual = annual.sortby("age_kaBP")
    return annual

def average_to_scale(annual_da: xr.DataArray, scale_years: int) -> xr.DataArray:
    if scale_years == 1:
        return annual_da

    age = annual_da["age_kaBP"].values
    bin_edges = np.arange(0, 12 + scale_years/1000.0, scale_years/1000.0)
    bin_centers = (bin_edges[:-1] + bin_edges[1:]) / 2

    binned = []
    for i in range(len(bin_centers)):
        mask = (age >= bin_edges[i]) & (age < bin_edges[i+1])
        if mask.sum() > 0:
            mean_slice = annual_da.isel(age_kaBP=mask).mean(dim="age_kaBP")
            binned.append(mean_slice)
        else:
            dummy = annual_da.isel(age_kaBP=0).copy()
            dummy[:] = np.nan
            binned.append(dummy)

    da_binned = xr.concat(binned, dim="age_kaBP")
    da_binned = da_binned.assign_coords(age_kaBP=("age_kaBP", bin_centers))
    return da_binned

def lat_weight_sqrtcos(lat: xr.DataArray) -> xr.DataArray:
    w = np.sqrt(np.cos(np.deg2rad(lat)))
    return xr.DataArray(w, coords={"lat": lat}, dims=("lat",))

def eof_from_annual_field(da: xr.DataArray, n_modes: int):
    anom = da - da.mean("age_kaBP")

    w_lat = lat_weight_sqrtcos(anom["lat"])
    anom_w = anom * w_lat

    Xw = anom_w.stack(space=("lat", "lon")).transpose("age_kaBP", "space").values
    Xw = np.asarray(Xw, dtype=np.float64)
    Xw -= np.nanmean(Xw, axis=0, keepdims=True)

    space_good_mask = None
    if np.isnan(Xw).any():
        good = np.isfinite(Xw).all(axis=0)
        Xw = Xw[:, good]
        space_good_mask = good
        print(f"[WARN] NaNs found. Removed {(~good).sum()} space points.")

    ntime = Xw.shape[0]
    if ntime < 10:
        raise RuntimeError("Too few time samples after annual averaging/time window selection.")

    C = (Xw.T @ Xw) / (ntime - 1.0)
    eigvals_all, eigvecs_all = np.linalg.eigh(C)
    idx = np.argsort(eigvals_all)[::-1]
    eigvals_all = eigvals_all[idx]
    eigvecs_all = eigvecs_all[:, idx]

    eigvals = eigvals_all[:n_modes]
    EOFw = eigvecs_all[:, :n_modes]
    PCs = Xw @ EOFw

    total = np.sum(eigvals_all[eigvals_all > 0])
    varfrac = eigvals / total if total > 0 else np.full_like(eigvals, np.nan)

    w2d = (w_lat * xr.ones_like(anom["lon"])).transpose("lat", "lon")
    w_flat = w2d.stack(space=("lat", "lon")).values.astype(np.float64)

    if space_good_mask is not None:
        EOFw_full = np.full((w_flat.size, n_modes), np.nan, dtype=np.float64)
        EOFw_full[space_good_mask, :] = EOFw
        EOFw = EOFw_full

    EOF_unw_flat = EOFw / w_flat[:, None]
    nlat = int(anom["lat"].size)
    nlon = int(anom["lon"].size)
    EOF_maps = EOF_unw_flat.T.reshape((n_modes, nlat, nlon))

    for k in range(n_modes):
        m = np.nanmean(EOF_maps[k])
        if np.isfinite(m) and m < 0:
            EOF_maps[k] *= -1
            PCs[:, k] *= -1

    age = da["age_kaBP"].values.astype(np.float64)
    return age, PCs, EOF_maps, eigvals, varfrac

def save_outputs_excel(out_xlsx: Path, results: dict, n_modes: int, scale_name: str):
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        meta = [
            ["OUT_PREFIX", CFG["OUT_PREFIX"]],
            ["DATA_DIR", str(CFG["DATA_DIR"])],
            ["OUT_DIR", str(out_xlsx.parent)],
            ["AGE_RANGE_KA", str(CFG["AGE_RANGE_KA"])],
            ["TIME_SCALE", scale_name],
            ["N_MODES", n_modes],
            ["VAR", CFG["VAR"]],
        ]
        pd.DataFrame(meta, columns=["key", "value"]).to_excel(writer, sheet_name="meta", index=False)

        for domain in results:
            rows = []
            for forcing in results[domain]:
                r = results[domain][forcing]
                for k in range(n_modes):
                    rows.append([forcing, f"EOF{k+1}", float(r["eigvals"][k]), float(r["varfrac"][k])])
            df_var = pd.DataFrame(rows, columns=["Forcing", "Mode", "Eigenvalue", "VarFraction"])
            df_var["VarPercent"] = df_var["VarFraction"] * 100.0
            df_var.to_excel(writer, sheet_name=f"{domain}_Var", index=False)

            for forcing in results[domain]:
                r = results[domain][forcing]
                age = r["age"]
                PCs = r["PCs"]
                df_pc = pd.DataFrame({"Age_kaBP": age})
                for k in range(n_modes):
                    df_pc[f"PC{k+1}"] = PCs[:, k]
                df_pc.to_excel(writer, sheet_name=f"{domain}_{forcing}", index=False)

def save_eof_patterns_netcdf(out_nc: Path, lat: np.ndarray, lon: np.ndarray, eof_maps: np.ndarray, scale_name: str):
    n_modes = eof_maps.shape[0]
    da = xr.DataArray(
        eof_maps.astype(np.float32),
        dims=("mode", "lat", "lon"),
        coords={"mode": np.arange(1, n_modes + 1), "lat": lat, "lon": lon},
        name="EOF_pattern",
        attrs={"note": f"Unweighted EOF patterns ({scale_name}) from annual TS anomalies with sqrt(cos(lat)) weighting."}
    )
    da.to_dataset().to_netcdf(out_nc)

def main():
    data_dir = CFG["DATA_DIR"]
    out_root = ensure_dir(CFG["OUT_DIR"] / CFG["OUT_DIRNAME"])
    age_min, age_max = CFG["AGE_RANGE_KA"]
    n_modes = int(CFG["N_MODES"])

    for scale_name, scale_years in CFG["TIME_SCALES"].items():
        print(f"\n=== Processing time scale: {scale_name} ({scale_years} year mean) ===")
        scale_root = ensure_dir(out_root / f"EOF_space_outputs_{scale_name}")

        results = {dom: {} for dom in CFG["DOMAINS"].keys()}

        for forcing, fname in CFG["FILES"].items():
            fpath = data_dir / fname
            if not fpath.exists():
                print(f"File not found: {fpath}")
                continue

            print(f"  Processing {forcing}: {fname}")
            ds = xr.open_dataset(fpath, decode_times=False)
            if CFG["VAR"] not in ds:
                print(f"Variable {CFG['VAR']} not found in {fname}")
                ds.close()
                continue

            ds2 = select_time_window(ds, age_min, age_max)
            ds3 = add_year_coord(ds2)
            annual = monthly_to_annual_mean(ds3[CFG["VAR"]])

            data_to_eof = average_to_scale(annual, scale_years)

            for dom, dcfg in CFG["DOMAINS"].items():
                lat_min = dcfg["lat_min"]
                lat_max = dcfg["lat_max"]
                data_dom = data_to_eof.sel(lat=slice(lat_min, lat_max))

                print(f"    Domain {dom}: lat[{lat_min},{lat_max}], shape={data_dom.shape}")

                try:
                    age, PCs, EOF_maps, eigvals, varfrac = eof_from_annual_field(data_dom, n_modes)
                except Exception as e:
                    print(f"    EOF computation failed ({forcing}, {dom}, {scale_name}): {e}")
                    continue

                results[dom][forcing] = {"age": age, "PCs": PCs, "eigvals": eigvals, "varfrac": varfrac}

                out_nc = scale_root / f"{CFG['OUT_PREFIX']}_{forcing}_{dom}_EOFpatterns_{age_min:g}-{age_max:g}ka.nc"
                save_eof_patterns_netcdf(out_nc, data_dom["lat"].values, data_dom["lon"].values, EOF_maps, scale_name)

                print(f"      saved EOF patterns nc: {out_nc.name}")
                print(f"      Var% (EOF1..{n_modes}): " + ", ".join([f"{100*v:.2f}%" for v in varfrac]))

            ds.close()

        out_xlsx = scale_root / f"{CFG['OUT_PREFIX']}_PCs_{age_min:g}-{age_max:g}ka_GLOBAL_NH.xlsx"
        save_outputs_excel(out_xlsx, results, n_modes, scale_name)

        print(f"  {scale_name} completed!")
        print(f"  Output folder: {scale_root}")
        print(f"  Excel PCs: {out_xlsx}")

    print("\n[All done]")
    print("The three time scales are saved in their corresponding folders.")

if __name__ == "__main__":
    main()