import pandas as pd
import numpy as np
import pymc as pm
import pytensor.tensor as at
import arviz as az
from pathlib import Path


def run_hbgp(exclude_stal=None):
    # ===== 1. Read data =====
    script_dir = Path(__file__).resolve().parent
    data_file = script_dir / "stalagmite_temp.xlsx"

    df = pd.read_excel(data_file)

    # Diagnostic prints
    print("Original df shape:", df.shape)
    print("df columns:", list(df.columns))
    print("First 5 values in Sample:", df["Sample"].head().tolist())

    # Keep only the four target stalagmites
    mask = df["Sample"].str.contains("SN35|SN42|SN24-4|SN24-3", na=False)
    data = df[mask].dropna(subset=["Age", "Temperature"]).copy()

    print("Filtered data shape:", data.shape)
    print("Unique Sample values after filtering:", data["Sample"].unique()[:5])  # Print only the first 5

    # Extract stalagmite group names: SN35, SN42, SN24-4, SN24-3
    stal_ids = data["Sample"].str.extract(r"(SN35|SN42|SN24-4|SN24-3)", expand=False)
    print("First 5 extracted stal_ids:", stal_ids.head().tolist())
    print("Number of non-null stal_ids:", stal_ids.notna().sum())

    data["stal_base"] = stal_ids

    # Filter data if one stalagmite is excluded
    if exclude_stal is not None:
        data = data[data["stal_base"] != exclude_stal]
        print(f"Excluding {exclude_stal}. Remaining stalagmites: {data['stal_base'].unique()}")

    # Optional sparse sampling (uncomment if needed; e.g., max_points_per_stal=40)
    # def subsample_per_stal(df, max_points=40):
    #     subsampled = []
    #     for stal in df['stal_base'].unique():
    #         sub = df[df['stal_base'] == stal].sort_values('Age')
    #         if len(sub) > max_points:
    #             indices = np.linspace(0, len(sub)-1, max_points, dtype=int)
    #             sub = sub.iloc[indices]
    #         subsampled.append(sub)
    #     return pd.concat(subsampled)
    # data = subsample_per_stal(data, 40)
    # print(f"Total observations after subsampling: {len(data)}")

    # Map to integer indices 0..J-1
    stal_unique = data["stal_base"].unique()
    stal_to_idx = {name: i for i, name in enumerate(stal_unique)}
    data["stal_idx"] = data["stal_base"].map(stal_to_idx).astype(int)

    print("Stalagmite mapping:", stal_to_idx)

    # ===== 2. Prepare numpy arrays =====
    age = data["Age"].values.astype(float)
    temp = data["Temperature"].values.astype(float)
    stal_idx = data["stal_idx"].values.astype(int)

    # For numerical stability: center Age and scale by 1000 years (~ka)
    age_mean = age.mean()
    age_scale = 1000.0
    x = (age - age_mean) / age_scale  # GP input variable

    X = x[:, None]  # (N, 1)
    y = temp  # (N,)
    J = len(stal_unique)  # Number of stalagmites (full=4, leave-one-out=3)

    # ===== 3. Build the lightweight hierarchical GP model =====
    with pm.Model() as model:
        # --- GP hyperparameters ---
        eta = pm.HalfNormal("eta", sigma=5.0)
        ell = pm.Gamma("ell", alpha=2.0, beta=1.0)

        # Stalagmite offsets: hierarchical prior
        sigma_a = pm.HalfNormal("sigma_a", sigma=2.0)
        a_raw = pm.Normal("a_raw", mu=0.0, sigma=1.0, shape=J)
        a = pm.Deterministic("a", a_raw * sigma_a)

        # Stalagmite-specific observation noise
        sigma_y = pm.HalfNormal("sigma_y", sigma=2.0, shape=J)

        # --- Gaussian process kernel: ExpQuad (RBF) ---
        cov_func = eta ** 2 * pm.gp.cov.ExpQuad(input_dim=1, ls=ell)

        # Define the latent GP (shared cave temperature curve f(t))
        gp = pm.gp.Latent(cov_func=cov_func)

        # Generate GP values f(t_i) at observation points
        f = gp.prior("f", X=X)

        # Mean at each observation = cave temperature + corresponding stalagmite offset
        mu = f + a[stal_idx]

        # Observation model
        y_obs = pm.Normal("y_obs", mu=mu, sigma=sigma_y[stal_idx], observed=y)

        # ===== 4. Lightweight sampling settings =====
        trace = pm.sample(
            draws=400,
            tune=400,
            target_accept=0.95,
            chains=2,
            cores=1,
            progressbar=True
        )

        # ===== 5. Predict the cave-level temperature curve on a regular time grid =====
        # Predict f(t) (shared cave component) without stalagmite offsets
        age_pred = np.linspace(0, 11.2, 200) * 1000  # Convert to years BP
        x_pred = (age_pred - age_mean) / age_scale
        x_pred = x_pred[:, None]

        f_pred = gp.conditional("f_pred", Xnew=x_pred)

        post_pred = pm.sample_posterior_predictive(
            trace,
            var_names=["f_pred"],
            progressbar=True
        )

    # ===== 6. Convert predictions back to Age (years) and write CSV =====
    f_samples = post_pred.posterior_predictive["f_pred"].values.reshape(-1, 200)  # (n_samples, 200)

    f_mean = f_samples.mean(axis=0)
    f_lower = np.percentile(f_samples, 2.5, axis=0)
    f_upper = np.percentile(f_samples, 97.5, axis=0)

    out = pd.DataFrame({
        "Age": age_pred,
        "CaveTemp_mean": f_mean,
        "CaveTemp_lower95": f_lower,
        "CaveTemp_upper95": f_upper
    })

    # Output filename
    if exclude_stal is None:
        filename = "ShennongCave_HierGP_full.csv"
    else:
        filename = f"ShennongCave_HierGP_leave_out_{exclude_stal}.csv"
    output_file = script_dir / filename
    out.to_csv(output_file, index=False)
    print(f"Output file generated: {output_file}")


if __name__ == "__main__":
    # Run the full model
    run_hbgp(exclude_stal=None)

    # Run leave-one-out later if needed (keep commented until the full model runs)
    # stal_to_exclude = ["SN35", "SN42", "SN24-3", "SN24-4"]
    # for stal in stal_to_exclude:
    #     run_hbgp(exclude_stal=stal)