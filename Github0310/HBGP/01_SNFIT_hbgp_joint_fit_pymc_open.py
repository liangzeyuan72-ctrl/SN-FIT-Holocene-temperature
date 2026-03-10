import pandas as pd
import numpy as np
import pymc as pm
from pathlib import Path

def main():
    # ===== 1. Read Excel data =====
    script_dir = Path(__file__).resolve().parent
    data_file = script_dir / "stalagmite_temp.xlsx"
    output_file = script_dir / "ShennongCave_HierGP_withSD_CI_PI.csv"

    df = pd.read_excel(data_file)
    df = df.dropna(subset=["Sample", "Age", "Temperature"]).copy()

    # Keep only the four target stalagmites
    mask = df["Sample"].str.contains("SN35|SN42|SN24-4|SN24-3", na=False)
    data = df[mask].copy()

    # Use the SD column if available
    if "SD" in data.columns:
        sd = data["SD"].fillna(data["SD"].mean()).values
    else:
        sd = np.ones(len(data)) * 0.2

    # Extract stalagmite IDs and map them to integer indices
    data["stal_base"] = data["Sample"].str.extract(r"(SN35|SN42|SN24-4|SN24-3)")[0]
    stal_unique = data["stal_base"].unique()
    stal_to_idx = {s: i for i, s in enumerate(stal_unique)}
    data["stal_idx"] = data["stal_base"].map(stal_to_idx).astype(int)

    print("Stalagmite mapping:", stal_to_idx)
    print("Total data points:", len(data))

    # ===== 2. Prepare arrays =====
    age = data["Age"].values.astype(float)
    temp = data["Temperature"].values.astype(float)
    stal_idx = data["stal_idx"].values.astype(int)

    # Center and scale Age (unit: ka)
    age_mean = age.mean()
    age_scale = 1000.0
    x = (age - age_mean) / age_scale
    X = x[:, None]

    # Weighting: smaller SD means larger weight
    weights = 1.0 / (sd ** 2)
    weights = weights / np.mean(weights)  # Normalize

    # ===== 3. Build the hierarchical GP model =====
    with pm.Model() as model:
        eta = pm.HalfNormal("eta", sigma=5.0)
        ell = pm.Gamma("ell", alpha=2.0, beta=1.0)
        sigma_a = pm.HalfNormal("sigma_a", sigma=2.0)
        a_raw = pm.Normal("a_raw", mu=0.0, sigma=1.0, shape=len(stal_unique))
        a = pm.Deterministic("a", a_raw * sigma_a)
        sigma_y = pm.HalfNormal("sigma_y", sigma=2.0, shape=len(stal_unique))

        cov_func = eta**2 * pm.gp.cov.ExpQuad(input_dim=1, ls=ell)
        gp = pm.gp.Latent(cov_func=cov_func)
        f = gp.prior("f", X=X)

        mu = f + a[stal_idx]

        # Add sample SD to the noise term
        y_obs = pm.Normal("y_obs",
                          mu=mu,
                          sigma=sigma_y[stal_idx] + sd,
                          observed=temp,
                          total_size=len(temp),
                          shape=len(temp),
                          )

        trace = pm.sample(
            draws=1000,
            tune=1000,
            target_accept=0.97,
            chains=1,
            cores=1,
            progressbar=True
        )

        # Prediction
        x_pred = np.linspace(x.min(), x.max(), 300)[:, None]
        f_pred = gp.conditional("f_pred", Xnew=x_pred)
        post_pred = pm.sample_posterior_predictive(
            trace, var_names=["f_pred"], return_inferencedata=False
        )

    # ===== 4. Extract predictions and calculate confidence intervals =====
    f_samples = post_pred["f_pred"]

    if f_samples.ndim == 3:
        f_mean = f_samples.mean(axis=(0, 1))
        f_lower = np.percentile(f_samples, 2.5, axis=(0, 1))
        f_upper = np.percentile(f_samples, 97.5, axis=(0, 1))
    else:
        f_mean = f_samples.mean(axis=0)
        f_lower = np.percentile(f_samples, 2.5, axis=0)
        f_upper = np.percentile(f_samples, 97.5, axis=0)

    age_pred = x_pred.ravel() * age_scale + age_mean

    # ===== 5. Write CSV output =====
    out = pd.DataFrame({
        "Age": age_pred,
        "CaveTemp_mean": f_mean,
        "CaveTemp_lower95": f_lower,
        "CaveTemp_upper95": f_upper
    })
    out.to_csv(output_file, index=False)
    print(f"Output written to: {output_file}")

if __name__ == "__main__":
    main()

