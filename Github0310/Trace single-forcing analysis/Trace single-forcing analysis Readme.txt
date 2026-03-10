Trace README

Files:
01_trace21k_point_loess_compact.py
02_trace21k_phase_weight_fit.py
03_trace21k_sn_global_coherence.py
04_trace21k_eof_global.py
05_trace21k_phase_difference_maps.py

Required files in the script directory:
TraCE-21K-ghg-only.monthly.TS.nc
TraCE-21K-orb-only.monthly.TS.nc
TraCE-21K-ice-only.monthly.TS.nc
TraCE-21K-fwf-only.monthly.TS.nc
TraCE21K_TS_point_28.7N_117.2E_annual_merged_loess0.20_compact.xlsx
SN-FIT BG.xlsx

Optional files:
World_countries.shp
World_countries.shx
World_countries.dbf

Python:
Python 3.10 or higher

Packages:
numpy
pandas
xarray
matplotlib
cartopy
openpyxl
scipy
statsmodels
shapely
scikit-misc

Install:
pip install numpy pandas xarray matplotlib cartopy openpyxl scipy statsmodels shapely scikit-misc

Run order:
1. 01_trace21k_point_loess_compact.py
2. 02_trace21k_phase_weight_fit.py
3. 03_trace21k_sn_global_coherence.py
4. 04_trace21k_eof_global.py
5. 05_trace21k_phase_difference_maps.py

Outputs and figure/table correspondence:
01: generates the compact LOESS Excel file used by 02
02: Fig. 4, Table 1, Fig. S11
03: Fig. S6, Table S1
03 + World_countries shapefile: Table S2, Table S3
04: Fig. S5
05: Fig. S7

Notes:
1. Script 01 generates the compact LOESS file for Script 02.
2. Script 02 can also run without Script 01 if TraCE21K_TS_point_28.7N_117.2E_annual_merged.xlsx is provided, because the script can build the compact LOESS series internally.
3. Scripts 03, 04, and 05 read the four TraCE-21k NetCDF files directly.
4. World_countries.shp/.shx/.dbf are only needed for the Land/Ocean masking part in Script 03, corresponding to Tables S2-S3. They are not required for the core global correlation map itself.
5. Keep all scripts and input files in the same folder.