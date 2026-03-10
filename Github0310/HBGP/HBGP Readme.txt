HBGP README

Files included:
01_SNFIT_hbgp_joint_fit_pymc_relative.py
02_snfit_hbgp_leave_one_out_relative.py

Place the following files in the same folder:
01_SNFIT_hbgp_joint_fit_pymc_relative.py
02_snfit_hbgp_leave_one_out_relative.py
stalagmite_temp.xlsx

Python version:
Python 3.10 or higher

Required packages:
numpy
pandas
pymc
pytensor
arviz
openpyxl

Install packages:
pip install numpy pandas pymc pytensor arviz openpyxl

Run:
python 01_SNFIT_hbgp_joint_fit_pymc_relative.py
python 02_snfit_hbgp_leave_one_out_relative.py

Notes:
1. Both scripts read stalagmite_temp.xlsx from the script directory by default.
2. Output files are also saved to the script directory.
3. No absolute file path is required.
4. If you rename the input file, update the filename in the scripts accordingly.