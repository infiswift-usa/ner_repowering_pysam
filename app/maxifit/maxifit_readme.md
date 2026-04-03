
## ⚙️ Prerequisites & Setup

* **Python Version:** Python 3.12 preferrably
* **OS:** Windows 10/11 (Required to run `MaxiFit.exe` and `pywinauto`).
* **MaxiFit App:** would mostly be installed at `C:\Program Files (x86)\MaxiFitPointVer5\過積載シミュレーション.exe` (by reading the docs)

# Windows GUI Automation
pywinauto==0.6.8
pywin32==306

# Data Processing & Export
pandas==2.2.0
pyarrow==15.0.0

pip install -r requirements.txt

### Inputs Used
The automation is driven by two required files located in the specs/ directory:

The Extracted Specs (specs/Mie Tsu_extracted.json)
This JSON payload defines the base solar plant configuration (typically generated from the document extraction phase). It dictates the location, hardware models, and base string math.

# JSON
{
  "source": "Mie Tsu.pdf",
  "location": {
    "area": "三重県",
    "point": "津"
  },
  "pcs_config": [
    {
      "pcs_type": "SG100CX-JP",
      "module_type": "NER132M625E-NGD",
      "modules_per_string": 16,
      "strings": 14,
      "tilt": 20,
      "pcs_count": 4
    }
  ]
}

# The Manifest (specs/manifest.csv)
A reference mapping file used by the GUI to select the correct <area> and <point> indices from the MaxiFit dropdown menus without having to read the UI screen manually.

### Outputs Generated

For every execution, the script creates a timestamped folder inside the output directory (e.g., output/simulation_runs/20260401_164833/ perm_00x.csv) : The raw PV generation CSVs dumped directly by the MaxiFit UI for each tested overload/tilt permutation.

# Execution Metadata:

permutations.json
run.json
run.log