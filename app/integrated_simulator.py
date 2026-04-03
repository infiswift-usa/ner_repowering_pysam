import os
import sys
import argparse
import subprocess
from pathlib import Path
from datetime import datetime
import json

# Add internal modules to pythonpath
BASE_DIR = Path(__file__).resolve().parent
sys.path.append(str(BASE_DIR))
sys.path.append(str(BASE_DIR / "cashflow_simulation"))
sys.path.append(str(BASE_DIR / "document_extraction"))

# Safe imports
try:
    from document_extraction.document_parsing import run_extraction as extract_pdf
    from cashflow_simulation.cashflow_simulator import run_simulation_pipeline 
except ImportError as e:
    print(f"Initialization Error: Ensure all inner modules are accessible. {e}")
    sys.exit(1)

def run_integration_pipeline(pdf_path: str, user_inputs_json: str = None):
    target_pdf = Path(pdf_path).resolve()
    
    if not target_pdf.exists():
        print(f"\n❌ Error: Target PDF file not found at: {target_pdf}")
        return

    print("\n" + "="*60)
    print(" 1. DATA EXTRACTOR (Skipped - Handled via Background SQL Updates) ")
    print("="*60)
    print("MySQL Database is acting as the single source of truth.")
    
    print("\n" + "="*60)
    print(" 2. PDF EXTRACTION & PARSING ")
    print("="*60)

    # 1. Tell the code exactly where to save the JSON file
    json_output_dir = BASE_DIR / "document_extraction"
    
    # 2. Run the extraction using the new arguments!
    print(f"Sending {target_pdf.name} to the Docling/Gemini Extractor...")
    extracted_json_path = extract_pdf(
        pdf_file_path=str(target_pdf), 
        output_directory=json_output_dir
    )
    
    if not extracted_json_path or not os.path.exists(extracted_json_path):
        print("\n❌ Error: PDF extraction failed or no JSON was generated.")
        return
        
    print(f"✅ Maxifit configuration generated at: {extracted_json_path}")
        
    print("\n" + "="*60)
    print(" 3. MAXIFIT APP AUTOMATION (BATCH PERMUTATIONS) ")
    print("="*60)
    
    demo_automation_script = BASE_DIR / "maxifit" / "maxifit_runner.py"
    MAIN_DIR = Path(__file__).resolve().parent
    manifest_csv_path =MAIN_DIR/ "specs" / "manifest.csv"
    
    print("Launching pywinauto robotic batch simulation in dedicated subprocess...")
    #automation_process = subprocess.run(
    #    [
    #        sys.executable, 
    #        str(demo_automation_script), 
    #        "--specs", str(extracted_json_path),
    #        "--manifest", str(manifest_csv_path)
    #    ],
    #    cwd=str(BASE_DIR)
    #)
    #
    #if automation_process.returncode != 0:
    #    print(f"\n⚠️ Warning: Maxifit Automation reported non-zero exit code ({automation_process.returncode}).")
    #else:
    #    print("\n✅ Maxifit Batch Automation Completed Successfully.")
    
    print("\n" + "="*60)
    print(" 4. FINANCIAL PRICING SIMULATION (BATCH LOOP) ")
    print("="*60)

    # Default Inputs (The baseline fallbacks)
    base_user_inputs = {
        'region': '中部',
        'ex_ac': 1000.00,
        'ex_dc': 1127.80,
        'rep_ac': 1000.00,
        'rep_dc': 1421.28,
        'ex_yield': 1433741.0,   
        'rep_yield': 2182388.74, 
        'ex_deg': 0.007,
        'rep_deg': 0.004,
        'fit_price': 32.0,
        'latest_price': 8.9,
        'op_start_date': datetime(2016, 8, 31),
        'mod_date': datetime(2025, 7, 31),
    }

    if user_inputs_json and os.path.exists(user_inputs_json):
        print(f"Loading custom user inputs from JSON: {user_inputs_json}")
        try:
            with open(user_inputs_json, 'r', encoding='utf-8') as f:
                loaded_inputs = json.load(f)
                base_user_inputs.update(loaded_inputs)
                
                for date_key in ['op_start_date', 'mod_date']:
                    if date_key in base_user_inputs and isinstance(base_user_inputs[date_key], str):
                        try:
                            base_user_inputs[date_key] = datetime.strptime(base_user_inputs[date_key], '%Y-%m-%d')
                        except Exception as de:
                            pass
        except Exception as e:
            print(f"Failed to load user_inputs_json: {e}")
            
    try:
        # 1. Locate the dynamic runs directory
        runs_dir = BASE_DIR / "output" / "simulation_runs"
        
        if not runs_dir.exists():
            print(f"❌ Error: The simulation_runs directory was not created.")
            return
            
        # 2. Find the newest timestamped folder
        all_run_folders = [f for f in runs_dir.iterdir() if f.is_dir()]
        if not all_run_folders:
            print("❌ Error: No timestamped run folders found.")
            return
            
        latest_run_folder = max(all_run_folders, key=os.path.getmtime)
        csv_files = sorted(latest_run_folder.glob("perm_*.csv"))
        
        if not csv_files:
            print(f"⚠️ Warning: No permutation CSVs found in {latest_run_folder.name}")
            return
            
        print(f"📂 Found {len(csv_files)} configurations in {latest_run_folder.name}. Starting financial loop...\n")

        # 3. Loop through every single CSV and run the financial math
        for i, csv_path in enumerate(csv_files, 1):
            print(f"   --- Processing Permutation {i}/{len(csv_files)}: {csv_path.name} ---")
            
            # Use .copy() so we always start with a clean baseline for each loop
            run_simulation_pipeline(base_user_inputs.copy(), csv_path=str(csv_path))
            
        print("\n✅ All financial permutations calculated successfully.")
        
    except Exception as e:
        print(f"\n❌ Error during Pricing Simulation Loop: {e}")

    print("\n" + "="*60)
    print(" 🎉 PIPELINE EXECUTION COMPLETE ")
    print("="*60)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Full Solar Pipeline Integrator")
    parser.add_argument("--pdf_path", type=str, help="Absolute or relative path to the Project PDF.", required=True)
    parser.add_argument("--user_inputs_json", type=str, help="Absolute or relative path to JSON file containing dynamic Price Calculator inputs.", default=None)
    
    args = parser.parse_args()
    run_integration_pipeline(
        pdf_path=args.pdf_path,
        user_inputs_json=args.user_inputs_json
    )