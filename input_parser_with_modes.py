import os
import json
import base64
from pathlib import Path
import io
from typing import TypedDict, List, Optional, Dict, Any
from dotenv import load_dotenv
# Docling Imports
from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import PdfPipelineOptions
from docling.document_converter import DocumentConverter, PdfFormatOption
from docling.backend.pypdfium2_backend import PyPdfiumDocumentBackend
# LangChain / LangGraph Imports
from langgraph.graph import StateGraph, END
from langgraph.types import RetryPolicy
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.messages import HumanMessage
from pydantic import BaseModel, Field
# Load API Key
load_dotenv()
import numpy as np
import re
import pandas as pd
import pvlib
from pvlib.modelchain import ModelChain
from pvlib.pvsystem import Array, FixedMount, PVSystem
from pvlib.temperature import TEMPERATURE_MODEL_PARAMETERS
from pvlib.bifacial import infinite_sheds
SIM_MODE = 'LEVEL_2'
# ==========================================
# STRICT PYDANTIC SCHEMAS (JSON Structure)
# ==========================================
class ProjectInfo(BaseModel):
    project_name: str = Field(..., description="Project name extracted from the filename.")
    date: str = Field(..., description="Date of the drawing (e.g., '2025.01.05').")
    drawing_number: str = Field(..., description="Drawing number (e.g., 'RP-0042-SL01-00').")
    scale: str = Field(..., description="Scale of the drawing (e.g., '1/400').")

class SolarModuleSpec(BaseModel):
    model_number: str = Field(..., description="Model number of the solar module (e.g., 'NER132M625E-NGD')")
    nominal_maximum_output_w: float = Field(..., description="公称最大出力 in W (e.g., 625)")
    nominal_max_power_operating_voltage_v: float = Field(..., description="公称最大出力動作電圧 in V (e.g., 41)")
    nominal_max_power_operating_current_a: float = Field(..., description="公称最大出力動作電流 in A (e.g., 14)")
    nominal_open_circuit_voltage_v: float = Field(..., description="公称開放電圧 in V (e.g., 49)")
    nominal_short_circuit_current_a: float = Field(..., description="公称短絡電流 in A (e.g., 16)")
    weight_kg: float = Field(..., description="質量 in kg (e.g., 32.8)")
    dimensions_mm: str = Field(..., description="寸法 (e.g., '2382x1134x30')")

class PCSGroup(BaseModel):
    group_name: str = Field(..., description="e.g., 'PCS 01 (1台)' or 'PCS 01~06(6台)'")
    pcs_model: str = Field(..., description="e.g., 'SG100CX-JP (三相3線100.0kW)'")
    modules_per_pcs: int = Field(..., description="Number of modules per PCS (e.g., 192)")
    modules_in_series: int = Field(..., description="Number of modules in series (e.g., 16)")
    strings_per_pcs: int = Field(..., description="Number of strings (e.g., 12)")
    module_output_kw: float = Field(..., description="Module output in kW (e.g., 120.00)")
    pcs_output_kw: float = Field(..., description="PCS output in kW (e.g., 90.00)")

class AreaDetails(BaseModel):
    area_name: Optional[str] = Field("Main Project Area", description="Name of the area (e.g., 'Area 1' or 'エリア1'). Use 'Main Project Area' if the document is not divided.")
    pcs_groups: List[PCSGroup] = Field(..., description="List of PCS blocks within this area.")
    #optional title angle
    tilt_angle: float = Field(..., description="The tilt angle for this specific area. For single areas, extract from '架台参考図' (e.g. 20). For multiple areas, extract from the label like '(アレイ角度10度)'.")    # These capture the totals at the bottom of the table
    total_modules: int = Field(..., description="Total module count for this area (モジュール合計枚数) (e.g., 544)")
    total_module_output_kw: float = Field(..., description="Total module output for area in kW (モジュール総出力) (e.g., 340.00)")
    total_system_output_kw: float = Field(..., description="Total system/PCS side output in kW (系統側総出力) (e.g., 270.00)")

class RackConfiguration(BaseModel):
    #config_name: str = Field(..., description="e.g., '20度傾斜架台_4段8列'")
    config_rows: str = Field(..., description="e.g., 4 from '4段8列'")
    config_columns: str = Field(..., description="e.g., 8 from '4段8列'")
    color_in_diagram: str = Field(..., description="The color of the text/rack in the diagram (e.g., 'red', 'green', 'blue').")
    unit_count: int = Field(..., description="Number of these rack structures/bases (e.g., 12, from '12基').")
    #modules_per_unit: int = Field(..., description="Number of modules per rack unit (e.g., 32, from '(32枚)').")
    width_mm: Optional[float] = Field(None, description="Width of this rack block in mm, found in the measurement lines of the main top-down panel layout drawing (e.g., 19126 for the red block). Longest side of rack.")
    pitch: Optional[float] = Field(None, description=" Pitch is typically the length of rack + spacing between 2 racks. Identify the large repeating distance between rows (e.g., 7300). Found in the measurement lines of the main top-down panel layout drawing.")
    length_mm: Optional[float] = Field(None, description="Length/depth of this rack block in mm, if found in the measurement lines of the main top-down panel layout drawing. Calculate this as (pitch - Row Spacing) if not explicitly labeled. Shorter side of rack.")

class RackProfile(BaseModel):
    #tilt_angle_degrees: Optional[float] = Field(None, description="Tilt angle of the solar panels (e.g., 20).if no side diagram, refer エリア.. and its bracket (アレイ角度...) in solar diagram for tilt angle")
    #array_pitch_mm: Optional[float] = Field(None, description="Pitch distance between the front of one array to the front of the next (e.g., 3181).")
    #foundation_width_mm: Optional[float] = Field(None, description="Total horizontal width of the foundation/base (e.g., 7500).")
    max_height_mm: Optional[float] = Field(None, description="Analyze the vertical dimensions in the side-profile diagram (架台参考図) from bottom to top:Ground to Bottom (Clearance): Identify the height from the ground to the lowest edge of the rack or ground clearance, (eg., 800).Bottom to Top (Rise): Identify the vertical rise of the rack structure itself from its ground clearance height to its highest point, (eg.,1572).Total Height: Calculate the total height from the ground to the top edge by summing these two values (eg., 800 + 1572 = 2372). ")
    #min_ground_clearance_mm: Optional[float] = Field(None, description="Minimum ground clearance at the lower edge (e.g., 800).")

class BlueprintExtraction(BaseModel):
    project_information: ProjectInfo
    module_specifications: SolarModuleSpec
    azimuth_angle: str = Field(..., description="The azimuth or true north angle extracted from the compass (e.g., '11 degrees West'). Also note down if there is no degree mentioned in the compass")
    area_breakdown: List[AreaDetails] = Field(..., description="List of all areas and their PCS/module totals.")
    rack_configurations: List[RackConfiguration] = Field(..., description="List of the different colored top-down rack configurations. check the measurement of specific colour. only refer nearby similar coloured rack if that specific colour doesnt have measurment ")
    rack_profile_measurements: Optional[RackProfile] = Field(None, description="Measurements extracted from the side-profile rack reference diagram (架台参考図). if no diagram, refer エリア.. and its bracket (アレイ角度...) in solar diagram for tilt angle")
    diagram_notes: List[str] = Field(..., description="Translations of special notes at the bottom of the page (e.g., '特記事項').")
    
    # Catch-all for unexpected lines/data
    additional_findings: Dict[str, Any] = Field(
        default_factory=dict, 
        description="Extract any OTHER technical specifications, equipment details, or notes found."
    )
# ==========================================
# STATE DEFINITION
# ==========================================
class ExtractionState(TypedDict):
    val_pdf_path: str
    raw_markdown: str
    page_images: List[str] 
    structured_data: dict
    error: str

# Globally define heavy objects
LLM = ChatGoogleGenerativeAI(
    model="gemini-3-flash-preview",
    api_key=os.environ.get("GOOGLE_API_KEY"),
    temperature=0
).with_structured_output(BlueprintExtraction) 

# ==========================================
# NODES
# ==========================================
def parse_document(state: ExtractionState):
    print(f"⚡ Parsing PDF & Capturing Diagrams: {state['val_pdf_path']}")
    
    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_ocr = False 
    pipeline_options.do_table_structure = True 
    pipeline_options.generate_page_images = True
    
    converter = DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(
                pipeline_options=pipeline_options,
                backend=PyPdfiumDocumentBackend
            )
        }
    )
    
    result = converter.convert(state['val_pdf_path'])
    markdown_text = result.document.export_to_markdown()
    
    images = []
    for page in result.pages:
        if page.image:
            buffered = io.BytesIO()
            page.image.save(buffered, format="JPEG", quality=85)
            img_str = base64.b64encode(buffered.getvalue()).decode()
            images.append(img_str)
            
    return {"raw_markdown": markdown_text, "page_images": images}

def route_after_parsing(state: ExtractionState):
    if state.get("error"): return END
    return "extractor"

def extraction_node(state: ExtractionState):
    print("🧠 Reasoning over Text + Diagrams...")
    
    # 1. Provide the filename directly as a hint for the prompt
    filename_hint = os.path.splitext(os.path.basename(state['val_pdf_path']))[0]
    
    content = [
        {
            "type": "text", 
            "text": f"""Analyze this solar blueprint. 
            
            CRITICAL INSTRUCTIONS:
            1. The file name is '{filename_hint}'. Use this for the project name.
            2. For the 'area_breakdown', map the Japanese tables (like Area 1/エリア1) directly to the PCS lists exactly as shown in the document.
            3. For 'rack_configurations', identify the text color (red, green, blue, etc.) used for '4 rows 8 columns', etc. Match dimensional measurements (widths/length/height) from the drawing to these specific racks. Analyze the dimension lines in the panel layout:
                - pitch: Identify the large repeating distance between rows (e.g., 7300).
                - Row Spacing: Identify the clear gap/walkway between the racks (e.g., 2981).
                - length_mm: Calculate this as (pitch - Row Spacing) if not explicitly labeled.
            4. To determine the 'azimuth_angle', examine the compass rose. 
                - If the True North arrow is perfectly vertical and aligned with the crosshairs with NO numeric degree offset written next to it, you MUST return '0 degrees'. 
                - DO NOT guess a degree if none is written. 
                - Only return 'X degrees West/East' if a specific number (like 5 or 11) is explicitly printed next to the compass or in the diagram notes.
                - If the True North arrow is tilted slightly to the Left of the vertical crosshair, identify this as West. If it is tilted to the Right, identify it as East.

            MARKDOWN EXTRACT:
            {state['raw_markdown']}"""
        }
    ]
    
    for img_data in state["page_images"]:
        content.append({
            "type": "image_url", 
            "image_url": {"url": f"data:image/jpeg;base64,{img_data}"}
        })
    try:
        out = LLM.invoke([HumanMessage(content=content)])
        # Pydantic classes natively convert to dicts via .model_dump()
        return {"structured_data": out.model_dump()} 
    except Exception as e:
        print(f"❌ Extraction Error: {e}")
        return {"error": str(e)}
    
# ==========================================
# BUILDING GRAPH
# ==========================================
retries = RetryPolicy(max_attempts=3, initial_interval=2.0)
workflow = StateGraph(ExtractionState)
workflow.add_node("parser", parse_document)
workflow.add_node("extractor", extraction_node, retry_policy=retries)
workflow.set_entry_point("parser")
workflow.add_conditional_edges("parser", route_after_parsing)
workflow.add_edge("extractor", END)
app = workflow.compile()

# ==========================================
# SIMULATOR
# ==========================================

# --- Dynamic Mapping & Calculation ---
def _normalize_project_name(raw_name: str) -> str:
    """Extracts trailing name segment from full filename strings.
    e.g. 'モジュール配置図_RP-0039-SL01-00_Mie Tsu' -> 'Mie Tsu'
    """
    parts = raw_name.rsplit("_", 1)
    return parts[-1].strip() if len(parts) > 1 else raw_name.strip()

SITE_LOOKUP: dict = {
    "Mie Tsu":   {"lat": 34.856, "lon": 136.452},
    "Mie Fukuo": {"lat": 34.856, "lon": 136.452},
}

def get_site_config(extracted_json: dict) -> tuple:
    """Maps extracted project name to coordinates and normalizes azimuth."""
    raw_name = extracted_json["project_information"]["project_name"]
    project_name = _normalize_project_name(raw_name)

    coords = SITE_LOOKUP.get(project_name)
    
    # Azimuth: 180 is South. '5 degrees West' -> 185
    az_str = extracted_json.get("azimuth_angle", "0 degrees")
    azimuth = 180.0
    val_match = re.search(r"(\d+)", az_str)
    offset = float(val_match.group(1)) if val_match else 0.0
    
    if "West" in az_str:
        azimuth += offset
    elif "East" in az_str:
        azimuth -= offset
        
    return coords, azimuth

def build_systems_from_json(extracted_json, azimuth):
    """Constructs PVSystems with dynamic PDC0 and Efficiency calculations."""
    temp_params = TEMPERATURE_MODEL_PARAMETERS["sapm"]["open_rack_glass_glass"]
    mod_spec = extracted_json["module_specifications"]
    
    module_params = {
        "pdc0": mod_spec["nominal_maximum_output_w"],
        "gamma_pdc": -0.0029  # Hardcoded from config
    }
    
    systems = []
    for area in extracted_json["area_breakdown"]:
        mount = FixedMount(surface_tilt=area["tilt_angle"], surface_azimuth=azimuth)
        
        for group in area["pcs_groups"]:
            match = re.search(r"(\d+)台", group["group_name"])
            if not match:
                raise ValueError(
                    f"Cannot parse unit count from group_name: '{group['group_name']}'. "
                    "Expected a pattern like '4台' or 'PCS 01~04 (4台)'."
                )
            num_units = int(match.group(1))
            
            # DYNAMIC CALCULATIONS:
            # 1. pdc0 = module_output_kw * 1000 (Convert kW to W)
            dynamic_pdc0 = group["module_output_kw"] * 1000
            
            # 2. Efficiency = pcs_output_kw / module_output_kw
            # Use extracted values to calculate efficiency dynamically
            #dynamic_eta = group["pcs_output_kw"] / group["module_output_kw"]
            dynamic_eta = 0.984
            array = Array(
                mount=mount,
                module_parameters=module_params,
                modules_per_string=group["modules_in_series"],
                strings=group["strings_per_pcs"],
                temperature_model_parameters=temp_params,
            )
            
            for _ in range(num_units):
                systems.append(
                    PVSystem(
                        arrays=[array],
                        inverter_parameters={
                            "pdc0": dynamic_pdc0, 
                            "eta_inv_nom": dynamic_eta
                        }
                    )
                )
    return systems

# --- Simulator Logic ---

def run_plant_simulation(df_weather, systems, location, extracted_json, azimuth):
    """Aggregates AC power using advanced POA models (Bifacial/Shading)."""
    # 1. Solar Position & DNI Correction
    times_mid = df_weather.index + pd.Timedelta(minutes=30)
    solar_pos = location.get_solarposition(times_mid)
    solar_pos.index = df_weather.index
    cos_zenith = np.cos(np.radians(solar_pos["zenith"].values))

    weather = pd.DataFrame(index=df_weather.index)
    weather['ghi'] = df_weather['GHI']
    weather['dhi'] = df_weather['DHI']
    weather['dni'] = (df_weather['DNI_horiz'] / cos_zenith).fillna(0).clip(0, 1500)
    weather['temp_air'] = df_weather['Temperature']
    weather['wind_speed'] = df_weather['WindSpeed']

    # 2. Geometry Retrieval (from JSON)
    rack_cfg = extracted_json.get("rack_configurations", [{}])[0]
    pitch_val = rack_cfg.get("pitch", 7300.0) / 1000 # 7300mm -> 7.3m
    length_val = rack_cfg.get("length_mm", 4319.0) / 1000 # 4319mm -> 4.319m
    # ==================================================
    # in pvlib: length_val=2.382 and pitch_val=length_val/gcr
    # ==================================================
    # Bottom-to-Top total height (800 + 1572 = 2372)
    h_max = (extracted_json.get("rack_profile_measurements") or {}).get("max_height_mm", 2372.0) / 1000
    
    # Mode Settings
    gcr = length_val / pitch_val if SIM_MODE == 'LEVEL_3' else 0.01
    #hub_height = h_max / 2 if SIM_MODE == 'LEVEL_3' else 10.0
    hub_height = h_max / 2
    tilt = extracted_json["area_breakdown"][0]["tilt_angle"]
    
    # 3. Calculate POA based on Mode
    if SIM_MODE != 'LEVEL_1':
        print(f"\n[Mode: {SIM_MODE}] Calculating Bifacial POA (Infinite Sheds)...")
        bifacial_irrad = infinite_sheds.get_irradiance(
            tilt, azimuth,
            solar_pos['zenith'].values, solar_pos['azimuth'].values,
            gcr, hub_height, pitch_val,
            weather['ghi'].values, weather['dhi'].values, weather['dni'].values,
            albedo=0.2, bifaciality=0.80, vectorize=True
        )
        weather['poa_global'] = bifacial_irrad['poa_global']
        weather['poa_direct'] = bifacial_irrad['poa_front_direct']
    else:
        print(f"\n[Mode: {SIM_MODE}] Calculating Mono-facial POA (Standard)...")
        poa = pvlib.irradiance.get_total_irradiance(tilt, azimuth, 
                                                    solar_pos['zenith'], solar_pos['azimuth'], 
                                                    weather['dni'], weather['ghi'], weather['dhi'])
        weather['poa_global'] = poa['poa_global']
        if poa["poa_beam"] :
            weather['poa_direct'] = poa["poa_beam"]
        else:
            weather['poa_direct'] = poa['poa_direct']

    weather['poa_diffuse'] = weather['poa_global'] - weather['poa_direct']

    # 4. ModelChain Execution
    ac_plant = pd.Series(0.0, index=weather.index, dtype=float)
    for system in systems:
        mc = ModelChain.with_pvwatts(system, location)
        mc.run_model_from_poa(weather) # Corrected to use pre-calculated POA
        ac_plant = ac_plant.add(mc.results.ac, fill_value=0)

    return ac_plant

# ==========================================
# RUN
# ==========================================
if __name__ == "__main__":
    pdf_file_path = r"C:\Users\ASUS\Downloads\モジュール配置図_RP-0039-SL01-00_Mie Tsu.pdf"
    weather_csv = r"D:\VS_CODE\Infiswift\metpv_11_automation\metpv11_clean_v2.csv"
    output_dir = Path(r"D:\VS_CODE\Infiswift")
    
    if os.path.exists(pdf_file_path):
        initial_input = {
            "val_pdf_path": pdf_file_path,
            "raw_markdown": "",
            "page_images": [],
            "structured_data": {},
            "error": ""
        }
        final_state = app.invoke(initial_input)
        
        if not final_state.get("error"):
            extracted_json = final_state["structured_data"]
            project_name = _normalize_project_name(extracted_json["project_information"]["project_name"])

            # Save JSON to Infiswift path
            json_filename = output_dir / f"{project_name}_extracted.json"
            with open(json_filename, "w", encoding="utf-8") as f:
                json.dump(extracted_json, f, indent=4, ensure_ascii=False)
            
            # Setup Simulation
            coords, azimuth = get_site_config(extracted_json)
            location = pvlib.location.Location(coords['lat'], coords['lon'], tz="Asia/Tokyo")
            systems = build_systems_from_json(extracted_json, azimuth)
            
            # Load Weather
            df_raw = pd.read_csv(weather_csv)
            df_raw['DateTime'] = pd.to_datetime(df_raw['DateTime'])
            df_raw.set_index('DateTime', inplace=True)
            df_raw.index = df_raw.index.tz_localize('Asia/Tokyo')
            
            # Run simulation with Mode logic
            ac_output = run_plant_simulation(df_raw, systems, location, extracted_json, azimuth)
            
            # Output
            monthly_kwh = ac_output.resample("ME").sum() / 1000
            yearly_kwh = ac_output.sum() / 1000
            
            print("\n" + "="*70)
            print(f"FINAL RESULTS (PVLIB WORKFLOW - {SIM_MODE})")
            print("="*70)
            print(f"Annual AC Energy: {yearly_kwh:,.0f} kWh")

            print("\nMonthly Production (kWh):")
            months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            for i, m in enumerate(months):
                print(f"  {m}:  {monthly_kwh.iloc[i]:>8,.0f} kWh")
            print("="*70)
        else:
           print(f"❌ Extraction failed: {final_state['error']}")
           raise SystemExit(1) 
    else:
        print(f"Error: PDF not found at {pdf_file_path}")
        raise SystemExit(1)