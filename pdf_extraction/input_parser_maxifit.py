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

# ==========================================
# STRICT PYDANTIC SCHEMAS (JSON Structure)
# ==========================================
class ProjectInfo(BaseModel):
    project_name: str = Field(..., description="Project name extracted from the filename.")
    date: str = Field(..., description="Date of the drawing (e.g., '2025.01.05').")
    drawing_number: str = Field(..., description="Drawing number (e.g., 'RP-0042-SL01-00').")
    
    raw_location_text: str = Field(..., description="Extract the exact location string written in the bottom right corner of the drawing (e.g., '-三重県津市-').")
    prefecture: str = Field(..., description="Extract just the prefecture name from the location text (e.g., '三重県').")
    subregion: str = Field(..., description="Extract the core city, town, or village name. IGNORING and STRIPPING OFF the prefecture, the district name , or trailing suffix . Examples: '長野県上伊那郡飯島町' -> '飯島'. '三重県津市' -> '津'. '鹿児島県薩摩川内市' -> '薩摩川内'.")

class SolarModuleSpec(BaseModel):
    model_number: str = Field(..., description="Model number of the solar module (e.g., 'NER132M625E-NGD')")
    nominal_maximum_output_w: float = Field(..., description="公称最大出力 in W (e.g., 625)")

class PVArrayConfig(BaseModel):
    """Directly structured for the MaxiFit Automation Script"""
    pcs_group_name: str = Field(..., description="Original name from the document (e.g., 'PCS 01~04 (4台)')")
    pcs: str = Field(..., description="Map the PCS model. E.g., 'SG100CX-JP' -> 'SunGrow SG100CX-JP', 'SUN2000-50KTL' -> 'HUAWEI SUN2000-50KTL-NHK3'")
    panel_type: str = Field(..., description="Model number of the solar module used (e.g., 'NER132M625E-NGD')")
    panel_series: int = Field(..., description="Number of modules in series (直列枚数) (e.g., 16)")
    panel_parallel: int = Field(..., description="Number of strings per PCS (系統数) (e.g., 14)")
    placement_angle: int = Field(..., description="Tilt angle for this specific array (e.g., 20).")
    # Force the model to think out loud before it answers!
    direction_reasoning: str = Field(..., description="Scan horizontally from left to right exactly across the degree text (e.g., '11°' or '5°'). Tell me the exact order of the visual elements from left to right. You MUST output one of these two phrases: 'Order: Plain Line -> Text -> Diamond Line' OR 'Order: Diamond Line -> Text -> Plain Line'..")
    #direction: int = Field(..., description="Azimuth angle as an integer. Use POSITIVE numbers for Left tilts of north arrow, and NEGATIVE numbers for Right tilts of north arrow. (e.g., if north arrow is x degree to right return -x or if north arrow is y degree to left return y).")
    direction: int = Field(..., description="Azimuth angle as an integer. If your reasoning order is 'Plain Line -> Text -> Diamond Line', the arrow is tilted RIGHT, so output a NEGATIVE number (e.g., -11). If your reasoning order is 'Diamond Line -> Text -> Plain Line', the arrow is tilted LEFT, so output a POSITIVE number (e.g., 11). If perfectly vertical, output 0.")
    backside_efficiency: int = Field(0, description="Always set to 0 unless specified.")
    num_arrays: int = Field(..., description="Number of PCS units in this group (e.g., extract 4 from '4台').")

class AreaDetails(BaseModel):
    area_name: str = Field(..., description="Name of the area.")
    pv_arrays: List[PVArrayConfig] = Field(..., description="List of PV arrays (PCS groups) within this area formatted for MaxiFit.")

class BlueprintExtraction(BaseModel):
    project_information: ProjectInfo
    module_specifications: SolarModuleSpec
    area_breakdown: List[AreaDetails] = Field(..., description="List of all areas and their PV array configurations.")

# ==========================================
# STATE DEFINITION
# ==========================================
class ExtractionState(TypedDict):
    val_pdf_path: str
    raw_markdown: str
    page_images: List[str] 
    structured_data: dict
    error: str

subregion_map={
    "津市": "津",
    "上伊那郡飯島町": "飯島",
    "薩摩川内市": "川内",
    "三重郡": "四日市"
}

PCS_MAP ={ 
    "SG100CX-JP": "SunGrow SG100CX-JP"
}
# Globally define heavy objects
LLM = ChatGoogleGenerativeAI(model="gemini-3-pro-preview", api_key=os.environ.get("GOOGLE_API_KEY"),temperature=0).with_structured_output(BlueprintExtraction) 

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
            2. LOCATION: Find location in bottom right corner. Seperate the 'prefecture' and 'subregion' (eg., in '三重県津市' as '三重県','津市') EXACTLY as written.
            3. AZIMUTH MAPPING: To determine the 'azimuth_angle', examine the compass rose. North arrow is the line with the diamond tip.
                - If the True North arrow is tilted to the LEFT of the vertical line: Output a POSITIVE integer (e.g., 5).
                - If the True North arrow is tilted to the RIGHT of the vertical line: Output a NEGATIVE integer (e.g., -11).
                - If the True North arrow is perfectly vertical and aligned with the crosshairs with NO numeric degree offset written next to it, you MUST return '0 degrees'. 
                - DO NOT guess a degree if none is written. 
            4. NUM ARRAYS: For 'num_arrays', extract the integer number of units from the PCS text (eg., extract 4 from 'PCS 01~04(4台)' since it has 4 pcs)

            
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

def _normalize_project_name(raw_name: str) -> str:
    """Extracts trailing name segment from full filename strings.
    e.g. 'モジュール配置図_RP-0039-SL01-00_Mie Tsu' -> 'Mie Tsu'
    """
    parts = raw_name.rsplit("_", 1)
    return parts[-1].strip() if len(parts) > 1 else raw_name.strip()

# ==========================================
# RUN
# ==========================================
if __name__ == "__main__":
    # Dynamically resolve project directory
    BASE_DIR = Path(__file__).resolve().parent
    PROJECT_ROOT = BASE_DIR.parent
    
    pdf_name = "モジュール配置図_RP-0040-SL01-00_Kagoshima iriki.pdf"
    
    pdf_file_path = PROJECT_ROOT / pdf_name
    output_dir = PROJECT_ROOT / "maxifit_output"
    manifest = BASE_DIR / "manifest.csv"

    if os.path.exists(pdf_file_path):
        initial_input = {
            "val_pdf_path": pdf_file_path,
            "raw_markdown": "",
            "page_images": [],
            "structured_data": {},
            "error": ""
        }
        final_state = app.invoke(initial_input)
        if final_state.get("error"):
            print(f"❌ Extraction failed: {final_state['error']}")
            raise SystemExit(1)
        else:
            print("\n✅ Final Extracted JSON:")
            print(json.dumps(final_state["structured_data"], indent=4, ensure_ascii=False))
            extracted_json=final_state["structured_data"]
            project_name = _normalize_project_name(extracted_json["project_information"]["project_name"])
            
            raw_pref = extracted_json["project_information"]['prefecture']
            raw_subreg=extracted_json["project_information"]["subregion"]
            mapped_subreg=raw_subreg

            try:
                df_manifest=pd.read_csv(manifest,encoding='utf-8')
                # Filter handling both directions (e.g., CSV="北海道(宗谷)" vs raw="北海道", or CSV="東京" vs raw="東京都")
                valid_ar=df_manifest[df_manifest['area'].str.contains(raw_pref,na=False)| df_manifest['area'].apply(lambda x: str(x) in raw_pref)]
                valid_pnt=valid_ar['point'].unique()

                # Step 1: Dynamic Match - is the valid point inside the extracted string? (e.g. "津" inside "津市")
                for point in valid_pnt:
                    if point in raw_subreg:
                        mapped_subreg=point
                        print("success")
                        break

                # Step 2: Fallback Override - use dictionary if dynamic match failed
                if mapped_subreg==raw_subreg:
                    for key,val in subregion_map.items(): # map subregion in pdf to value maxifit needs
                        if key in raw_subreg:
                            mapped_subreg=val
                            print("failed")
                            break
            except Exception as e:
                print(f"⚠️ Warning: Dynamic subregion matching failed ({e}).")

            flat_pv_arrays=[]
            for area in extracted_json["area_breakdown"]:
                for array in area["pv_arrays"]:
                    raw_pcs=array["pcs"]
                    for key,val in PCS_MAP.items(): # map "SG100CX-JP" to "SunGrow SG100CX-JP"
                        if key in raw_pcs:
                            array["pcs"] = val
                            break
                    array.pop("pcs_group_name",None) #only needed to count pcs and not for maxifit
                    flat_pv_arrays.append(array)

            maxifit_payload={
                "prefecture": extracted_json["project_information"]["prefecture"],
                "subregion": mapped_subreg,
                "system_efficiency": 95,
                "power_efficiency": 1.0,
                "pv_arrays": flat_pv_arrays,
                "output_files": {
                    "output_directory": str(output_dir),
                    "csv_filename": f"MAXIFIT_csv_output_{project_name}",
                    "print_filename": f"MAXIFIT_output_print_{project_name}",
                    "config_filename": f"MAXIFITconfig_{project_name}",
                    "overwrite_existing": False
                }
            }
            json_save = BASE_DIR
            json_filename = json_save / f"{project_name}_extracted.json"
            with open(json_filename, "w", encoding="utf-8") as f:
                json.dump(maxifit_payload, f, indent=4, ensure_ascii=False)
            print(f"✅ JSON saved to: {json_filename}")
