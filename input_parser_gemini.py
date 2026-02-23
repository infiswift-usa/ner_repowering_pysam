import os
import json
import base64
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
    
    # These capture the totals at the bottom of the table
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
    width_mm: Optional[float] = Field(None, description="Width of this rack block in mm, found in the measurement lines of the main top-down panel layout drawing (e.g., 19126 for the red block).")
    length_mm: Optional[float] = Field(None, description="Length/depth of this rack block in mm, if found in the measurement lines of the main top-down panel layout drawing.")

class RackProfile(BaseModel):
    tilt_angle_degrees: Optional[float] = Field(None, description="Tilt angle of the solar panels (e.g., 20).")
    array_pitch_mm: Optional[float] = Field(None, description="Pitch distance between the front of one array to the front of the next (e.g., 3181).")
    foundation_width_mm: Optional[float] = Field(None, description="Total horizontal width of the foundation/base (e.g., 7500).")
    max_height_mm: Optional[float] = Field(None, description="Maximum height of the array structure at the top edge (e.g., 1572).")
    min_ground_clearance_mm: Optional[float] = Field(None, description="Minimum ground clearance at the lower edge (e.g., 800).")

class BlueprintExtraction(BaseModel):
    project_information: ProjectInfo
    module_specifications: SolarModuleSpec
    azimuth_angle: str = Field(..., description="The azimuth or true north angle extracted from the compass (e.g., '11 degrees West'). Also note down if there is no degree mentioned in the compass")
    area_breakdown: List[AreaDetails] = Field(..., description="List of all areas and their PCS/module totals.")
    rack_configurations: List[RackConfiguration] = Field(..., description="List of the different colored top-down rack configurations.")
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
    model="gemini-2.5-flash", # Use 2.5-flash as it is the most modern vision model right now
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
            3. For 'rack_configurations', identify the text color (red, green, blue, etc.) used for '4 rows 8 columns', etc. Match dimensional measurements (widths/pitch) from the drawing to these specific racks.
            4. Look carefully at the compass rose to determine the 'azimuth_angle'.
            
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
# RUN
# ==========================================
if __name__ == "__main__":
    pdf_file_path = r"C:\Users\ASUS\Downloads\モジュール配置図_RP-0039-SL01-00_Mie Tsu.pdf"
    
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
            print(f"Failed: {final_state['error']}")
        else:
            print("\n✅ Final Extracted JSON:")
            print(json.dumps(final_state["structured_data"], indent=4, ensure_ascii=False))
    else:
        print(f"Error: File not found.")