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
# STATE DEFINITION
# ==========================================
class ExtractionState(TypedDict):
    val_pdf_path: str
    raw_markdown: str
    page_images: List[str] # Base64 strings of the drawing
    structured_data: dict
    error: str

class DynamicExtraction(BaseModel):
    data: Dict[str, Any] = Field(
        description="A comprehensive dictionary containing all specs, measurements, and notes discovered in the document."
    )
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
    
    # Capture the first page as an image for Vision reasoning
    images = []
    #avoid result.document.pages to avoid AttributeError: 'int' object has no attribute 'image'. Did you mean: 'imag'?
    for page in result.pages:
        # Convert PIL image to base64
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
    
    # We use Flash for speed, but with 'Vision' capabilities
    llm = ChatGoogleGenerativeAI(
        model="gemini-3-flash-preview", 
        api_key=os.environ.get("GOOGLE_API_KEY"),
        temperature=0
    ).with_structured_output(DynamicExtraction) 
    #llm.with_structured_output(dict) would raise ValueError: Unsupported schema type <class 'type'>
    # Construct Multimodal Message
    # We provide the Markdown for the specs and the Image for the geometry
    content = [
        {
            "type": "text", 
            "text": f"""Analyze this solar blueprint. 
            1. Use the Markdown text for precise table data (Watts, Model numbers).
            2. Use the attached Image to understand the diagram specs (tilt angles, heights, clearance).
            3. Translate Japanese terms to English keys.
            
            MARKDOWN:
            {state['raw_markdown']}"""
        }
    ]
    
    # Add the images of the pages to the prompt
    for img_data in state["page_images"]:
        content.append({
            "type": "image_url", 
            "image_url": {"url": f"data:image/jpeg;base64,{img_data}"}
        })

    try:
        out = llm.invoke([HumanMessage(content=content)])
        # result is currently a 'DynamicSolarExtraction' object
        # FIX: Convert the object to a dictionary to solve TypeError: Object of type DynamicExtraction is not JSON serializable
        # If you defined 'data' as the field name in the Pydantic class:
        res=out.data
        return {"structured_data": res}
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
    pdf_file_path = r"C:\Users\ASUS\Downloads\モジュール配置図_RP-0007-SL03-00_Mie Fukuo.pdf"
    
    
    if os.path.exists(pdf_file_path):
        initial_input = {
            "val_pdf_path": pdf_file_path,
            "raw_markdown": "",
            "page_images": [],
            "structured_data": {}
        }
        final_state = app.invoke(initial_input)
        print("\n✅ Final Extracted JSON:")
        print(json.dumps(final_state["structured_data"], indent=4, ensure_ascii=False))
    else:
        print(f"Error: File not found.")