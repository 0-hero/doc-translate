import sys
import logging
from unstructured_inference.models.base import get_model
from unstructured_inference.inference.layout import DocumentLayout
from .utils.file_utils import parse_layout
from .utils.llm_utils import translate_text
from tqdm import tqdm

logging.basicConfig(level=logging.INFO)

def init_model(model_name="chipperv1"):
    '''
    Initializes the unstructured-io model for inference.
    '''
    logging.info(f"Initializing model {model_name}")
    model = get_model(model_name)
    return model

def get_layout(input_file, model, file_type="pdf"):
    '''
    Gets the layout of the input file.
    '''
    logging.info(f"Getting layout for {input_file} of type {file_type}")
    if file_type == "pdf":
        layout = DocumentLayout.from_file(input_file, extract_tables= True, detection_model=model)
    elif file_type == "image":
        layout = DocumentLayout.from_image_file(input_file, extract_tables= True, detection_model=model)
    else:
        logging.error("Invalid file type.")
        sys.exit(1)
    parsed_layout = parse_layout(layout)
    return parsed_layout

def translate_layout(parsed_layout, source_lang, target_lang, provider="openai"):
    '''
    Translate text components of the layout file
    '''
    logging.info("Translating layout")
    for page in tqdm(parsed_layout["document"]["pages"]):
        for element in page["layout_elements"]:
            translated_text = translate_text(element["text"], source_lang, target_lang, provider)
            element["translated_text"] = translated_text
    return parsed_layout