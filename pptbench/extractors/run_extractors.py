# pptbench/extractors/run_extractors.py

import os
from json import dumps
import logging

from pptx import Presentation

from pptbench.extractors.ppt_extractor import PowerPointShapeExtractor

# Configure logging to include DEBUG level
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def run_extractors(pptx_path: str, measurement_unit: str = "pt") -> dict:
    if not pptx_path:
        raise ValueError("pptx_path is required")
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"File not found: {pptx_path}")
    ppt = Presentation(pptx_path)
    shape_extractor = PowerPointShapeExtractor(ppt, measurement_unit)
    extracted_info = shape_extractor.extract_ppt()
    return extracted_info

# Example usage
if __name__ == "__main__":
    pptx_path = "../../dataset/pptx/FBNIGNWBP6W7JONNO2JVH7YFA2SSPXDN.pptx"
    extracted_info = run_extractors(pptx_path, measurement_unit="pt")
    print(dumps(extracted_info, indent=4))
