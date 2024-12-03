# pptbench/extractors/run_extractors.py

import os
from json import dumps

from pptx import Presentation

from .ppt_extractor import PowerPointShapeExtractor

def run_extractors(pptx_path: str, measurement_unit: str = "emu") -> dict:
    if not pptx_path:
        raise ValueError("pptx_path is required")
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"File not found: {pptx_path}")
    ppt = Presentation(pptx_path)
    shape_extractor = PowerPointShapeExtractor(ppt, measurement_unit)
    extracted_info = shape_extractor.extract_ppt()
    return extracted_info

# Example usage
# if __name__ == "__main__":
#     from sys import argv
#
#     if len(argv) != 2:
#         print("Usage: python run_extractors.py <path_to_pptx>")
#         exit(1)
#
#     pptx_path = argv[1]
#     extracted_info = run_extractors(pptx_path, measurement_unit="pt")
#     print(dumps(extracted_info, indent=4))
