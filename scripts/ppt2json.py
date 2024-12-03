# Enhanced Extraction Script with Debugging

import os
import logging
from json import dumps
from tqdm import tqdm
from pptbench.extractors.run_extractors import run_extractors

# Configure Logging
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
LOG_FILE = "extraction_debug.log"

logging.basicConfig(
    level=logging.DEBUG,  # Set to DEBUG to capture all levels of logs
    format=LOG_FORMAT,
    handlers=[
        logging.FileHandler(LOG_FILE, mode='w'),
        logging.StreamHandler()
    ]
)


def setup_directories(pptx_dir: str, output_dir: str):
    """
    Ensure that the output directory exists.
    """
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            logging.info(f"Created output directory: {output_dir}")
        except Exception as e:
            logging.error(f"Failed to create output directory '{output_dir}': {e}")
            raise


def process_pptx_files(pptx_dir: str, output_dir: str, measurement_unit: str = "emu"):
    """
    Process all .pptx files in the specified directory and extract information.
    """
    setup_directories(pptx_dir, output_dir)

    pptx_files = [f for f in os.listdir(pptx_dir) if f.endswith(".pptx")]
    total_files = len(pptx_files)

    if total_files == 0:
        logging.warning(f"No .pptx files found in directory: {pptx_dir}")
        return

    logging.info(f"Starting extraction of {total_files} .pptx files from '{pptx_dir}'")

    success_count = 0
    failure_count = 0

    for pptx in tqdm(pptx_files, desc="Processing PPTX files", unit="file"):
        pptx_path = os.path.join(pptx_dir, pptx)
        output_filename = pptx.replace(".pptx", ".json")
        output_path = os.path.join(output_dir, output_filename)

        logging.debug(f"Processing file: {pptx_path}")

        try:
            # Run extractors
            extracted_data = run_extractors(pptx_path, measurement_unit=measurement_unit)
            logging.debug(f"Extraction successful for file: {pptx_path}")

            # Write JSON output
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(dumps(extracted_data, indent=4))
            logging.info(f"Successfully extracted and saved to: {output_path}")

            success_count += 1
        except Exception as e:
            logging.error(f"Failed to process '{pptx_path}': {e}")
            failure_count += 1

    logging.info(f"Extraction completed: {success_count} succeeded, {failure_count} failed.")


if __name__ == "__main__":
    # Define directories
    PPTX_DIR = "../dataset/pptx"
    OUTPUT_DIR = "../dataset/json"

    # Optional: Change measurement_unit if needed (e.g., "pt", "emu")
    MEASUREMENT_UNIT = "emu"

    # Start processing
    process_pptx_files(PPTX_DIR, OUTPUT_DIR, measurement_unit=MEASUREMENT_UNIT)
