import os
import toml
from modules.file_processing import *
from modules.data_consolidation import *
from modules.utilities import *
# from modules.utilities import course_code



if __name__ == "__main__":

    config_path = "config.toml"  # Specify the path to your TOML configuration file

    # Load paths from configuration
    config = toml.load(config_path)
    input_folder_path = config["input_folder"]["path"]
    consolidated_excel_output_path = config["consolidated_excel_output"]["path"]
    pass_list_pdf_output_path = config["pass_list_pdf_output"]["path"]
    supp_list_pdf_output_path = config["supp_list_pdf_output"]["path"]
    senate_doc_pdf_output_path = config["senate_documents_output"]["path"]
    mechatronics_units_path = config["mechatronics_engineering_units"]["path"]
    
    setup_logging()
    # Log the output to the file and print to the console
    log_print("Starting the program...")

    consolidate_mark_sheet(mechatronics_units_path, input_folder_path, consolidated_excel_output_path, pass_list_pdf_output_path, supp_list_pdf_output_path, senate_doc_pdf_output_path, config_path)

    log_print(f"Mark sheet consolidated and saved as '{consolidated_excel_output_path}'.")
