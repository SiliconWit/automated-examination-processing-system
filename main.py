import os
import toml
from modules.file_processing import *
from modules.data_consolidation import *
# from modules.utilities import *

if __name__ == "__main__":

    config_path = "config.toml"  # Specify the path to your TOML configuration file

    # Load paths from configuration
    config = toml.load(config_path)
    input_folder_path = config["input_folder"]["path"]
    output_excel_path = config["output_excel"]["path"]

    consolidate_mark_sheet(input_folder_path, output_excel_path, config_path)

    print(f"Mark sheet consolidated and saved as '{output_excel_path}'.")
