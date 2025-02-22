import os
import pandas as pd
import json
import shutil
import argparse

def find_excel_files(directory):
    """
    Recursively search for the first Excel file in each subfolder of the given directory.
    Once an Excel file is found, stop searching in that subfolder and move to the next.
    """
    excel_files = {}  # Dictionary to store folder names and their first-found Excel file
    for subfolder in os.listdir(directory):
        for root, dirs, files in os.walk(os.path.join(directory, subfolder)):
            for file in files:
                if file.endswith(".xlsx") or file.endswith(".xls"):
                    if subfolder not in excel_files:  # Only take the first file found in each subfolder
                        excel_files[subfolder] = os.path.join(root, file)
                    break  # Stop searching deeper in this folder once an Excel file is found
    return excel_files

def convert_excel_to_json(excel_path, save_folder):
    """
    Converts the given Excel file to JSON and saves it in the specified folder.
    """
    try:
        xls = pd.ExcelFile(excel_path)

        # Read dataset names
        dataset_info = xls.parse("Datasets")  # Assuming the sheet is named "Datasets"
        dataset_dict = dict(zip(dataset_info.iloc[:, 0].dropna(), dataset_info["Label"]))

        # JSON structure
        json_output = {
            "datasets": [],
            "standard": {"product": "sdtmig", "version": "3-3"},
            "codelists": []
        }

        # Process each dataset
        for dataset_name, dataset_label in dataset_dict.items():
            if dataset_name in xls.sheet_names:
                df = xls.parse(dataset_name, header=None)

                # Extract metadata from the first 4 rows
                column_names = df.iloc[0].astype(str).tolist()
                column_labels = df.iloc[1].astype(str).tolist()
                column_types = df.iloc[2].astype(str).tolist()
                column_lengths = df.iloc[3].fillna(0).astype(int).tolist()

                # Extract records from the 5th row onward
                records = df.iloc[4:].reset_index(drop=True)
                max_records = len(records)
                formatted_records = {}

                for i, col in enumerate(column_names):
                    col_type = column_types[i]
                    col_values = records.iloc[:, i].tolist()

                    formatted_values = []
                    for value in col_values:
                        if pd.isna(value):  # Handle missing values
                            formatted_values.append("" if col_type == "Char" else None)
                        else:
                            if col_type == "Char":
                                formatted_values.append(str(value))  # Store as string
                            elif col_type == "Num":
                                try:
                                    num_value = int(value) if str(value).isdigit() else value
                                except ValueError:
                                    num_value = value  # Keep as is if it cannot be converted
                                formatted_values.append(num_value)

                    # Ensure all columns have the same number of records (preserve empty rows)
                    while len(formatted_values) < max_records:
                        formatted_values.append("" if col_type == "Char" else None)

                    formatted_records[col] = formatted_values

                dataset_json = {
                    "filename": dataset_name,
                    "label": dataset_label,
                    "domain": dataset_name.split('.')[0].upper(),
                    "variables": [
                        {
                            "name": column_names[i],
                            "label": column_labels[i],
                            "type": column_types[i],
                            "length": column_lengths[i]
                        }
                        for i in range(len(column_names))
                    ],
                    "records": formatted_records
                }

                json_output["datasets"].append(dataset_json)

        # Ensure the save directory exists
        os.makedirs(save_folder, exist_ok=True)
        json_filename = os.path.join(save_folder, "converted_dataset.json")

        # Save JSON
        with open(json_filename, "w", encoding="utf-8") as f:
            json.dump(json_output, f, indent=4)

        print(f"✅ JSON saved: {json_filename}")

        # Check for XML file in the same directory as the Excel file
        xml_filename = os.path.splitext(excel_path)[0] + ".xml"
        if os.path.exists(xml_filename):
            shutil.copy(xml_filename, save_folder)
            print(f"✅ XML copied: {xml_filename}")

    except Exception as e:
        print(f"❌ Error processing {excel_path}: {e}")

def find_excel_for_rule(rule_id, datasets_directory):
    """
    Finds the Excel file corresponding to the given rule_id in the datasets_directory.
    """
    for root, dirs, files in os.walk(datasets_directory):
        if rule_id in dirs:
            rule_folder = os.path.join(root, rule_id)
            return find_excel_files(rule_folder)
    return {}

def process_rules_directory(published_rules_directory, datasets_directory):
    """
    Processes all rule folders in the given published_rules_directory,
    finds the corresponding Excel files in the datasets_directory,
    converts them to JSON, and saves them in 'json_datasets' under a folder named after the rule_id.
    """
    json_datasets_folder = "json_datasets"
    for rule_id in os.listdir(published_rules_directory):
        rule_folder = os.path.join(published_rules_directory, rule_id)
        if os.path.isdir(rule_folder):
            excel_files = find_excel_for_rule(rule_id, datasets_directory)
            for folder_name, excel_path in excel_files.items():
                save_folder = os.path.join(json_datasets_folder, rule_id)
                convert_excel_to_json(excel_path, save_folder)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert Excel files to JSON.")
    parser.add_argument("--rules", required=True, help="Path to the directory containing rule files.")
    parser.add_argument("--datasets", required=True, help="Path to the directory containing dataset files.")
    args = parser.parse_args()

    process_rules_directory(args.rules, args.datasets)