# ExcelToJsonConverter
Converts CDISC Rules Engine Excel datasets to JSON for TEST command

## Introduction
The `ExcelToJsonConverter` is a Python script designed to convert CDISC Rules Engine Excel datasets into JSON format. This tool is particularly useful for preparing datasets for the TEST command.

## Usage
To use the `ExcelToJsonConverter`, you need to provide the paths to the directories containing the rule files and dataset files. The script will process these directories, find the corresponding Excel files, convert them to JSON, and save the JSON files in a specified folder.

### Command-Line Arguments
- `--rules`: Path to the directory containing rule files.
- `--datasets`: Path to the directory containing dataset files.

### Example
```sh
python XSLX-to-JSON.py --rules path/to/rules --datasets path/to/datasets
```
## Directory Structure
The script expects the following directory structure for the rule and dataset folders:

### Rules Directory
The rules directory should contain subfolders, each named after a specific rule ID. Each subfolder can contain multiple rule files.

```
rules_directory/
├── rule1/
│   ├── rule_file1.json
│   └── rule_file2.json
├── rule2/
│   ├── rule_file1.json
│   └── rule_file2.json
└── rule3/
    ├── rule_file1.json
    └── rule_file2.json

```

### Datasets Directory
The datasets directory should contain subfolders, each named after a specific dataset. Each subfolder can contain multiple dataset files.

```
datasets_directory/
├── dataset1/
│   ├── dataset_file1.xlsx
│   └── dataset_file2.xlsx
├── dataset2/
│   ├── dataset_file1.xlsx
│   └── dataset_file2.xlsx
└── dataset3/
    ├── dataset_file1.xlsx
    └── dataset_file2.xlsx

```
### Output
The converted JSON files will be saved in a folder named json_datasets, with subfolders named after the rule IDs.

```
json_datasets/
├── rule1/
│   └── converted_dataset.json
├── rule2/
│   └── converted_dataset.json
└── rule3/
    └── converted_dataset.json

```

## Dependencies
1. pandas
2. argparse
3. json
4. shutil
5. os

You can install the required dependencies using pip:
