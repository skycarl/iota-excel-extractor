# Purpose

# How to run

1. Open a terminal (e.g., Command Prompt or Git Bash)
2. In the command prompt, change directory into the base repository directory, e.g. 

```bash
cd "D:\some_folder"
```

3. Run script, providing an input directory and an output file name:

```bash
poetry run python excel_extractor/main.py <input directory here> <output file name here>
```

Example:

```bash
poetry run python excel_extractor/main.py "D:\some_folder\some_subfolder" output_file_name.xlsx
```

Output is written into the `output/` directory of the repository. 

