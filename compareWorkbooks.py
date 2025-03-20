from utilities import clean_parameters
import xlwings as xw
import json
import re
import os
import difflib
import shutil


class CompareWorkbooks():
    def __init__(self, wb_new_path, wb_old_path):
        self.wb_new_path = wb_new_path
        self.wb_old_path = wb_old_path
        self.wb_new = xw.Book(self.wb_new_path)
        self.wb_old = xw.Book(self.wb_old_path)
        
        # Remove folders if they exist
        self.cleanup_folders()

    def cleanup_folders(self):
        # List of folders to check and remove
        folders_to_remove = ["Funzioni estratte new", "Funzioni estratte old", "difference"]

        for folder in folders_to_remove:
            if os.path.exists(folder):
                print(f"Removing existing folder: {folder}")
                shutil.rmtree(folder)

    def extractFunctionsInfo(self, new_old, write_json=True):
        
        if new_old == "New":
            wb = self.wb_new 
        elif new_old == "Old":
            wb = self.wb_old
        else:
            raise ValueError("Parameter new/old must be either 'New' or 'Old'")
        
        functions = []
        function_pattern = re.compile(r'function\s+(\w+)\s*\((.*)', re.IGNORECASE)

        for component in wb.api.VBProject.VBComponents:
            code_module = component.CodeModule
            total_lines = code_module.CountOfLines
            line_num = 1

            while line_num <= total_lines:
                line_content = code_module.Lines(line_num, 1).strip()
                if line_content.startswith("Function"):
                    declaration_lines = [line_content]
                    while ")" not in line_content and line_num < total_lines:
                        line_num += 1
                        line_content = code_module.Lines(line_num, 1).strip()
                        declaration_lines.append(line_content)

                    full_declaration = " ".join(declaration_lines)
                    match = function_pattern.search(full_declaration)

                    if match:
                        function_name = match.group(1)
                        params_raw = full_declaration.split('(', 1)[1].rsplit(')', 1)[0]
                        cleaned_params = clean_parameters(params_raw)
                        functions.append({
                            "module": component.Name,
                            "line": line_num,
                            "function_name": function_name,
                            "parameters": cleaned_params
                        })
                line_num += 1

        if write_json:
            with open("results.json", "w", encoding="utf-8") as out_file:
                json.dump(functions, out_file, indent=6)
            print(" Results have been stored in results.json file!")
        else:
            return functions

    def extractFunctionsCode(self, new: bool, old: bool, write_bas=True):
        wb = self.wb_new if new else self.wb_old
        output_dir = "Funzioni estratte new" if new else "Funzioni estratte old"
        functions_dict = {}

        for component in wb.api.VBProject.VBComponents:
            code_module = component.CodeModule
            lines_num = code_module.CountOfLines
            if lines_num == 0:
                continue

            try:
                code_lines = code_module.Lines(1, lines_num).splitlines()
            except Exception as e:
                print(f" Error in module '{component.Name}': {e}")
                continue

            inside_function = False
            current_function_name = ""
            current_function_code = []

            for line in code_lines:
                check_line = line.lstrip()
                if check_line.startswith(("Sub ", "Function ")):
                    inside_function = True
                    current_function_name = check_line.split()[1].split("(")[0]
                    current_function_code = [line]
                elif inside_function:
                    current_function_code.append(line)
                    if check_line.startswith(("End Sub", "End Function")):
                        functions_dict[current_function_name] = "\n".join(current_function_code)
                        inside_function = False
                        current_function_name = ""
                        current_function_code = []

        if write_bas:
            os.makedirs(output_dir, exist_ok=True)
            for func_name, code in functions_dict.items():
                safe_func_name = "".join(c for c in func_name if c.isalnum() or c in (' ', '_')).rstrip()
                filename = os.path.join(output_dir, f"{safe_func_name}.bas")
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(code)
            print(f" Results have been stored in the folder '{output_dir}'")
        else:
            return functions_dict

    def showDifferences(self):
        # Extract functions first
        self.extractFunctionsCode(new=True, old=False)
        self.extractFunctionsCode(new=False, old=True)

        folder_new = "Funzioni estratte new"
        folder_old = "Funzioni estratte old"
        output_folder = "difference"

        # Check if folders exist
        if not os.path.exists(folder_new):
            print(f"Folder '{folder_new}' not found.")
            return
        if not os.path.exists(folder_old):
            print(f"Folder '{folder_old}' not found.")
            return

        # Make sure output folder exists
        os.makedirs(output_folder, exist_ok=True)

        # Get list of .bas files in both folders
        files_new = set(f for f in os.listdir(folder_new) if f.endswith(".bas"))
        files_old = set(f for f in os.listdir(folder_old) if f.endswith(".bas"))

        # Find common files to compare
        common_files = files_new & files_old

        if not common_files:
            print("No common .bas files to compare.")
            return

        for filename in common_files:
            file_new_path = os.path.join(folder_new, filename)
            file_old_path = os.path.join(folder_old, filename)

            with open(file_new_path, "r", encoding="utf-8") as f1:
                lines_new = f1.readlines()
            with open(file_old_path, "r", encoding="utf-8") as f2:
                lines_old = f2.readlines()

            diff = difflib.unified_diff(
                lines_old, lines_new,
                fromfile=f"{folder_old}/{filename}",
                tofile=f"{folder_new}/{filename}",
                lineterm="\n",
                n = 1000
            )

            # Check if there are any differences
            diff_lines = list(diff)
            if diff_lines:
                diff_output_path = os.path.join(output_folder, f"diff_{filename}")
                
                with open(diff_output_path, "w", encoding="utf-8") as diff_file:
                    for line in diff_lines:
                        diff_file.write(line + "\n")
                print(f"Diff for '{filename}' saved in '{diff_output_path}'")
            else:
                print(f"No differences found for '{filename}'.")

        only_in_new = files_new - files_old
        only_in_old = files_old - files_new

        if only_in_new:
            print(f"Files only in '{folder_new}': {only_in_new}")
        if only_in_old:
            print(f"Files only in '{folder_old}': {only_in_old}")



    def close_workbooks(self):
        self.wb_new.close()
        self.wb_old.close()


if __name__ == "__main__":
    
    new_file = "Aggiornamento_B4_COLG4_v5_CIB.xlsm"
    old_file = "Aggiornamento_B4_primo_rilascio_v2_CIB.xlsm"

    comparer = CompareWorkbooks(new_file, old_file)
    comparer.showDifferences()
    comparer.close_workbooks()
