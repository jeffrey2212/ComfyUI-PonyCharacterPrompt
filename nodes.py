import random
import openpyxl
import os

class PonyCharacterPromptPicker:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "excel_path": ("STRING", {"default": os.path.join("resource", "pony_char_list.xlsx")}),
                "sheet_name": ("STRING", {"default": "Female character list"}),
                "column_letter": ("STRING", {"default": "A"}),
            },
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "pick_prompt"
    CATEGORY = "prompt"

    def pick_prompt(self, excel_path, sheet_name, column_letter):
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")

        workbook = openpyxl.load_workbook(excel_path, read_only=True)
        
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in the Excel file")

        sheet = workbook[sheet_name]
        column = sheet[column_letter]

        # Start from row 3 (index 2 in Python) to the end
        valid_cells = [cell.value for cell in column[2:] if cell.value is not None]

        if not valid_cells:
            raise ValueError(f"No valid values found in column {column_letter} starting from row 3")

        selected_prompt = random.choice(valid_cells)
        return (selected_prompt,)
