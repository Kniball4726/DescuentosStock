import os
import pdfplumber
import re

folder = "Mayoristas"
pdf_files = [f for f in os.listdir(folder) if f.endswith('.pdf')]

all_products = {}

for file_name in pdf_files:
    file_path = os.path.join(folder, file_name)
    print(f"\nProcessing {file_name}...")
    with pdfplumber.open(file_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue
            lines = text.split("\n")
            in_detalle = False
            for line in lines:
                line_stripped = line.strip()
                if "DETALLE" in line_stripped:
                    in_detalle = True
                    continue
                if in_detalle and ("Cant. Articulos:" in line_stripped or "OBSERVACIONES" in line_stripped):
                    in_detalle = False
                
                if in_detalle:
                    tokens = line_stripped.split()
                    if not tokens:
                        continue
                    # Check if the line looks like a product row: 
                    # starts with a code (integer) and ends with quantity (integer)
                    if tokens[0].isdigit() and tokens[-1].isdigit():
                        code = int(tokens[0])
                        qty = int(tokens[-1])
                        all_products[code] = all_products.get(code, 0) + qty
                        print(f"  Parsed: Code={code}, Qty={qty} | Line: '{line_stripped}'")

print("\nSummary of all parsed products:")
for code, qty in sorted(all_products.items()):
    print(f"Code: {code:4d} | Total Qty: {qty:3d}")
