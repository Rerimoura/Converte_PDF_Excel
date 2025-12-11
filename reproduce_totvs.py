
import sys
import os
import PyPDF2
from extractors import RedeBizExtractor, MondelezExtractor

def dump_text():
    file_path = "rede_biz.pdf"
    if not os.path.exists(file_path):
        print(f"File {file_path} not found.")
        return

    print(f"Dumping text from {file_path}...")
    texto_completo = ""
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text = page.extract_text()
            if text:
                texto_completo += text + "\n"
    
    with open("debug_totvs_text.txt", "w", encoding="utf-8") as f:
        f.write(texto_completo)
    print("Text dumped to debug_totvs_text.txt")

    # Run extraction and save output
    extractor = RedeBizExtractor()
    dfs = extractor.extract(file_path)
    with open("debug_totvs_result.txt", "w", encoding="utf-8") as f:
        f.write(f"Found {len(dfs)} tables.\n")
        for i, df in enumerate(dfs):
            f.write(f"Table {i+1} Shape: {df.shape}\n")
            f.write(f"Columns: {df.columns.tolist()}\n")
            f.write(f"Dtypes:\n{df.dtypes}\n")
            f.write(df.to_string())
            f.write("\n" + "-" * 20 + "\n")
    print("Extraction result saved to debug_totvs_result.txt")

if __name__ == "__main__":
    dump_text()
