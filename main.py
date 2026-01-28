import os
import win32com.client as win32
from pathlib import Path

# ==============================
# CONFIGURAÇÃO
# ==============================
FULL_LOCAL_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = r"\import"   # pasta com .doc
OUTPUT_DIR = r"\output"   # pasta para .docx

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==============================
# INICIALIZA WORD
# ==============================

word = win32.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0  # evita popups

# ==============================
# CONVERSÃO
# ==============================

docs = list(Path(FULL_LOCAL_DIR+INPUT_DIR).glob("*"))
print(f"{Path(FULL_LOCAL_DIR+INPUT_DIR)} ")

if not docs:
    print("⚠️ Nenhum arquivo .doc encontrado")
else:
    for doc_path in docs:
        try:
            print(f"📄 Processando: {doc_path.name}")

            doc = word.Documents.Open(str(doc_path))

            output_path = Path(FULL_LOCAL_DIR+OUTPUT_DIR) / (doc_path.stem + ".docx")

            # 16 = wdFormatXMLDocument (.docx)
            doc.SaveAs(str(output_path), FileFormat=16)

            doc.Close(False)

            print("✅ Convertido com sucesso")

        except Exception as e:
            print(f"❌ Erro em {doc_path.name}: {e}")

# ==============================
# FINALIZA
# ==============================

word.Quit()
print("\n🎉 Conversão finalizada")
