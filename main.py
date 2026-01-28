import os
import win32com.client as win32
from pathlib import Path

# ==============================
# CONFIGURAÇÃO
# ==============================

INPUT_DIR = r"C:\Users\escri\Documents\Repositorio\convertdoctodocx\import"   # pasta com .doc
OUTPUT_DIR = r"C:\Users\escri\Documents\Repositorio\convertdoctodocx\output"   # pasta para .docx

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

docs = list(Path(INPUT_DIR).glob("*.doc"))

if not docs:
    print("⚠️ Nenhum arquivo .doc encontrado")
else:
    for doc_path in docs:
        try:
            print(f"📄 Processando: {doc_path.name}")

            doc = word.Documents.Open(str(doc_path))

            output_path = Path(OUTPUT_DIR) / (doc_path.stem + ".docx")

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
