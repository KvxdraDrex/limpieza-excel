import argparse
import pandas as pd
import re

# === Limpieza (incluye efecto ESPACIOS) ===
def limpiar_texto(texto: str) -> str:
    if pd.isna(texto):
        return texto
    s = str(texto)

    # Normalización de caracteres (tus reglas)
    rep = {
        "Á":"A","Ä":"A","Ā":"A",
        "É":"E",
        "Í":"I","Ì":"I",
        "Ó":"O","Ö":"O","Õ":"O","Ø":"O",
        "Ü":"U","Ú":"U",
        "Ñ":"N",
        "Ç":"C","ç":"c",
        "ó":"o","ö":"o","õ":"o",
        "í":"i","ü":"u",
        "š":"s","ý":"y",
    }
    for a,b in rep.items():
        s = s.replace(a,b)

    # Limpieza de símbolos (macro)
    s = (s.replace("$","S").replace(":","").replace("(","").replace(")","")
           .replace("#","\t").replace("°","").replace("-"," ").replace("/"," ")
           .replace('"',"").replace("”","").replace("“","")
           .replace("'","\t").replace("_","\t").replace(",","\t")
           .replace("&","AND"))

    # Efecto ESPACIOS: quitar espacios en extremos y dejar uno solo entre palabras
    # (También elimina espacios no separables \u00A0)
    s = s.replace("\u00A0", " ")
    s = s.strip()
    s = re.sub(r" {2,}", " ", s)  # colapsa 2+ espacios a 1 (sin tocar tabs)
    return s

def main():
    ap = argparse.ArgumentParser(description="Elimina fila 1 y limpia columnas A y B (incluye efecto ESPACIOS).")
    ap.add_argument("--in",  dest="infile",  required=True)
    ap.add_argument("--out", dest="outfile", required=True)
    ap.add_argument("--sheet", dest="sheet", default=None)  # primera hoja por defecto
    args = ap.parse_args()

    # Leer SIN encabezado para poder eliminar la fila 1 real del archivo
    df = pd.read_excel(args.infile, sheet_name=args.sheet, header=None, dtype=str)

    # 1) Eliminar la fila 1 (índice 0)
    if len(df) > 0:
        df = df.iloc[1:, :].reset_index(drop=True)

    # 2) Limpiar columnas A y B si existen
    if df.shape[1] >= 1:
        df.iloc[:, 0] = df.iloc[:, 0].apply(limpiar_texto)
    if df.shape[1] >= 2:
        df.iloc[:, 1] = df.iloc[:, 1].apply(limpiar_texto)

    # 3) Guardar sin encabezados ni columnas extra
    df.to_excel(args.outfile, index=False, header=False)
    print("Proceso de limpieza completado (Python).")

if __name__ == "__main__":
    main()
