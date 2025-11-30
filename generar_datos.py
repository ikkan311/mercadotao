import pandas as pd
import math
import json

EXCEL_FILE = "Remito20251127_svg_convertido.xlsx"
OUTPUT_JSON = "productos.json"

def etiqueta_stock(valor):
    if pd.isna(valor):
        return "A PEDIDO"
    try:
        n = float(valor)
    except (ValueError, TypeError):
        return "A PEDIDO"
    if n == 0:
        return "SIN STOCK"
    return f"EN STOCK ({int(n) if n.is_integer() else n})"

def main():
    df = pd.read_excel(EXCEL_FILE)

    # Nos quedamos solo con filas que tengan nombre y categoría
    df = df[df["NOMBRE"].notna() & df["CATEGORIA"].notna()]

    # Orden por ORDEN CATALOGO
    if "ORDEN CATALOGO" in df.columns:
        df["ORDEN CATALOGO"] = pd.to_numeric(df["ORDEN CATALOGO"], errors="coerce")
        df = df.sort_values("ORDEN CATALOGO")
    else:
        df = df.sort_values("ORDEN")

    productos = []
    for _, row in df.iterrows():
        try:
            precio = float(row["PRECIO"])
        except Exception:
            precio = 0.0

        prod = {
            "codigo": int(row["ORDEN"]) if not pd.isna(row["ORDEN"]) else None,
            "orden_catalogo": (
                int(row["ORDEN CATALOGO"])
                if "ORDEN CATALOGO" in row and not pd.isna(row["ORDEN CATALOGO"])
                else None
            ),
            "nombre": str(row["NOMBRE"]).strip(),
            "categoria": str(row["CATEGORIA"]).strip(),
            "precio": precio,
            "stock_numero": None if pd.isna(row["STOCK"]) else row["STOCK"],
            "stock_texto": etiqueta_stock(row["STOCK"]),
            "novedad": "" if pd.isna(row.get("NOVEDAD", "")) else str(row["NOVEDAD"]).strip(),
            "disponibilidad": "" if pd.isna(row.get("DISPONIBILIDAD", "")) else str(row["DISPONIBILIDAD"]).strip(),
            "promocion": "" if pd.isna(row.get("PROMOCION", "")) else str(row["PROMOCION"]).strip(),
            "descripcion": "" if pd.isna(row.get("Descripción", "")) else str(row["Descripción"]).strip(),
            "imagen": "" if pd.isna(row.get("@imagen", "")) else str(row["@imagen"]).strip(),
        }
        productos.append(prod)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(productos, f, ensure_ascii=False, indent=2)

    print(f"Generado {OUTPUT_JSON} con {len(productos)} productos.")

if __name__ == "__main__":
    main()
