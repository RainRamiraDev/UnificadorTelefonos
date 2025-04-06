import os
import sys
import pandas as pd

# Nombre del archivo maestro
ARCHIVO_MAESTRO = 'maestro.xlsx'

# 🔁 Si el argumento es "reset", se vacía el archivo maestro
if len(sys.argv) > 1 and sys.argv[1].lower() == 'reset':
    df_vacio = pd.DataFrame(columns=['Telefono'])
    df_vacio.to_excel(ARCHIVO_MAESTRO, index=False)
    print("⚠️ El archivo maestro fue reseteado.")
    exit()

# Cargar archivo maestro
if os.path.exists(ARCHIVO_MAESTRO):
    df_maestro = pd.read_excel(ARCHIVO_MAESTRO)
else:
    df_maestro = pd.DataFrame(columns=['Telefono'])

print(f"📖 Maestro cargado con {len(df_maestro)} registros.")

# Obtener lista de archivos Excel en la carpeta
CARPETA_EXCELS = 'Excels'
archivos = [os.path.join(CARPETA_EXCELS, f) for f in os.listdir(CARPETA_EXCELS) if f.endswith('.xlsx')]

print(f"🗂 Archivos encontrados: {archivos}")

datos_nuevos = []

# Función para encontrar la columna que contiene "telefono"
def encontrar_columna_telefono(df):
    for i, row in df.iterrows():
        for col in df.columns:
            valor = str(row[col]).strip().lower()
            if valor == 'telefono':
                return i, col
    return None, None

for archivo in archivos:
    print(f"🔍 Leyendo: {archivo}")
    try:
        df = pd.read_excel(archivo, header=None)

        fila_titulo, col_telefono = encontrar_columna_telefono(df)
        if fila_titulo is None:
            print(f"⚠️ No se encontró columna 'Telefono' en {archivo}")
            continue

        col_idx = df.columns.get_loc(col_telefono)
        telefonos = df.iloc[fila_titulo+1:, col_idx].dropna()
        telefonos = telefonos.astype(str).str.strip()
        datos_nuevos.extend(telefonos.tolist())

    except Exception as e:
        print(f"❌ Error leyendo {archivo}: {e}")

telefonos_existentes = df_maestro['Telefono'].astype(str).str.strip().tolist()
telefonos_para_agregar = [t for t in datos_nuevos if t not in telefonos_existentes]

if telefonos_para_agregar:
    df_nuevos = pd.DataFrame({'Telefono': telefonos_para_agregar})
    df_maestro = pd.concat([df_maestro, df_nuevos], ignore_index=True)
    df_maestro.to_excel(ARCHIVO_MAESTRO, index=False)
    print(f"✅ {len(telefonos_para_agregar)} teléfonos agregados al maestro.")
else:
    print("ℹ️ No se encontraron datos nuevos para agregar.")
