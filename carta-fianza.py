import pandas as pd
from thefuzz import process, fuzz
import re

# ==========================================
# CONFIGURACIÓN DE ARCHIVOS Y HOJAS
# ==========================================
NOMBRE_ARCHIVO = 'Cuestionario_ServBCP (Carta Fianza) - Noviembre.xlsx'
HOJA_INPUT = 'Credicorp'
HOJA_BD = 'BD'

# ==========================================
# 1. CARGA DE DATOS
# ==========================================
print(f"Leyendo archivo: {NOMBRE_ARCHIVO}...")
try:
    df_input = pd.read_excel(NOMBRE_ARCHIVO, sheet_name=HOJA_INPUT)
    df_bd = pd.read_excel(NOMBRE_ARCHIVO, sheet_name=HOJA_BD)
except FileNotFoundError:
    print("❌ ERROR: No encuentro el archivo. Verifica que esté en la misma carpeta.")
    exit()

# ==========================================
# 2. LIMPIEZA DE DATOS (NORMALIZACIÓN)
# ==========================================
print("Limpiando nombres y estandarizando países...")

def limpiar_nombre(nombre):
    if pd.isna(nombre): return ""
    nombre = str(nombre).lower().strip()
    # Quitamos S.A., S.A.C, etc. para comparar solo el nombre real
    nombre = re.sub(r'\b(s\.?a\.?|s\.?a\.?c\.?|e\.?i\.?r\.?l\.?|ltd|inc|y filiales|spa)\b', '', nombre)
    nombre = re.sub(r'[^\w\s]', '', nombre) # Quita signos raros
    return nombre.strip()

# Diccionario para convertir el país del Input al código de tu BD (CHI, PER, etc.)
mapa_paises = {
    'perú': 'PER', 'peru': 'PER',
    'chile': 'CHI', 
    'colombia': 'COL',
    'bolivia': 'BOL'
}

# Aplicamos limpieza en copias temporales (para no dañar los datos originales del reporte)
df_input['Empresa_Limpia'] = df_input['Nombre de la empresa'].apply(limpiar_nombre)
df_input['Pais_Norm'] = df_input['Pais'].str.lower().map(mapa_paises).fillna(df_input['Pais'])

df_bd['Cliente_Limpio'] = df_bd['CLIENTE'].apply(limpiar_nombre)
df_bd['PAIS_BD'] = df_bd['PAIS'].astype(str).str.strip()

# ==========================================
# 3. MOTOR DE EMPAREJAMIENTO (FUZZY MATCHING)
# ==========================================
print("🤖 Buscando coincidencias en la Base de Datos...")

def buscar_match(row):
    nombre_buscado = row['Empresa_Limpia']
    pais_buscado = str(row['Pais_Norm']).strip()
    
    # Filtramos la BD para buscar solo en el país correcto (Más rápido y preciso)
    bd_filtrada = df_bd[df_bd['PAIS_BD'] == pais_buscado]
    lista_candidatos = bd_filtrada['Cliente_Limpio'].unique()
    
    if len(lista_candidatos) == 0:
        return "SIN DATA EN PAIS", 0
    
    # Buscamos el nombre más parecido
    resultado = process.extractOne(nombre_buscado, choices=lista_candidatos, scorer=fuzz.token_sort_ratio)
    if resultado is None:
        return "SIN COINCIDENCIA", 0
    mejor_match, puntaje = resultado
    return mejor_match, puntaje

# Ejecutamos la búsqueda fila por fila
df_input[['MATCH_EN_BD', 'PORCENTAJE']] = df_input.apply(
    lambda x: pd.Series(buscar_match(x)), axis=1
)

# ==========================================
# 4. PREPARAR HOJA "REPORTE"
# ==========================================
print("Armando el reporte final...")

def obtener_color(puntaje):
    if puntaje >= 85: return 'VERDE'
    elif puntaje >= 50: return 'MORADO'
    else: return 'ROJO'

df_input['SEMAFORO'] = df_input['PORCENTAJE'].apply(obtener_color)

# Seleccionamos las columnas que pide tu reporte + Las de validación
# Usamos los nombres exactos que me diste para la hoja Reporte
cols_reporte = [
    'Pais', 
    'Nombre de la empresa', 
    'IDC', 
    'Nemonico', 
    'Se ha prestado servicio de carta fianza?'
]

# Creamos el DataFrame final con las columnas originales + Resultados
df_final = df_input[cols_reporte].copy()
df_final['NOMBRE_ENCONTRADO_BD'] = df_input['MATCH_EN_BD']
df_final['%_COINCIDENCIA'] = df_input['PORCENTAJE']
df_final['ESTADO'] = df_input['SEMAFORO']

# ==========================================
# 5. EXPORTAR CON COLORES (Estilo Excel)
# ==========================================
archivo_salida = 'Reporte_Final_Procesado.xlsx'

# Función para pintar las celdas según el valor
def colorear_celdas(val):
    color = ''
    if val == 'VERDE':
        color = 'background-color: #C6EFCE; color: #006100' # Verde Excel
    elif val == 'MORADO':
        color = 'background-color: #E6E6FA; color: #4B0082' # Morado suave
    elif val == 'ROJO':
        color = 'background-color: #FFC7CE; color: #9C0006' # Rojo Excel
    return color

# Aplicamos el estilo y guardamos
print(f"Guardando {archivo_salida} con colores...")
with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
    # Convertimos a Excel aplicando la función de colores a la columna ESTADO
    df_final.style.map(colorear_celdas, subset=['ESTADO']).to_excel(writer, sheet_name='Reporte', index=False)

print("✅ ¡LISTO! Abre 'Reporte_Final_Procesado.xlsx'. La hoja 'Reporte' ya tiene los colores.")