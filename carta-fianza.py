import pandas as pd
from thefuzz import process, fuzz
import re

# ==========================================
# PARAMETRIZACION DE ARCHIVOS Y HOJAS
# ==========================================
NOMBRE_ARCHIVO = 'Carta_Fianza_Plantilla.xlsx'
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
    print("ERROR: No se encontro el archivo. Verifica que este en la misma carpeta.")
    raise SystemExit

# ==========================================
# 2. LIMPIEZA DE DATOS (NORMALIZACION)
# ==========================================
print("Limpiando nombres y estandarizando paises...")

def limpiar_nombre(nombre):
    if pd.isna(nombre): 
        return ""
    nombre = str(nombre).lower().strip()
    # Quitamos sufijos comunes: S.A., S.A.C, SPA, etc.
    nombre = re.sub(r'\b(s\.?a\.?c?\.?|e\.?i\.?r\.?l\.?|ltd|inc|y filiales|spa)\b', ' ', nombre)
    # Quitamos "serie X"
    nombre = re.sub(r'\bserie\s*"?[a-z0-9]+"?\b', ' ', nombre)
    # Quitamos signos raros y normalizamos espacios
    nombre = re.sub(r'[^\w\s]', ' ', nombre)
    nombre = re.sub(r'\s+', ' ', nombre).strip()
    return nombre

# Diccionario para convertir el pais del Input al codigo en BD (CHI, PER, etc.)
mapa_paises = {
    'perú': 'PER', 'peru': 'PER',
    'chile': 'CHI', 
    'colombia': 'COL',
    'bolivia': 'BOL'
}

# Input
df_input['Empresa_Limpia'] = df_input['Nombre de la empresa'].apply(limpiar_nombre)
df_input['Pais_Norm'] = df_input['Pais'].astype(str).str.lower().map(mapa_paises).fillna(df_input['Pais'])

# BD
df_bd['CLIENTE'] = df_bd['CLIENTE'].astype(str)
df_bd['Cliente_Limpio'] = df_bd['CLIENTE'].apply(limpiar_nombre)
df_bd['PAIS_BD'] = df_bd['PAIS'].astype(str).str.strip()

# Quitamos filas vacías de BD
df_bd = df_bd[df_bd['Cliente_Limpio'] != ''].copy()

# ==========================================
# 3. ALGORITMO DE EMPAREJAMIENTO (FUZZY MATCHING)
# ==========================================
print("Buscando coincidencias en la Base de Datos...")

def calcular_similitud_palabras(nombre1, nombre2):
    """
    Calcula similitud basada en palabras compartidas.
    Penaliza cuando no hay palabras en comun.
    """
    palabras1 = set(nombre1.split())
    palabras2 = set(nombre2.split())
    
    # Quitamos palabras muy cortas (de, el, la, etc.)
    palabras1 = {p for p in palabras1 if len(p) > 2}
    palabras2 = {p for p in palabras2 if len(p) > 2}
    
    if not palabras1 or not palabras2:
        return 0
    
    # Palabras exactas en comun
    comunes = palabras1 & palabras2
    if comunes:
        return len(comunes) / max(len(palabras1), len(palabras2)) * 100
    
    # Si no hay palabras exactas, buscar palabras parciales (que empiecen igual)
    parciales = 0
    for p1 in palabras1:
        for p2 in palabras2:
            # Si una palabra contiene a la otra o viceversa (minimo 4 chars)
            if len(p1) >= 4 and len(p2) >= 4:
                if p1.startswith(p2[:4]) or p2.startswith(p1[:4]):
                    parciales += 1
                    break
    
    return parciales / max(len(palabras1), len(palabras2)) * 50  # Max 50% por parciales

def buscar_match(row):
    nombre_buscado = row['Empresa_Limpia']
    
    if not nombre_buscado:
        return "SIN DATA", 0, "", ""
    
    lista_candidatos = df_bd['Cliente_Limpio'].unique()
    if len(lista_candidatos) == 0:
        return "SIN DATA", 0, "", ""

    # Paso 1: Obtener top 5 candidatos con ratio basico
    top_candidatos = process.extract(
        nombre_buscado, 
        choices=lista_candidatos, 
        scorer=fuzz.ratio,  # Comparacion directa caracter por caracter
        limit=5
    )
    
    if not top_candidatos:
        return "SIN COINCIDENCIA", 0, "", ""
    
    # Paso 2: Re-evaluar con nuestra logica de palabras
    mejor_match = None
    mejor_puntaje = 0
    
    for candidato, puntaje_fuzz in top_candidatos:
        # Calcular similitud por palabras compartidas
        puntaje_palabras = calcular_similitud_palabras(nombre_buscado, candidato)
        
        # Puntaje combinado: 40% fuzzy + 60% palabras
        puntaje_final = (puntaje_fuzz * 0.4) + (puntaje_palabras * 0.6)
        
        # Penalizacion si el candidato es mucho mas corto
        ratio_longitud = len(candidato) / len(nombre_buscado) if len(nombre_buscado) > 0 else 0
        if ratio_longitud < 0.5:  # Si el candidato es menos de la mitad de largo
            puntaje_final *= 0.5  # Penalizar 50%
        
        if puntaje_final > mejor_puntaje:
            mejor_puntaje = puntaje_final
            mejor_match = candidato
    
    if mejor_match is None or mejor_puntaje < 10:
        return "SIN COINCIDENCIA", 0, "", ""

    # Recuperamos el registro original
    registro_bd = df_bd[df_bd['Cliente_Limpio'] == mejor_match].iloc[0]
    cliente_original = registro_bd['CLIENTE']
    codunicocli = registro_bd['CODUNICOCLI'] if 'CODUNICOCLI' in registro_bd else ""
    pais_match = registro_bd['PAIS_BD']

    return cliente_original, int(mejor_puntaje), codunicocli, pais_match

# Ejecutamos la busqueda fila por fila
df_input[['MATCH_EN_BD', 'PORCENTAJE', 'CODUNICOCLI_BD', 'PAIS_MATCH']] = df_input.apply(
    lambda x: pd.Series(buscar_match(x)), axis=1
)

# ==========================================
# 4. PREPARAR HOJA "REPORTE"
# ==========================================
print("Armando el reporte final...")

def obtener_color(puntaje):
    # <<< CORREGIDO: usar >= (Python) en vez de &gt;= (HTML)
    if puntaje >= 85:
        return 'VERDE'
    elif puntaje >= 50:
        return 'MORADO'
    else:
        return 'ROJO'

df_input['SEMAFORO'] = df_input['PORCENTAJE'].apply(obtener_color)

# Creamos el DataFrame final con las columnas del reporte
df_final = pd.DataFrame()
df_final['Pais'] = df_input['Pais']
df_final['Nombre de la empresa'] = df_input['Nombre de la empresa']

# IDC: Si es VERDE, poner el CODUNICOCLI de la BD, si no, dejar el original
df_final['IDC'] = df_input.apply(
    lambda row: row['CODUNICOCLI_BD'] if row['SEMAFORO'] == 'VERDE' else row.get('IDC', ''), axis=1
)

df_final['Nemonico'] = df_input.get('Nemonico', '')

# Se ha prestado carta fianza: SI cuando es VERDE, vacío en otros casos
df_final['Se ha prestado servicio de carta fianza?'] = df_input['SEMAFORO'].apply(
    lambda x: 'SI' if x == 'VERDE' else ''
)

df_final['NOMBRE_ENCONTRADO_BD'] = df_input['MATCH_EN_BD']
df_final['%_COINCIDENCIA'] = df_input['PORCENTAJE']
df_final['ESTADO'] = df_input['SEMAFORO']
df_final['PAIS_MATCH'] = df_input['PAIS_MATCH']  # para transparencia

# ==========================================
# 5. EXPORTAR AL EXCEL CON COLORES
# ==========================================
archivo_salida = 'Reporte_Final_Procesado.xlsx'

def colorear_celdas(val):
    if val == 'VERDE':
        return 'background-color: #C6EFCE; color: #006100' # Verde Excel
    elif val == 'MORADO':
        return 'background-color: #E6E6FA; color: #4B0082' # Morado suave
    elif val == 'ROJO':
        return 'background-color: #FFC7CE; color: #9C0006' # Rojo Excel
    return ''

print(f"Guardando {archivo_salida} ...")
with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
    df_final.style.map(colorear_celdas, subset=['ESTADO']).to_excel(
        writer, sheet_name='Reporte', index=False
    )

print("Reporte Listo! Abre 'Reporte_Final_Procesado.xlsx'. La hoja 'Reporte' ya tiene los colores con sus resultados.")