import pandas as pd
from thefuzz import process, fuzz
import re

# ==========================================
# PARAMETRIZACION DE ARCHIVOS Y HOJAS
# ==========================================
NOMBRE_ARCHIVO = 'prueba.xlsx'
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

# Palabras comunes que NO identifican a una empresa (stopwords)
STOPWORDS = {
    # Sufijos legales
    'sa', 'sac', 'saa', 'eirl', 'ltd', 'inc', 'spa', 'corp', 'group', 'grupo',
    # Articulos y preposiciones
    'de', 'del', 'la', 'el', 'los', 'las', 'y', 'e', 'en', 'a', 'the', 'of',
    # Paises
    'peru', 'chile', 'colombia', 'bolivia', 'panama', 'per', 'chi', 'col', 'brasil', 'mexico',
    # Tipos de empresa genericos
    'empresa', 'empresas', 'compania', 'sociedad', 'corporacion', 'inversiones', 'holding', 'holdings',
    'banco', 'bank', 'financial', 'financiera', 'financiero',
    # Sectores/industrias (muy genericos)
    'minera', 'mineras', 'minas', 'mineros', 'mining',
    'energia', 'energy', 'generacion', 'distribucion', 'electrica', 'electric',
    'construccion', 'construcciones', 'constructora',
    'servicios', 'service', 'services', 'comercial', 'industrial',
    'retail', 'internacional', 'international', 'sucursal',
    # Otras palabras genericas que causan falsos positivos
    'diagnostico', 'instituto', 'interconexion', 'operadores', 'operador',
    'open', 'plaza', 'mall', 'centro', 'tienda', 'tiendas'
}

def extraer_palabras_clave(nombre):
    """Extrae las palabras significativas de un nombre (no stopwords)."""
    palabras = nombre.lower().split()
    # Filtrar stopwords y palabras muy cortas
    clave = [p for p in palabras if p not in STOPWORDS and len(p) >= 3]
    return clave

def obtener_palabra_distintiva(palabras_clave):
    """
    Obtiene la palabra MAS distintiva (la mas larga y unica).
    Esta es la que DEBE coincidir para un match alto.
    """
    if not palabras_clave:
        return None
    # La palabra mas larga suele ser la mas distintiva
    return max(palabras_clave, key=len)

def calcular_score_avanzado(nombre_input, nombre_bd):
    """
    Calcula un score de similitud mas inteligente.
    Prioriza coincidencias de palabras clave distintivas.
    """
    palabras_input = extraer_palabras_clave(nombre_input)
    palabras_bd = extraer_palabras_clave(nombre_bd)
    
    if not palabras_input or not palabras_bd:
        return 0, False
    
    # Palabra distintiva del input (la mas importante)
    palabra_distintiva = obtener_palabra_distintiva(palabras_input)
    
    # 1. Buscar palabras clave exactas en comun
    comunes_exactas = set(palabras_input) & set(palabras_bd)
    
    # 2. Verificar si la palabra distintiva coincide EXACTAMENTE
    distintiva_coincide_exacta = palabra_distintiva in comunes_exactas if palabra_distintiva else False
    
    # 3. Verificar coincidencia parcial de palabra distintiva (prefijo 5+ chars)
    distintiva_coincide_parcial = False
    if palabra_distintiva and not distintiva_coincide_exacta:
        for p_bd in palabras_bd:
            if len(palabra_distintiva) >= 5 and len(p_bd) >= 5:
                if palabra_distintiva[:5] == p_bd[:5]:
                    distintiva_coincide_parcial = True
                    break
    
    # 4. Buscar coincidencias parciales de otras palabras
    comunes_parciales = 0
    for p_in in palabras_input:
        if p_in in comunes_exactas:
            continue
        for p_bd in palabras_bd:
            if p_bd in comunes_exactas:
                continue
            min_len = min(len(p_in), len(p_bd))
            if min_len >= 4:
                prefijo = min(4, min_len)
                if p_in[:prefijo] == p_bd[:prefijo]:
                    comunes_parciales += 0.5
                    break
    
    total_comunes = len(comunes_exactas) + comunes_parciales
    max_palabras = max(len(palabras_input), len(palabras_bd))
    
    # Score base por palabras clave
    score_palabras = (total_comunes / max_palabras) * 100 if max_palabras > 0 else 0
    
    # BONUS/PENALIZACION por palabra distintiva
    if distintiva_coincide_exacta:
        score_palabras = min(100, score_palabras + 30)  # Bonus grande
    elif distintiva_coincide_parcial:
        score_palabras = min(100, score_palabras + 15)  # Bonus medio
    else:
        # PENALIZAR si la palabra distintiva NO coincide
        score_palabras = score_palabras * 0.6  # Penalizacion fuerte
    
    return score_palabras, distintiva_coincide_exacta

def buscar_match(row):
    nombre_buscado = row['Empresa_Limpia']
    
    if not nombre_buscado or len(nombre_buscado.strip()) < 2:
        return "SIN DATA", 0, "", "", False
    
    lista_candidatos = df_bd['Cliente_Limpio'].unique().tolist()
    if len(lista_candidatos) == 0:
        return "SIN DATA", 0, "", "", False

    # PASO 1: Obtener top 30 candidatos usando token_set_ratio
    top_candidatos = process.extract(
        nombre_buscado, 
        choices=lista_candidatos, 
        scorer=fuzz.token_set_ratio,
        limit=30
    )
    
    if not top_candidatos:
        return "SIN COINCIDENCIA", 0, "", "", False
    
    # Obtener palabra distintiva del input para verificar si es corta
    palabras_input = extraer_palabras_clave(nombre_buscado)
    palabra_distintiva = obtener_palabra_distintiva(palabras_input)
    palabra_distintiva_corta = palabra_distintiva and len(palabra_distintiva) < 4
    
    # PASO 2: Re-evaluar con nuestro score de palabras clave
    mejor_match = None
    mejor_puntaje = 0
    mejor_distintiva = False
    
    for candidato, puntaje_fuzz in top_candidatos:
        # Calcular score por palabras clave
        score_palabras, distintiva_coincide = calcular_score_avanzado(nombre_buscado, candidato)
        
        # Puntaje combinado: 40% fuzzy token_set + 60% palabras clave
        puntaje_final = (puntaje_fuzz * 0.4) + (score_palabras * 0.6)
        
        # BONUS si palabra distintiva coincide exacta
        if distintiva_coincide:
            puntaje_final = min(100, puntaje_final + 10)
        
        if puntaje_final > mejor_puntaje:
            mejor_puntaje = puntaje_final
            mejor_match = candidato
            mejor_distintiva = distintiva_coincide
    
    if mejor_match is None or mejor_puntaje < 15:
        return "SIN COINCIDENCIA", 0, "", "", False

    # Recuperamos el registro original de la BD
    registro_bd = df_bd[df_bd['Cliente_Limpio'] == mejor_match].iloc[0]
    cliente_original = registro_bd['CLIENTE']
    codunicocli = registro_bd['CODUNICOCLI'] if 'CODUNICOCLI' in registro_bd else ""
    pais_match = registro_bd['PAIS_BD']

    return cliente_original, int(mejor_puntaje), codunicocli, pais_match, palabra_distintiva_corta

# Ejecutamos la busqueda fila por fila
df_input[['MATCH_EN_BD', 'PORCENTAJE', 'CODUNICOCLI_BD', 'PAIS_MATCH', 'DISTINTIVA_CORTA']] = df_input.apply(
    lambda x: pd.Series(buscar_match(x)), axis=1
)

# ==========================================
# 4. PREPARAR HOJA "REPORTE"
# ==========================================
print("Armando el reporte final...")

def obtener_color(puntaje, palabra_distintiva_corta=False):
    # VERDE: Solo si estamos MUY seguros (>= 95%)
    # MORADO: Revisar manualmente (50-94%)
    # ROJO: No encontrado (< 50%)
    # Si la palabra distintiva es muy corta (<4 chars), forzar MORADO
    # EXCEPTO si es match perfecto (100%)
    if puntaje >= 100:
        return 'VERDE'  # Match perfecto, siempre confiable
    elif puntaje >= 95:
        if palabra_distintiva_corta:
            return 'MORADO'  # Palabra muy corta, requiere revision
        return 'VERDE'
    elif puntaje >= 50:
        return 'MORADO'
    else:
        return 'ROJO'

df_input['SEMAFORO'] = df_input.apply(
    lambda row: obtener_color(row['PORCENTAJE'], row['DISTINTIVA_CORTA']), axis=1
)

# Creamos el DataFrame final con las columnas del reporte
df_final = pd.DataFrame()
df_final['Pais'] = df_input['Pais']
df_final['Nombre de la empresa'] = df_input['Nombre de la empresa']

# IDC: Si es VERDE, poner el CODUNICOCLI de la BD, si no, dejar el original
df_final['IDC'] = df_input.apply(
    lambda row: row['CODUNICOCLI_BD'] if row['SEMAFORO'] == 'VERDE' else row.get('IDC', ''), axis=1
)

df_final['Nemonico'] = df_input.get('Nemonico', '')

# Se ha prestado carta fianza: SI cuando es VERDE, NO cuando es ROJO, vacio en MORADO
df_final['Se ha prestado servicio de carta fianza?'] = df_input['SEMAFORO'].apply(
    lambda x: 'SI' if x == 'VERDE' else ('NO' if x == 'ROJO' else '')
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