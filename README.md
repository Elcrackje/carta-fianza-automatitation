problmmatica
Contexto: Actualmente, el proceso de validación para determinar si se ha prestado el servicio de "Carta Fianza" a clientes corporativos implica el cruce de información entre dos fuentes: un reporte externo (Excel de solicitud) y la Base de Datos interna (Query BD).

# Carta Fianza - Procesador de Excel

Script para comparar nombres de empresas entre hojas de Excel usando fuzzy matching.

## 📋 Requisitos

- Python 3.11 o superior

## 🔧 Instalación de dependencias

Ejecuta el siguiente comando en la terminal:

```bash
pip install pandas thefuzz openpyxl jinja2
```

### Detalle de cada librería:

| Librería | Descripción |
|----------|-------------|
| `pandas` | Manipulación y análisis de datos en DataFrames |
| `thefuzz` | Fuzzy matching para comparar strings similares |
| `openpyxl` | Lectura y escritura de archivos Excel (.xlsx) |
| `jinja2` | Necesario para aplicar estilos/colores en Excel |

## 📁 Estructura de archivos

```
carta-fianza/
├── carta-fianza.py                                    # Script principal
├── Cuestionario_ServBCP (Carta Fianza) - Noviembre.xlsx  # Archivo de entrada
├── Reporte_Final_Procesado.xlsx                       # Archivo de salida (generado)
└── README.md                                          # Este archivo
```

## 🚀 Uso

1. Asegúrate de que el archivo Excel de entrada esté en la misma carpeta
2. Ejecuta el script:

```bash
python carta-fianza.py
```

3. Se generará `Reporte_Final_Procesado.xlsx` con los resultados

## 📊 Hojas del Excel de entrada

El archivo Excel debe tener las siguientes hojas:

- **Credicorp**: Datos de entrada con las empresas a buscar
- **BD**: Base de datos de clientes para comparar

## 🚦 Semáforo de resultados

| Color | Porcentaje | Significado |
|-------|------------|-------------|
| 🟢 Verde | ≥ 85% | Alta coincidencia |
| 🟣 Morado | 50% - 84% | Coincidencia media (revisar) |
| 🔴 Rojo | < 50% | Baja coincidencia |
