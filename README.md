# Carta Fianza - Procesador de Excel

Script para comparar nombres de empresas entre hojas de Excel usando fuzzy matching.

## Requisitos

- Python 3.11 o superior

## Instalacion de dependencias

Ejecuta el siguiente comando en la terminal:

```bash
pip install pandas thefuzz openpyxl jinja2
```

### Detalle de cada libreria:

| Libreria | Descripcion |
|----------|-------------|
| `pandas` | Manipulacion y analisis de datos en DataFrames |
| `thefuzz` | Fuzzy matching para comparar strings similares |
| `openpyxl` | Lectura y escritura de archivos Excel (.xlsx) |
| `jinja2` | Necesario para aplicar estilos/colores en Excel |

## Estructura de archivos

```
carta-fianza/
├── carta-fianza.py                                    # Script principal
├── Cuestionario_ServBCP (Carta Fianza) - Noviembre.xlsx  # Archivo de entrada
├── Reporte_Final_Procesado.xlsx                       # Archivo de salida (generado)
└── README.md                                          # Este archivo
```

## Uso

1. Asegurate de que el archivo Excel de entrada este en la misma carpeta
2. Ejecuta el script:

```bash
python carta-fianza.py
```

3. Se generara `Reporte_Final_Procesado.xlsx` con los resultados

## Hojas del Excel de entrada

El archivo Excel debe tener las siguientes hojas:

- **Credicorp**: Datos de entrada con las empresas a buscar
- **BD**: Base de datos de clientes para comparar

## Semaforo de resultados

| Color | Porcentaje | Significado |
|-------|------------|-------------|
| Verde | >= 85% | Alta coincidencia |
| Morado | 50% - 84% | Coincidencia media (revisar) |
| Rojo | < 50% | Baja coincidencia |
