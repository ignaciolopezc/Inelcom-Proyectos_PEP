# Inelcom-Proyectos_PEP

## Descripción
Sistema automatizado para la cuadratura y análisis de proyectos PEP (Project Element Planning), integrando datos de múltiples fuentes: NewBacklog, transacciones SAP (CN43N), y datos conformados de Oracle.

## Requisitos

### Python
- Python 3.14 o superior
- Entorno virtual (`rpa_env`)

### Dependencias
```bash
pip install -r requirements.txt
```

**Librerías principales:**
- `pandas==2.3.3` - Manipulación y análisis de datos
- `openpyxl==3.1.5` - Lectura/escritura de archivos Excel
- `oracledb` - Conexión a base de datos Oracle
- `numpy==2.3.5` - Operaciones numéricas

## Estructura del Proyecto

```
REQ001/
├── test.py                     # Script principal de procesamiento
├── connectionODBC.py           # Clase de conexión a Oracle
├── requirements.txt            # Dependencias del proyecto
├── README.md                   # Este archivo
├── rpa_env/                    # Entorno virtual
└── Archivos de entrada:
    ├── NewBacklog_20251030.xlsx
    ├── CONFORMADO_EXP.VW_GR_PROYECTO_PEP_GIT.xlsx
    ├── BD_CN43N_INV_20251113.XLSX
    ├── BD_CN43N_VTA_20251113.XLSX
    └── CJI3_INV_2009_20015.xlsx
```

## Funcionalidad Principal

### Script: `test.py`

#### Función: `funcion_cuadratura()`

Procesa y consolida datos de proyectos desde múltiples fuentes:

**1. Procesamiento NewBacklog:**
- Carga datos de proyectos desde `NewBacklog_20251030.xlsx`
- Extrae año y mes de la columna `TERMINOJP`
- Filtra proyectos del año 2025
- Estados válidos: 'Cancelado', 'Terminado JP', 'Terminado'
- Filtra proyectos externos (`EXT == 1`)
- Elimina columnas no necesarias (INSTPEND, RTAPEND, etc.)

**2. Procesamiento Base Ventas (CN43N):**
- Carga transacciones de ventas desde `BD_CN43N_VTA_20251113.XLSX`
- Extrae código SISON (13 caracteres si empieza con 'CHI-', 9 caracteres en caso contrario)
- Extrae código de Sociedad del campo `Elemento PEP` (posiciones 2-6)
- Marca tipo como 'VTA'

**3. Procesamiento Base Inversiones (CN43N):**
- Carga transacciones de inversión desde `BD_CN43N_INV_20251113.XLSX`
- Similar procesamiento que ventas
- Marca tipo como 'INV'

**4. Consolidación CN43N:**
- Combina datos de Ventas e Inversiones en un solo DataFrame

**5. Datos Conformados:**
- Carga datos de Oracle desde `CONFORMADO_EXP.VW_GR_PROYECTO_PEP_GIT.xlsx`
- Renombra columna `NROPEP` a `Elemento PEP`

**6. Datos CJI3:**
- Carga primeras filas de `CJI3_INV_2009_20015.xlsx`

**7. Resultado:**
- Realiza merge INNER entre NewBacklog y CN43N usando columna `SISON`
- Genera archivo de salida `Resultado.xlsx` con múltiples hojas:
  - **NewBacklog**: Proyectos filtrados
  - **CN43N**: Consolidado de Ventas e Inversiones
  - **Conformado**: Datos conformados de Oracle
  - **data CJI3**: Datos de CJI3
  - **Resultado**: Cruce entre NewBacklog y CN43N

### Script: `connectionODBC.py`

Clase `ConexionODBC` para gestión de conexiones a Oracle Database.

**Parámetros de conexión:**
- **Host**: `smt-scan.tchile.local`
- **Puerto**: `1521`
- **Servicio**: `explota`
- **Usuario**: `SRV_MKTB2B`
- **Password**: `Mkt_chile_2025`

**Métodos principales:**
- `conectar()`: Establece conexión a Oracle
- `ejecutar_query(query)`: Ejecuta consultas SELECT
- `ejecutar_comando(comando)`: Ejecuta INSERT/UPDATE/DELETE
- `desconectar()`: Cierra conexión
- Soporte para context manager (`with` statement)

**Ejemplo de uso:**
```python
from connectionODBC import ConexionODBC

# Uso con context manager
with ConexionODBC() as conn:
    resultados = conn.ejecutar_query("SELECT * FROM tabla")
    print(resultados)
```

## Instalación y Configuración

### 1. Clonar repositorio
```bash
git clone https://github.com/ignaciolopezc/Inelcom-Proyectos_PEP.git
cd REQ001
```

### 2. Crear y activar entorno virtual
```powershell
py -m venv rpa_env
.\rpa_env\Scripts\Activate.ps1
```

### 3. Instalar dependencias
```powershell
pip install -r requirements.txt
```

### 4. Configurar archivos de entrada
Asegúrate de tener los siguientes archivos Excel en el directorio raíz:
- `NewBacklog_20251030.xlsx`
- `CONFORMADO_EXP.VW_GR_PROYECTO_PEP_GIT.xlsx`
- `BD_CN43N_INV_20251113.XLSX`
- `BD_CN43N_VTA_20251113.XLSX`
- `CJI3_INV_2009_20015.xlsx`

## Ejecución

```powershell
# Activar entorno virtual (si no está activo)
.\rpa_env\Scripts\Activate.ps1

# Ejecutar script principal
py test.py
```

**Salida esperada:**
```
Inicio Cargando datos de NewBacklog...
Hora de inicio: 2025-11-21 10:30:00
Hora de fin: 2025-11-21 10:30:15
Tiempo transcurrido: 0:00:15
Fin Proceso completado.
```

**Archivo generado:**
- `Resultado.xlsx` (con 5 hojas de resultados)

## Filtros y Validaciones

### NewBacklog
- **Año**: 2025
- **Estados**: Cancelado, Terminado JP, Terminado
- **Tipo**: Externos (EXT = 1)

### CN43N (Ventas e Inversiones)
- **SISON**: Extracción condicional según prefijo 'CHI-'
- **Sociedad**: Extraída de posiciones 2-6 del Elemento PEP

### Resultado
- **Join**: INNER merge sobre columna `SISON`
- **Sufijos**: `_Newbacklog` y `_CN43N` para columnas duplicadas

## Notas Técnicas

### Extracción de Fechas
```python
# Formato de TERMINOJP: DD/MM/YYYY
año = TERMINOJP[6:10]   # Posiciones 6-9
mes = TERMINOJP[3:5]    # Posiciones 3-4
```

### Extracción SISON
```python
# Si empieza con 'CHI-': primeros 13 caracteres
# Si no: primeros 9 caracteres
SISON = Denominación[0:13] if starts_with('CHI-') else Denominación[0:9]
```

### Extracción Sociedad
```python
# Del campo Elemento PEP, posiciones 2-6
Sociedad = Elemento_PEP[2:6]
```

## Mantenimiento

### Actualizar dependencias
```powershell
pip freeze > requirements.txt
```

### Actualizar archivos de entrada
Los nombres de archivos están definidos al inicio de `test.py`. Actualízalos según sea necesario:
```python
Libro_Newbacklog = 'NewBacklog_YYYYMMDD.xlsx'
Libro_BaseInversion = 'BD_CN43N_INV_YYYYMMDD.XLSX'
Libro_BaseVentas = 'BD_CN43N_VTA_YYYYMMDD.XLSX'
```

## Autor
Ignacio López
- GitHub: [@ignaciolopezc](https://github.com/ignaciolopezc)

## Licencia
Proyecto interno - Inelcom - Telefonica Chile

