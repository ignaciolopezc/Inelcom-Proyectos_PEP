import pandas as pd
import numpy as np
from datetime import datetime

# Definir nombres de archivos
Libro_Newbacklog = 'NewBacklog_20251030.xlsx'
Libro_Conformado = 'CONFORMADO_EXP.VW_GR_PROYECTO_PEP_GIT.xlsx'
Libro_BaseInversion = 'BD_CN43N_INV_20251113.XLSX'
Libro_BaseVentas = 'BD_CN43N_VTA_20251113.XLSX'
Libro_Resultado = 'Resultado.xlsx'

Libro_CJI3_INV_2009 = 'CJI3_INV_2009_20015.xlsx'

# Crear DataFrame vacío para almacenar resultados
df_newbacklog = pd.DataFrame()
df_baseVentas = pd.DataFrame()
df_baseInversion = pd.DataFrame()
df_CN34N = pd.DataFrame()
df_conformado = pd.DataFrame()
df_resultado = pd.DataFrame()

df_datosCJI3 = pd.DataFrame()

# cargar todos los libros
# filtrar fechas / terminados
# obtener los datos de proyectos (chi y pep)
# crear un cruce entre vta y newbacklog

def funcion_cuadratura():
    try:
        # Manejo de datos Newbacklog --------------------------------------------------
        df_newbacklog = pd.read_excel(Libro_Newbacklog, sheet_name=0)
        df_newbacklog = df_newbacklog.drop(columns=['INSTPEND', 'RTAPEND', 'IRPEND', 'DIVISA', 'FCV', 'NAV', 'DURACION_CONTRATO','FECESTADOGIT','CREACION','FECPROPUESTA','FECPLANIFICADA','JEFEAREA','SUPERVISOR','JP','ASIGNAJP','JPDATOS','ASIGNAJP_DATOS','TERMINOJP_DATOS','FECACTUALIZACION','PLANORIGINAL','CAUSAREPLANIFICACION','OBSREPLANIFICACION','CANTIDADREP','TIPOQUIEBRE','QUIEBRE','FECQUIEBRE','TIPOLOGIA_OPORTUNIDAD','TIPO_OPORTUNIDAD','COMPLEJIDAD','PORCENTAJEREAL','AM','JC','SGTE','GTE','MESPLAZOFACTURACION','TIEMPODETENIDO','NIVELMETODOLOGICO','DESTECNICA'], errors='ignore')
        df_newbacklog['año'] = df_newbacklog['TERMINOJP'].astype(str).str[6:6 + 4]
        df_newbacklog['mes'] = df_newbacklog['TERMINOJP'].astype(str).str[3:3 + 2]
        df_newbacklog = df_newbacklog[df_newbacklog['año'] == '2025']
        # df_newbacklog = df_newbacklog[df_newbacklog['mes'] == '10']
        df_newbacklog = df_newbacklog[df_newbacklog['ESTADOGIT'].isin(['Cancelado', 'Terminado JP', 'Terminado'])]
        df_newbacklog = df_newbacklog[df_newbacklog['EXT'] == 1]
        df_newbacklog['Archivo'] = 'NewBacklog'


        # Manejo de datos Base Ventas CN43N -------------------------------------------
        df_baseVentas = pd.read_excel(Libro_BaseVentas, sheet_name=0)
        df_baseVentas = df_baseVentas.drop(columns=['Status','Responsable'], errors='ignore')
        
        # Crear columnas SISON y validacion basadas en si empieza con 'CHI-'
        chiVentas = df_baseVentas['Denominación'].astype(str).str.startswith('CHI-')
        df_baseVentas['SISON'] = df_baseVentas['Denominación'].astype(str).str[0:13].where(chiVentas, df_baseVentas['Denominación'].astype(str).str[0:9])        
        df_baseVentas['Sociedad'] = df_baseVentas['Elemento PEP'].astype(str).str[2:6]
        df_baseVentas['Tipo'] = 'VTA'
        df_baseVentas['Archivo'] = 'Ventas'


        # Manejo de datos Base Inversiones CN43N  -------------------------------------
        df_baseInversion = pd.read_excel(Libro_BaseInversion, sheet_name=0)
        df_baseInversion = df_baseInversion.drop(columns=['Status','Responsable'], errors='ignore')
        chiInversion = df_baseInversion['Denominación'].astype(str).str.startswith('CHI-')
        df_baseInversion['SISON'] = df_baseInversion['Denominación'].astype(str).str[0:13].where(chiInversion, df_baseInversion['Denominación'].astype(str).str[0:9])   
        df_baseInversion['Sociedad'] = df_baseInversion['Elemento PEP'].astype(str).str[2:2 +4]
        df_baseInversion['Tipo'] = 'INV'
        df_baseInversion['Archivo'] = 'Inversion'


        # # Consolidado de Ventas e Inversiones CN43N ----------------------------------
        df_CN34N = pd.concat([df_baseVentas, df_baseInversion], ignore_index=True)


        # Manejo de datos Conformado -------------------------------------------------
        df_conformado = pd.read_excel(Libro_Conformado, sheet_name=0)
        df_conformado = df_conformado.rename(columns={'NROPEP': 'Elemento PEP'})
        df_conformado['Archivo'] = 'Conformado'

        # Reporte salida -------------------------------------------------------------
        df_resultado = df_newbacklog.merge(df_CN34N, how='inner', on=['SISON'], suffixes=('_Newbacklog', '_CN43N'))


        # Carga de datos CJI3 --------------------------------------------------------
        df_datosCJI3 = pd.read_excel(Libro_CJI3_INV_2009, sheet_name=0)
        df_datosCJI3 = df_datosCJI3.head()
        df_datosCJI3['Archivo'] = 'CJI3'


        


        # print(df_newbacklog.head())
        with pd.ExcelWriter(Libro_Resultado, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_newbacklog.to_excel(writer, sheet_name='NewBacklog', index=False)
            df_CN34N.to_excel(writer, sheet_name='CN43N', index=False)
            df_conformado.to_excel(writer, sheet_name='Conformado', index=False)
            df_datosCJI3.to_excel(writer, sheet_name='data CJI3', index=False)
            df_resultado.to_excel(writer, sheet_name='Resultado', index=False)







    except FileNotFoundError:
        print(f"Error: No se encontró el archivo.")
        return None
    except Exception as e:
        print(f"Error al cargar el archivo: {str(e)}")
        return None
    
    pass



# carga del newbacklog_libro
    # filtros
        # agregar columna año (derivado de columna terminadosJP)
        # agregar columna mes (derivado de columna terminadosJP)
        # fecha de termino (2025-10)
        # proyectos terminados / termindados jp / cancelados
        # CHI terminados en -1
        # revisar cuales son los PEP que no tienen Segmento
        # revisar las clases de documento que deben ser WE - WL


# carga archivos CN43N VTA / INV
    # Filtros:
        # agregar columna año (derivado de columna terminadosJP)
        # agregar columna mes (derivado de columna terminadosJP)
        # fecha de termino (2025-10)








if __name__ == "__main__":
    print("Inicio Cargando datos de NewBacklog...")
    hora_inicio = datetime.now()
    print(f"Hora de inicio: {hora_inicio.strftime('%Y-%m-%d %H:%M:%S')}")
    
    funcion_cuadratura()
    
    hora_fin = datetime.now()
    tiempo_transcurrido = hora_fin - hora_inicio
    print(f"Hora de fin: {hora_fin.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Tiempo transcurrido: {tiempo_transcurrido}")

    print("Fin Proceso completado.")

