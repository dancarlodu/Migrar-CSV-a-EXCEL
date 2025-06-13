import pandas as pd
import glob
import os

def crear_directorios_necesarios():
    """Crea las carpetas input y output si no existen"""
    os.makedirs('input', exist_ok=True)
    os.makedirs('output', exist_ok=True)

def convertir_csv_a_excel():
    """Convierte todos los archivos CSV en input/ a XLSX en output/"""
    crear_directorios_necesarios()
    
    # Obtener lista de archivos CSV en la carpeta input
    csv_files = glob.glob('input/*.csv')
    
    if not csv_files:
        print("No se encontraron archivos CSV en la carpeta 'input/'")
        return
    
    for csv_file in csv_files:
        try:
            # Leer CSV
            # df = pd.read_csv(csv_file)          
            df = pd.read_csv(csv_file, encoding='latin-1')  # o encoding='cp1252'

            
            # Obtener el nombre base del archivo (sin extensión ni carpeta)
            base_name = os.path.basename(csv_file)
            file_name_without_ext = os.path.splitext(base_name)[0]
            
            # Crear nombre para el archivo Excel con prefijo
            excel_file = f"output/datos_comerciales_{file_name_without_ext}.xlsx"
            
            # Crear un escritor de Excel para poder renombrar la hoja
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Hoja1', index=False)
                
            print(f"Convertido: {csv_file} -> {excel_file}")
            
        except Exception as e:
            print(f"Error procesando {csv_file}: {str(e)}")
    
    print(f"\nProceso completado. Se convirtieron {len(csv_files)} archivos CSV a XLSX")

if __name__ == "__main__":
    convertir_csv_a_excel()