from glob import glob
from tqdm import tqdm
import pandas as pd
import os
import sys

db_path = "./base de datos.xlsx"
empty_db_path = "./base de datos PENDIENTES.xlsx"
archivos_dir = "./Archivos/"
filled_dir = "./Rellenados/"

saltar = 4 # salta las primeras x
index_n_columns = 2 # first n index
data_n_columns = 8 # next k columns for data to take into account

def get_pendientes(file_df, pendiente_df = None, update=False):
    # get filas pendientes por rellenar
    tit = str(titulos_de_interes[index_n_columns])
    file_pendientes_df = file_df.loc[file_df[tit] == "?"]
    file_pendientes_df2 = file_df[file_df[tit].isnull()]
    list_df = [file_pendientes_df, file_pendientes_df2]
    if update:
        list_df = [pendiente_df,] + list_df
    pendiente_df = pd.concat(list_df).drop_duplicates()
    
    return pendiente_df

def rellenar_excels(titulos_de_interes, titulos, db_path):
    if not os.path.exists(db_path):
        print("Error: Falta la base de datos!")
        return 1
    if not os.path.exists(filled_dir):
        os.makedirs(filled_dir)
        
    db_df = pd.read_excel(db_path)
    idx_tit = titulos_de_interes[:index_n_columns]
    short_db = db_df[idx_tit]
    for file_path in tqdm(archivos):
        file_df = pd.read_excel(file_path)
        pendiente_df = get_pendientes(file_df)

        for index, row in pendiente_df.iterrows():
            lista_busqueda = short_db.isin(list(row[idx_tit])).all(1)
            if lista_busqueda.any():
                for i, b in enumerate(lista_busqueda):
                    if b:
                        fila_db = db_df.iloc[i]
                        break
                for title in titulos_de_interes[index_n_columns:]:
                    file_df.loc[index, title] = fila_db[[title]].values
        
        filename = os.path.basename(file_path)
        filename, extension = os.path.splitext(filename)
        file_df = file_df[titulos]        
        file_df.sort_index().to_excel(filled_dir + filename + extension.lower(), index=False, merge_cells=False, engine='openpyxl')

def clean_db(db_path, titulos_de_interes):
    #quita de pendientes aquellas lineas que hayan entrado en la base de datos más adelante
    if not os.path.exists(db_path):
        print("Error: Falta la base de datos!")
        return 1

    db_df = pd.read_excel(db_path)
    pendiente_df = pd.read_excel(empty_db_path)
    idx_tit = titulos_de_interes[:index_n_columns]
    short_db = db_df[idx_tit]

    fp_index = []
    for index, row in pendiente_df.iterrows():
        if short_db.isin(list(row[idx_tit])).all(1).any():
            fp_index.append(index)
    if fp_index:
        pendiente_df.drop(fp_index, inplace = True)

    pendiente_df.sort_index().to_excel(empty_db_path, merge_cells=False, index=False, engine='openpyxl')    

def generar_db(titulos_de_interes, titulos):

    #crear df del db y poner headers
    db_df = pd.DataFrame(columns=titulos_de_interes)
    pendiente_df = pd.DataFrame(columns=titulos_de_interes)

    for file_path in tqdm(archivos):
        file_df = pd.read_excel(file_path)

        file_df = file_df[titulos_de_interes] # solo datos de interes

        pendiente_df = get_pendientes(file_df, pendiente_df, True)
        
        tit = str(titulos_de_interes[index_n_columns])
        uncmplt_row_index = file_df[file_df[tit] == "?"].index
        uncmplt_row_index2 = file_df[file_df[tit].isnull()].index
        uncmplt_row_index = uncmplt_row_index.append(uncmplt_row_index2)
        file_df.drop(uncmplt_row_index, inplace = True)
        db_df = pd.concat([db_df, file_df]).drop_duplicates()
        db_df = db_df[~db_df.index.duplicated(keep='first')]


    db_df.sort_index().to_excel(db_path, merge_cells=False, index=False, engine='openpyxl')
    pendiente_df.sort_index().to_excel(empty_db_path, merge_cells=False, index=False, engine='openpyxl')

def get_titles():
    if not archivos:
        print("Error: No hay excels en la carpeta Archivos")
        return 1,1
        
    file_df = pd.read_excel(archivos[0])
    titulos = list(file_df.columns)# todos los titulos
    titulos_de_interes = list(titulos[ saltar : saltar + index_n_columns + data_n_columns])

    return titulos_de_interes, titulos

# main
if os.path.exists(archivos_dir):
    archivos = glob(archivos_dir+"*")
    titulos_de_interes, titulos = get_titles()
    if titulos_de_interes != 1:
        if len(sys.argv) > 1:
            print("\n Rellenando excels ...\n", flush=True)
            rellenar_excels(titulos_de_interes, titulos, db_path)
        else:
            print("\n Generando base de datos ...\n", flush=True)
            generar_db(titulos_de_interes, titulos)
            print("\n Limpiando base de datos PENDIENTE ...", flush=True)
            clean_db(db_path, titulos_de_interes)
            print("\n Rellenando excels ...\n", flush=True)
            rellenar_excels(titulos_de_interes, titulos, db_path)

        print("\n ------------- FIN -------------\n", flush=True)
else:
    os.makedirs(archivos_dir)
    print("\nCarpeta: '"+str(archivos_dir)+"' generada")
    print("Mete los excels en esta nueva carpeta y vuelve a ejecutar el programa.\n")