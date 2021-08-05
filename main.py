from glob import glob
from tqdm import tqdm
import pandas as pd
import os
import sys

db_path = "./base de datos.xlsx"
empty_db_path = "./base de datos PENDIENTES.xlsx"
archivos_dir = "./Archivos/"
filled_dir = "./Rellenados/"

saltar = 13
index_n_columns = 2
data_n_columns = 6

def rellenar_excels(titulos_de_interes, titutlos, db_path):
    if not os.path.exists(db_path):
        print("Error: Falta la base de datos!")
        return
    if not os.path.exists(filled_dir):
        os.makedirs(filled_dir)
        
    db_df = pd.read_excel(db_path)
    db_df.set_index(titulos_de_interes[:index_n_columns], inplace=True)
    for file_path in tqdm(archivos):
        file_df = pd.read_excel(file_path)
        file_df.set_index(titulos_de_interes[:index_n_columns], inplace=True)
        
        pendientes_df = file_df[file_df.isnull().any(axis=1)]
        for index, row in pendientes_df.iterrows():
            if index in db_df.index:
                file_df.at[index] = db_df.loc[index]
        
        filename = file_path.split("/")[-1].split('\\')[-1]
        file_df.reset_index(inplace=True)
        file_df = file_df[titulos]        
        file_df.to_excel(filled_dir + filename, index=False)

def generar_db(titulos_de_interes, titulos):

    #crear df del db y poner headers
    db_df = pd.DataFrame(columns=titulos_de_interes)
    db_df.set_index(titulos_de_interes[:index_n_columns], inplace=True)

    pendiente_df = pd.DataFrame(columns=titulos_de_interes)
    pendiente_df.set_index(titulos_de_interes[:index_n_columns], inplace=True)

    for file_path in tqdm(archivos):
        file_df = pd.read_excel(file_path)

        file_df = file_df[titulos_de_interes] # solo datos de interes
        file_df.set_index(titulos_de_interes[:index_n_columns], inplace=True)

        file_pendientes_df = file_df[file_df.isnull().any(axis=1)]
        pendiente_df = pd.concat([pendiente_df,file_pendientes_df]).drop_duplicates()
        
        file_df.dropna(inplace=True)
        db_df = pd.concat([db_df, file_df]).drop_duplicates()

    db_df.to_excel(db_path)
    pendiente_df.to_excel(empty_db_path)
    return titulos_de_interes, titulos

def get_titles():
    if not archivos:
        print("Error: No hay excels en la carpeta Archivos")
        return 
        
    file_df = pd.read_excel(archivos[0])
    titulos = list(file_df.columns)
    titulos_de_interes = list(titulos[ saltar : saltar + index_n_columns + data_n_columns])

    return titulos_de_interes, titulos

# main
if os.path.exists(archivos_dir):
    archivos = glob(archivos_dir+"*")
    titulos_de_interes, titulos = get_titles()
    
    if len(sys.argv) > 1:
        #if sys.argv[1] == "2":
        print("\n Rellenando excels ...\n", flush=True)
        rellenar_excels(titulos_de_interes, titulos, db_path)
        #else:
            #print("Error: no existe ese comando")
    else:
        print("\n Generando base de datos ...\n", flush=True)
        generar_db(titulos_de_interes, titulos)
        print("\n Rellenando excels ...\n", flush=True)
        rellenar_excels(titulos_de_interes, titulos, db_path)

    print("\n ------------- FIN -------------\n", flush=True)
else:
    os.makedirs(archivos_dir)
    print("\nCarpeta: '"+str(archivos_dir)+"' generada")
    print("Mete los excels en esta nueva carpeta y vuelve a ejecutar el programa.\n")