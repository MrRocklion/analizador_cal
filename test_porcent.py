import pandas as pd
import math
from openpyxl import load_workbook
import time
from PySide6.QtWidgets import QApplication, QMessageBox
import math
def calcularPorcent(data,porcent,final_path,condition,cal,estratificado):
    df = pd.DataFrame(data)
    longitud = len(df)
    muestra_generada = []
    aux_porcent = porcent 
    if condition == 2:
        N_poblation = longitud
        error = 0.05 
        confianza = 1.96 
        total_porcent = (N_poblation *(confianza**2)*0.5*0.5)/((error**2)*(N_poblation-1)+(confianza**2)*0.5*0.5)/longitud
        if estratificado:
            muestra_generada = generaMuestreoEstratificado(longitud=longitud,aux_porcent=total_porcent,df=df).to_dict(orient='records')
        else:
            df_ordenado = df.sample(frac=total_porcent).sort_values(by=['year', 'mes'], ascending=[True, True])
            muestra_generada = df_ordenado.to_dict(orient='records')
    elif condition == 1:
        df_ordenado = df.sort_values(by=['year', 'mes'], ascending=[True, True])
        muestra_generada = df_ordenado.to_dict(orient='records')
    else :
        if estratificado:
            muestra_generada = generaMuestreoEstratificado(longitud=longitud,aux_porcent=aux_porcent,df=df).to_dict(orient='records')
        else:
            df_ordenado = df.sample(frac=aux_porcent).sort_values(by=['year', 'mes'], ascending=[True, True])
            muestra_generada = df_ordenado.to_dict(orient='records')
   


def generaMuestreoEstratificado(longitud,aux_porcent,df):
    aux_estratificado =  pd.DataFrame()
    years = df['year'].unique()
    for y in years:
        for i in range(1,13):
            try:
                n_samples =math.floor((longitud * aux_porcent)/24)
                print(f'numeros de muestras {n_samples}')
                filtro = df.query(f'year == {y} and mes == {i}')
                aux_data = filtro.sample(n=n_samples)
                aux_estratificado = pd.concat([aux_estratificado, aux_data])
            except ValueError:
                Alerta(f'insuficientes datos en el año {y} y mes {i}')
                time.sleep(2)
                break
    return(aux_estratificado)
         

def Alerta(mensaje):
    app = QApplication([])
    alerta = QMessageBox()
    alerta.setIcon(QMessageBox.Warning)
    alerta.setText(mensaje)
    alerta.setWindowTitle("Error")
    alerta.setStandardButtons(QMessageBox.Ok)
    alerta.exec()
    app.quit()

data = pd.read_csv('030.csv')
calcularPorcent(data=data,porcent=0.2,final_path='test.xlsx',condition=2,cal=20,estratificado=False)