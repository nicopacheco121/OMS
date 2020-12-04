import ws_connection
import ws_dolar_utils
import xlwings as xw
import keys

""" 
ANTES DE USAR SE DEBERA HACER LO SIGUIENTE:
------------------------------------------
- crear un archivo python llamado 'keys.py' colocando: 
    OMS_URL 
    API_KEY_ID 
    API_KEY_SECRET 
    OMS_USER 
    OMS_PASSWORD
- colocar en la misma carpeta donde corre el 1 excel con las especies, la primer fila con los nombres de la 
nominacion (ars, mep, ccl) y debajo los tickers en ese mismo orden: ARS / MEP / CCL
    
- colocar en la misma carpeta donde corre el archivo un excel llamado 'streaming_excel_dolar.xlsx' con una hoja 
llamada 'dolar'
"""

# Se instancia la clase diccionario que leerá los precios y actualizará un diccionario
precios = ws_dolar_utils.Diccionario()

# Se toman las especies solicitadas. Debe haber un archivo llamado especies.xslx en la carpeta
tupla_tickers = ws_dolar_utils.tickers()

# Comienza la conexion del web socket
handler = ws_connection.ConnectionHandler(
                            OMS_URL=keys.OMS_URL,API_KEY_ID=keys.API_KEY_ID,API_KEY_SECRET=keys.API_KEY_SECRET,
                            OMS_USER=keys.OMS_USER,OMS_PASSWORD=keys.OMS_PASSWORD,
                            tickers = tupla_tickers[1])

# Excel
excel = ws_dolar_utils.Excel(hoja=xw.Book('streaming_excel_dolar.xlsx').sheets('dolar'),
                             precios_ci=precios.precios_ci, precios_48=precios.precios_48,
                             tickers_en_lista=tupla_tickers[0], tickers_pesos=tupla_tickers[2])