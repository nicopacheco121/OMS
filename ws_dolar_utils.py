import pandas as pd
import ws_queue
import threading
import logging
import xlwings as xw
import time

logging.basicConfig(level=logging.INFO, format='{asctime} {levelname} ({threadName:11s}) {message}', style='{')

def tickers(archivo='especies.xlsx'):
    """ Lee el archivo de especies y entrega una tupla con lo siguiente:
        - Lista de listas agrupando las especies de distinta demoninacion
        - Lista con todas la especies sin agrupar
        - Lista especies solo en pesos """

    datos = pd.read_excel(archivo)
    resultado = datos.values.tolist()

    # Lista agrupada
    tickers_en_lista = resultado.copy()

    # Lista desagrupada
    funcion = lambda x: [item for sublist in x for item in sublist]
    tickers = funcion(resultado)

    # Listado de tickers en pesos
    tickers_pesos = datos.iloc[:, 0].tolist()

    return tickers_en_lista,tickers,tickers_pesos


class Diccionario():

    """ Esta clase va tomando los precios actualizados de la cola y los guarda en un diccionario """

    def __init__(self):
        self.precios_ci = {}
        self.precios_48 = {}
        self.queueManager = ws_queue.QueueManager.getInstance()

        mdThread = threading.Thread(name="ProcessMD", target=self.processMD)
        mdThread.start()

    def processMD(self):
        while True:
            message = self.queueManager.readMarketData()
            #logging.info(f"marketdata {message}")

            ticker = message["instrumentId"]["symbol"]
            self.guardaPrecios(message=message,ticker=ticker)

    def guardaPrecios(self,message,ticker):
        bid = message["marketData"]["BI"][0]["price"]
        off = message["marketData"]["OF"][0]["price"]

        try:
            type = message['instrumentId']['settlementType']
            if type == '1':
                self.precios_ci[ticker] = {'bid' : bid , 'off' : off}
            if type == '3':
                self.precios_48[ticker] = {'bid': bid, 'off': off}

        except:
            pass

    def ver_datos(self):
        return self.precios_ci,self.precios_48


class Excel:

    """ Recibe un archivo con una hoja llamada dolar, escribe algunos parametros para entender el archivo,
    calcula los precios ccl y mep y escribe en el excel en tiempo real """

    def __init__(self, hoja, precios_ci, precios_48, tickers_en_lista,tickers_pesos):
        self.hoja = hoja
        self.tickers_en_lista = tickers_en_lista
        self.tickers_pesos = tickers_pesos
        self.dolares = {'MEP': {'CI': {}, '48': {}}, 'CCL': {'CI': {}, '48': {}}}
        self.filas = {}
        self.precios_ci = precios_ci
        self.precios_48 = precios_48
        self.queueManager = ws_queue.QueueManager.getInstance()



        self.columna = {'M_CI_C': 'C', 'M_CI_V': 'D', 'M_48_C': 'E', 'M_48_V': 'F',
                        'C_CI_C': 'G', 'C_CI_V': 'H', 'C_48_C': 'I', 'C_48_V': 'J', }

        # Doy formato a la hoja excel
        hoja = xw.Book('streaming_excel_dolar.xlsx').sheets('dolar')

        # Colores
        hoja.range('C3:F3').color = xw.utils.rgb_to_int((208, 206, 206))  # MEP
        hoja.range('G3:J3').color = xw.utils.rgb_to_int((174, 170, 170))  # CCL
        hoja.range(('C4:D4'), ('G4:H4')).color = xw.utils.rgb_to_int((214, 220, 228))  # CI
        hoja.range('E4:F4').color = xw.utils.rgb_to_int((172, 185, 202))  # 48
        hoja.range('I4:J4').color = xw.utils.rgb_to_int((172, 185, 202))  # 48
        hoja.range('C5:J5').color = xw.utils.rgb_to_int((255, 242, 204))  # Compra
        hoja.range('D5').color = xw.utils.rgb_to_int((255, 230, 153))  # Venta
        hoja.range('F5').color = xw.utils.rgb_to_int((255, 230, 153))  # Venta
        hoja.range('H5').color = xw.utils.rgb_to_int((255, 230, 153))  # Venta
        hoja.range('I5').color = xw.utils.rgb_to_int((255, 230, 153))  # Venta
        # Palabras
        hoja.range('D3').value = 'MEP'
        hoja.range('H3').value = 'CCL'
        hoja.range('C4').value = 'CI'
        hoja.range('E4').value = '48'
        hoja.range('G4').value = 'CI'
        hoja.range('I4').value = '48'
        hoja.range('C5:I5').value = 'Compra'
        hoja.range('D5').value = 'Venta'
        hoja.range('F5').value = 'Venta'
        hoja.range('H5').value = 'Venta'
        hoja.range('J5').value = 'Venta'

        # Especies
        for t in range(len(self.tickers_pesos)):

            especie = self.tickers_pesos[t]

            # Diccionario de filas y escritura en excel
            self.filas[especie] = t + 6
            fila = str(self.filas[especie])
            string = 'B' + fila
            hoja.range(string).value = especie

            # Diccionario con clave especie principal y valores con mep y ccl
            self.dolares['MEP']['CI'][especie] = ["", ""]
            self.dolares['MEP']['48'][especie] = ["", ""]
            self.dolares['CCL']['CI'][especie] = ["", ""]
            self.dolares['CCL']['48'][especie] = ["", ""]

        # Corro hilo para procesar informacion
        mdThread = threading.Thread(name="ProcessMDExcel", target=self.processMD)
        mdThread.start()

    def processMD(self):
        while True:

            for i in range(len(self.tickers_en_lista)):

                ticker_principal = self.tickers_en_lista[i][0]
                ticker_mep = self.tickers_en_lista[i][1]
                ticker_ccl = self.tickers_en_lista[i][2]

                # MEP - CI
                try:
                    precio_compra = round(self.precios_ci[ticker_principal]['bid'] / self.precios_ci[ticker_mep]['off'],
                                          2)
                    precio_venta = round(self.precios_ci[ticker_principal]['off'] / self.precios_ci[ticker_mep]['bid'],
                                         2)

                    if precio_compra != self.dolares['MEP']['CI'][ticker_principal][0]:
                        self.changeColor(ticker=ticker_principal)
                        self.dolares['MEP']['CI'][ticker_principal][0] = precio_compra
                        self.escribeExcel(precio=precio_compra, ticker=ticker_principal, columna=self.columna['M_CI_C'])

                    if precio_venta != self.dolares['MEP']['CI'][ticker_principal][1]:
                        self.changeColor(ticker=ticker_principal)
                        self.dolares['MEP']['CI'][ticker_principal][1] = precio_venta
                        self.escribeExcel(precio=precio_venta, ticker=ticker_principal, columna=self.columna['M_CI_V'])
                except:
                    pass

                # MEP - 48
                try:
                    precio_compra = round(self.precios_48[ticker_principal]['bid'] / self.precios_48[ticker_mep]['off'],
                                          2)
                    precio_venta = round(self.precios_48[ticker_principal]['off'] / self.precios_48[ticker_mep]['bid'],
                                         2)

                    if precio_compra != self.dolares['MEP']['48'][ticker_principal][0]:
                        self.changeColor(ticker=ticker_principal)
                        self.dolares['MEP']['48'][ticker_principal][0] = precio_compra
                        self.escribeExcel(precio=precio_compra, ticker=ticker_principal, columna=self.columna['M_48_C'])

                    if precio_venta != self.dolares['MEP']['48'][ticker_principal][1]:
                        self.changeColor(ticker=ticker_principal)
                        self.dolares['MEP']['48'][ticker_principal][1] = precio_venta
                        self.escribeExcel(precio=precio_venta, ticker=ticker_principal, columna=self.columna['M_48_V'])


                except:
                    pass

                # CCL - CI
                try:
                    precio_compra = round(self.precios_ci[ticker_principal]['bid'] / self.precios_ci[ticker_ccl]['off'],
                                          2)
                    precio_venta = round(self.precios_ci[ticker_principal]['off'] / self.precios_ci[ticker_ccl]['bid'],
                                         2)

                    if precio_compra != self.dolares['CCL']['CI'][ticker_principal][0]:
                        self.changeColor(ticker=ticker_principal)
                        self.dolares['CCL']['CI'][ticker_principal][0] = precio_compra
                        self.escribeExcel(precio=precio_compra, ticker=ticker_principal, columna=self.columna['C_CI_C'])

                    if precio_venta != self.dolares['CCL']['CI'][ticker_principal][1]:
                        self.changeColor(ticker=ticker_principal)
                        self.dolares['CCL']['CI'][ticker_principal][1] = precio_venta
                        self.escribeExcel(precio=precio_venta, ticker=ticker_principal, columna=self.columna['C_CI_V'])
                except:
                    pass

                # CCL - 48
                try:
                    precio_compra = round(self.precios_48[ticker_principal]['bid'] / self.precios_48[ticker_ccl]['off'],
                                          2)
                    precio_venta = round(self.precios_48[ticker_principal]['off'] / self.precios_48[ticker_ccl]['bid'],
                                         2)

                    if precio_compra != self.dolares['CCL']['48'][ticker_principal][0]:
                        self.changeColor(ticker=ticker_principal)
                        self.dolares['CCL']['48'][ticker_principal][0] = precio_compra
                        self.escribeExcel(precio=precio_compra, ticker=ticker_principal, columna=self.columna['C_48_C'])

                    if precio_venta != self.dolares['CCL']['48'][ticker_principal][1]:
                        self.changeColor(ticker=ticker_principal)
                        self.dolares['CCL']['48'][ticker_principal][1] = precio_venta
                        self.escribeExcel(precio=precio_venta, ticker=ticker_principal, columna=self.columna['C_48_V'])

                except:
                    pass

    def changeColor(self, ticker):

        num = self.filas[ticker]
        string = 'B' + str(num) + ':J' + str(num)

        hoja = xw.Book('streaming_excel_dolar.xlsx').sheets('dolar')
        hoja.range(string).color = xw.utils.rgb_to_int((220, 220, 220))
        time.sleep(0.005)
        hoja.range(string).color = xw.utils.rgb_to_int((255, 255, 255))

    def escribeExcel(self, precio, ticker, columna):
        hoja = xw.Book('streaming_excel_dolar.xlsx').sheets('dolar')

        string = columna + str(self.filas[ticker])
        hoja.range(string).value = precio

    def obtener_ubicacion(self):
        pass



