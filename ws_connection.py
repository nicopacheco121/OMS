import sys
sys.path.insert(0, "./lib")
import requests
from requests.auth import HTTPBasicAuth
import ssl
import json
import keys
import logging
from web_socket_stomp_app import WebSocketStompApp
from web_socket_stomp_app import MessageHandler
from ws_queue import QueueManager

logging.basicConfig(level=logging.INFO, format='{asctime} {levelname} ({threadName:11s}) {message}', style='{')

""" Codigo para conectarse al websocket, filtra los mensajes segun las especies requeridas y los envia a una cola """

class BYMAMarketDataMessageHandler(MessageHandler):

    def __init__(self,tickers):
        self.tickers = tickers

    def url(self):
        return "/intercon/stomp/bc/md/byma"

    def callback(self, message):

        message = json.loads(message)
        ticker_filtro = message["instrumentId"]["symbol"]

        if ticker_filtro in self.tickers[1]:
            #print(ticker_filtro)
            QueueManager.getInstance().sendMarketData(message)


class ConnectionHandler:

    def __init__(self,OMS_URL,API_KEY_ID,API_KEY_SECRET,OMS_USER,OMS_PASSWORD,tickers):

        logging.info(f'Iniciando conexion con usuario {OMS_USER}')

        #TOKEN
        r = requests.post('https://'+OMS_URL+'/generic-oauth-core/oauth/token', verify=False,
                  auth=HTTPBasicAuth(API_KEY_ID, API_KEY_SECRET),
                  data = {'grant_type':'password','username':OMS_USER,'password':OMS_PASSWORD})

        token = r.json()['access_token']

        ws = WebSocketStompApp("wss://"+keys.OMS_URL+"/vanoms-be-core/rest/api/intercon/stomp/websocket",keys.OMS_USER,
                       header={"Authorization": "Bearer " + token})
        ws.register_handler(BYMAMarketDataMessageHandler(tickers=tickers))
        ws.run_forever( sslopt={"cert_reqs": ssl.CERT_NONE, "check_hostname": True, "ssl_version": ssl.PROTOCOL_TLSv1})

        logging.info(f'Iniciando correctamente la conexion')