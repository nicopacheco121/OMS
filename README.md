# OMS
Código para conectarse a la api OMS de BYMA y extraer valores de dolares MEP y CCL.

ANTES DE USAR SE DEBERA HACER LO SIGUIENTE:
------------------------------------------
- crear un archivo python llamado 'keys.py' colocando: 
    OMS_URL 
    API_KEY_ID 
    API_KEY_SECRET 
    OMS_USER 
    OMS_PASSWORD
- colocar en la misma carpeta donde corre el excel con las especies, la primer fila con los nombres de la 
nominacion (ars, mep, ccl) y debajo los tickers en ese mismo orden: ARS / MEP / CCL
    
- colocar en la misma carpeta donde corre el archivo un excel llamado 'streaming_excel_dolar.xlsx' con una hoja 
llamada 'dolar'

