from bardapi import Bard
import os, time
import pandas as pd
# Ruta del archivo Excel
ruta_archivo = './BASE DE DATOS.xlsx'
os.environ['_BARD_API_KEY']='Zgge7WM_cTpbv3D2adQn4Yv6a-ofI2mZQdoo4bXxf2vpzrg_a0aTGbNWhnx9pBg3EV26Rw.'
data_frame = pd.read_excel(ruta_archivo, sheet_name='estudiantes')
datos_columna = data_frame.iloc[:, 1].tolist()

for nombre in datos_columna:
    
	bard = Bard(timeout=100)
	MATEA = nombre
	response = bard.get_answer(f"{MATEA} Dirias que es nombre de hombre o de mujer")['content']
	# respuesta = response.rstrip('.').rsplit(" ", 1)[-1]
	print(response)
	time.sleep(1)
 
 ## Problema: luego de varios intentos, no se pueden obtener las respuestas esperadas, hay veces que no devuelve la respuesta indicada, por lo que se uso respuesta, ya que daba la palabra hombre o mujer en la ultima palabra, para eso el split, mas luego de hacer varios intentos, la api devuelve el error Exception: Response code not 200. Response Status is 429, lo que esta indicando que no se pueden hacer mas solicitudes lo que en el navegador aparece: Nuestros sistemas han detectado tráfico inusual procedente de tu red de ordenadores. En esta página se comprueba si eres tú quien envía las solicitudes en lugar de un robot. y toca resolver un captcha. de lo contrario seria completamente funciona, y el paso siguiente seria escribir cada respuesta en el excel, en la columna indicada