from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
import json


# Obtengo las credenciales de afip
with open('credenciales.json', 'r') as file:
    credentials = json.load(file)

cuit = credentials['cuit']
clave = credentials['clave']
cuitReceptor = credentials['cuit_receptor']



def automateInvoiceCreation(daysToAdd, data):


	# CHROME
	# ----------------------------------------------------------------------------------------------

	# Ruta del ChromeDriver
	chrome_driver_path = "C:/Users/ag/Desktop/chromedriver.exe"

	# Configuro opciones
	options = webdriver.ChromeOptions()
	options.add_argument('--start-maximized') # Inicia el navegador maximizado

	# Creo servicio con el driver
	service = Service(chrome_driver_path)

	# Creo instancia del navegador
	browser = webdriver.Chrome(service=service, options=options)

	# Abro el sitio de afip
	browser.get('https://www.afip.gob.ar/landing/default.asp')

	# Hago click en el botón "Iniciar sesión" de la landing
	btnIniciarSesion = browser.find_element(By.LINK_TEXT, 'Iniciar sesión')
	btnIniciarSesion.click()

	# Espero unos segundos para que se abra la nueva pestaña
	time.sleep(2)

	# Trato nueva pestaña
	tabs = browser.window_handles # Obtengo una lista de todas las pestañas abiertas
	browser.switch_to.window(tabs[-1]) # Cambio el foco a la nueva pestaña (la última)

	# Relleno CUIT y presiono "Siguiente"
	inputCuit = browser.find_element(By.ID, 'F1:username')
	btnSiguiente = browser.find_element(By.ID, 'F1:btnSiguiente')
	inputCuit.send_keys(cuit)
	btnSiguiente.click()

	# Relleno clave y presiono "Ingresar"
	inputClave = browser.find_element(By.ID, 'F1:password')
	btnIngresar = browser.find_element(By.ID, 'F1:btnIngresar')
	inputClave.send_keys(clave)
	btnIngresar.click()

	# Hago click en "Comprobantes en línea"
	btnComprobantes = WebDriverWait(browser, 10).until(
	    EC.element_to_be_clickable((By.XPATH, "//a//*[text()='Comprobantes en línea']"))
	)																			  # Espera hasta que el botón sea clickeable
	browser.execute_script("arguments[0].scrollIntoView(true);", btnComprobantes) # Desplazar la ventana para hacer visible el elemento
	btnComprobantes.click()

	time.sleep(1)

	# Cambio a la nueva pestaña
	tabs = browser.window_handles # Obtengo una lista de todas las pestañas abiertas
	browser.switch_to.window(tabs[-1]) # Cambio el foco a la nueva pestaña (la última)

	# Hago click en el botón de la empresa
	btnEmpresa = WebDriverWait(browser, 10).until(
	    EC.element_to_be_clickable((By.CLASS_NAME, 'btn_empresa'))
	)
	btnEmpresa.click()

	# Hago click en "Generar Comprobantes"
	btnGenerarComprobantes = WebDriverWait(browser, 10).until(
	    EC.element_to_be_clickable((By.LINK_TEXT, 'Generar Comprobantes'))
	)
	btnGenerarComprobantes.click()
	# btnGenerarComprobantes = browser.find_element(By.LINK_TEXT, 'Generar Comprobantes')

	# Hago click en el select de "Punto de Ventas a utilizar", escojo el option correcto y hago click en "Continuar"
	selectPuntoVenta = WebDriverWait(browser, 10).until(
	    EC.element_to_be_clickable((By.ID, 'puntodeventa'))
	)
	select = Select(selectPuntoVenta) # Creo un objeto Select
	select.select_by_index(1) # Selecciono por índice (empieza en 0)
	time.sleep(1) # Espero que se cargue "Tipo de Comprobante" después de seleccionar el punto de venta
	btnContinuar = browser.find_element(By.XPATH, "//input[@type='button' and @value='Continuar >']")
	btnContinuar.click()

	time.sleep(1)

	# Pantalla: DATOS DE EMISIÓN (PASO 1 DE 4)
	# ------------------------------------------------------------------------------------------------
	selectConceptos = WebDriverWait(browser, 10).until(
	    EC.element_to_be_clickable((By.ID, 'idconcepto'))
	)
	select = Select(selectConceptos) # Creo un objeto Select
	select.select_by_value('3') # Selecciona la opción que muestra "Productos y Servicios".
	time.sleep(1) # Espero que se cargue la sección "Período Facturado" que contiene "Vto. para el Pago"
	inputFecha = browser.find_element(By.ID, 'vencimientopago')
	currentDate = inputFecha.get_attribute('value')

	# Agrego los días a la fecha actual
	dateObj = datetime.strptime(currentDate, '%d/%m/%Y') # Convierto el string de fecha actual a un objeto datetime
	newDateObj = dateObj + timedelta(days=daysToAdd) # Agrego los días
	dueDate = newDateObj.strftime('%d/%m/%Y') # Convierto el objeto datetime de nuevo a un string
	
	# Pongo la nueva fecha en el input
	inputFecha.clear()
	inputFecha.send_keys(dueDate)

	btnContinuar = browser.find_element(By.XPATH, "//input[@type='button' and @value='Continuar >']")
	btnContinuar.click()

	time.sleep(1)

	# Pantalla: DATOS DEL RECEPTOR (PASO 2 DE 4)
	# ------------------------------------------------------------------------------------------------

	# Select "Condición frente al IVA"
	selectCondicion = WebDriverWait(browser, 10).until(
	    EC.element_to_be_clickable((By.ID, 'idivareceptor'))
	)
	select = Select(selectCondicion)
	select.select_by_value('1') # Selecciono "IVA Responsable Inscripto"

	# Input "CUIT"
	inputCuitReceptor = browser.find_element(By.ID, 'nrodocreceptor')
	inputCuitReceptor.send_keys(cuitReceptor)
	inputCuitReceptor.send_keys(Keys.ENTER) # Para forzar la carga de los datos del receptor
	time.sleep(1)

	# Checkbox "Cuenta Corriente"
	checkboxCC = browser.find_element(By.ID, 'formadepago4')
	checkboxCC.click()

	# Botón "Continuar"
	btnContinuar = browser.find_element(By.XPATH, "//input[@type='button' and @value='Continuar >']")
	btnContinuar.click()

	time.sleep(1)

	# Pantalla: DATOS DE LA OPERACIÓN (PASO 3 DE 4)
	# ------------------------------------------------------------------------------------------------
	
	for index, row in enumerate(data):
		# Relleno "Producto/Servicio"
		inputProSer = WebDriverWait(browser, 10).until(
		    EC.element_to_be_clickable((By.ID, f'detalle_descripcion{index + 1}'))
		)
		inputProSer.send_keys(data[index]['prodServ'])

		# Selecciono "U. Medida"
		select = Select(browser.find_element(By.ID, f'detalle_medida{index + 1}'))
		select.select_by_visible_text('unidades')

		# Relleno "Prec. Unitario"
		inputPrecio = browser.find_element(By.ID, f'detalle_precio{index + 1}')
		inputPrecio.send_keys(data[index]['precUnit'])

		# Presiono "Agregar línea descripción" si no es la última fila
		if index < len(data) - 1:
			btnAgregar = browser.find_element(By.CSS_SELECTOR, 'input[value="Agregar línea descripción"]')
			btnAgregar.click()


	
	input()


# automateInvoiceCreation(15)



