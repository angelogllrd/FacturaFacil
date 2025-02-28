from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime, timedelta
import json


# Obtengo credenciales desde el archivo JSON
with open('credenciales.json', 'r') as file:
	credentials = json.load(file)

CUIT = credentials['cuit']
CLAVE = credentials['clave']
CUIT_RECEPTOR = credentials['cuit_receptor']

# Ruta del ChromeDriver
CHROME_DRIVER_PATH = "C:/Users/ag/Desktop/chromedriver.exe"


def automateInvoiceCreation(days, data, callback):
	"""Automatiza la creación de una factura en el sistema de AFIP."""

	callback(0, 'Entrando al inicio de sesión de AFIP...')
	
	# Configuración del navegador
	options = webdriver.ChromeOptions()
	options.add_experimental_option('detach', True) # Evita que el navegador se cierre al terminar
	options.add_argument('--start-maximized') # Inicia el navegador maximizado
	service = Service(CHROME_DRIVER_PATH)
	browser = webdriver.Chrome(service=service, options=options)

	try:
		# Abro el sitio de AFIP y selecciono iniciar sesión
		browser.get('https://www.afip.gob.ar/landing/default.asp')
		WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, 'Iniciar sesión'))).click()

		# Cambio a la nueva pestaña
		WebDriverWait(browser, 10).until(lambda b: len(b.window_handles) > 1)
		browser.switch_to.window(browser.window_handles[-1])
	except Exception as e:
		raise Exception(f'Error entrando al sitio de inicio de sesión: {e}') # Propago el error con contexto
	
	callback(11, 'Ingresando credenciales...')

	try:
		# Ingreso CUIT y clave
		WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, 'F1:username'))).send_keys(CUIT)
		browser.find_element(By.ID, 'F1:btnSiguiente').click()
		WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, 'F1:password'))).send_keys(CLAVE)
		browser.find_element(By.ID, 'F1:btnIngresar').click()
	except Exception as e:
		raise Exception(f'Error en el logueo: {e}')

	callback(22, 'Seleccionando "Comprobantes en línea"...')

	try:
		# Accedo a "Comprobantes en línea"
		btnComprobantes = WebDriverWait(browser, 10).until(
			EC.element_to_be_clickable((By.XPATH, '//a//*[text()="Comprobantes en línea"]'))
		)
		browser.execute_script('arguments[0].scrollIntoView(true);', btnComprobantes) # Scroll para hacer visible el elemento
		btnComprobantes.click()

		# Cambio a la nueva pestaña
		WebDriverWait(browser, 10).until(lambda b: len(b.window_handles) > 2)
		browser.switch_to.window(browser.window_handles[-1])
	except Exception as e:
		raise Exception(f'Error en selección de "Comprobantes en línea": {e}')


	# Se sigue con el sistema RCEL (acrónimo de Registro de Comprobantes en Línea)
	# ------------------------------------------------------------------------------------------------

	callback(33, 'Seleccionando empresa...')

	try:
		# Selecciono la empresa
		WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn_empresa'))).click()
	except Exception as e:
		raise Exception(f'Error en selección de empresa: {e}')

	callback(44, 'Seleccionando "Generar Comprobantes"...')

	try:
		# Selecciono "Generar Comprobantes"
		WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, 'Generar Comprobantes'))).click()
	except Exception as e:
		raise Exception(f'Error en selección de "Generar Comprobantes": {e}')

	callback(55, 'Seleccionando punto de venta y tipo de factura...')

	try:
		# Selecciono el punto de ventas
		selectPuntoVenta = WebDriverWait(browser, 10).until(
			EC.element_to_be_clickable((By.ID, 'puntodeventa'))
		)
		Select(selectPuntoVenta).select_by_index(1)

		# Selecciono tipo de factura
		WebDriverWait(browser, 10).until(
			EC.presence_of_element_located((By.XPATH, "//select[@id='universocomprobante']/option[text()='Factura A']")) # Espero que se carguen las options
		)
		Select(browser.find_element(By.ID, 'universocomprobante')).select_by_visible_text('Factura A') # Selecciono "Factura A"

		# Selecciono "Continuar"
		browser.find_element(By.XPATH, "//input[@type='button' and @value='Continuar >']").click()
	except Exception as e:
		raise Exception(f'Error en selección de punto de venta y tipo de factura: {e}')


	# DATOS DE EMISIÓN (PASO 1 DE 4)
	# ------------------------------------------------------------------------------------------------

	callback(66, 'Seleccionando conceptos y fecha de vencimiento...')

	try:
		# Selecciono los conceptos a incluir
		selectConceptos = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, 'idconcepto')))
		Select(selectConceptos).select_by_value('3') # Selecciono "Productos y Servicios"
		
		# Ajuste de fecha de vencimiento
		inputFecha = WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.ID, 'vencimientopago')))
		currentDateStr = inputFecha.get_attribute('value')
		currentDateObj = datetime.strptime(currentDateStr, '%d/%m/%Y') # Convierto el string de fecha actual a un objeto datetime
		dueDateObj = currentDateObj + timedelta(days=days) # Agrego los días
		dueDateStr = dueDateObj.strftime('%d/%m/%Y') # Convierto el objeto datetime de nuevo a un string
		inputFecha.clear()
		inputFecha.send_keys(dueDateStr)

		# Selecciono "Continuar"
		browser.find_element(By.XPATH, "//input[@type='button' and @value='Continuar >']").click()
	except Exception as e:
		raise Exception(f'Error en selección de conceptos y fecha de vencimiento: {e}')


	# DATOS DEL RECEPTOR (PASO 2 DE 4)
	# ------------------------------------------------------------------------------------------------

	callback(77, 'Ingresando datos del receptor...')

	try:
		# Selecciono la condición frente al IVA
		selectCondicion = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.ID, 'idivareceptor')))
		Select(selectCondicion).select_by_value('1') # Selecciono "IVA Responsable Inscripto"

		# Ingreso "CUIT"
		browser.find_element(By.ID, 'nrodocreceptor').send_keys(CUIT_RECEPTOR, Keys.ENTER)

		# Espero a que se cargue la información del receptor
		WebDriverWait(browser, 10).until(
		    lambda b: b.find_element(By.ID, 'razonsocialreceptor').get_attribute('value').strip() != '' # Uso una condición personalizada
		)

		# Selecciono "Cuenta Corriente"
		browser.find_element(By.ID, 'formadepago4').click()

		# Selecciono "Continuar"
		browser.find_element(By.XPATH, "//input[@type='button' and @value='Continuar >']").click()
	except Exception as e:
		raise Exception(f'Error en el ingreso de datos del receptor: {e}')


	# DATOS DE LA OPERACIÓN (PASO 3 DE 4)
	# ------------------------------------------------------------------------------------------------

	callback(88, 'Rellenando líneas de factura...')
	
	try:
		numItems = len(data)
		for index, item in enumerate(data):
			callback(88, f'Rellenando líneas de factura ({index + 1}/{numItems})...')

			# Relleno "Producto/Servicio"
			WebDriverWait(browser, 10).until(
				EC.element_to_be_clickable((By.ID, f'detalle_descripcion{index + 1}'))
			).send_keys(item['prodServ'])

			# Selecciono "U. Medida"
			Select(browser.find_element(By.ID, f'detalle_medida{index + 1}')).select_by_visible_text('unidades')

			# Relleno "Prec. Unitario"
			browser.find_element(By.ID, f'detalle_precio{index + 1}').send_keys(item['precUnit'])

			# Selecciono "Agregar línea descripción" si no es la última fila
			if index < len(data) - 1:
				browser.find_element(By.CSS_SELECTOR, 'input[value="Agregar línea descripción"]').click()
	except Exception as e:
		raise Exception(f'Error rellenando líneas de factura: {e}')

	callback(100, 'Factura generada')