import pyperclip
import re


# Indices de cada columna
colToInd = {
	'ot': 0,
	'fecha': 1,
	'solicitante': 2,
	'sector': 3,
	'descripcion': 4,
	'entrega': 5,
	'siniva': 6,
	'coniva': 7
}


def formatClipboard():
	"""Formatea el texto del portapapeles para obtener la tabla en un formato manejable."""

	# Obtengo el contenido del portapapeles
	cb = pyperclip.paste()

	# Reemplazo saltos de línea dentro de las celdas por un espacio
	cb = re.sub(r'(\t".*?"\t)', lambda m: m.group(1).replace('\n', ' '), cb, flags=re.DOTALL)

	# Quito las comillas dobles que se agregan al texto de celdas con saltos de linea
	cb = re.sub(r'(?<=\t)"', '', cb) # Lookbehind positivo
	cb = re.sub(r'"(?=\t)', '', cb) # Lookahead positivo

	# Divido en filas y columnas
	rows = cb.splitlines() # Separo en filas
	table = [[col.strip() for col in row.split('\t')] for row in rows] # Separo en columnas (quitando espacios de más)

	# # Imprimo la tabla
	# for row in table:
	# 	print(row)

	return table


def checkClipboard():
	"""Verifica si lo copiado en el portapapeles es de la hoja de Google Sheets."""

	# Obtengo la tabla formateada
	table = formatClipboard()

	# Verifico si hay texto en el portapapeles
	if not table:
		return False, 'No se copió nada'

	# Verifico cantidad correcta de columnas (7 hasta "sin IVA", 8 hasta "con IVA")
	for row in table:
		if len(row) < 7 or len(row) > 8: 
			return False, 'No se reconoce lo copiado\n(copiar de "OT" a "Valor S/IVA")'

	# Verifico contenido de cada columna
	for row in table:
		# Columna "OT"
		ot = row[colToInd['ot']]
		if not ot.isdecimal() and ot != '#N/A':
			return False, f'"OT" con formato inválido: {ot}'

		# Columna "Fecha"
		fecha = row[colToInd['fecha']]
		dateRegex = r'''
			\d{4}[-/.]\d{1,2}[-/.]\d{1,2} # Captura fechas con el año primero
			|							  # o...
			\d{1,2}[-/.]\d{1,2}[-/.]\d{4} # Captura fechas con el año al final
		'''
		if not re.match(dateRegex, fecha, re.VERBOSE) and fecha != '#N/A':
			return False, f'"Fecha" con formato inválido: {fecha}'

		# Columna "Solicitado por"
		solicitante = row[colToInd['solicitante']]
		if not solicitante.isalpha() and len(solicitante) > 10 and solicitante != '#N/A':
			return False, f'"Solicitado por" con formato inválido: {solicitante}'

		# Columna "Sector"
		sector = row[colToInd['sector']]
		if not sector.isalpha() and len(sector) > 10 and sector != '#N/A':
			return False, f'"Sector" con formato inválido: {sector}'

		# Columnas "Descripción" y "Entrega" no las controlo

		# Columna "Valor S/IVA"
		siniva = row[colToInd['siniva']]
		sinivaLimpio = siniva.replace('$', '').replace('.', '')
		if not sinivaLimpio.isdecimal() and sinivaLimpio.count(',') != 1:
			return False, f'"Valor S/IVA" con formato inválido: {siniva}'

		# Columna "Valor c/IVA"
		if len(row) == 8:
			coniva = row[colToInd['coniva']]
			conivaLimpio = coniva.replace('$', '').replace('.', '')
			if not conivaLimpio.isdecimal() and conivaLimpio.count(',') != 1:
				return False, f'"Valor C/IVA" con formato inválido: {coniva}'

	return True, 'Listo para generar factura'