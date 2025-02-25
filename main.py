from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox
from PyQt5.uic import loadUi
from PyQt5.QtCore import QTimer
import sys
import clipboard_handler as ch
import selenium_handler as sh
import pyperclip


def cargarStylesheet(filename):
	"""Carga la hoja de estilos de la app."""

	with open(filename, 'r') as file:
		return file.read()


class MainWindow(QMainWindow):
	def __init__(self):
		super().__init__()

		# Cargo el archivo .ui
		loadUi('ui/app.ui', self)

		# Configuro un QTimer para ejecutar checkClipboard() cada 300ms
		self.timer = QTimer(self)
		self.timer.timeout.connect(self.checkClipboard)
		self.timer.start(300) # Ejecuta cada 300ms

		self.pushButton.setEnabled(False)

		self.oldContent = None
		self.lastValidTable = None

		self.pushButton.clicked.connect(self.generateInvoice)


	def checkClipboard(self):
		"""Ejecuta checkClipboard() cada 300ms."""

		# Verifico cambios en el portapapeles, para evitar procesamiento innecesario
		self.newContent = pyperclip.paste()
		if self.oldContent != self.newContent:
			# Resguardo el nuevo contenido
			self.oldContent = self.newContent

			# Proceso el contenido y hago cambios en la interfaz en consecuencia
			state, detail = ch.checkClipboard()
			self.label_estado.setText(detail)
			if state:
				self.label_estado.setStyleSheet('color: green; font-weight: bold')
				self.pushButton.setEnabled(True)

				# Resguardo la tabla validada
				self.lastValidTable = ch.formatClipboard()
			else:
				self.label_estado.setStyleSheet('color: red; font-weight: bold')
				self.pushButton.setEnabled(False)


	def generateInvoice(self):
		"""Inicia el proceso de creación de factura."""

		# Verifico si el OCL está vacio y muestro un mensaje
		if not self.lineEdit_ocl.text():
			msg = QMessageBox(self)
			msg.setIcon(QMessageBox.Warning)
			msg.setWindowTitle('Confirmación')
			msg.setText('¿Desea continuar sin OCL?')
			msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

			# Cambio el texto de los botones a "Sí" y "No"
			msg.button(QMessageBox.Yes).setText('Sí')
			msg.button(QMessageBox.No).setText('No')

			# Cambio el tamaño de la letra del mensaje
			msg.setStyleSheet('QLabel {font-size: 11pt}')

			answer = msg.exec_()

			if answer == QMessageBox.No:
				return
			
		# Formateo datos a usar en la pantalla "DATOS DE LA OPERACIÓN (PASO 3 DE 4)"
		data = []
		for row in self.lastValidTable:
			# Formo lo que va en el input de Producto/Servicio
			descPart = row[ch.colToInd['descripcion']]
			otPart = f'{' ' if descPart.endswith('.') else '. '}OT {row[ch.colToInd['ot']]}'
			oclPart = f'{'. OCL ' + self.lineEdit_ocl.text() if self.lineEdit_ocl.text() else ''}'
			prodServ = f'{descPart}{otPart}{oclPart}'

			# Formo lo que va en el input de "Prec. Unitario"
			precUnit = row[ch.colToInd['siniva']].replace('$', '').replace('.', '')

			data.append({'prodServ': prodServ, 'precUnit': precUnit})

		# for fila in data:
		# 	print(fila)

		# Llamo a la función que automatiza la creación de la factura
		sh.automateInvoiceCreation(self.spinBox_dias.value(), data)




			


if __name__ == "__main__":
	app = QApplication(sys.argv)

	# Cargo y aplico la hoja de estilos de la aplicación
	stylesheet = cargarStylesheet('resources/styles.css')
	app.setStyleSheet(stylesheet)

	window = MainWindow()
	window.show()
	sys.exit(app.exec_())