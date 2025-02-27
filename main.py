from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox
from PyQt5.QtCore import QTimer, QThread, pyqtSignal, Qt
from PyQt5.uic import loadUi
import sys
import clipboard_utils as cu
import afip_automation as aa
import pyperclip


def cargarStylesheet(filename):
	"""Carga la hoja de estilos de la app."""
	with open(filename, 'r') as file:
		return file.read()


class InvoiceAutomationThread(QThread):
	progress = pyqtSignal(int, str) # Emitirá (porcentaje, mensaje)
	finished = pyqtSignal() # Señal cuando termina el proceso

	# En PyQt5, los QMessageBox deben mostrarse desde el hilo principal, por eso emito señales para mostrar los mensajes ahí
	error = pyqtSignal(str) # Señal para errores
	success = pyqtSignal(str) # Señal para mensajes exitosos

	def __init__(self, days, data):
		super().__init__()
		self.days = days
		self.data = data


	def run(self):
		"""Ejecuta la automatización de la factura en un hilo separado."""
		try:
			aa.automateInvoiceCreation(self.days, self.data, self.reportProgress)
			self.success.emit('La factura se generó correctamente.') # Emito mensaje exitoso
		except Exception as e:
			self.error.emit(str(e)) # Emito mensaje de error
		finally:
			self.finished.emit() # Emito señal de finalización


	def reportProgress(self, percent, message):
		"""Método que emite la señal de progreso."""
		self.progress.emit(percent, message)



class MainWindow(QMainWindow):
	def __init__(self):
		super().__init__()

		# Cargo el archivo .ui
		loadUi('ui/main.ui', self)

		# Mantengo la ventana siempre encima
		self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)

		# Configuro un QTimer para ejecutar checkClipboard() cada 300ms
		self.timer = QTimer(self)
		self.timer.timeout.connect(self.checkClipboard)
		self.timer.start(300) # Ejecuta cada 300ms

		# Botón
		self.pushButton.setEnabled(False)
		self.pushButton.clicked.connect(self.generateInvoice)

		# Barra de progreso
		self.progressBar.hide() # También podría usar self.progressBar.setVisible(False)

		# Inicializo variables
		self.oldContent = None
		self.lastValidTable = None


	def checkClipboard(self):
		"""Ejecuta checkClipboard() cada 300ms."""

		# Verifico cambios en el portapapeles, para evitar procesamiento innecesario
		self.newContent = pyperclip.paste()
		if self.oldContent != self.newContent:
			# Resguardo el nuevo contenido
			self.oldContent = self.newContent

			# Proceso el contenido y hago cambios en la UI en consecuencia
			state, detail = cu.checkClipboard()
			self.label_estado.setText(detail)
			if state:
				self.label_estado.setStyleSheet('color: green; font-weight: bold')
				self.pushButton.setEnabled(True)

				# Resguardo la tabla validada
				self.lastValidTable = cu.formatClipboard()
			else:
				self.label_estado.setStyleSheet('color: red; font-weight: bold')
				self.pushButton.setEnabled(False)


	def generateInvoice(self):
		"""Inicia el proceso de creación de factura."""

		# Paro el timer para que no se siga analizando el portapapeles
		self.timer.stop()

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
				self.timer.start(300) # Reanudo el timer que había parado
				return
			
		# Formateo datos a usar en la pantalla "DATOS DE LA OPERACIÓN (PASO 3 DE 4)"
		data = []
		for row in self.lastValidTable:
			# Formo lo que va en el input de Producto/Servicio
			descPart = row[cu.colToInd['descripcion']]
			otPart = f'{' ' if descPart.endswith('.') else '. '}OT {row[cu.colToInd['ot']]}'
			oclPart = f'{'. OCL ' + self.lineEdit_ocl.text() if self.lineEdit_ocl.text() else ''}'
			prodServ = f'{descPart}{otPart}{oclPart}'

			# Formo lo que va en el input de "Prec. Unitario"
			precUnit = row[cu.colToInd['siniva']].replace('$', '').replace('.', '')

			data.append({'prodServ': prodServ, 'precUnit': precUnit})

		# Creo el hilo para ejecutar Selenium sin bloquear la UI, y conecto señales
		self.thread = InvoiceAutomationThread(self.spinBox_dias.value(), data)
		self.thread.progress.connect(self.updateProgressBar) # Conecto señal de progreso
		self.thread.finished.connect(self.onInvoiceFinished) # Conecto señal de finalización
		self.thread.error.connect(self.showErrorMessage) # Conecto señal de error
		self.thread.success.connect(self.showSuccessMessage) # Conecto señal de éxito
		self.thread.start()

		# Deshabilito botón y muestro barra de progreso
		self.progressBar.show()
		self.pushButton.setEnabled(False)
		self.label_estado.setStyleSheet('color: orange; font-weight: bold')


	def showErrorMessage(self, message):
		"""Muestra un mensaje de error en el hilo principal."""
		QMessageBox.critical(self, 'Error', message)


	def showSuccessMessage(self, message):
		"""Muestra un mensaje de éxito en el hilo principal."""
		QMessageBox.information(self, 'Proceso terminado', message)


	def updateProgressBar(self, value, message):
		"""Actualiza la barra de progreso con el valor y mensaje."""
		self.progressBar.setValue(value)
		self.label_estado.setText(message)


	def onInvoiceFinished(self):
		"""Se ejecuta independientemente de si se creó la factura o hubo error."""
		self.progressBar.hide()
		self.oldContent = '' # Fuerzo análisis del clipboard actual para producir los cambios apropiados en la UI
		self.timer.start(300)



if __name__ == "__main__":
	app = QApplication(sys.argv)

	# Cargo y aplico la hoja de estilos de la aplicación
	stylesheet = cargarStylesheet('resources/styles.css')
	app.setStyleSheet(stylesheet)

	window = MainWindow()
	window.show()
	sys.exit(app.exec_())