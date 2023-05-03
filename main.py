import win32print
import win32ui

# Nome da impressora que será calibrada
printer_name = "ZEBRA"

# Comando ZPL para calibrar a impressora

#calibrate_cmd = '^XA^FO20,20^ADN,36,20^FDOii Principe^FS^XZ'

calibrate_cmd = '~JC' #ZD220
#calibrate_cmd = '''^XA
                    ^JUS
                    ^XZ
                  '''

# Abre uma conexão com a impressora selecionada
printer_handle = win32print.OpenPrinter(printer_name)

# Obtém informações sobre a impressora selecionada
printer_info = win32print.GetPrinter(printer_handle, 2)
printer_name = printer_info["pPrinterName"]

# Cria um objeto DC (Device Context) para a impressora
dc = win32ui.CreateDC()

# Define o objeto DC para a impressora selecionada
dc.CreatePrinterDC(printer_name)

# Envia o comando de calibração para a impressora
job = win32print.StartDocPrinter(printer_handle, 1, ("Calibrate Zebra", None, None))
win32print.StartPagePrinter(printer_handle)
win32print.WritePrinter(printer_handle, calibrate_cmd.encode())
win32print.EndPagePrinter(printer_handle)
win32print.EndDocPrinter(printer_handle)

# Fecha a conexão com a impressora
win32print.ClosePrinter(printer_handle)
