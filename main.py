# main.py

import tkinter as tk
from snmp_helper import get_printer_info
from printer_info_app import PrinterInfoApp
from email_helper import send_email_alert, send_email_alert_serial_number_change
from snmp_helper import get_printer_info

if __name__ == '__main__':
    root = tk.Tk()
    app = PrinterInfoApp(root)
    root.mainloop()
