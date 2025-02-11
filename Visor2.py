import os
import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser
import subprocess

class FileViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Búsqueda de Información de Pago")
        self.root.geometry("800x600")
        
        # Variables
        self.file_path = tk.StringVar()
        self.directory = tk.StringVar()
        
        # Interfaz
        tk.Label(root, text="Número de Cédula:").pack(pady=10)
        self.entry_cedula = tk.Entry(root, width=30)
        self.entry_cedula.pack()
        
        self.lbl_status = tk.Label(root, text="", fg="blue")
        self.lbl_status.pack(pady=5)
        
        tk.Button(root, text="Buscar", command=self.search_file).pack(pady=10)
        tk.Button(root, text="Abrir en Directorio", command=self.open_directory).pack(pady=5)
        tk.Button(root, text="Copiar Ruta al Portapapeles", command=self.copy_to_clipboard).pack(pady=5)
        tk.Button(root, text="Cerrar", command=self.root.quit).pack(pady=10)
        
    def search_file(self):
        cedula = self.entry_cedula.get().strip()
        if not cedula:
            messagebox.showwarning("Error", "Debe ingresar un número de cédula.")
            return
        
        # Buscar archivo en el directorio seleccionado
        directory = filedialog.askdirectory(title="Seleccione el directorio de búsqueda")
        if not directory:
            return
        
        self.directory.set(directory)
        self.lbl_status.config(text="Buscando...")
        file_name = f"{cedula}.html"
        
        for root, _, files in os.walk(directory):
            if file_name in files:
                self.file_path.set(os.path.join(root, file_name))
                self.lbl_status.config(text=f"Archivo encontrado: {file_name}")
                webbrowser.open(self.file_path.get())  # Abrir en el navegador
                return
        
        self.lbl_status.config(text="Archivo no encontrado.")
        messagebox.showinfo("Resultado", "No se encontró el archivo con la cédula especificada.")
    
    def open_directory(self):
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showwarning("Error", "No hay archivo seleccionado para visualizar.")
            return
        
        # Abrir el directorio del archivo en el explorador de archivos
        directory = os.path.dirname(file_path)
        try:
            if os.name == 'nt':  # Windows
                os.startfile(directory)
            elif os.name == 'posix':  # macOS o Linux
                subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', directory])
            else:
                messagebox.showwarning("Error", "No se puede abrir el directorio en este sistema operativo.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el directorio. Detalle: {e}")
    
    def copy_to_clipboard(self):
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showwarning("Error", "No hay archivo seleccionado para copiar la ruta.")
            return
        
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(file_path)
            self.root.update()  # Actualizar el portapapeles
            messagebox.showinfo("Éxito", "La ruta del archivo se copió al portapapeles.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo copiar la ruta. Detalle: {e}")

