import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
import threading
import sys
import os
from datetime import datetime

# Importar el módulo principal de OCR
from ocr_market import (
    main as ocr_main,
    set_log_callback,
    obtener_ventana_juego,
    capturar_ventana,
    extraer_precios,
    init_ocr
)


class OCRMainGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Legend of Ymir - OCR Market Extractor")
        self.root.geometry("900x700")
        
        # Configurar estilo
        self.root.configure(bg="#2b2b2b")
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Treeview', background="#3c3c3c", foreground="#ffffff")
        
        # Variable para el thread
        self.ocr_thread = None
        self.is_running = False
        
        # Crear la interfaz
        self.create_widgets()
        
        # Establecer el callback para logs
        set_log_callback(self.log_message)
        
        # Vincular el cierre de la ventana
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def create_widgets(self):
        """Crea todos los widgets de la interfaz"""
        
        # ====== ENCABEZADO ======
        header_frame = tk.Frame(self.root, bg="#404040", height=60)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="Legend of Ymir - OCR Market Extractor",
            font=("Arial", 14, "bold"),
            bg="#404040",
            fg="#00d4ff"
        )
        title_label.pack(pady=10)
        
        # ====== AREA CENTRAL ======
        main_frame = tk.Frame(self.root, bg="#2b2b2b")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Etiqueta para logs
        log_label = tk.Label(
            main_frame,
            text="📋 Registro de actividades:",
            font=("Arial", 10, "bold"),
            bg="#2b2b2b",
            fg="#ffffff"
        )
        log_label.pack(anchor="w", pady=(0, 5))
        
        # Area de texto para logs (scrolled)
        self.log_text = scrolledtext.ScrolledText(
            main_frame,
            height=25,
            width=100,
            bg="#1e1e1e",
            fg="#00ff00",
            insertbackground="#00ff00",
            font=("Consolas", 9),
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Configurar etiquetas de colores para el texto
        self.log_text.tag_config("info", foreground="#00ff00")
        self.log_text.tag_config("error", foreground="#ff3333")
        self.log_text.tag_config("warning", foreground="#ffaa00")
        self.log_text.tag_config("success", foreground="#00ff66")
        self.log_text.tag_config("debug", foreground="#0088ff")
        
        # ====== BARRA DE PROGRESO ======
        progress_frame = tk.Frame(main_frame, bg="#2b2b2b")
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        progress_label = tk.Label(
            progress_frame,
            text="Progreso:",
            font=("Arial", 9),
            bg="#2b2b2b",
            fg="#ffffff"
        )
        progress_label.pack(anchor="w", pady=(0, 5))
        
        self.progress = ttk.Progressbar(
            progress_frame,
            length=400,
            mode='indeterminate'
        )
        self.progress.pack(fill=tk.X)
        
        # ====== BARRA DE ESTADO ======
        status_frame = tk.Frame(self.root, bg="#404040", height=30)
        status_frame.pack(fill=tk.X, padx=0, pady=0)
        status_frame.pack_propagate(False)
        
        self.status_label = tk.Label(
            status_frame,
            text="Estado: Listo",
            font=("Arial", 9),
            bg="#404040",
            fg="#00d4ff",
            anchor="w",
            padx=10
        )
        self.status_label.pack(fill=tk.BOTH, expand=True)
        
        # ====== BOTONES ======
        button_frame = tk.Frame(self.root, bg="#2b2b2b")
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Estilo de botones
        self.create_button_style()
        
        # Botón de inicio
        self.start_btn = tk.Button(
            button_frame,
            text="▶ Iniciar Extracción",
            command=self.start_ocr,
            font=("Arial", 10, "bold"),
            bg="#00aa44",
            fg="#ffffff",
            padx=20,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        # Botón de limpiar logs
        self.clear_btn = tk.Button(
            button_frame,
            text="🗑️ Limpiar Logs",
            command=self.clear_logs,
            font=("Arial", 10),
            bg="#555555",
            fg="#ffffff",
            padx=15,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Botón de salir
        self.exit_btn = tk.Button(
            button_frame,
            text="❌ Salir",
            command=self.on_closing,
            font=("Arial", 10),
            bg="#aa0000",
            fg="#ffffff",
            padx=20,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        self.exit_btn.pack(side=tk.RIGHT, padx=5)
        
        # Log inicial
        self.log_message("")
        self.log_message("="*60)
        self.log_message("  BIENVENIDO AL EXTRACTOR OCR DE YMIR MARKET")
        self.log_message("="*60)
        self.log_message("")
        self.log_message("▸ Asegúrese de que:")
        self.log_message("  1. El juego Legend of Ymir está abierto")
        self.log_message("  2. El archivo rois.json está configurado correctamente")
        self.log_message("  3. Presione INICIAR para comenzar")
        self.log_message("")
        self.log_message("El proceso extraerá precios de:")
        self.log_message("  • Market WEMIX")
        self.log_message("  • Market Diamantes")
        self.log_message("")
    
    def create_button_style(self):
        """Crea estilos para los botones"""
        pass
    
    def log_message(self, msg, tag="info"):
        """Agrega un mensaje al área de logs con timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Determinar etiqueta según el contenido del mensaje
        if "[ERROR]" in msg:
            tag = "error"
        elif "[WARN]" in msg:
            tag = "warning"
        elif "[OK]" in msg:
            tag = "success"
        elif "[DEBUG]" in msg:
            tag = "debug"
        
        # Agregar mensaje al texto
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n", tag)
        self.log_text.see(tk.END)  # Scroll automático hacia el final
        self.log_text.config(state=tk.DISABLED)
        
        # Actualizar la ventana
        self.root.update()
    
    def clear_logs(self):
        """Limpia el área de logs"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log_message("Logs limpios")
    
    def update_status(self, status, color="#00d4ff"):
        """Actualiza la barra de estado"""
        self.status_label.config(text=f"Estado: {status}", fg=color)
    
    def start_ocr(self):
        """Inicia el proceso OCR en un thread separado"""
        if self.is_running:
            messagebox.showwarning("Advertencia", "El proceso ya está en ejecución")
            return
        
        # Verificar que el juego esté abierto
        ventana = obtener_ventana_juego()
        if not ventana:
            messagebox.showerror("Error", "No se encontró la ventana del juego.\nAsegúrese de que 'Legend of Ymir' está abierto.")
            return
        
        # Desactivar botones
        self.start_btn.config(state=tk.DISABLED)
        self.clear_btn.config(state=tk.DISABLED)
        
        # Actualizar estado
        self.is_running = True
        self.update_status("En ejecución...", "#ffaa00")
        self.progress.start()
        
        # Iniciar el thread del OCR
        self.ocr_thread = threading.Thread(target=self.run_ocr, daemon=True)
        self.ocr_thread.start()
    
    def run_ocr(self):
        """Ejecuta el proceso OCR (en thread)"""
        try:
            # Log inicial
            self.log_message("")
            self.log_message("="*60)
            self.log_message("Iniciando extracción...", "info")
            self.log_message("="*60)
            self.log_message("")
            
            # Ejecutar el main del OCR
            success = ocr_main()
            
            if success:
                self.log_message("")
                self.log_message("✓ Proceso completado exitosamente", "success")
                self.update_status("Completado", "#00ff66")
            else:
                self.log_message("")
                self.log_message("✗ El proceso fue cancelado o falló", "error")
                self.update_status("Error/Cancelado", "#ff3333")
        
        except Exception as e:
            self.log_message("")
            self.log_message(f"✗ Error durante la ejecución: {str(e)}", "error")
            import traceback
            self.log_message(traceback.format_exc(), "error")
            self.update_status("Error", "#ff3333")
        
        finally:
            # Reactivar botones
            self.is_running = False
            self.progress.stop()
            self.start_btn.config(state=tk.NORMAL)
            self.clear_btn.config(state=tk.NORMAL)
            self.log_message("")
            self.log_message("Presione 'Iniciar Extracción' para ejecutar nuevamente")
    
    def on_closing(self):
        """Maneja el cierre de la ventana"""
        if self.is_running:
            if messagebox.askyesno("Cancelar", "El proceso está en ejecución.\n¿Desea cancelarlo y salir?"):
                self.is_running = False
                self.root.quit()
        else:
            self.root.quit()


def main():
    root = tk.Tk()
    app = OCRMainGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
