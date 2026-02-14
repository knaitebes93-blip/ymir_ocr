import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog, ttk
import json
import os
import queue
import threading
from datetime import datetime
import keyboard

from ocr_market import (
    append_capture_to_excel,
    capturar_ventana,
    detectar_tipo_market,
    extraer_precios,
    init_ocr,
    obtener_ventana_juego,
    set_log_callback,
    sync_to_google_sheet,
)


class OCRMainGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Legend of Ymir - OCR Market Extractor")
        self.root.geometry("980x720")

        self.root.configure(bg="#2b2b2b")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#3c3c3c", foreground="#ffffff")

        self.worker_thread = None
        self.processor_running = False
        self.capture_queue = queue.Queue()
        self.job_counter = 0
        self.processed_jobs = 0
        self.failed_jobs = 0
        self.hotkeys_registered = False

        self.create_widgets()
        set_log_callback(self.log_message)
        self.setup_hotkeys()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        header_frame = tk.Frame(self.root, bg="#404040", height=60)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)

        title_label = tk.Label(
            header_frame,
            text="Legend of Ymir - OCR Market Extractor",
            font=("Arial", 14, "bold"),
            bg="#404040",
            fg="#00d4ff",
        )
        title_label.pack(pady=10)

        main_frame = tk.Frame(self.root, bg="#2b2b2b")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        log_label = tk.Label(
            main_frame,
            text="Registro de actividades:",
            font=("Arial", 10, "bold"),
            bg="#2b2b2b",
            fg="#ffffff",
        )
        log_label.pack(anchor="w", pady=(0, 5))

        self.log_text = scrolledtext.ScrolledText(
            main_frame,
            height=25,
            width=100,
            bg="#1e1e1e",
            fg="#00ff00",
            insertbackground="#00ff00",
            font=("Consolas", 9),
            wrap=tk.WORD,
            state=tk.DISABLED,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.log_text.tag_config("info", foreground="#00ff00")
        self.log_text.tag_config("error", foreground="#ff3333")
        self.log_text.tag_config("warning", foreground="#ffaa00")
        self.log_text.tag_config("success", foreground="#00ff66")
        self.log_text.tag_config("debug", foreground="#0088ff")

        progress_frame = tk.Frame(main_frame, bg="#2b2b2b")
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        progress_label = tk.Label(
            progress_frame,
            text="Progreso:",
            font=("Arial", 9),
            bg="#2b2b2b",
            fg="#ffffff",
        )
        progress_label.pack(anchor="w", pady=(0, 5))

        self.progress = ttk.Progressbar(progress_frame, length=400, mode="indeterminate")
        self.progress.pack(fill=tk.X)

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
            padx=10,
        )
        self.status_label.pack(fill=tk.BOTH, expand=True)

        button_frame = tk.Frame(self.root, bg="#2b2b2b")
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        button_row_top = tk.Frame(button_frame, bg="#2b2b2b")
        button_row_top.pack(fill=tk.X, pady=(0, 6))

        button_row_bottom = tk.Frame(button_frame, bg="#2b2b2b")
        button_row_bottom.pack(fill=tk.X)

        self.start_btn = tk.Button(
            button_row_top,
            text="Iniciar Procesador",
            command=self.start_processor,
            font=("Arial", 10, "bold"),
            bg="#00aa44",
            fg="#ffffff",
            padx=16,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2,
        )
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = tk.Button(
            button_row_top,
            text="Detener Procesador",
            command=self.stop_processor,
            font=("Arial", 10),
            bg="#aa6600",
            fg="#ffffff",
            padx=16,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2,
            state=tk.DISABLED,
        )
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.capture_btn = tk.Button(
            button_row_top,
            text="Capturar",
            command=self.capture_and_enqueue,
            font=("Arial", 10),
            bg="#1266d4",
            fg="#ffffff",
            padx=16,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2,
        )
        self.capture_btn.pack(side=tk.LEFT, padx=5)

        self.clear_btn = tk.Button(
            button_row_bottom,
            text="Limpiar Logs",
            command=self.clear_logs,
            font=("Arial", 10),
            bg="#555555",
            fg="#ffffff",
            padx=15,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2,
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)

        self.add_roi_btn = tk.Button(
            button_row_bottom,
            text="Agregar ROI Columna",
            command=self.add_column_roi,
            font=("Arial", 10),
            bg="#2d6ca2",
            fg="#ffffff",
            padx=15,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2,
        )
        self.add_roi_btn.pack(side=tk.LEFT, padx=5)

        self.add_tab_switch_roi_btn = tk.Button(
            button_row_bottom,
            text="Config ROI TAB SWITCH",
            command=lambda: self.add_column_roi(preset_key="tab_switch"),
            font=("Arial", 10),
            bg="#5a4fb3",
            fg="#ffffff",
            padx=15,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2,
        )
        self.add_tab_switch_roi_btn.pack(side=tk.LEFT, padx=5)

        self.exit_btn = tk.Button(
            button_row_bottom,
            text="Salir",
            command=self.on_closing,
            font=("Arial", 10),
            bg="#aa0000",
            fg="#ffffff",
            padx=20,
            pady=10,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2,
        )
        self.exit_btn.pack(side=tk.RIGHT, padx=5)

        self.log_message("")
        self.log_message("=" * 60)
        self.log_message("  BIENVENIDO AL EXTRACTOR OCR DE YMIR MARKET")
        self.log_message("=" * 60)
        self.log_message("1) Inicia el procesador")
        self.log_message("2) Usa Capturar (detecta automaticamente WEMIX/DIAMANTES por switch)")
        self.log_message("3) El worker analiza en segundo plano y guarda en Excel")
        self.log_message("Configura ROIs: item/price/sales y tab_switch (deteccion auto de tab)")
        self.log_message("Atajos: F8 = Capturar")
        self.log_message("")

    def setup_hotkeys(self):
        if self.hotkeys_registered:
            return
        try:
            keyboard.add_hotkey("f8", lambda: self.root.after(0, self.capture_and_enqueue))
            self.hotkeys_registered = True
            self.log_message("[OK] Atajo activo: F8 (Capturar)")
        except Exception as exc:
            self.log_message(f"[WARN] No se pudieron activar atajos globales: {exc}")

    def log_message(self, msg, tag="info"):
        if threading.current_thread() is not threading.main_thread():
            self.root.after(0, self.log_message, msg, tag)
            return

        timestamp = datetime.now().strftime("%H:%M:%S")
        if "[ERROR]" in msg:
            tag = "error"
        elif "[WARN]" in msg:
            tag = "warning"
        elif "[OK]" in msg:
            tag = "success"
        elif "[DEBUG]" in msg:
            tag = "debug"

        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n", tag)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def clear_logs(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.log_message("Logs limpios")

    def update_status(self, status, color="#00d4ff"):
        if threading.current_thread() is not threading.main_thread():
            self.root.after(0, self.update_status, status, color)
            return
        self.status_label.config(text=f"Estado: {status}", fg=color)

    def start_processor(self):
        if self.processor_running:
            messagebox.showwarning("Advertencia", "El procesador ya esta en ejecucion")
            return

        ventana = obtener_ventana_juego()
        if not ventana:
            messagebox.showerror(
                "Error",
                "No se encontro la ventana del juego. Asegurate que 'Legend of Ymir' este abierto.",
            )
            return

        self.processor_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.add_roi_btn.config(state=tk.DISABLED)
        self.add_tab_switch_roi_btn.config(state=tk.DISABLED)
        self.progress.start()
        self.update_status("Procesador activo", "#ffaa00")

        self.worker_thread = threading.Thread(target=self.run_worker, daemon=True)
        self.worker_thread.start()
        self.log_message("[OK] Procesador OCR iniciado")

    def stop_processor(self):
        if not self.processor_running:
            return

        self.processor_running = False
        self.stop_btn.config(state=tk.DISABLED)
        self.log_message("[INFO] Deteniendo procesador. Se terminara al vaciar la cola...")

    def capture_and_enqueue(self):
        if not self.processor_running:
            self.start_processor()
            if not self.processor_running:
                return

        self.log_message("[*] Capturando pantalla...")
        img = capturar_ventana()
        if img is None:
            messagebox.showerror("Error", "No se pudo capturar la ventana del juego")
            return

        self.job_counter += 1
        job = {
            "id": self.job_counter,
            "img": img,
            "captured_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        self.capture_queue.put(job)

        queue_size = self.capture_queue.qsize()
        self.log_message(f"[OK] Captura #{job['id']} encolada. Pendientes: {queue_size}")
        self.update_status(f"Procesador activo | Cola: {queue_size}", "#ffaa00")
        self.refocus_gui()

    def refocus_gui(self):
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(120, lambda: self.root.attributes("-topmost", False))
            self.root.focus_force()
        except Exception as exc:
            self.log_message(f"[WARN] No se pudo devolver foco a la GUI: {exc}")

    def run_worker(self):
        try:
            init_ocr()
            while self.processor_running or not self.capture_queue.empty():
                try:
                    job = self.capture_queue.get(timeout=0.4)
                except queue.Empty:
                    continue

                self.process_job(job)
                self.capture_queue.task_done()
        except Exception as exc:
            self.failed_jobs += 1
            self.log_message(f"[ERROR] Worker detenido por error: {exc}")
        finally:
            self.root.after(0, self.on_worker_finished)

    def process_job(self, job):
        job_id = job["id"]

        try:
            self.log_message(f"[*] Procesando captura #{job_id}...")
            tipo_detectado = detectar_tipo_market(job["img"])
            if tipo_detectado not in ("WEMIX", "DIAMANTES"):
                self.log_message(f"[WARN] Captura #{job_id}: switch no reconocido ({tipo_detectado}). Se usa WEMIX.")
                tipo_detectado = "WEMIX"
            else:
                self.log_message(f"[OK] Captura #{job_id}: switch detectado -> {tipo_detectado}")

            items = extraer_precios(job["img"], debug_tag=f"{tipo_detectado.lower()}_{job_id}")
            self.log_message(f"[OK] Captura #{job_id} ({tipo_detectado}): filas detectadas {len(items)}")

            append_capture_to_excel(items, tipo_detectado, timestamp=job["captured_at"])

            if tipo_detectado == "WEMIX":
                sync_to_google_sheet(items, [])
            else:
                sync_to_google_sheet([], items)

            self.processed_jobs += 1
            pending = self.capture_queue.qsize()
            self.update_status(
                f"Procesador activo | Pendientes: {pending} | OK: {self.processed_jobs}",
                "#00ff66" if pending == 0 else "#ffaa00",
            )
        except Exception as exc:
            self.failed_jobs += 1
            self.log_message(f"[ERROR] Fallo en captura #{job_id}: {exc}")

    def on_worker_finished(self):
        self.progress.stop()
        self.processor_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.add_roi_btn.config(state=tk.NORMAL)
        self.add_tab_switch_roi_btn.config(state=tk.NORMAL)
        self.update_status(
            f"Listo | Procesadas: {self.processed_jobs} | Fallidas: {self.failed_jobs}",
            "#00d4ff",
        )
        self.log_message("[OK] Procesador detenido")

    def add_column_roi(self, preset_key=None):
        if self.processor_running:
            messagebox.showwarning(
                "Advertencia",
                "No se puede editar ROIs mientras el procesador esta en ejecucion.",
            )
            return

        if preset_key:
            key = str(preset_key).strip().lower().replace(" ", "_")
        else:
            column_name = simpledialog.askstring(
                "Nueva Columna",
                "Nombre de la columna para la ROI (ej: item, price, sales, tab_switch):",
                parent=self.root,
            )
            if not column_name:
                return
            key = column_name.strip().lower().replace(" ", "_")

        if not key:
            messagebox.showwarning("Advertencia", "El nombre de columna no es valido.")
            return

        rois_path = "rois.json"
        existing_rois = {}
        if os.path.exists(rois_path):
            try:
                with open(rois_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if isinstance(loaded, dict):
                    existing_rois = loaded
            except Exception as exc:
                messagebox.showerror("Error", f"No se pudo leer rois.json:\n{exc}")
                return

        if key in existing_rois:
            overwrite = messagebox.askyesno(
                "Columna existente",
                f"La columna '{key}' ya tiene una ROI. Deseas reemplazarla?",
            )
            if not overwrite:
                return

        self.log_message(f"[*] Capturando ventana para definir ROI de '{key}'...")
        img = capturar_ventana()
        if img is None:
            messagebox.showerror("Error", "No se pudo capturar la ventana del juego.")
            return

        try:
            from PIL import Image
            from select_rois_tk import ROISelector
        except Exception as exc:
            messagebox.showerror("Error", f"Faltan dependencias para selector de ROI:\n{exc}")
            return

        img_rgb = Image.fromarray(img[:, :, ::-1])
        selector = ROISelector(img_rgb, labels=[key], parent=self.root)
        new_rois = selector.run()

        if not new_rois or key not in new_rois:
            self.log_message("[WARN] No se guardo ninguna ROI nueva.")
            return

        existing_rois[key] = new_rois[key]
        try:
            with open(rois_path, "w", encoding="utf-8") as f:
                json.dump(existing_rois, f, indent=2, ensure_ascii=False)
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo guardar rois.json:\n{exc}")
            return

        self.log_message(f"[OK] ROI guardada para la columna '{key}' en {rois_path}")
        messagebox.showinfo("ROI guardada", f"Se guardo la ROI de la columna '{key}'.")

    def on_closing(self):
        if self.processor_running or not self.capture_queue.empty():
            should_close = messagebox.askyesno(
                "Salir",
                "Hay tareas en proceso o pendientes. Desea detener y salir?",
            )
            if not should_close:
                return
            self.processor_running = False
        try:
            if self.hotkeys_registered:
                keyboard.clear_all_hotkeys()
        except Exception:
            pass

        self.root.quit()


def main():
    root = tk.Tk()
    OCRMainGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
