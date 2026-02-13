import json
import os
import sys
import math
import traceback
from ocr_market import capturar_ventana, obtener_ventana_juego

try:
    import tkinter as tk
    from PIL import Image, ImageTk
except Exception as e:
    print('Tkinter/Pillow no disponibles:', e)
    print('Instalá Pillow (`pip install pillow`) y asegurate que Python tenga soporte Tk.')
    sys.exit(1)


class ROISelector:
    def __init__(self, pil_image, labels=['item','price','sales']):
        self.root = tk.Tk()
        self.root.title('ROI Selector')
        self.labels = labels
        self.current = 0
        self.pil_image = pil_image
        self.img_w, self.img_h = pil_image.size

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        # scale down only if excessively larger than screen; otherwise keep full size and enable scroll
        max_w = screen_w - 100
        max_h = screen_h - 150
        if self.img_w > max_w or self.img_h > max_h:
            scale = min(max_w / self.img_w, max_h / self.img_h)
        else:
            scale = 1.0
        self.scale = scale
        disp_w = int(self.img_w * scale)
        disp_h = int(self.img_h * scale)
        self.display_img = pil_image.resize((disp_w, disp_h), Image.LANCZOS)
        self.tkimg = ImageTk.PhotoImage(self.display_img)

        # canvas with scrollbars to allow viewing full image
        frame = tk.Frame(self.root)
        frame.pack(fill='both', expand=True)
        hbar = tk.Scrollbar(frame, orient='horizontal')
        vbar = tk.Scrollbar(frame, orient='vertical')
        self.canvas = tk.Canvas(frame, width=min(disp_w, max_w), height=min(disp_h, max_h), xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        hbar.config(command=self.canvas.xview)
        vbar.config(command=self.canvas.yview)
        hbar.pack(side='bottom', fill='x')
        vbar.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)
        self.canvas.config(scrollregion=(0,0,disp_w,disp_h))
        self.canvas_img = self.canvas.create_image(0,0,anchor='nw',image=self.tkimg)

        self.label_var = tk.StringVar()
        self.label_var.set(f'Select column: {self.labels[self.current]} (drag rectangle)')
        self.lbl = tk.Label(self.root, textvariable=self.label_var)
        self.lbl.pack(fill='x')

        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack(fill='x')
        self.ok_btn = tk.Button(self.btn_frame, text='Confirm and Next', command=self.confirm)
        self.ok_btn.pack(side='left')
        self.cancel_btn = tk.Button(self.btn_frame, text='Cancel', command=self.cancel)
        self.cancel_btn.pack(side='right')

        self.rect = None
        self.start_x = None
        self.start_y = None
        self.rois = {}

        self.canvas.bind('<ButtonPress-1>', self.on_button_press)
        self.canvas.bind('<B1-Motion>', self.on_move_press)
        self.canvas.bind('<ButtonRelease-1>', self.on_button_release)

    def on_button_press(self, event):
        # convert event coords to canvas coords (account for scrolling)
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        if self.rect:
            self.canvas.delete(self.rect)
            self.rect = None

    def on_move_press(self, event):
        curX, curY = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if self.rect:
            self.canvas.coords(self.rect, self.start_x, self.start_y, curX, curY)
        else:
            self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, curX, curY, outline='red', width=2)

    def on_button_release(self, event):
        # finalize
        pass

    def confirm(self):
        if not self.rect:
            tk.messagebox.showwarning('No selection','Draw a rectangle before confirming')
            return
        x0, y0, x1, y1 = [float(v) for v in self.canvas.coords(self.rect)]
        # normalize and map back to original image coords
        if x1 < x0:
            x0, x1 = x1, x0
        if y1 < y0:
            y0, y1 = y1, y0
        x0_img = x0 / self.scale
        x1_img = x1 / self.scale
        y0_img = y0 / self.scale
        y1_img = y1 / self.scale
        # store relative bounds (0..1) for both X and Y
        rel_left = max(0.0, min(1.0, x0_img / self.img_w))
        rel_right = max(0.0, min(1.0, x1_img / self.img_w))
        rel_top = max(0.0, min(1.0, y0_img / self.img_h))
        rel_bottom = max(0.0, min(1.0, y1_img / self.img_h))
        label = self.labels[self.current]
        self.rois[label] = {'x': [rel_left, rel_right], 'y': [rel_top, rel_bottom]}
        self.current += 1
        if self.current >= len(self.labels):
            self.root.quit()
            return
        self.label_var.set(f'Select column: {self.labels[self.current]} (drag rectangle)')
        # remove previous rect and continue
        if self.rect:
            self.canvas.delete(self.rect)
            self.rect = None

    def cancel(self):
        self.rois = {}
        self.root.quit()

    def run(self):
        self.root.mainloop()
        return self.rois


def main():
    print('Capturando ventana del juego...')
    img = capturar_ventana()
    if img is None:
        print('No se pudo capturar la ventana del juego. Asegurate de que esté visible.')
        return
    # convert BGR numpy to PIL Image RGB
    from PIL import Image
    img_rgb = Image.fromarray(img[:,:,::-1])
    sel = ROISelector(img_rgb)
    rois = sel.run()
    if not rois:
        print('No se guardaron ROIs.')
        return
    out = 'rois.json'
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(rois, f, indent=2, ensure_ascii=False)
    print('ROIs guardadas en', out)
    print(rois)

if __name__ == '__main__':
    try:
        main()
    except Exception:
        traceback.print_exc()
