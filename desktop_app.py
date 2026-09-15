import os
import sys
import threading
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import slide_manager

class PowerPointAIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI PowerPoint Precision Converter v2.5")
        self.root.geometry("680x560")
        self.root.minsize(600, 500)
        self.root.configure(bg="#F8FAFC")
        
        self.available_folders = slide_manager.get_available_folders()
        self.folder_names = list(self.available_folders.keys())
        
        self.setup_ui()
        
    def setup_ui(self):
        # Header
        header_frame = tk.Frame(self.root, bg="#0A3B75", height=85)
        header_frame.pack(fill="x", side="top")
        
        lbl_title = tk.Label(
            header_frame, 
            text="🎯 AI PowerPoint Precision Converter", 
            font=("Segoe UI", 16, "bold"), 
            fg="#FFFFFF", 
            bg="#0A3B75"
        )
        lbl_title.pack(pady=(15, 2))
        
        lbl_sub = tk.Label(
            header_frame, 
            text="Slaydlarni toza oq listdan boshlab professional PPTX ga aylantirish", 
            font=("Segoe UI", 9), 
            fg="#CBD5E1", 
            bg="#0A3B75"
        )
        lbl_sub.pack(pady=(0, 15))
        
        # Body
        body_frame = tk.Frame(self.root, bg="#F8FAFC", padx=25, pady=16)
        body_frame.pack(fill="both", expand=True)
        
        # Preset selector
        lbl_preset = tk.Label(body_frame, text="1. Taqdimotni Tanlang (Google Drive / Desktop):", font=("Segoe UI", 10, "bold"), fg="#1E293B", bg="#F8FAFC")
        lbl_preset.pack(anchor="w")
        
        self.preset_var = tk.StringVar(value=self.folder_names[0] if self.folder_names else "")
        cb_preset = ttk.Combobox(body_frame, textvariable=self.preset_var, values=self.folder_names, state="readonly", font=("Segoe UI", 10))
        cb_preset.pack(fill="x", pady=(4, 12))
        cb_preset.bind("<<ComboboxSelected>>", self.on_preset_changed)
        
        # Custom Folder
        lbl_folder = tk.Label(body_frame, text="yoki Boshqa Papka Manzilini Kiriting:", font=("Segoe UI", 10, "bold"), fg="#1E293B", bg="#F8FAFC")
        lbl_folder.pack(anchor="w")
        
        folder_box = tk.Frame(body_frame, bg="#F8FAFC")
        folder_box.pack(fill="x", pady=(4, 15))
        
        initial_path = self.available_folders.get(self.preset_var.get(), r"C:\Users\user\Desktop\5M_Of_Advertising")
        self.folder_var = tk.StringVar(value=initial_path)
        self.entry_folder = tk.Entry(folder_box, textvariable=self.folder_var, font=("Segoe UI", 10), bg="#FFFFFF", relief="solid", bd=1)
        self.entry_folder.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 8))
        
        btn_browse = tk.Button(
            folder_box, 
            text="📁 Tanlash...", 
            font=("Segoe UI", 9, "bold"), 
            bg="#E2E8F0", 
            fg="#1E293B", 
            relief="flat", 
            padx=12,
            command=self.browse_folder
        )
        btn_browse.pack(side="right")
        
        # Action Card
        card_frame = tk.Frame(body_frame, bg="#FFFFFF", relief="solid", bd=1, padx=20, pady=16)
        card_frame.pack(fill="both", expand=True, pady=4)
        
        # 1. Main Action: Direct Convert
        btn_convert = tk.Button(
            card_frame, 
            text="🚀 Taqdimotni Noldan Yaratish (Generate PPTX)", 
            font=("Segoe UI", 11, "bold"), 
            bg="#0A3B75", 
            fg="#FFFFFF", 
            activebackground="#F15A24",
            activeforeground="#FFFFFF",
            relief="flat", 
            cursor="hand2",
            pady=10,
            command=self.run_conversion_thread
        )
        btn_convert.pack(fill="x", pady=(0, 8))
        
        # 2. Web UI Launcher Action
        btn_web = tk.Button(
            card_frame, 
            text="🌐 Veb-Dashboardni Ochish (Slaydlarni Taqqoslash)", 
            font=("Segoe UI", 10, "bold"), 
            bg="#F15A24", 
            fg="#FFFFFF", 
            activebackground="#D03801",
            activeforeground="#FFFFFF",
            relief="flat", 
            cursor="hand2",
            pady=8,
            command=self.launch_web_app
        )
        btn_web.pack(fill="x", pady=(0, 8))
        
        # 3. Open in PowerPoint
        btn_ppt = tk.Button(
            card_frame, 
            text="📂 PowerPoint Dasturida Ochish", 
            font=("Segoe UI", 9, "bold"), 
            bg="#E2E8F0", 
            fg="#0A3B75", 
            relief="flat", 
            cursor="hand2",
            pady=6,
            command=self.open_powerpoint
        )
        btn_ppt.pack(fill="x")
        
        # Status Bar
        self.lbl_status = tk.Label(
            self.root, 
            text="Holat: Tayyor", 
            font=("Segoe UI", 9), 
            fg="#64748B", 
            bg="#E2E8F0", 
            anchor="w", 
            padx=15, 
            pady=6
        )
        self.lbl_status.pack(fill="x", side="bottom")

    def on_preset_changed(self, event=None):
        name = self.preset_var.get()
        if name in self.available_folders:
            self.folder_var.set(self.available_folders[name])

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder_var.get())
        if folder:
            self.folder_var.set(folder)

    def browse_images(self):
        files = filedialog.askopenfilenames(
            title="Slayd Rasmlarini Tanlang",
            filetypes=[("Rasm Fayllari", "*.png;*.jpg;*.jpeg;*.webp")]
        )
        if files:
            first_dir = os.path.dirname(files[0])
            self.folder_var.set(first_dir)
            self.set_status(f"{len(files)} ta rasm tanlandi.", color="#0A3B75")

    def set_status(self, text, color="#64748B"):
        self.lbl_status.config(text=f"Holat: {text}", fg=color)

    def run_conversion_thread(self):
        t = threading.Thread(target=self.run_conversion)
        t.daemon = True
        t.start()

    def run_conversion(self):
        folder = self.folder_var.get()
        if not os.path.exists(folder):
            messagebox.showerror("Xatolik", f"Papka topilmadi:\n{folder}")
            return
            
        self.set_status("Taqdimot yaratilmoqda, kuting...", color="#0A3B75")
        
        success, msg = slide_manager.convert_presentation(folder)
        if success:
            self.set_status("✅ Taqdimot 100% muvaffaqiyatli yaratildi!", color="#16A34A")
            if messagebox.askyesno("Tayyor!", "Taqdimot muvaffaqiyatli yaratildi!\n\nUni hoziroq PowerPoint'da ochishni xohlaysizmi?"):
                self.open_powerpoint()
        else:
            self.set_status("❌ Xatolik yuz berdi!", color="#DC2626")
            messagebox.showerror("Xatolik", msg)

    def launch_web_app(self):
        self.set_status("Veb-ilova brauzerda ochilmoqda...", color="#0A3B75")
        subprocess.Popen([sys.executable, "-m", "uvicorn", "backend.main:app", "--port", "8000"], creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        webbrowser.open("http://localhost:8000")

    def open_powerpoint(self):
        folder = self.folder_var.get()
        pptx_path = slide_manager.get_folder_pptx_path(folder)
        if pptx_path and os.path.exists(pptx_path):
            try:
                os.startfile(pptx_path)
                self.set_status("PowerPoint dasturida ochildi.", color="#16A34A")
            except Exception as e:
                messagebox.showerror("Xatolik", f"PowerPoint ochilmadi: {e}")
        else:
            messagebox.showwarning("Ogohlantirish", "Avval 'Taqdimotni Noldan Yaratish' tugmasini bosing!")

if __name__ == "__main__":
    root = tk.Tk()
    app = PowerPointAIApp(root)
    root.mainloop()
