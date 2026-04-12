import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import json
import csv
import os
import threading
import time
import sys
import platform

ver = "0.1.3"
build_ver = 4

# Kiểm tra modules
try:
    from plyer import notification
    PLYER_AVAILABLE = True
except ImportError:
    PLYER_AVAILABLE = False
    print("Cài đặt plyer để có thông báo: pip install plyer")

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Windows auto-start (optional)
try:
    import winshell
    from win32com.client import Dispatch
    WINSHELL_AVAILABLE = True
except ImportError:
    WINSHELL_AVAILABLE = False

class TaskManager:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Task Manager ({ver})")
        self.root.geometry("1100x750")
        
        # Dữ liệu tasks
        self.tasks = []
        self.current_filter = "Tất cả"
        self.sort_by = "priority"
        self.sort_order = "asc"
        self.data_file = "checklist_data.json"
        self.backup_folder = "backups"
        self.auto_save = True
        self.auto_backup = True
        self.last_save_time = datetime.now()
        
        # Cấu hình dark mode, thông báo, auto-start
        self.config_file = "config.json"
        self.load_config()
        
        # Tạo thư mục backup
        if not os.path.exists(self.backup_folder):
            os.makedirs(self.backup_folder)
        
        # Màu sắc ưu tiên (chế độ sáng)
        self.priority_colors = {
            "Cao": "#FF6161",
            "Trung bình": "#FFCA81", 
            "Thấp": "#8AFF8A"
        }
        
        # Load dữ liệu công việc
        self.load_data()
        
        # Khởi tạo categories
        self.categories = ["Công việc", "Học tập", "Cá nhân", "Sức khỏe", "Khác"]
        
        # Bắt đầu các thread nền
        self.start_background_tasks()
        
        # Tạo giao diện
        self.setup_ui()
        
        # Áp dụng theme (sáng/tối)
        self.apply_theme()
        
        # Refresh hiển thị
        self.refresh_task_list()
        
        # Bind sự kiện đóng cửa sổ
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def do_nothing(self):
        """Hàm này thực sự không làm gì cả"""
        return
    
    # ---------- Cấu hình ----------
    def load_config(self):
        self.dark_mode = False
        self.background_notifications = True
        self.auto_start = False
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.dark_mode = config.get('dark_mode', False)
                    self.background_notifications = config.get('background_notifications', True)
                    self.auto_start = config.get('auto_start', False)
            except:
                pass
    
    def save_config(self):
        config = {
            'dark_mode': self.dark_mode,
            'background_notifications': self.background_notifications,
            'auto_start': self.auto_start
        }
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
    
    # ---------- Dark Mode ----------
    def apply_theme(self):
        if self.dark_mode:
            bg_color = "#2b2b2b"
            fg_color = "#ffffff"
            select_bg = "#404040"
            tree_bg = "#1e1e1e"
            tree_fg = "#e0e0e0"
            heading_bg = "#3c3c3c"
            heading_fg = "#ffffff"
            entry_bg = "#3c3c3c"
            entry_fg = "#ffffff"
            tab_bg = "#2b2b2b"
            tab_fg = "#ffffff"
        else:
            bg_color = "#f0f0f0"
            fg_color = "#000000"
            select_bg = "#0078d7"
            tree_bg = "#ffffff"
            tree_fg = "#000000"
            heading_bg = "#d9d9d9"
            heading_fg = "#000000"
            entry_bg = "#ffffff"
            entry_fg = "#000000"
            tab_bg = "#f0f0f0"
            tab_fg = "#000000"
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(".", background=bg_color, foreground=fg_color)
        style.configure("TFrame", background=bg_color)
        style.configure("TLabel", background=bg_color, foreground=fg_color)
        style.configure("TLabelframe", background=bg_color, foreground=fg_color)
        style.configure("TLabelframe.Label", background=bg_color, foreground=fg_color)
        style.configure("TButton", background=entry_bg, foreground=fg_color)
        style.map("TButton", background=[('active', select_bg)])
        style.configure("TEntry", fieldbackground=entry_bg, foreground=entry_fg)
        style.configure("TCombobox", fieldbackground=entry_bg, foreground=entry_fg)
        style.configure("Treeview", background=tree_bg, foreground=tree_fg, fieldbackground=tree_bg)
        style.configure("Treeview.Heading", background=heading_bg, foreground=heading_fg)
        style.map("Treeview", background=[('selected', select_bg)])
        style.configure("TNotebook", background=tab_bg)
        style.configure("TNotebook.Tab", background=tab_bg, foreground=tab_fg)
        
        self.root.configure(bg=bg_color)
        
        if hasattr(self, 'stats_text'):
            if self.dark_mode:
                self.stats_text.configure(bg="#1e1e1e", fg="#e0e0e0", insertbackground="white")
            else:
                self.stats_text.configure(bg="white", fg="black", insertbackground="black")
        
        if hasattr(self, 'backup_listbox'):
            if self.dark_mode:
                self.backup_listbox.configure(bg="#3c3c3c", fg="white", selectbackground="#404040")
            else:
                self.backup_listbox.configure(bg="white", fg="black", selectbackground="#0078d7")
        
        if hasattr(self, 'dark_mode_btn'):
            self.dark_mode_btn.config(text="🌙 Dark Mode" if not self.dark_mode else "☀️ Light Mode")
    
    def toggle_dark_mode(self):
        self.dark_mode = not self.dark_mode
        self.apply_theme()
        self.save_config()
        self.refresh_task_list()
    
    # ---------- Giao diện ----------
    def setup_ui(self):
        # Notebook (tab)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Các tab
        self.main_tab = ttk.Frame(self.notebook)
        self.io_tab = ttk.Frame(self.notebook)
        self.stats_tab = ttk.Frame(self.notebook)
        self.settings_tab = ttk.Frame(self.notebook)
        self.infomation_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.main_tab, text="📋 Quản lý công việc")
        self.notebook.add(self.io_tab, text="📁 Import/Export")
        self.notebook.add(self.stats_tab, text="📊 Thống kê & Báo cáo")
        self.notebook.add(self.settings_tab, text="⚙️ Cài đặt")
        self.notebook.add(self.infomation_tab, text="🗒️ Credit")
        
        self.setup_main_tab()
        self.setup_io_tab()
        self.setup_stats_tab()
        self.setup_settings_tab()
        self.setup_infomation_tab()
    
    def setup_main_tab(self):
        # Thanh trên cùng
        top_frame = ttk.Frame(self.main_tab)
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(top_frame, text="💾 Lưu ngay", command=self.manual_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="🔄 Backup ngay", command=self.manual_backup).pack(side=tk.LEFT, padx=5)
        self.auto_save_label = ttk.Label(top_frame, text="✅ Auto-save: BẬT", foreground="green")
        self.auto_save_label.pack(side=tk.LEFT, padx=20)
        
        # Khung nhập liệu
        input_frame = ttk.LabelFrame(self.main_tab, text="➕ Thêm công việc mới", padding="15")
        input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(input_frame, text="Tên công việc:*").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.task_name_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.task_name_var, width=40).grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Mô tả:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.task_desc_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.task_desc_var, width=40).grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Danh mục:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_var = tk.StringVar(value="Công việc")
        ttk.Combobox(input_frame, textvariable=self.category_var, values=self.categories, width=15).grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Ưu tiên:").grid(row=2, column=2, sticky=tk.W, padx=5, pady=5)
        self.priority_var = tk.StringVar(value="Trung bình")
        ttk.Combobox(input_frame, textvariable=self.priority_var, values=["Cao", "Trung bình", "Thấp"], width=15).grid(row=2, column=3, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Deadline:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.deadline_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.deadline_var, width=20).grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(input_frame, text="(HH:MM DD/MM/YYYY)", font=('Arial', 8)).grid(row=3, column=2, sticky=tk.W)
        
        ttk.Label(input_frame, text="Tags:").grid(row=3, column=3, sticky=tk.W, padx=5, pady=5)
        self.tags_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.tags_var, width=20).grid(row=3, column=4, sticky=tk.W, padx=5, pady=5)
        
        ttk.Button(input_frame, text="➕ Thêm công việc", command=self.add_task).grid(row=4, column=0, columnspan=5, pady=15)
        
        # Khung lọc và sắp xếp
        control_frame = ttk.Frame(self.main_tab)
        control_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(control_frame, text="🔍 Lọc theo:").pack(side=tk.LEFT, padx=5)
        self.filter_var = tk.StringVar(value="Tất cả")
        filter_combo = ttk.Combobox(control_frame, textvariable=self.filter_var,
                                   values=["Tất cả", "Đang thực hiện", "Hoàn thành", "Quá hạn", "Theo danh mục"],
                                   width=18, state="readonly")
        filter_combo.pack(side=tk.LEFT, padx=5)
        filter_combo.bind('<<ComboboxSelected>>', lambda e: self.on_filter_change())
        
        self.category_filter_var = tk.StringVar(value="Công việc")
        self.category_filter_combo = ttk.Combobox(control_frame, textvariable=self.category_filter_var,
                                                values=self.categories, width=15, state="disabled")
        self.category_filter_combo.pack(side=tk.LEFT, padx=5)
        self.category_filter_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_task_list())
        
        ttk.Label(control_frame, text="📊 Sắp xếp theo:").pack(side=tk.LEFT, padx=5)
        self.sort_var = tk.StringVar(value="priority")
        sort_combo = ttk.Combobox(control_frame, textvariable=self.sort_var,
                                 values=["priority", "deadline", "name", "category", "status"],
                                 width=15, state="readonly")
        sort_combo.pack(side=tk.LEFT, padx=5)
        sort_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_task_list())
        
        self.sort_order_btn = ttk.Button(control_frame, text="⬆️ Tăng dần", width=10, command=self.toggle_sort_order)
        self.sort_order_btn.pack(side=tk.LEFT, padx=5)
        
        # Bảng hiển thị công việc
        list_frame = ttk.LabelFrame(self.main_tab, text="📋 Danh sách công việc", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ("status", "name", "category", "priority", "deadline", "tags", "progress", "actions")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        self.tree.heading("status", text="Đã xong")
        self.tree.heading("name", text="Tên công việc")
        self.tree.heading("category", text="Danh mục")
        self.tree.heading("priority", text="Ưu tiên")
        self.tree.heading("deadline", text="Deadline")
        self.tree.heading("tags", text="Tags")
        self.tree.heading("progress", text="Tiến độ")
        self.tree.heading("actions", text="Thao tác")
        
        self.tree.column("status", width=50, anchor="center")
        self.tree.column("name", width=250)
        self.tree.column("category", width=100, anchor="center")
        self.tree.column("priority", width=80, anchor="center")
        self.tree.column("deadline", width=130, anchor="center")
        self.tree.column("tags", width=120, anchor="center")
        self.tree.column("progress", width=100, anchor="center")
        self.tree.column("actions", width=80, anchor="center")
        
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        self.tree.bind('<ButtonRelease-1>', self.on_tree_click)
    
    def setup_io_tab(self):
        main_frame = ttk.Frame(self.io_tab, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Export
        export_frame = ttk.LabelFrame(main_frame, text="📤 Export dữ liệu", padding="15")
        export_frame.pack(fill=tk.X, pady=10)
        ttk.Label(export_frame, text="Chọn định dạng xuất:").pack(anchor=tk.W, pady=5)
        format_frame = ttk.Frame(export_frame)
        format_frame.pack(fill=tk.X, pady=5)
        self.export_format = tk.StringVar(value="json")
        ttk.Radiobutton(format_frame, text="JSON", variable=self.export_format, value="json").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(format_frame, text="CSV", variable=self.export_format, value="csv").pack(side=tk.LEFT, padx=10)
        if PANDAS_AVAILABLE:
            ttk.Radiobutton(format_frame, text="Excel", variable=self.export_format, value="excel").pack(side=tk.LEFT, padx=10)
        if REPORTLAB_AVAILABLE:
            ttk.Radiobutton(format_frame, text="PDF", variable=self.export_format, value="pdf").pack(side=tk.LEFT, padx=10)
        ttk.Button(export_frame, text="💾 Export dữ liệu", command=self.export_data, width=20).pack(pady=10)
        
        # Import
        import_frame = ttk.LabelFrame(main_frame, text="📥 Import dữ liệu", padding="15")
        import_frame.pack(fill=tk.X, pady=10)
        ttk.Label(import_frame, text="Chọn file để import:").pack(anchor=tk.W, pady=5)
        ttk.Button(import_frame, text="📂 Import từ file", command=self.import_data, width=20).pack(pady=5)
        ttk.Label(import_frame, text="⚠️ Lưu ý: Import sẽ thêm dữ liệu mới vào danh sách hiện tại", foreground="orange").pack(anchor=tk.W, pady=5)
        
        # Backup/Restore
        backup_frame = ttk.LabelFrame(main_frame, text="🔄 Backup & Restore", padding="15")
        backup_frame.pack(fill=tk.X, pady=10)
        btn_frame = ttk.Frame(backup_frame)
        btn_frame.pack()
        ttk.Button(btn_frame, text="💾 Tạo Backup", command=self.manual_backup, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📂 Khôi phục", command=self.restore_from_backup, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(backup_frame, text="Danh sách backups:").pack(anchor=tk.W, pady=(10,5))
        self.backup_listbox = tk.Listbox(backup_frame, height=5)
        self.backup_listbox.pack(fill=tk.X, pady=5)
        self.refresh_backup_list()
    
    def setup_stats_tab(self):
        stats_display = ttk.Frame(self.stats_tab, padding="20")
        stats_display.pack(fill=tk.BOTH, expand=True)
        self.stats_text = tk.Text(stats_display, height=20, width=80)
        self.stats_text.pack(fill=tk.BOTH, expand=True)
        ttk.Button(stats_display, text="🔄 Làm mới thống kê", command=self.update_detailed_stats).pack(pady=10)
        if REPORTLAB_AVAILABLE:
            ttk.Button(stats_display, text="📊 Xuất báo cáo PDF", command=self.export_pdf).pack(pady=5)
        self.update_detailed_stats()
    
    def setup_settings_tab(self):
        settings_frame = ttk.Frame(self.settings_tab, padding="20")
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # Auto-save
        auto_save_frame = ttk.LabelFrame(settings_frame, text="⚙️ Cài đặt tự động", padding="15")
        auto_save_frame.pack(fill=tk.X, pady=10)
        self.auto_save_var = tk.BooleanVar(value=self.auto_save)
        ttk.Checkbutton(auto_save_frame, text="Bật Auto-save (lưu sau mỗi 30 giây)", variable=self.auto_save_var, command=self.toggle_auto_save).pack(anchor=tk.W)
        self.auto_backup_var = tk.BooleanVar(value=self.auto_backup)
        ttk.Checkbutton(auto_save_frame, text="Bật Auto-backup (backup mỗi giờ)", variable=self.auto_backup_var, command=self.toggle_auto_backup).pack(anchor=tk.W, pady=5)
        
        # Notifications background
        notif_bg_frame = ttk.LabelFrame(settings_frame, text="🔔 Thông báo nền", padding="15")
        notif_bg_frame.pack(fill=tk.X, pady=10)
        self.notif_bg_var = tk.BooleanVar(value=self.background_notifications)
        ttk.Checkbutton(notif_bg_frame, text="Bật kiểm tra deadline và gửi thông báo (chạy nền)",
                        variable=self.notif_bg_var, command=self.toggle_background_notifications).pack(anchor=tk.W)
        ttk.Label(notif_bg_frame, text="⚠️ Cần thư viện plyer: pip install plyer", font=('Arial', 8), foreground="orange").pack(anchor=tk.W)
        
        # Auto start with system
        autostart_frame = ttk.LabelFrame(settings_frame, text="🚀 Tự động khởi động", padding="15")
        autostart_frame.pack(fill=tk.X, pady=10)
        self.auto_start_var = tk.BooleanVar(value=self.auto_start)
        ttk.Checkbutton(autostart_frame, text="Tự động chạy Checklist Pro khi khởi động máy tính",
                        variable=self.auto_start_var, command=self.toggle_auto_start).pack(anchor=tk.W)
        ttk.Label(autostart_frame, text="⚠️ Cần quyền admin khi bật (Windows) hoặc tự động tạo file .desktop (Linux)", 
                  foreground="orange", font=('Arial', 8)).pack(anchor=tk.W, pady=5)
        
        # Dark mode
        theme_frame = ttk.LabelFrame(settings_frame, text="🎨 Giao diện (thử nghiệm)", padding="15")
        theme_frame.pack(fill=tk.X, pady=10)
        self.dark_mode_btn = ttk.Button(theme_frame, text="🌙 Dark Mode" if not self.dark_mode else "☀️ Light Mode",
                                       command=self.toggle_dark_mode, width=20)
        self.dark_mode_btn.pack(pady=5)
        
        # Data management
        data_frame = ttk.LabelFrame(settings_frame, text="🗄️ Quản lý dữ liệu", padding="15")
        data_frame.pack(fill=tk.X, pady=10)
        ttk.Button(data_frame, text="🗑️ Xóa tất cả dữ liệu", command=self.clear_all_data, width=20).pack(pady=5)
        ttk.Button(data_frame, text="📁 Mở thư mục dữ liệu", command=self.open_data_folder, width=20).pack(pady=5)
    
    def setup_infomation_tab(self):
        deffont = ('Arial', 10)
        credit_frame = ttk.Frame(self.infomation_tab, padding="30")
        credit_frame.pack(fill=tk.BOTH, expand=True)
        
        # Khung thông tin
        info_frame = ttk.LabelFrame(credit_frame, text="📋 Thông tin ứng dụng", padding="15")
        info_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(info_frame, text=f"Phiên bản: {ver} (build {build_ver})", font=deffont).pack(anchor=tk.W, pady=2)
        ttk.Label(info_frame, text=f"Thư mục backup: {self.backup_folder}", font=deffont).pack(anchor=tk.W, pady=2)
        ttk.Label(info_frame, text=f"Mã nguồn app: github.com/loivn-bl/taskmanager", font=deffont).pack(anchor=tk.W, pady=2)
        
        # Khung tác giả
        author_frame = ttk.LabelFrame(credit_frame, text="👨 Tác giả", padding="15")
        author_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(author_frame, text="Phát triển bởi: Lưu Gia Lợi", font=deffont).pack(anchor=tk.W, pady=2)
        ttk.Label(author_frame, text="Email: luugialoi37@gmail.com", font=deffont).pack(anchor=tk.W, pady=2)
        
        # Khung thư viện
        module_frame = ttk.LabelFrame(credit_frame, text="📑 Thư viện", padding="15")
        module_frame.pack(fill=tk.X, pady=10)

        module_text = """Ứng dụng được xây dựng với các thư viện mã nguồn mở:
        • Python, Tkinter
        • plyer (thông báo desktop)
        • pandas, openpyxl (xuất Excel)
        • reportlab (xuất PDF)"""

        module_label = ttk.Label(module_frame, text=module_text, wraplength=600, justify=tk.LEFT)
        module_label.pack(anchor=tk.W)

        # Khung cảm ơn
        thanks_frame = ttk.LabelFrame(credit_frame, text="❤️ Lời cảm ơn", padding="15")
        thanks_frame.pack(fill=tk.X, pady=10)
        
        thanks_text = """Cảm ơn bạn đã sử dụng ứng dụng Task Manager!
Mọi đóng góp ý kiến xin gửi về email/link Github trên.
Chúc bạn quản lý công việc hiệu quả!"""
        
        thanks_label = ttk.Label(thanks_frame, text=thanks_text, wraplength=600, justify=tk.LEFT)
        thanks_label.pack(anchor=tk.W)

    # ---------- Nhiệm vụ chính ----------
    def add_task(self):
        name = self.task_name_var.get().strip()
        if not name:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập tên công việc!")
            return
        
        description = self.task_desc_var.get().strip()
        category = self.category_var.get()
        priority = self.priority_var.get()
        tags = self.tags_var.get().strip()
        deadline_str = self.deadline_var.get().strip()
        
        if deadline_str:
            try:
                deadline = datetime.strptime(deadline_str, "%H:%M %d/%m/%Y")
                if deadline < datetime.now():
                    messagebox.showwarning("Cảnh báo", "Deadline không thể ở quá khứ!")
                    return
            except ValueError:
                messagebox.showwarning("Cảnh báo", "Định dạng deadline không đúng!\nVui lòng dùng: HH:MM DD/MM/YYYY")
                return
        
        task = {
            "id": len(self.tasks) + 1,
            "name": name,
            "description": description,
            "category": category,
            "priority": priority,
            "deadline": deadline_str,
            "tags": tags,
            "completed": False,
            "progress": 0,
            "created_at": datetime.now().strftime("%H:%M %d/%m/%Y"),
            "updated_at": datetime.now().strftime("%H:%M %d/%m/%Y")
        }
        self.tasks.append(task)
        self.save_data()
        
        self.task_name_var.set("")
        self.task_desc_var.set("")
        self.deadline_var.set("")
        self.tags_var.set("")
        self.refresh_task_list()
        messagebox.showinfo("Thành công", f"Đã thêm công việc: {name}")
    
    def get_task_deadline_status(self, task):
        """Trả về trạng thái deadline: 'normal', 'approaching', 'overdue'"""
        if task["completed"]:
            return "normal"
        if not task.get("deadline"):
            return "normal"
        try:
            deadline = datetime.strptime(task["deadline"], "%H:%M %d/%m/%Y")
            now = datetime.now()
            if deadline < now:
                return "overdue"
            elif (deadline - now).total_seconds() <= 86400 * 2:  # 2 ngày
                return "approaching"
            else:
                return "normal"
        except:
            return "normal"
    
    def refresh_task_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        filtered = self.filter_tasks()
        filtered = self.sort_tasks(filtered)
        
        # Màu priority cho dark mode
        if self.dark_mode:
            priority_colors = {"Cao": "#5c2e2e", "Trung bình": "#5c4a2e", "Thấp": "#2e5c3e"}
        else:
            priority_colors = self.priority_colors
        
        for task in filtered:
            status, is_overdue = self.get_task_status(task)
            deadline_status = self.get_task_deadline_status(task)
            
            # Chọn icon cho cột status dựa trên deadline_status
            if task["completed"]:
                status_icon = "✅"
            elif deadline_status == "overdue":
                status_icon = "❌"  # quá hạn
            elif deadline_status == "approaching":
                status_icon = "⚠️"  # gần hạn
            else:
                status_icon = "⬜"
            
            item_id = self.tree.insert("", tk.END, values=(
                status_icon,
                task["name"],
                task["category"],
                task["priority"],
                task["deadline"] if task["deadline"] else "Không có",
                task.get("tags", ""),
                f"{task.get('progress', 0)}%",
                "⚙️"
            ))
            color = priority_colors.get(task["priority"], "#3c3c3c" if self.dark_mode else "#FFFFFF")
            self.tree.tag_configure(f"color_{task['id']}", background=color)
            self.tree.item(item_id, tags=(f"color_{task['id']}",))
            
            # Thêm màu nền riêng cho trạng thái deadline
            if not task["completed"]:
                if deadline_status == "overdue":
                    overdue_color = "#8b3a3a" if self.dark_mode else "#FFCCCC"
                    self.tree.tag_configure(f"deadline_{task['id']}", background=overdue_color)
                    self.tree.item(item_id, tags=(f"deadline_{task['id']}",))
                elif deadline_status == "approaching":
                    approaching_color = "#8b6b3a" if self.dark_mode else "#FFFFCC"
                    self.tree.tag_configure(f"deadline_{task['id']}", background=approaching_color)
                    self.tree.item(item_id, tags=(f"deadline_{task['id']}",))
    
    def get_task_status(self, task):
        if task["completed"]:
            return "✅", False
        if task["deadline"]:
            try:
                deadline = datetime.strptime(task["deadline"], "%H:%M %d/%m/%Y")
                if deadline < datetime.now():
                    return "⚠️", True
            except:
                pass
        return "⬜", False
    
    def filter_tasks(self):
        filter_type = self.filter_var.get()
        if filter_type == "Đang thực hiện":
            return [t for t in self.tasks if not t["completed"]]
        elif filter_type == "Hoàn thành":
            return [t for t in self.tasks if t["completed"]]
        elif filter_type == "Quá hạn":
            return [t for t in self.tasks if not t["completed"] and self.get_task_deadline_status(t) == "overdue"]
        elif filter_type == "Theo danh mục":
            cat = self.category_filter_var.get()
            return [t for t in self.tasks if t["category"] == cat]
        else:
            return self.tasks.copy()
    
    def sort_tasks(self, tasks):
        key = self.sort_var.get()
        reverse = (self.sort_order == "desc")
        if key == "priority":
            order = {"Cao":1, "Trung bình":2, "Thấp":3}
            return sorted(tasks, key=lambda x: order.get(x["priority"],2), reverse=reverse)
        elif key == "deadline":
            return sorted(tasks, key=lambda x: x["deadline"] if x["deadline"] else "9999-12-31", reverse=reverse)
        elif key == "name":
            return sorted(tasks, key=lambda x: x["name"].lower(), reverse=reverse)
        elif key == "category":
            return sorted(tasks, key=lambda x: x["category"], reverse=reverse)
        elif key == "status":
            return sorted(tasks, key=lambda x: x["completed"], reverse=reverse)
        else:
            return tasks
    
    def on_filter_change(self):
        if self.filter_var.get() == "Theo danh mục":
            self.category_filter_combo.config(state="readonly")
        else:
            self.category_filter_combo.config(state="disabled")
        self.refresh_task_list()
    
    def toggle_sort_order(self):
        self.sort_order = "desc" if self.sort_order == "asc" else "asc"
        self.sort_order_btn.config(text="⬇️ Giảm dần" if self.sort_order == "desc" else "⬆️ Tăng dần")
        self.refresh_task_list()
    
    def update_stats(self):
        pass
    
    def update_detailed_stats(self):
        self.stats_text.delete(1.0, tk.END)
        total = len(self.tasks)
        completed = sum(1 for t in self.tasks if t["completed"])
        pending = total - completed
        
        # Thống kê trạng thái deadline
        normal = 0
        approaching = 0
        overdue = 0
        for t in self.tasks:
            if t["completed"]:
                continue
            status = self.get_task_deadline_status(t)
            if status == "overdue":
                overdue += 1
            elif status == "approaching":
                approaching += 1
            else:
                normal += 1
        
        priority_stats = {"Cao":0, "Trung bình":0, "Thấp":0}
        cat_stats = {}
        for t in self.tasks:
            priority_stats[t["priority"]] += 1
            cat = t["category"]
            cat_stats.setdefault(cat, {"total":0, "completed":0})
            cat_stats[cat]["total"] += 1
            if t["completed"]:
                cat_stats[cat]["completed"] += 1
        
        stats = f"""
╔══════════════════════════════════════════════════════════════════╗
║                     BÁO CÁO THỐNG KÊ CHI TIẾT                    ║
╚══════════════════════════════════════════════════════════════════╝

📈 TỔNG QUAN:
   • Tổng số công việc: {total}
   • Hoàn thành: {completed}
   • Tỷ lệ hoàn thành: {0 if total<=0 else completed/total*100:.1f}%
   • Đang thực hiện: {pending}

⏰ TRẠNG THÁI DEADLINE (công việc chưa hoàn thành):
   • Bình thường: {normal}
   • Gần deadline (≤2 ngày): {approaching}
   • Quá hạn: {overdue}

⭐ THEO ĐỘ ƯU TIÊN:
   • Cao: {priority_stats['Cao']}
   • Trung bình: {priority_stats['Trung bình']}
   • Thấp: {priority_stats['Thấp']}

📁 THEO DANH MỤC:
"""
        for cat, data in cat_stats.items():
            percent = data['completed']/data['total']*100 if data['total']>0 else 0
            stats += f"   • {cat}: {data['completed']}/{data['total']} ({percent:.1f}%)\n"
        
        stats += f"""
📝 CÔNG VIỆC GẦN ĐÂY:
"""
        recent = sorted(self.tasks, key=lambda x: x.get('created_at',''), reverse=True)[:5]
        for t in recent:
            status_icon = "✅" if t["completed"] else "⏳"
            stats += f"   • {status_icon} {t['name']} - {t.get('created_at','N/A')}\n"
        
        self.stats_text.insert(1.0, stats)
    
    # ---------- Import/Export ----------
    def export_data(self):
        fmt = self.export_format.get()
        if fmt == "json":
            self.export_json()
        elif fmt == "csv":
            self.export_csv()
        elif fmt == "excel" and PANDAS_AVAILABLE:
            self.export_excel()
        elif fmt == "pdf" and REPORTLAB_AVAILABLE:
            self.export_pdf()
        else:
            messagebox.showwarning("Cảnh báo", f"Định dạng {fmt} không khả dụng")
    
    def export_json(self):
        fn = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON","*.json")],
                                          initialfile=f"checklist_{datetime.now():%H-%M-%S_%d-%m-%Y}.json")
        if fn:
            with open(fn, 'w', encoding='utf-8') as f:
                json.dump(self.tasks, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Thành công", f"Đã export {len(self.tasks)} công việc")
    
    def export_csv(self):
        fn = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")],
                                          initialfile=f"checklist_{datetime.now():%H-%M-%S_%d-%m-%Y}.csv")
        if fn:
            with open(fn, 'w', newline='', encoding='utf-8-sig') as f:
                fieldnames = ['id','name','description','category','priority','deadline','tags','completed','progress','created_at']
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for t in self.tasks:
                    row = {k: t.get(k,'') for k in fieldnames}
                    writer.writerow(row)
            messagebox.showinfo("Thành công", f"Đã export CSV")
    
    def export_excel(self):
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Lỗi", "Cần pip install pandas openpyxl")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")],
                                          initialfile=f"checklist_{datetime.now():%H-%M-%S_%d-%m-%Y}.xlsx")
        if fn:
            df = pd.DataFrame(self.tasks)
            df.to_excel(fn, index=False)
            messagebox.showinfo("Thành công", "Đã export Excel")
    
    def export_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Lỗi", "Cần pip install reportlab")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF","*.pdf")],
                                          initialfile=f"report_{datetime.now():%H-%M-%S_%d-%m-%Y}.pdf")
        if fn:
            doc = SimpleDocTemplate(fn, pagesize=A4)
            story = []
            styles = getSampleStyleSheet()
            title = Paragraph(f"Báo cáo công việc - {datetime.now().strftime('%d/%m/%Y')}", styles['Title'])
            story.append(title)
            story.append(Spacer(1,12))
            data = [['Thống kê','Số lượng']]
            data.append(['Tổng số', len(self.tasks)])
            data.append(['Hoàn thành', sum(1 for t in self.tasks if t['completed'])])
            data.append(['Đang thực hiện', sum(1 for t in self.tasks if not t['completed'])])
            table = Table(data)
            table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.grey),
                                       ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                                       ('ALIGN',(0,0),(-1,-1),'CENTER'),
                                       ('GRID',(0,0),(-1,-1),1,colors.black)]))
            story.append(table)
            doc.build(story)
            messagebox.showinfo("Thành công", "Đã xuất PDF")
    
    def import_data(self):
        fn = filedialog.askopenfilename(filetypes=[("JSON/CSV","*.json *.csv"),("All","*.*")])
        if not fn:
            return
        try:
            if fn.endswith('.json'):
                with open(fn, 'r', encoding='utf-8') as f:
                    new = json.load(f)
                    for t in new:
                        t['id'] = len(self.tasks) + 1
                        self.tasks.append(t)
            elif fn.endswith('.csv'):
                with open(fn, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        row['id'] = len(self.tasks) + 1
                        self.tasks.append(row)
            self.save_data()
            self.refresh_task_list()
            messagebox.showinfo("Thành công", "Import hoàn tất")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
    
    # ---------- Backup ----------
    def manual_save(self):
        self.save_data()
        messagebox.showinfo("Thành công", "Đã lưu dữ liệu")
    
    def manual_backup(self):
        ts = datetime.now().strftime("%H-%M-%S_%d-%m-%Y")
        backup_file = os.path.join(self.backup_folder, f"backup_{ts}.json")
        with open(backup_file, 'w', encoding='utf-8') as f:
            json.dump(self.tasks, f, ensure_ascii=False, indent=2)
        self.refresh_backup_list()
        messagebox.showinfo("Thành công", f"Backup tại {backup_file}")
    
    def restore_from_backup(self):
        sel = self.backup_listbox.curselection()
        if not sel:
            messagebox.showwarning("Cảnh báo", "Chọn một backup")
            return
        backup_file = self.backup_listbox.get(sel[0])
        path = os.path.join(self.backup_folder, backup_file)
        if messagebox.askyesno("Xác nhận", "Restore sẽ thay thế dữ liệu hiện tại. Tiếp tục?"):
            with open(path, 'r', encoding='utf-8') as f:
                self.tasks = json.load(f)
            self.save_data()
            self.refresh_task_list()
            messagebox.showinfo("Thành công", "Restore hoàn tất")
    
    def refresh_backup_list(self):
        self.backup_listbox.delete(0, tk.END)
        if os.path.exists(self.backup_folder):
            backups = sorted([f for f in os.listdir(self.backup_folder) if f.endswith('.json')], reverse=True)
            for b in backups:
                self.backup_listbox.insert(tk.END, b)
    
    # ---------- Cài đặt ----------
    def toggle_auto_save(self):
        self.auto_save = self.auto_save_var.get()
        status = "BẬT" if self.auto_save else "TẮT"
        self.auto_save_label.config(text=f"✅ Auto-save: {status}", foreground="green" if self.auto_save else "red")
    
    def toggle_auto_backup(self):
        self.auto_backup = self.auto_backup_var.get()
    
    def toggle_background_notifications(self):
        self.background_notifications = self.notif_bg_var.get()
        self.save_config()
        if self.background_notifications:
            self.start_notification_thread()
        messagebox.showinfo("Thông báo", f"Đã {'bật' if self.background_notifications else 'tắt'} thông báo nền")
    
    def toggle_auto_start(self):
        self.set_auto_start(self.auto_start_var.get())
        messagebox.showinfo("Thông báo", f"Đã {'bật' if self.auto_start_var.get() else 'tắt'} tự động khởi động cùng hệ thống")
    
    def set_auto_start(self, enabled):
        self.auto_start = enabled
        self.save_config()
        
        if platform.system() == "Windows":
            if not WINSHELL_AVAILABLE:
                messagebox.showwarning("Cảnh báo", "Thiếu thư viện winshell và pywin32. Cài đặt: pip install winshell pywin32")
                return
            startup_folder = winshell.startup()
            shortcut_path = os.path.join(startup_folder, "ChecklistPro.lnk")
            if enabled:
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = sys.executable
                shortcut.Arguments = f'"{os.path.abspath(__file__)}"'
                shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
                shortcut.IconLocation = sys.executable
                shortcut.save()
            else:
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
        elif platform.system() == "Linux":
            autostart_dir = os.path.expanduser("~/.config/autostart")
            if not os.path.exists(autostart_dir):
                os.makedirs(autostart_dir)
            desktop_file = os.path.join(autostart_dir, "checklistpro.desktop")
            if enabled:
                content = f"""[Desktop Entry]
Type=Application
Name=Checklist Pro
Exec={sys.executable} {os.path.abspath(__file__)}
Hidden=false
NoDisplay=false
X-GNOME-Autostart-enabled=true
"""
                with open(desktop_file, 'w') as f:
                    f.write(content)
            else:
                if os.path.exists(desktop_file):
                    os.remove(desktop_file)
        # MacOS chưa hỗ trợ
    
    def clear_all_data(self):
        if messagebox.askyesno("Cảnh báo", "Xóa tất cả dữ liệu? Không thể hoàn tác!"):
            self.tasks = []
            self.save_data()
            self.refresh_task_list()
            messagebox.showinfo("Thành công", "Đã xóa toàn bộ")
    
    def open_data_folder(self):
        folder = os.path.dirname(os.path.abspath(self.data_file))
        if platform.system() == "Windows":
            os.startfile(folder)
        else:
            import subprocess
            subprocess.Popen(["xdg-open", folder])
    
    # ---------- Xử lý sự kiện trên tree ----------
    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)
            if item:
                values = self.tree.item(item, "values")
                if values:
                    task_name = values[1]
                    task = next((t for t in self.tasks if t["name"] == task_name), None)
                    if task:
                        if col == "#7":
                            self.update_progress_dialog(task["id"])
                        elif col == "#8":
                            menu = tk.Menu(self.root, tearoff=0)
                            menu.add_command(label="✅ Đánh dấu hoàn thành", command=lambda: self.toggle_complete(task["id"]))
                            menu.add_command(label="📊 Cập nhật tiến độ", command=lambda: self.update_progress_dialog(task["id"]))
                            menu.add_separator()
                            menu.add_command(label="✏️ Chỉnh sửa", command=lambda: self.edit_task(task["id"]))
                            menu.add_separator()
                            menu.add_command(label="🗑️ Xóa", command=lambda: self.delete_task(task["id"]))
                            menu.post(event.x_root, event.y_root)
    
    def toggle_complete(self, task_id):
        for t in self.tasks:
            if t["id"] == task_id:
                t["completed"] = not t["completed"]
                if t["completed"]:
                    t["progress"] = 100
                else:
                    if t["progress"] == 100:
                        t["progress"] = 0
                t["updated_at"] = datetime.now().strftime("%H:%M %d/%m/%Y")
                break
        self.save_data()
        self.refresh_task_list()
    
    def update_progress_dialog(self, task_id):
        task = next((t for t in self.tasks if t["id"] == task_id), None)
        if not task:
            return
        win = tk.Toplevel(self.root)
        win.title("📊 Cập nhật tiến độ")
        win.geometry("350x200")
        win.transient(self.root)
        win.grab_set()
        
        frame = ttk.Frame(win, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=f"Công việc: {task['name']}").pack(pady=5)
        ttk.Label(frame, text="Tiến độ:").pack()
        
        progress_var = tk.IntVar(value=task.get("progress", 0))
        scale = ttk.Scale(frame, from_=0, to=100, variable=progress_var, orient=tk.HORIZONTAL, length=250)
        scale.pack(pady=10)
        
        label = ttk.Label(frame, text=f"{progress_var.get()}%")
        label.pack()
        
        def update_label(*args):
            val = progress_var.get()
            label.config(text=f"{val}%")
            if val == 100 and not task["completed"]:
                task["completed"] = True
            elif val < 100 and task["completed"]:
                task["completed"] = False
        
        progress_var.trace('w', update_label)
        
        def save():
            task["progress"] = progress_var.get()
            task["completed"] = (task["progress"] == 100)
            task["updated_at"] = datetime.now().strftime("%H:%M %d/%m/%Y")
            self.save_data()
            self.refresh_task_list()
            win.destroy()
            messagebox.showinfo("Thành công", "Đã cập nhật tiến độ")
        
        ttk.Button(frame, text="Lưu", command=save).pack(pady=10)
        ttk.Button(frame, text="Hủy", command=win.destroy).pack()
    
    def edit_task(self, task_id):
        task = next((t for t in self.tasks if t["id"] == task_id), None)
        if not task:
            return
        win = tk.Toplevel(self.root)
        win.title("✏️ Chỉnh sửa công việc")
        win.geometry("500x550")
        frame = ttk.Frame(win, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Tên công việc:").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_var = tk.StringVar(value=task["name"])
        ttk.Entry(frame, textvariable=name_var, width=40).grid(row=0, column=1, pady=5)
        
        ttk.Label(frame, text="Mô tả:").grid(row=1, column=0, sticky=tk.W, pady=5)
        desc_var = tk.StringVar(value=task.get("description",""))
        ttk.Entry(frame, textvariable=desc_var, width=40).grid(row=1, column=1, pady=5)
        
        ttk.Label(frame, text="Danh mục:").grid(row=2, column=0, sticky=tk.W, pady=5)
        cat_var = tk.StringVar(value=task["category"])
        ttk.Combobox(frame, textvariable=cat_var, values=self.categories, width=20).grid(row=2, column=1, pady=5)
        
        ttk.Label(frame, text="Ưu tiên:").grid(row=3, column=0, sticky=tk.W, pady=5)
        pri_var = tk.StringVar(value=task["priority"])
        ttk.Combobox(frame, textvariable=pri_var, values=["Cao","Trung bình","Thấp"], width=20).grid(row=3, column=1, pady=5)
        
        ttk.Label(frame, text="Deadline:").grid(row=4, column=0, sticky=tk.W, pady=5)
        dl_var = tk.StringVar(value=task.get("deadline",""))
        ttk.Entry(frame, textvariable=dl_var, width=30).grid(row=4, column=1, pady=5)
        ttk.Label(frame, text="(YYYY-MM-DD HH:MM)", font=('Arial',8)).grid(row=5, column=1, sticky=tk.W)
        
        ttk.Label(frame, text="Tags:").grid(row=6, column=0, sticky=tk.W, pady=5)
        tags_var = tk.StringVar(value=task.get("tags",""))
        ttk.Entry(frame, textvariable=tags_var, width=30).grid(row=6, column=1, pady=5)
        
        # Progress
        ttk.Label(frame, text="Tiến độ:").grid(row=7, column=0, sticky=tk.W, pady=5)
        progress_var = tk.IntVar(value=task.get("progress", 0))
        progress_scale = ttk.Scale(frame, from_=0, to=100, variable=progress_var, orient=tk.HORIZONTAL, length=200)
        progress_scale.grid(row=7, column=1, sticky=tk.W, pady=5)
        progress_label = ttk.Label(frame, text=f"{progress_var.get()}%")
        progress_label.grid(row=7, column=2, padx=5)
        
        def update_progress_label(*args):
            val = progress_var.get()
            progress_label.config(text=f"{val}%")
            if val == 100:
                completed_var.set(True)
            elif val < 100 and completed_var.get():
                completed_var.set(False)
        
        progress_var.trace('w', update_progress_label)
        
        # Completed checkbox
        ttk.Label(frame, text="Trạng thái:").grid(row=8, column=0, sticky=tk.W, pady=5)
        completed_var = tk.BooleanVar(value=task["completed"])
        completed_check = ttk.Checkbutton(frame, text="Hoàn thành", variable=completed_var)
        completed_check.grid(row=8, column=1, sticky=tk.W, pady=5)
        
        def on_completed_toggle():
            if completed_var.get():
                progress_var.set(100)
            else:
                if progress_var.get() == 100:
                    progress_var.set(0)
        completed_check.config(command=on_completed_toggle)
        
        def save_edit():
            new_name = name_var.get().strip()
            if not new_name:
                messagebox.showwarning("Cảnh báo", "Tên không được để trống")
                return
            new_dl = dl_var.get().strip()
            if new_dl:
                try:
                    datetime.strptime(new_dl, "%H:%M %d/%m/%Y")
                except:
                    messagebox.showwarning("Cảnh báo", "Sai định dạng deadline")
                    return
            task["name"] = new_name
            task["description"] = desc_var.get()
            task["category"] = cat_var.get()
            task["priority"] = pri_var.get()
            task["deadline"] = new_dl
            task["tags"] = tags_var.get()
            task["progress"] = progress_var.get()
            task["completed"] = completed_var.get()
            task["updated_at"] = datetime.now().strftime("%H:%M %d/%m/%Y")
            self.save_data()
            self.refresh_task_list()
            win.destroy()
            messagebox.showinfo("Thành công", "Đã cập nhật")
        
        ttk.Button(frame, text="💾 Lưu", command=save_edit).grid(row=9, column=0, columnspan=2, pady=20)
        ttk.Button(frame, text="❌ Hủy", command=win.destroy).grid(row=10, column=0, columnspan=2)
    
    def delete_task(self, task_id):
        if messagebox.askyesno("Xác nhận", "Xóa công việc này?"):
            self.tasks = [t for t in self.tasks if t["id"] != task_id]
            self.save_data()
            self.refresh_task_list()
    
    # ---------- Nền ----------
    def check_deadlines(self):
        """Kiểm tra deadline và gửi thông báo (chạy trong thread nền)"""
        while True:
            if self.background_notifications and PLYER_AVAILABLE:
                now = datetime.now()
                for task in self.tasks:
                    if not task.get("completed", False) and task.get("deadline"):
                        try:
                            deadline = datetime.strptime(task["deadline"], "%H:%M %d/%m/%Y")
                            time_diff = (deadline - now).total_seconds()
                            
                            if 0 < time_diff <= 3600:  # 1 giờ
                                self.send_notification(
                                    "⏰ Sắp đến deadline!",
                                    f"Công việc '{task['name']}' sẽ hết hạn trong 1 giờ tới!"
                                )
                            elif 0 < time_diff <= 86400:  # 1 ngày
                                self.send_notification(
                                    "⚠️ Nhắc nhở deadline",
                                    f"Công việc '{task['name']}' sẽ hết hạn trong 1 ngày nữa!"
                                )
                            elif time_diff < 0:  # Quá hạn
                                self.send_notification(
                                    "❗ Công việc đã quá hạn!",
                                    f"Công việc '{task['name']}' đã quá deadline!"
                                )
                        except:
                            pass
            time.sleep(1800)  # 30 phút
    
    def send_notification(self, title, message):
        """Gửi thông báo desktop nếu plyer có sẵn"""
        if not self.background_notifications:
            return
        try:
            if PLYER_AVAILABLE:
                notification.notify(
                    title=title,
                    message=message,
                    app_name="Checklist Pro",
                    timeout=10
                )
            else:
                print(f"🔔 {title}: {message}")
        except Exception as e:
            print(f"Không thể gửi thông báo: {e}")
    
    def start_notification_thread(self):
        """Khởi động thread kiểm tra deadline"""
        if not hasattr(self, 'notification_thread') or self.notification_thread is None or not self.notification_thread.is_alive():
            self.notification_thread = threading.Thread(target=self.check_deadlines, daemon=True)
            self.notification_thread.start()
    
    def start_background_tasks(self):
        # Auto-save thread
        def auto_save_worker():
            while True:
                time.sleep(30)
                if self.auto_save:
                    self.save_data()
        # Auto-backup thread
        def auto_backup_worker():
            while True:
                time.sleep(3600)
                if self.auto_backup:
                    self.manual_backup()
        threading.Thread(target=auto_save_worker, daemon=True).start()
        threading.Thread(target=auto_backup_worker, daemon=True).start()
        
        # Notification thread
        self.start_notification_thread()
    
    def save_data(self):
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.tasks, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"Lỗi lưu: {e}")
            return False
    
    def load_data(self):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    self.tasks = json.load(f)
            except:
                self.tasks = []
    
    def on_closing(self):
        self.save_data()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = TaskManager(root)
    root.mainloop()