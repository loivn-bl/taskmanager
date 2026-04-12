import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import json
import csv
import os
import threading
import time

# Kiểm tra thư viện tùy chọn
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

class TaskManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Task Manager (0.1.1)")
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
        
        # Tạo thư mục backup nếu chưa có
        if not os.path.exists(self.backup_folder):
            os.makedirs(self.backup_folder)
        
        # Màu sắc
        self.priority_colors = {
            "Cao": "#FFE5E5",
            "Trung bình": "#FFF4E5", 
            "Thấp": "#E5F5E5"
        }
        
        # Load dữ liệu
        self.load_data()
        
        # Khởi tạo categories
        self.categories = ["Công việc", "Học tập", "Cá nhân", "Sức khỏe", "Khác"]
        
        # Bắt đầu các thread nền
        self.start_background_tasks()
        
        # Tạo giao diện
        self.setup_ui()
        
        # Refresh hiển thị
        self.refresh_task_list()
        
        # Bind sự kiện đóng cửa sổ
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def do_nothing(self):
        """Hàm này thực sự không làm gì cả"""
        return
    
    def setup_ui(self):
        """Tạo giao diện chính"""
        # Style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Tạo notebook (tab)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Quản lý công việc
        self.main_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="📋 Quản lý công việc")
        
        # Tab 2: Import/Export
        self.io_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.io_tab, text="📁 Import/Export")
        
        # Tab 3: Thống kê & Báo cáo
        self.stats_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.stats_tab, text="📊 Thống kê & Báo cáo")
        
        # Tab 4: Cài đặt
        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="⚙️ Cài đặt")
        
        # Setup từng tab
        self.setup_main_tab()
        self.setup_io_tab()
        self.setup_stats_tab()
        self.setup_settings_tab()
    
    def setup_main_tab(self):
        """Thiết lập tab quản lý chính"""
        # Khung trên cùng
        top_frame = ttk.Frame(self.main_tab)
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        
        quick_save_btn = ttk.Button(top_frame, text="💾 Lưu ngay", command=self.manual_save)
        quick_save_btn.pack(side=tk.LEFT, padx=5)
        
        backup_btn = ttk.Button(top_frame, text="🔄 Backup ngay", command=self.manual_backup)
        backup_btn.pack(side=tk.LEFT, padx=5)
        
        self.auto_save_label = ttk.Label(top_frame, text="✅ Auto-save: BẬT", foreground="green")
        self.auto_save_label.pack(side=tk.LEFT, padx=20)
        
        # Khung nhập liệu
        input_frame = ttk.LabelFrame(self.main_tab, text="➕ Thêm công việc mới", padding="15")
        input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(input_frame, text="Tên công việc:*").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.task_name_var = tk.StringVar()
        self.task_name_entry = ttk.Entry(input_frame, textvariable=self.task_name_var, width=40)
        self.task_name_entry.grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Mô tả:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.task_desc_var = tk.StringVar()
        self.task_desc_entry = ttk.Entry(input_frame, textvariable=self.task_desc_var, width=40)
        self.task_desc_entry.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Danh mục:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_var = tk.StringVar(value="Công việc")
        category_combo = ttk.Combobox(input_frame, textvariable=self.category_var, 
                                     values=self.categories, width=15)
        category_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Ưu tiên:").grid(row=2, column=2, sticky=tk.W, padx=5, pady=5)
        self.priority_var = tk.StringVar(value="Trung bình")
        priority_combo = ttk.Combobox(input_frame, textvariable=self.priority_var, 
                                     values=["Cao", "Trung bình", "Thấp"], width=15)
        priority_combo.grid(row=2, column=3, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Deadline:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.deadline_var = tk.StringVar()
        self.deadline_entry = ttk.Entry(input_frame, textvariable=self.deadline_var, width=20)
        self.deadline_entry.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(input_frame, text="(HH:MM DD/MM/YYYY)", font=('Arial', 8)).grid(row=3, column=2, sticky=tk.W)
        
        ttk.Label(input_frame, text="Tags:").grid(row=3, column=3, sticky=tk.W, padx=5, pady=5)
        self.tags_var = tk.StringVar()
        self.tags_entry = ttk.Entry(input_frame, textvariable=self.tags_var, width=20)
        self.tags_entry.grid(row=3, column=4, sticky=tk.W, padx=5, pady=5)
        
        add_btn = ttk.Button(input_frame, text="➕ Thêm công việc", command=self.add_task)
        add_btn.grid(row=4, column=0, columnspan=5, pady=15)
        
        # Khung điều khiển
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
        
        self.sort_order_btn = ttk.Button(control_frame, text="⬆️ Tăng dần", width=10,
                                        command=self.toggle_sort_order)
        self.sort_order_btn.pack(side=tk.LEFT, padx=5)
        
        # Khung danh sách công việc
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
        """Thiết lập tab Import/Export"""
        main_io_frame = ttk.Frame(self.io_tab, padding="20")
        main_io_frame.pack(fill=tk.BOTH, expand=True)
        
        # Export
        export_frame = ttk.LabelFrame(main_io_frame, text="📤 Export dữ liệu", padding="15")
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
        import_frame = ttk.LabelFrame(main_io_frame, text="📥 Import dữ liệu", padding="15")
        import_frame.pack(fill=tk.X, pady=10)
        ttk.Label(import_frame, text="Chọn file để import:").pack(anchor=tk.W, pady=5)
        ttk.Button(import_frame, text="📂 Import từ file", command=self.import_data, width=20).pack(pady=5)
        ttk.Label(import_frame, text="⚠️ Lưu ý: Import sẽ thêm dữ liệu mới vào danh sách hiện tại",
                 foreground="orange").pack(anchor=tk.W, pady=5)
        
        # Backup/Restore
        backup_frame = ttk.LabelFrame(main_io_frame, text="🔄 Backup & Restore", padding="15")
        backup_frame.pack(fill=tk.X, pady=10)
        btn_frame = ttk.Frame(backup_frame)
        btn_frame.pack()
        ttk.Button(btn_frame, text="💾 Tạo Backup", command=self.manual_backup, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📂 Restore từ Backup", command=self.restore_from_backup, width=15).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(backup_frame, text="Danh sách backups:").pack(anchor=tk.W, pady=(10, 5))
        self.backup_listbox = tk.Listbox(backup_frame, height=5)
        self.backup_listbox.pack(fill=tk.X, pady=5)
        self.refresh_backup_list()
    
    def setup_stats_tab(self):
        """Thiết lập tab thống kê"""
        stats_display = ttk.Frame(self.stats_tab, padding="20")
        stats_display.pack(fill=tk.BOTH, expand=True)
        
        self.stats_text = tk.Text(stats_display, height=20, width=80, font=("JetBrainsMono Nerd Font", 10))
        self.stats_text.pack(fill=tk.BOTH, expand=True)
        
        ttk.Button(stats_display, text="🔄 Làm mới thống kê", command=self.update_detailed_stats).pack(pady=10)
        if REPORTLAB_AVAILABLE:
            ttk.Button(stats_display, text="📊 Xuất báo cáo PDF", command=self.export_pdf).pack(pady=5)
        
        self.update_detailed_stats()
    
    def setup_settings_tab(self):
        """Thiết lập tab cài đặt"""
        settings_frame = ttk.Frame(self.settings_tab, padding="20")
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # Auto-save
        auto_save_frame = ttk.LabelFrame(settings_frame, text="⚙️ Cài đặt tự động", padding="15")
        auto_save_frame.pack(fill=tk.X, pady=10)
        
        self.auto_save_var = tk.BooleanVar(value=self.auto_save)
        ttk.Checkbutton(auto_save_frame, text="Bật Auto-save (lưu sau mỗi 30 giây)", 
                       variable=self.auto_save_var, command=self.toggle_auto_save).pack(anchor=tk.W)
        
        self.auto_backup_var = tk.BooleanVar(value=self.auto_backup)
        ttk.Checkbutton(auto_save_frame, text="Bật Auto-backup (backup mỗi giờ)", 
                       variable=self.auto_backup_var, command=self.toggle_auto_backup).pack(anchor=tk.W, pady=5)
        
        # Notification
        notif_frame = ttk.LabelFrame(settings_frame, text="🔔 Cài đặt thông báo", padding="15")
        notif_frame.pack(fill=tk.X, pady=10)
        self.notification_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(notif_frame, text="Bật thông báo desktop", variable=self.notification_var).pack(anchor=tk.W)
        
        # Data management
        data_frame = ttk.LabelFrame(settings_frame, text="🗄️ Quản lý dữ liệu", padding="15")
        data_frame.pack(fill=tk.X, pady=10)
        ttk.Button(data_frame, text="🗑️ Xóa tất cả dữ liệu", command=self.clear_all_data, width=20).pack(pady=5)
        ttk.Button(data_frame, text="📁 Mở thư mục dữ liệu", command=self.open_data_folder, width=20).pack(pady=5)
        
        # Info
        info_frame = ttk.LabelFrame(settings_frame, text="ℹ️ Thông tin", padding="15")
        info_frame.pack(fill=tk.X, pady=10)
        ttk.Label(info_frame, text="Version: 0.1.1 (build 2)").pack(anchor=tk.W)
    
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
        
        deadline = None
        if deadline_str:
            try:
                deadline = datetime.strptime(deadline_str, "%H:%M %d/%m/%y")
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
            "created_at": datetime.now().strftime("%H:%M %d/%m/%y"),
            "updated_at": datetime.now().strftime("%H:%M %d/%m/%y")
        }
        self.tasks.append(task)
        self.save_data()
        
        self.task_name_var.set("")
        self.task_desc_var.set("")
        self.deadline_var.set("")
        self.tags_var.set("")
        self.refresh_task_list()
        messagebox.showinfo("Thành công", f"Đã thêm công việc: {name}")
    
    def refresh_task_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        filtered_tasks = self.filter_tasks()
        filtered_tasks = self.sort_tasks(filtered_tasks)
        
        for task in filtered_tasks:
            status, is_overdue = self.get_task_status(task)
            item_id = self.tree.insert("", tk.END, values=(
                status,
                task["name"],
                task["category"],
                task["priority"],
                task["deadline"] if task["deadline"] else "Không có",
                task.get("tags", ""),
                f"{task.get('progress', 0)}%",
                "⚙️"
            ))
            color = self.priority_colors.get(task["priority"], "#FFFFFF")
            self.tree.tag_configure(f"color_{task['id']}", background=color)
            self.tree.item(item_id, tags=(f"color_{task['id']}",))
            if is_overdue:
                self.tree.tag_configure(f"overdue_{task['id']}", background="#FFCCCC")
                self.tree.item(item_id, tags=(f"overdue_{task['id']}",))
        
        self.update_stats()
    
    def get_task_status(self, task):
        if task["completed"]:
            return "✅", False
        if task["deadline"]:
            try:
                deadline = datetime.strptime(task["deadline"], "%H:%M %d/%m/%y")
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
            return [t for t in self.tasks if not t["completed"] and t["deadline"]]
        elif filter_type == "Theo danh mục":
            category = self.category_filter_var.get()
            return [t for t in self.tasks if t["category"] == category]
        else:
            return self.tasks.copy()
    
    def sort_tasks(self, tasks):
        sort_key = self.sort_var.get()
        reverse = (self.sort_order == "desc")
        if sort_key == "priority":
            priority_order = {"Cao": 1, "Trung bình": 2, "Thấp": 3}
            return sorted(tasks, key=lambda x: priority_order.get(x["priority"], 2), reverse=reverse)
        elif sort_key == "deadline":
            return sorted(tasks, key=lambda x: x["deadline"] if x["deadline"] else "9999-12-31", reverse=reverse)
        elif sort_key == "name":
            return sorted(tasks, key=lambda x: x["name"].lower(), reverse=reverse)
        elif sort_key == "category":
            return sorted(tasks, key=lambda x: x["category"], reverse=reverse)
        elif sort_key == "status":
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
        total = len(self.tasks)
        completed = sum(1 for t in self.tasks if t["completed"])
        pending = total - completed
        overdue = 0
        for t in self.tasks:
            if not t["completed"] and t["deadline"]:
                try:
                    deadline = datetime.strptime(t["deadline"], "%H:%M %d/%m/%y")
                    if deadline < datetime.now():
                        overdue += 1
                except:
                    pass
        # Có thể hiển thị ở đâu đó nếu muốn, nhưng không bắt buộc
        pass
    
    def update_detailed_stats(self):
        self.stats_text.delete(1.0, tk.END)
        total = len(self.tasks)
        completed = sum(1 for t in self.tasks if t["completed"])
        pending = total - completed
        
        # Thống kê theo ưu tiên
        priority_stats = {"Cao": 0, "Trung bình": 0, "Thấp": 0}
        for t in self.tasks:
            priority_stats[t["priority"]] += 1
        
        # Thống kê theo danh mục
        category_stats = {}
        for t in self.tasks:
            cat = t["category"]
            if cat not in category_stats:
                category_stats[cat] = {"total": 0, "completed": 0}
            category_stats[cat]["total"] += 1
            if t["completed"]:
                category_stats[cat]["completed"] += 1
        
        # Xây dựng chuỗi thống kê
        stats = f"""
╔══════════════════════════════════════════════════════════════════╗
║                    📊 BÁO CÁO THỐNG KÊ CHI TIẾT                  ║
╚══════════════════════════════════════════════════════════════════╝

📈 TỔNG QUAN:
   • Tổng số công việc: {total}
   • Hoàn thành: {completed} 
"""
        if total > 0:
            stats += f"  • Tỷ lệ hoàn thành: {completed/total*100:.1f}%\n"
        else:
            stats += "  • Tỷ lệ hoàn thành: 0%\n"
        
        stats += f"""   • Đang thực hiện: {pending}

⭐ THEO ĐỘ ƯU TIÊN:
   • Cao: {priority_stats['Cao']} công việc
   • Trung bình: {priority_stats['Trung bình']} công việc
   • Thấp: {priority_stats['Thấp']} công việc

📁 THEO DANH MỤC:
"""
        for cat, data in category_stats.items():
            if data['total'] > 0:
                percent = data['completed']/data['total']*100
                stats += f"   • {cat}: {data['completed']}/{data['total']} ({percent:.1f}%)\n"
            else:
                stats += f"   • {cat}: {data['completed']}/{data['total']} (0%)\n"
        
        # Thống kê deadline
        upcoming = 0
        overdue = 0
        for t in self.tasks:
            if not t["completed"] and t["deadline"]:
                try:
                    deadline = datetime.strptime(t["deadline"], "%H:%M %d/%m/%y")
                    if deadline < datetime.now():
                        overdue += 1
                    elif deadline - datetime.now() <= timedelta(days=3):
                        upcoming += 1
                except:
                    pass
        
        stats += f"""
⏰ THEO DEADLINE:
   • Quá hạn: {overdue} công việc
   • Sắp đến hạn (3 ngày tới): {upcoming} công việc

📝 CÔNG VIỆC GẦN ĐÂY:
"""
        recent_tasks = sorted(self.tasks, key=lambda x: x.get('created_at', ''), reverse=True)[:5]
        for task in recent_tasks:
            status = "✅" if task["completed"] else "⏳"
            stats += f"   • {status} {task['name']} - {task.get('created_at', 'N/A')}\n"
        
        self.stats_text.insert(1.0, stats)
    
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
            messagebox.showwarning("Cảnh báo", f"Định dạng {fmt} không khả dụng (thiếu thư viện)")
    
    def export_json(self):
        filename = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")],
                                                initialfile=f"checklist_{datetime.now():%H%M%S_%d%m%Y}.json")
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(self.tasks, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Thành công", f"Đã export {len(self.tasks)} công việc ra JSON")
    
    def export_csv(self):
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                                initialfile=f"checklist_{datetime.now():%H%M%S_%d%m%Y}.csv")
        if filename:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                fieldnames = ['id', 'name', 'description', 'category', 'priority', 'deadline', 'tags', 'completed', 'progress', 'created_at']
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for task in self.tasks:
                    row = {k: task.get(k, '') for k in fieldnames}
                    writer.writerow(row)
            messagebox.showinfo("Thành công", f"Đã export {len(self.tasks)} công việc ra CSV")
    
    def export_excel(self):
        if not PANDAS_AVAILABLE:
            messagebox.showerror("Lỗi", "Cần cài pandas: pip install pandas openpyxl")
            return
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                                initialfile=f"checklist_{datetime.now():%H%M%S_%d%m%Y}.xlsx")
        if filename:
            df = pd.DataFrame(self.tasks)
            df.to_excel(filename, index=False)
            messagebox.showinfo("Thành công", f"Đã export ra Excel")
    
    def export_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Lỗi", "Cần cài reportlab: pip install reportlab")
            return
        filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
                                                initialfile=f"report_{datetime.now():%H%M%S_%d%m%Y}.pdf")
        if filename:
            doc = SimpleDocTemplate(filename, pagesize=A4)
            story = []
            styles = getSampleStyleSheet()
            title = Paragraph(f"Báo cáo công việc - {datetime.now().strftime('%d/%m/%Y')}", styles['Title'])
            story.append(title)
            story.append(Spacer(1, 12))
            data = [['Thống kê', 'Số lượng']]
            data.append(['Tổng số', len(self.tasks)])
            data.append(['Hoàn thành', sum(1 for t in self.tasks if t['completed'])])
            data.append(['Đang thực hiện', sum(1 for t in self.tasks if not t['completed'])])
            table = Table(data)
            table.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.grey),
                                       ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                                       ('ALIGN',(0,0),(-1,-1),'CENTER'),
                                       ('GRID',(0,0),(-1,-1),1,colors.black)]))
            story.append(table)
            doc.build(story)
            messagebox.showinfo("Thành công", "Đã xuất báo cáo PDF")
    
    def import_data(self):
        filename = filedialog.askopenfilename(filetypes=[("JSON/CSV", "*.json *.csv"), ("All", "*.*")])
        if not filename:
            return
        try:
            if filename.endswith('.json'):
                with open(filename, 'r', encoding='utf-8') as f:
                    new_tasks = json.load(f)
                    for t in new_tasks:
                        t['id'] = len(self.tasks) + 1
                        self.tasks.append(t)
            elif filename.endswith('.csv'):
                with open(filename, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        row['id'] = len(self.tasks) + 1
                        self.tasks.append(row)
            self.save_data()
            self.refresh_task_list()
            messagebox.showinfo("Thành công", "Import hoàn tất")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
    
    def manual_save(self):
        self.save_data()
        messagebox.showinfo("Thành công", "Đã lưu dữ liệu")
    
    def manual_backup(self):
        timestamp = datetime.now().strftime("%H%M%S_%d%m%Y")
        backup_file = os.path.join(self.backup_folder, f"backup_{timestamp}.json")
        with open(backup_file, 'w', encoding='utf-8') as f:
            json.dump(self.tasks, f, ensure_ascii=False, indent=2)
        self.refresh_backup_list()
        messagebox.showinfo("Thành công", f"Backup lưu tại {backup_file}")
    
    def restore_from_backup(self):
        sel = self.backup_listbox.curselection()
        if not sel:
            messagebox.showwarning("Cảnh báo", "Chọn một backup để restore")
            return
        backup_file = self.backup_listbox.get(sel[0])
        backup_path = os.path.join(self.backup_folder, backup_file)
        if messagebox.askyesno("Xác nhận", "Restore sẽ thay thế dữ liệu hiện tại. Tiếp tục?"):
            with open(backup_path, 'r', encoding='utf-8') as f:
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
    
    def clear_all_data(self):
        if messagebox.askyesno("Cảnh báo", "Xóa tất cả dữ liệu? Không thể hoàn tác!"):
            self.tasks = []
            self.save_data()
            self.refresh_task_list()
            messagebox.showinfo("Thành công", "Đã xóa toàn bộ dữ liệu")
    
    def open_data_folder(self):
        import subprocess, platform
        folder = os.path.dirname(os.path.abspath(self.data_file))
        if platform.system() == "Windows":
            os.startfile(folder)
        else:
            subprocess.Popen(["xdg-open", folder])
    
    def toggle_auto_save(self):
        self.auto_save = self.auto_save_var.get()
        status = "BẬT" if self.auto_save else "TẮT"
        self.auto_save_label.config(text=f"✅ Auto-save: {status}", foreground="green" if self.auto_save else "red")
    
    def toggle_auto_backup(self):
        self.auto_backup = self.auto_backup_var.get()
    
    def start_background_tasks(self):
        def auto_save_worker():
            while True:
                time.sleep(30)
                if self.auto_save:
                    self.save_data()
        def auto_backup_worker():
            while True:
                time.sleep(3600)
                if self.auto_backup:
                    self.manual_backup()
        threading.Thread(target=auto_save_worker, daemon=True).start()
        threading.Thread(target=auto_backup_worker, daemon=True).start()
    
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
    
    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)
            if item and col == "#8":
                values = self.tree.item(item, "values")
                if values:
                    task_name = values[1]
                    task = next((t for t in self.tasks if t["name"] == task_name), None)
                    if task:
                        menu = tk.Menu(self.root, tearoff=0)
                        menu.add_command(label="✅ Đánh dấu hoàn thành", 
                                        command=lambda: self.toggle_complete(task["id"]))
                        menu.add_command(label="✏️ Chỉnh sửa", 
                                        command=lambda: self.edit_task(task["id"]))
                        menu.add_separator()
                        menu.add_command(label="🗑️ Xóa", 
                                        command=lambda: self.delete_task(task["id"]))
                        menu.add_command(label="❌ Đóng", 
                                        command=lambda: self.do_nothing())
                        menu.post(event.x_root, event.y_root)
    
    def toggle_complete(self, task_id):
        for task in self.tasks:
            if task["id"] == task_id:
                task["completed"] = not task["completed"]
                task["progress"] = 100 if task["completed"] else 0
                task["updated_at"] = datetime.now().strftime("%H:%M %d/%m/%y")
                break
        self.save_data()
        self.refresh_task_list()
    
    def edit_task(self, task_id):
        task = next((t for t in self.tasks if t["id"] == task_id), None)
        if not task:
            return
        edit_win = tk.Toplevel(self.root)
        edit_win.title("✏️ Chỉnh sửa công việc")
        edit_win.geometry("500x450")
        frame = ttk.Frame(edit_win, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Tên công việc:").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_var = tk.StringVar(value=task["name"])
        ttk.Entry(frame, textvariable=name_var, width=40).grid(row=0, column=1, pady=5)
        
        ttk.Label(frame, text="Mô tả:").grid(row=1, column=0, sticky=tk.W, pady=5)
        desc_var = tk.StringVar(value=task.get("description", ""))
        ttk.Entry(frame, textvariable=desc_var, width=40).grid(row=1, column=1, pady=5)
        
        ttk.Label(frame, text="Danh mục:").grid(row=2, column=0, sticky=tk.W, pady=5)
        cat_var = tk.StringVar(value=task["category"])
        ttk.Combobox(frame, textvariable=cat_var, values=self.categories, width=20).grid(row=2, column=1, pady=5)
        
        ttk.Label(frame, text="Ưu tiên:").grid(row=3, column=0, sticky=tk.W, pady=5)
        pri_var = tk.StringVar(value=task["priority"])
        ttk.Combobox(frame, textvariable=pri_var, values=["Cao","Trung bình","Thấp"], width=20).grid(row=3, column=1, pady=5)
        
        ttk.Label(frame, text="Deadline:").grid(row=4, column=0, sticky=tk.W, pady=5)
        dl_var = tk.StringVar(value=task.get("deadline", ""))
        ttk.Entry(frame, textvariable=dl_var, width=30).grid(row=4, column=1, pady=5)
        ttk.Label(frame, text="(HH:MM DD/MM/YYYY)", font=('Arial',8)).grid(row=5, column=1, sticky=tk.W)
        
        ttk.Label(frame, text="Tags:").grid(row=6, column=0, sticky=tk.W, pady=5)
        tags_var = tk.StringVar(value=task.get("tags", ""))
        ttk.Entry(frame, textvariable=tags_var, width=30).grid(row=6, column=1, pady=5)
        
        def save_edit():
            new_name = name_var.get().strip()
            if not new_name:
                messagebox.showwarning("Cảnh báo", "Tên không được để trống")
                return
            new_dl = dl_var.get().strip()
            if new_dl:
                try:
                    datetime.strptime(new_dl, "%H:%M %d/%m/%y")
                except:
                    messagebox.showwarning("Cảnh báo", "Sai định dạng deadline")
                    return
            task["name"] = new_name
            task["description"] = desc_var.get()
            task["category"] = cat_var.get()
            task["priority"] = pri_var.get()
            task["deadline"] = new_dl
            task["tags"] = tags_var.get()
            task["updated_at"] = datetime.now().strftime("%H:%M %d/%m/%y")
            self.save_data()
            self.refresh_task_list()
            edit_win.destroy()
            messagebox.showinfo("Thành công", "Đã cập nhật")
        
        ttk.Button(frame, text="💾 Lưu", command=save_edit).grid(row=7, column=0, columnspan=2, pady=20)
        ttk.Button(frame, text="❌ Hủy", command=edit_win.destroy).grid(row=8, column=0, columnspan=2)
    
    def delete_task(self, task_id):
        if messagebox.askyesno("Xác nhận", "Xóa công việc này?"):
            self.tasks = [t for t in self.tasks if t["id"] != task_id]
            self.save_data()
            self.refresh_task_list()
    
    def on_closing(self):
        self.save_data()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = TaskManager(root)
    root.mainloop()