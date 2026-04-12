import tkinter as tk
from tkinter import ttk, messagebox, colorchooser, font
from datetime import datetime, timedelta
import json
import os
from plyer import notification
import threading
import time

class AdvancedChecklistApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Task Manager (0.1.0)")
        self.root.geometry("1200x700")
        
        # Dữ liệu tasks
        self.tasks = []
        self.current_filter = "Tất cả"
        self.sort_by = "priority"  # priority, deadline, name, status
        self.sort_order = "asc"
        self.data_file = "checklist_data.json"
        
        # Màu sắc ưu tiên
        self.priority_colors = {
            "Cao": "#FF6B6B",    # Đỏ nhạt
            "Trung bình": "#FFD93D", # Vàng
            "Thấp": "#6BCB77"    # Xanh lá
        }
        
        # Load dữ liệu
        self.load_data()
        
        # Kiểm tra deadline trong background
        self.checking_deadlines = True
        self.start_deadline_checker()
        
        # Tạo giao diện
        self.setup_ui()
        
        # Refresh hiển thị
        self.refresh_task_list()
    
    def do_nothing(self):
        """Hàm này thực sự không làm gì cả"""
        return

    def setup_ui(self):
        """Tạo giao diện người dùng"""
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Khung chính
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Tiêu đề
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        
        title_label = ttk.Label(title_frame, text="QUẢN LÝ CÔNG VIỆC", 
                                font=('Arial', 20, 'bold'))
        title_label.pack()
        
        # Khung input
        input_frame = ttk.LabelFrame(main_frame, text="➕ Thêm công việc mới", padding="10")
        input_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Tên công việc
        ttk.Label(input_frame, text="Tên công việc:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.task_name_var = tk.StringVar()
        self.task_name_entry = ttk.Entry(input_frame, textvariable=self.task_name_var, width=40)
        self.task_name_entry.grid(row=0, column=1, columnspan=3, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Mô tả
        ttk.Label(input_frame, text="Mô tả:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.task_desc_var = tk.StringVar()
        self.task_desc_entry = ttk.Entry(input_frame, textvariable=self.task_desc_var, width=40)
        self.task_desc_entry.grid(row=1, column=1, columnspan=3, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0))
        
        # Ưu tiên
        ttk.Label(input_frame, text="Ưu tiên:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.priority_var = tk.StringVar(value="Trung bình")
        priority_combo = ttk.Combobox(input_frame, textvariable=self.priority_var, 
                                      values=["Cao", "Trung bình", "Thấp"], width=15)
        priority_combo.grid(row=2, column=1, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        
        # Deadline
        ttk.Label(input_frame, text="Deadline:").grid(row=2, column=2, sticky=tk.W, padx=(10, 10), pady=(5, 0))
        self.deadline_var = tk.StringVar()
        self.deadline_entry = ttk.Entry(input_frame, textvariable=self.deadline_var, width=20)
        self.deadline_entry.grid(row=2, column=3, sticky=tk.W, pady=(5, 0))
        ttk.Label(input_frame, text="(HH:MM DD/MM/YYYY)", font=('Arial', 8)).grid(row=3, column=3, sticky=tk.W)
        
        # Nút thêm
        add_btn = ttk.Button(input_frame, text="➕ Thêm công việc", command=self.add_task)
        add_btn.grid(row=4, column=0, columnspan=4, pady=(10, 0))
        
        # Khung điều khiển
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Bộ lọc
        ttk.Label(control_frame, text="Lọc theo:").pack(side=tk.LEFT, padx=(0, 5))
        self.filter_var = tk.StringVar(value="Tất cả")
        filter_combo = ttk.Combobox(control_frame, textvariable=self.filter_var, 
                                   values=["Tất cả", "Đang thực hiện", "Hoàn thành", "Quá hạn"],
                                   width=15, state="readonly")
        filter_combo.pack(side=tk.LEFT, padx=(0, 20))
        filter_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_task_list())
        
        # Sắp xếp theo
        ttk.Label(control_frame, text="Sắp xếp theo:").pack(side=tk.LEFT, padx=(0, 5))
        self.sort_var = tk.StringVar(value="priority")
        sort_combo = ttk.Combobox(control_frame, textvariable=self.sort_var,
                                 values=["priority", "deadline", "name", "status"],
                                 width=15, state="readonly")
        sort_combo.pack(side=tk.LEFT, padx=(0, 5))
        sort_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_task_list())
        
        # Nút đảo thứ tự
        self.sort_order_btn = ttk.Button(control_frame, text="⬆️", width=3, 
                                        command=self.toggle_sort_order)
        self.sort_order_btn.pack(side=tk.LEFT)
        
        # Khung danh sách công việc
        list_frame = ttk.LabelFrame(main_frame, text="📋 Danh sách công việc", padding="10")
        list_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Tạo Treeview
        columns = ("status", "name", "description", "priority", "deadline", "actions")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        
        # Định nghĩa các cột
        self.tree.heading("status", text="Đã xong")
        self.tree.heading("name", text="Tên công việc")
        self.tree.heading("description", text="Mô tả")
        self.tree.heading("priority", text="Ưu tiên")
        self.tree.heading("deadline", text="Deadline")
        self.tree.heading("actions", text="Thao tác")
        
        self.tree.column("status", width=50, anchor="center")
        self.tree.column("name", width=200)
        self.tree.column("description", width=200)
        self.tree.column("priority", width=100, anchor="center")
        self.tree.column("deadline", width=150, anchor="center")
        self.tree.column("actions", width=150, anchor="center")
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Bind sự kiện click vào nút action
        self.tree.bind('<ButtonRelease-1>', self.on_tree_click)
        
        # Khung thống kê
        stats_frame = ttk.LabelFrame(main_frame, text="📊 Thống kê", padding="10")
        stats_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        self.stats_label = ttk.Label(stats_frame, text="", font=('Arial', 10))
        self.stats_label.pack()
        
        # Cấu hình grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
    
    def add_task(self):
        """Thêm công việc mới"""
        name = self.task_name_var.get().strip()
        if not name:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập tên công việc!")
            return
        
        description = self.task_desc_var.get().strip()
        priority = self.priority_var.get()
        deadline_str = self.deadline_var.get().strip()
        
        # Kiểm tra deadline
        deadline = None
        if deadline_str:
            try:
                deadline = datetime.strptime(deadline_str, "%H:%M %d/%m/%Y")
                if deadline < datetime.now():
                    messagebox.showwarning("Cảnh báo", "Deadline không thể ở quá khứ!")
                    return
            except ValueError:
                messagebox.showwarning("Cảnh báo", "Định dạng deadline không đúng!\nVui lòng dùng: HH:MM DD/MM/YYYY")
                return
        
        # Tạo task mới
        task = {
            "id": len(self.tasks) + 1,
            "name": name,
            "description": description,
            "priority": priority,
            "deadline": deadline_str,
            "deadline_obj": deadline,
            "completed": False,
            "created_at": datetime.now().strftime("%H:%M %d/%m/%Y")
        }
        
        self.tasks.append(task)
        self.save_data()
        
        # Xóa form
        self.task_name_var.set("")
        self.task_desc_var.set("")
        self.deadline_var.set("")
        
        # Refresh danh sách
        self.refresh_task_list()
        messagebox.showinfo("Thành công", f"Đã thêm công việc: {name}")
    
    def toggle_complete(self, task_id):
        """Đánh dấu hoàn thành/ chưa hoàn thành"""
        for task in self.tasks:
            if task["id"] == task_id:
                task["completed"] = not task["completed"]
                break
        
        self.save_data()
        self.refresh_task_list()
    
    def delete_task(self, task_id):
        """Xóa công việc"""
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa công việc này?"):
            self.tasks = [task for task in self.tasks if task["id"] != task_id]
            self.save_data()
            self.refresh_task_list()
    
    def edit_task(self, task_id):
        """Chỉnh sửa công việc"""
        task = next((t for t in self.tasks if t["id"] == task_id), None)
        if not task:
            return
        
        # Tạo cửa sổ edit
        edit_window = tk.Toplevel(self.root)
        edit_window.title("✏️ Chỉnh sửa công việc")
        edit_window.geometry("500x400")
        
        # Form chỉnh sửa
        frame = ttk.Frame(edit_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Tên công việc:").grid(row=0, column=0, sticky=tk.W, pady=5)
        name_var = tk.StringVar(value=task["name"])
        name_entry = ttk.Entry(frame, textvariable=name_var, width=40)
        name_entry.grid(row=0, column=1, pady=5)
        
        ttk.Label(frame, text="Mô tả:").grid(row=1, column=0, sticky=tk.W, pady=5)
        desc_var = tk.StringVar(value=task["description"])
        desc_entry = ttk.Entry(frame, textvariable=desc_var, width=40)
        desc_entry.grid(row=1, column=1, pady=5)
        
        ttk.Label(frame, text="Ưu tiên:").grid(row=2, column=0, sticky=tk.W, pady=5)
        priority_var = tk.StringVar(value=task["priority"])
        priority_combo = ttk.Combobox(frame, textvariable=priority_var, 
                                     values=["Cao", "Trung bình", "Thấp"], width=20)
        priority_combo.grid(row=2, column=1, pady=5)
        
        ttk.Label(frame, text="Deadline:").grid(row=3, column=0, sticky=tk.W, pady=5)
        deadline_var = tk.StringVar(value=task["deadline"])
        deadline_entry = ttk.Entry(frame, textvariable=deadline_var, width=30)
        deadline_entry.grid(row=3, column=1, pady=5)
        ttk.Label(frame, text="(HH:MM DD/MM/YYYY)", font=('Arial', 8)).grid(row=4, column=1, sticky=tk.W)
        
        def save_changes():
            new_name = name_var.get().strip()
            if not new_name:
                messagebox.showwarning("Cảnh báo", "Tên công việc không được để trống!")
                return
            
            # Kiểm tra deadline
            new_deadline = deadline_var.get().strip()
            if new_deadline:
                try:
                    deadline_obj = datetime.strptime(new_deadline, "%H:%M %d/%m/%Y")
                    if deadline_obj < datetime.now():
                        messagebox.showwarning("Cảnh báo", "Deadline không thể ở quá khứ!")
                        return
                except ValueError:
                    messagebox.showwarning("Cảnh báo", "Định dạng deadline không đúng!")
                    return
            
            task["name"] = new_name
            task["description"] = desc_var.get()
            task["priority"] = priority_var.get()
            task["deadline"] = new_deadline
            task["deadline_obj"] = datetime.strptime(new_deadline, "%H:%M %d/%m/%Y") if new_deadline else None
            
            self.save_data()
            self.refresh_task_list()
            edit_window.destroy()
            messagebox.showinfo("Thành công", "Đã cập nhật công việc!")
        
        ttk.Button(frame, text="💾 Lưu thay đổi", command=save_changes).grid(row=5, column=0, columnspan=2, pady=20)
        ttk.Button(frame, text="❌ Hủy", command=edit_window.destroy).grid(row=6, column=0, columnspan=2)
    
    def refresh_task_list(self):
        """Làm mới danh sách công việc"""
        # Xóa tất cả
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Lọc tasks
        filtered_tasks = self.filter_tasks()
        
        # Sắp xếp
        filtered_tasks = self.sort_tasks(filtered_tasks)
        
        # Hiển thị
        for task in filtered_tasks:
            # Xác định trạng thái quá hạn
            is_overdue = False
            if task["deadline"] and not task["completed"]:
                deadline = datetime.strptime(task["deadline"], "%H:%M %d/%m/%Y")
                if deadline < datetime.now():
                    is_overdue = True
            
            # Status
            if task["completed"]:
                status = "✅"
            elif is_overdue:
                status = "⚠️"
            else:
                status = "⬜"
            
            # Màu sắc theo priority
            priority_color = self.priority_colors.get(task["priority"], "#FFFFFF")
            
            # Thêm vào tree
            item_id = self.tree.insert("", tk.END, values=(
                status,
                task["name"],
                task["description"][:50] + ("..." if len(task["description"]) > 50 else ""),
                task["priority"],
                task["deadline"] if task["deadline"] else "Không có",
                "⚙️"
            ))
            
            # Đổi màu dòng
            self.tree.tag_configure(f"color_{task['id']}", background=priority_color)
            self.tree.item(item_id, tags=(f"color_{task['id']}",))
        
        # Cập nhật thống kê
        self.update_stats()
    
    def filter_tasks(self):
        """Lọc tasks theo điều kiện"""
        filter_type = self.filter_var.get()
        
        if filter_type == "Đang thực hiện":
            return [t for t in self.tasks if not t["completed"]]
        elif filter_type == "Hoàn thành":
            return [t for t in self.tasks if t["completed"]]
        elif filter_type == "Quá hạn":
            return [t for t in self.tasks if not t["completed"] and t["deadline"] and 
                   datetime.strptime(t["deadline"], "%H:%M %d/%m/%Y") < datetime.now()]
        else:
            return self.tasks.copy()
    
    def sort_tasks(self, tasks):
        """Sắp xếp tasks"""
        sort_key = self.sort_var.get()
        reverse = (self.sort_order == "desc")
        
        if sort_key == "priority":
            priority_order = {"Cao": 1, "Trung bình": 2, "Thấp": 3}
            return sorted(tasks, key=lambda x: priority_order.get(x["priority"], 2), reverse=reverse)
        elif sort_key == "deadline":
            return sorted(tasks, key=lambda x: x["deadline"] if x["deadline"] else "9999-12-31", reverse=reverse)
        elif sort_key == "name":
            return sorted(tasks, key=lambda x: x["name"].lower(), reverse=reverse)
        elif sort_key == "status":
            return sorted(tasks, key=lambda x: x["completed"], reverse=reverse)
        else:
            return tasks
    
    def toggle_sort_order(self):
        """Đảo thứ tự sắp xếp"""
        self.sort_order = "desc" if self.sort_order == "asc" else "asc"
        self.sort_order_btn.config(text="⬇️" if self.sort_order == "desc" else "⬆️")
        self.refresh_task_list()
    
    def update_stats(self):
        """Cập nhật thống kê"""
        total = len(self.tasks)
        completed = sum(1 for t in self.tasks if t["completed"])
        pending = total - completed
        
        # Đếm quá hạn
        overdue = 0
        for t in self.tasks:
            if not t["completed"] and t["deadline"]:
                deadline = datetime.strptime(t["deadline"], "%H:%M %d/%m/%Y")
                if deadline < datetime.now():
                    overdue += 1
        
        stats_text = f"📈 Tổng: {total} | ✅ Hoàn thành: {completed} | ⏳ Đang thực hiện: {pending} | ⚠️ Quá hạn: {overdue}"
        
        # Tính phần trăm hoàn thành
        if total > 0:
            percent = (completed / total) * 100
            stats_text += f" | 📊 Tiến độ: {percent:.1f}%"
        
        self.stats_label.config(text=stats_text)
    
    def on_tree_click(self, event):
        """Xử lý click vào treeview"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)
            
            if item and column == "#6":  # Cột Actions
                # Lấy task id
                values = self.tree.item(item, "values")
                if values:
                    # Tìm task tương ứng
                    task_name = values[1]
                    task = next((t for t in self.tasks if t["name"] == task_name), None)
                    
                    if task:
                        # Tạo menu popup
                        menu = tk.Menu(self.root, tearoff=0)
                        menu.add_command(label="✅ Đánh dấu hoàn thành", 
                                       command=lambda: self.toggle_complete(task["id"]))
                        menu.add_command(label="✏️ Chỉnh sửa", 
                                       command=lambda: self.edit_task(task["id"]))
                        menu.add_separator()
                        menu.add_command(label="🗑️ Xóa", 
                                       command=lambda: self.delete_task(task["id"]))
                        menu.add_command(label="❌ Hủy", 
                                       command=lambda: self.do_nothing())
                        
                        menu.post(event.x_root, event.y_root)
    
    def check_deadlines(self):
        """Kiểm tra deadline và gửi thông báo"""
        while self.checking_deadlines:
            now = datetime.now()
            
            for task in self.tasks:
                if not task["completed"] and task["deadline"]:
                    deadline = datetime.strptime(task["deadline"], "%H:%M %d/%m/%Y")
                    time_diff = deadline - now
                    
                    # Thông báo khi còn 1 giờ, 1 ngày, hoặc quá hạn
                    if 0 < time_diff.total_seconds() <= 3600:  # 1 giờ
                        self.send_notification(
                            "⏰ Sắp đến deadline!",
                            f"Công việc '{task['name']}' sẽ hết hạn trong 1 giờ tới!"
                        )
                    elif 0 < time_diff.total_seconds() <= 86400:  # 1 ngày
                        self.send_notification(
                            "⚠️ Nhắc nhở deadline",
                            f"Công việc '{task['name']}' sẽ hết hạn trong 1 ngày nữa!"
                        )
                    elif time_diff.total_seconds() < 0:  # Quá hạn
                        self.send_notification(
                            "❗ Công việc đã quá hạn!",
                            f"Công việc '{task['name']}' đã quá deadline!"
                        )
            
            # Kiểm tra mỗi 30 phút
            time.sleep(1800)
    
    def send_notification(self, title, message):
        """Gửi thông báo desktop"""
        try:
            notification.notify(
                title=title,
                message=message,
                app_name="Advanced Checklist",
                timeout=10
            )
        except:
            # Nếu không có thư viện plyer, hiển thị messagebox
            print(f"🔔 {title}: {message}")
    
    def start_deadline_checker(self):
        """Bắt đầu thread kiểm tra deadline"""
        checker_thread = threading.Thread(target=self.check_deadlines, daemon=True)
        checker_thread.start()
    
    def save_data(self):
        """Lưu dữ liệu vào file JSON"""
        data_to_save = []
        for task in self.tasks:
            task_copy = task.copy()
            if "deadline_obj" in task_copy:
                del task_copy["deadline_obj"]
            data_to_save.append(task_copy)
        
        with open(self.data_file, "w", encoding="utf-8") as f:
            json.dump(data_to_save, f, ensure_ascii=False, indent=2)
    
    def load_data(self):
        """Tải dữ liệu từ file JSON"""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, "r", encoding="utf-8") as f:
                    loaded_data = json.load(f)
                    
                    # Khôi phục deadline_obj
                    for task in loaded_data:
                        if task["deadline"]:
                            try:
                                task["deadline_obj"] = datetime.strptime(task["deadline"], "%H:%M %d/%m/%Y")
                            except:
                                task["deadline_obj"] = None
                        else:
                            task["deadline_obj"] = None
                    
                    self.tasks = loaded_data
            except:
                self.tasks = []
        else:
            # Tạo dữ liệu mẫu
            self.tasks = [
                {
                    "id": 1,
                    "name": "1",
                    "description": "placeholder",
                    "priority": "Cao",
                    "deadline": (datetime.now() + timedelta(days=2)).strftime("%H:%M %d/%m/%Y"),
                    "deadline_obj": datetime.now() + timedelta(days=2),
                    "completed": False,
                    "created_at": datetime.now().strftime("%H:%M %d/%m/%Y")
                },
                {
                    "id": 2,
                    "name": "2",
                    "description": "placeholder",
                    "priority": "Trung bình",
                    "deadline": (datetime.now() + timedelta(days=5)).strftime("%H:%M %d/%m/%Y"),
                    "deadline_obj": datetime.now() + timedelta(days=5),
                    "completed": False,
                    "created_at": datetime.now().strftime("%H:%M %d/%m/%Y")
                },
                {
                    "id": 3,
                    "name": "3",
                    "description": "placeholder",
                    "priority": "Thấp",
                    "deadline": "",
                    "deadline_obj": None,
                    "completed": True,
                    "created_at": datetime.now().strftime("%H:%M %d/%m/%Y")
                }
            ]

# Chạy ứng dụng
if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedChecklistApp(root)
    root.mainloop()