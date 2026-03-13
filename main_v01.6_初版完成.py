import customtkinter as ctk
from tkinter import filedialog, messagebox, Menu
import re
import os
import json

ctk.set_appearance_mode("dark")

class DocumentWriterV16(ctk.CTk):
    def __init__(self):
        super().__init__()
        # 修正 2：開啟顯示名稱
        self.app_title = "文件撰寫工具_v01.6"
        self.title(self.app_title)
        self.geometry("1300x850")

        # --- 狀態變數 ---
        self.current_file_path = None
        self.is_dirty = False
        self.raw_content = "" 
        self.project_list = [] 
        self.temp_edit_content = "" 

        # --- 佈局配置 ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- 左側側邊欄 ---
        self.sidebar = ctk.CTkFrame(self, width=240)
        self.sidebar.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        self.btn_import = ctk.CTkButton(self.sidebar, text="➕匯入文件", height=40, font=("Microsoft JhengHei", 15, "bold"), 
                                        fg_color="#A44A3F", hover_color="#8C3A30", command=self.import_document)
        self.btn_import.pack(pady=(20, 5), padx=20, fill="x")

        self.btn_new_file = ctk.CTkButton(self.sidebar, text="📝新增文件", height=40, font=("Microsoft JhengHei", 15, "bold"), 
                                          fg_color="#3D5A80", hover_color="#293D55", command=self.create_new_file_in_project)
        self.btn_new_file.pack(pady=5, padx=20, fill="x")

        ctk.CTkLabel(self.sidebar, text="🎨 排版設定", font=("Microsoft JhengHei", 14)).pack(pady=(10, 0))
        
        self.mode_switch = ctk.CTkSwitch(self.sidebar, text="切換模式", command=self.toggle_mode)
        self.mode_switch.pack(pady=10)

        self.indent_switch = ctk.CTkSwitch(self.sidebar, text="預設縮排", command=self.refresh_reading_view)
        self.indent_switch.deselect() 
        self.indent_switch.configure(state="disabled") 
        self.indent_switch.pack(pady=10)

        ctk.CTkLabel(self.sidebar, text="📁 專案區域 (右鍵編輯)", font=("Microsoft JhengHei", 14, "bold")).pack(pady=(20, 5))
        self.project_frame = ctk.CTkScrollableFrame(self.sidebar, height=350, fg_color="#262626")
        self.project_frame.pack(pady=5, padx=10, fill="both", expand=True)

        self.setup_export_zone()

        # --- 右側主編輯區 ---
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(1, weight=1)

        self.edit_toolbar = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.edit_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        
        btn_w = 100
        ctk.CTkButton(self.edit_toolbar, text="📂開啟檔案", width=btn_w, command=self.open_new_project, fg_color="#4361EE").pack(side="left", padx=2)
        ctk.CTkButton(self.edit_toolbar, text="💾儲存檔案", width=btn_w, command=self.save_current_file).pack(side="left", padx=2)
        ctk.CTkButton(self.edit_toolbar, text="💾另存新檔", width=btn_w, command=self.save_as_file).pack(side="left", padx=2)
        ctk.CTkButton(self.edit_toolbar, text="➕追加檔案", width=btn_w, command=self.append_file, fg_color="#B56576").pack(side="left", padx=2)
        ctk.CTkButton(self.edit_toolbar, text="🗑️清除內容", width=btn_w, command=self.clear_all, fg_color="#6C757D").pack(side="left", padx=2)

        self.textbox = ctk.CTkTextbox(self.main_container, font=("Microsoft JhengHei", 18), undo=True, wrap="word", padx=45, pady=35)
        self.textbox.grid(row=1, column=0, sticky="nsew")
        self.textbox.bind("<<Modified>>", self.on_content_changed)

        self.setup_context_menu()
        self.right_clicked_item_path = None

    # --- 核心邏輯 ---

    def toggle_mode(self):
        is_read_mode = self.mode_switch.get()
        if is_read_mode:
            self.edit_toolbar.grid_forget()
            self.temp_edit_content = self.textbox.get("0.0", "end-1c")
            self.indent_switch.configure(state="normal")
            self.refresh_reading_view()
            self.textbox.configure(state="disabled", fg_color="#2B2B2B")
            self.mode_switch.configure(text="閱讀模式")
        else:
            self.edit_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 5))
            self.indent_switch.configure(state="disabled")
            self.textbox.configure(state="normal", fg_color="#1E1E1E")
            self.textbox.delete("0.0", "end")
            self.textbox.insert("0.0", self.temp_edit_content)
            self.mode_switch.configure(text="編輯模式")

    def apply_formatting(self, text):
        text = re.sub(r'【([一二三四五六七八九十]+)】', r'第\1章', text)
        lines = text.split('\n')
        formatted_lines = []
        for line in lines:
            stripped = line.strip()
            if stripped:
                if "=====" in stripped: 
                    formatted_lines.append("\n　　◈ ◈ ◈ ◈ ◈ ◈\n")
                elif self.indent_switch.get(): 
                    formatted_lines.append(f"　　{stripped}")
                else: 
                    formatted_lines.append(stripped)
            else: 
                formatted_lines.append("")
        return "\n".join(formatted_lines)

    def refresh_reading_view(self):
        if self.mode_switch.get():
            self.textbox.configure(state="normal")
            clean_text = self.apply_formatting(self.temp_edit_content)
            self.textbox.delete("0.0", "end")
            self.textbox.insert("0.0", clean_text)
            self.textbox.configure(state="disabled")

    def update_title(self):
        name = os.path.basename(self.current_file_path) if self.current_file_path else "未命名"
        star = "*" if self.is_dirty else ""
        self.title(f"{name}{star} - {self.app_title}")

    # --- 檔案操作 ---

    def clear_all(self):
        if not self.check_unsaved(): return
        self.textbox.delete("0.0", "end")
        self.current_file_path = None
        self.raw_content = ""; self.is_dirty = False
        self.update_title()

    def on_content_changed(self, event=None):
        if self.textbox.edit_modified():
            current_text = self.textbox.get("0.0", "end-1c")
            self.is_dirty = (current_text != self.raw_content)
            self.update_title()
            self.textbox.edit_modified(False)

    def import_document(self):
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if path:
            self.add_to_project_ui(os.path.basename(path), path)
            self.load_any_file(path)

    def create_new_file_in_project(self):
        if not self.check_unsaved(): return
        name = f"新文件_{len(self.project_list)+1}.txt"
        self.textbox.delete("0.0", "end")
        self.current_file_path = None; self.raw_content = ""; self.is_dirty = False
        self.add_to_project_ui(name, "")
        self.update_title()

    def load_any_file(self, path):
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f: content = f.read()
            self.raw_content = content
            self.textbox.delete("0.0", "end")
            self.textbox.insert("0.0", content)
            self.current_file_path = path; self.is_dirty = False
            self.textbox.edit_modified(False); self.update_title()
        except: pass

    def save_current_file(self):
        if not self.current_file_path: return self.save_as_file()
        content = self.textbox.get("0.0", "end-1c") if not self.mode_switch.get() else self.temp_edit_content
        with open(self.current_file_path, "w", encoding="utf-8") as f: f.write(content)
        self.raw_content = content; self.is_dirty = False
        self.update_title(); messagebox.showinfo("成功", "已儲存檔案")

    def save_as_file(self):
        path = filedialog.asksaveasfilename(defaultextension=".txt")
        if path: self.current_file_path = path; self.save_current_file()

    def check_unsaved(self):
        if self.is_dirty:
            res = messagebox.askyesnocancel("警告", "尚未儲存，是否存檔？")
            if res is True: self.save_current_file(); return True
            return res is False
        return True

    def open_new_project(self):
        if not self.check_unsaved(): return
        path = filedialog.askopenfilename()
        if path: self.load_any_file(path)

    def switch_project_file(self, path):
        if not self.check_unsaved(): return
        self.load_any_file(path)

    # --- UI 輔助與專案管理 ---

    def setup_context_menu(self):
        self.context_menu = Menu(self, tearoff=0, bg="#2B2B2B", fg="white", activebackground="#4361EE")
        self.context_menu.add_command(label="移除專案", command=self.menu_remove)
        self.context_menu.add_command(label="修改名稱", command=self.menu_rename)
        self.context_menu.add_command(label="複製專案", command=self.menu_duplicate)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="上移", command=lambda: self.menu_move(-1))
        self.context_menu.add_command(label="下移", command=lambda: self.menu_move(1))

    def show_context_menu(self, event, path):
        self.right_clicked_item_path = path
        self.context_menu.post(event.x_root, event.y_root)

    def add_to_project_ui(self, name, path):
        btn = ctk.CTkButton(self.project_frame, text=f"📄{name}", anchor="w", fg_color="transparent", hover_color="#3D3D3D")
        btn.configure(command=lambda p=path: self.switch_project_file(p))
        btn.bind("<Button-3>", lambda e, p=path: self.show_context_menu(e, p))
        btn.pack(fill="x", pady=2)
        self.project_list.append({"name": name, "path": path, "widget": btn})

    def menu_remove(self):
        self.project_list = [i for i in self.project_list if i["path"] != self.right_clicked_item_path]
        self.rebuild_project_ui()

    def menu_rename(self):
        if not self.right_clicked_item_path: return
        new_name = ctk.CTkInputDialog(text="新檔名:", title="修改名稱").get_input()
        if new_name:
            if not new_name.endswith(".txt"): new_name += ".txt"
            old_path = self.right_clicked_item_path
            new_path = os.path.join(os.path.dirname(old_path), new_name)
            try:
                os.rename(old_path, new_path)
                for item in self.project_list:
                    if item["path"] == old_path: item.update({"name": new_name, "path": new_path})
                self.rebuild_project_ui()
            except Exception as e: messagebox.showerror("錯誤", f"重新命名失敗: {e}")

    def menu_duplicate(self):
        target = next((i for i in self.project_list if i["path"] == self.right_clicked_item_path), None)
        if target:
            self.load_any_file(target["path"])
            self.current_file_path = None; self.is_dirty = True
            self.add_to_project_ui(f"副本_{target['name']}", "")

    def menu_move(self, dir):
        idx = next(i for i, v in enumerate(self.project_list) if v["path"] == self.right_clicked_item_path)
        new_idx = idx + dir
        if 0 <= new_idx < len(self.project_list):
            self.project_list[idx], self.project_list[new_idx] = self.project_list[new_idx], self.project_list[idx]
            self.rebuild_project_ui()

    def rebuild_project_ui(self):
        for w in self.project_frame.winfo_children(): w.destroy()
        temp = self.project_list.copy(); self.project_list = []
        for item in temp: self.add_to_project_ui(item["name"], item["path"])

    def setup_export_zone(self):
        ctk.CTkLabel(self.sidebar, text="📤 匯出選項", font=("Microsoft JhengHei", 12)).pack(pady=(10, 0))
        ctk.CTkButton(self.sidebar, text="Word (.docx)", height=24, command=self.export_word, fg_color="#2D6A4F").pack(pady=2, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="專案 (.json)", height=24, command=self.save_json, fg_color="#7209B7").pack(pady=2, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="原始稿 (.txt)", height=24, command=self.export_txt, fg_color="#555555").pack(pady=2, padx=20, fill="x")

    def export_word(self):
        # 修正 1：匯出名稱
        p = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=f"{self.app_title}.docx")
        if p:
            from docx import Document
            doc = Document()
            content = self.temp_edit_content if self.mode_switch.get() else self.textbox.get("0.0", "end-1c")
            for line in self.apply_formatting(content).split('\n'): doc.add_paragraph(line)
            doc.save(p); messagebox.showinfo("成功", "Word 匯出成功")

    def save_json(self):
        p = filedialog.asksaveasfilename(defaultextension=".json", initialfile=f"{self.app_title}.json")
        if p:
            with open(p, "w", encoding="utf-8") as f:
                json.dump({"content": self.textbox.get("0.0", "end-1c")}, f, ensure_ascii=False)

    def export_txt(self):
        p = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=f"{self.app_title}.txt")
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write(self.textbox.get("0.0", "end-1c"))

    def append_file(self):
        p = filedialog.askopenfilename()
        if p:
            try:
                with open(p, "r", encoding="utf-8") as f: self.textbox.insert("end", "\n" + f.read())
                self.is_dirty = True; self.update_title()
            except: pass

if __name__ == "__main__":
    app = DocumentWriterV16()
    app.mainloop()