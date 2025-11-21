import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import json
import os
from datetime import datetime

class NovelWriterPro:
    def __init__(self, root):
        self.root = root
        self.root.title("NovelWritingSoftware - 小説執筆ソフト")
        self.root.geometry("1400x900")
        
        # データ構造
        self.projects = []  # 複数プロジェクト管理
        self.current_project_idx = None
        self.current_project = {
            "name": "新規プロジェクト",
            "chapters": [],
            "characters": [],
            "settings": [],
            "current_chapter": None,
            "current_episode": None,
            "writing_goal": 2000,
            "versions": {}  # 下書き・校正版管理
        }
        
        self.auto_save_timer = None
        
        self.setup_ui()
        self.load_projects()
        self.start_auto_save()
    
    def setup_ui(self):
        # メニューバー
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        file_menu.add_command(label="新規プロジェクト", command=self.new_project)
        file_menu.add_command(label="プロジェクトを開く", command=self.open_project_dialog)
        file_menu.add_command(label="保存", command=self.save_all)
        file_menu.add_command(label="エクスポート", command=self.export_project)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.root.quit)
        
        version_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="バージョン", menu=version_menu)
        version_menu.add_command(label="下書きとして保存", command=lambda: self.save_version("draft"))
        version_menu.add_command(label="校正版として保存", command=lambda: self.save_version("proofread"))
        version_menu.add_command(label="バージョン履歴", command=self.show_version_history)
        
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="表示", menu=view_menu)
        view_menu.add_command(label="集中モード", command=self.focus_mode)
        view_menu.add_command(label="通常モード", command=self.normal_mode)
        
        # メインコンテナ
        self.main_container = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # 左サイドバー
        self.setup_sidebar()
        
        # 中央エリア
        self.setup_editor()
        
        # 右サイドバー
        self.setup_tools()
        
        # ステータスバー
        self.setup_statusbar()
    
    def setup_sidebar(self):
        sidebar = ttk.Frame(self.main_container, width=280)
        self.main_container.add(sidebar, weight=0)
        
        # プロジェクト選択
        project_frame = ttk.LabelFrame(sidebar, text="プロジェクト", padding=10)
        project_frame.pack(fill=tk.X, padx=5, pady=5)
        
        btn_row = ttk.Frame(project_frame)
        btn_row.pack(fill=tk.X, pady=5)
        ttk.Button(btn_row, text="新規", command=self.new_project, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="切替", command=self.switch_project, width=8).pack(side=tk.LEFT, padx=2)
        
        self.project_combo = ttk.Combobox(project_frame, state="readonly")
        self.project_combo.pack(fill=tk.X, pady=5)
        self.project_combo.bind("<<ComboboxSelected>>", self.on_project_select)
        
        # タブ切り替え
        tab_control = ttk.Notebook(sidebar)
        tab_control.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 章・話管理タブ
        structure_tab = ttk.Frame(tab_control)
        tab_control.add(structure_tab, text="章・話")
        
        ttk.Label(structure_tab, text="章一覧", font=("", 11, "bold")).pack(pady=5)
        
        chapter_btn = ttk.Frame(structure_tab)
        chapter_btn.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(chapter_btn, text="+ 章追加", command=self.add_chapter).pack(side=tk.LEFT, padx=2)
        ttk.Button(chapter_btn, text="削除", command=self.delete_chapter).pack(side=tk.LEFT, padx=2)
        
        self.chapter_listbox = tk.Listbox(structure_tab, height=8)
        self.chapter_listbox.pack(fill=tk.X, padx=5, pady=5)
        self.chapter_listbox.bind("<<ListboxSelect>>", self.on_chapter_select)
        
        ttk.Separator(structure_tab, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        ttk.Label(structure_tab, text="話一覧", font=("", 11, "bold")).pack(pady=5)
        
        episode_btn = ttk.Frame(structure_tab)
        episode_btn.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(episode_btn, text="+ 話追加", command=self.add_episode).pack(side=tk.LEFT, padx=2)
        ttk.Button(episode_btn, text="削除", command=self.delete_episode).pack(side=tk.LEFT, padx=2)
        
        self.episode_listbox = tk.Listbox(structure_tab, height=8)
        self.episode_listbox.pack(fill=tk.X, padx=5, pady=5)
        self.episode_listbox.bind("<<ListboxSelect>>", self.load_episode)
        
        # キャラクター管理タブ
        character_tab = ttk.Frame(tab_control)
        tab_control.add(character_tab, text="キャラクター")
        
        ttk.Label(character_tab, text="登場人物", font=("", 11, "bold")).pack(pady=5)
        
        char_btn_frame = ttk.Frame(character_tab)
        char_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(char_btn_frame, text="+ 追加", command=self.add_character).pack(side=tk.LEFT, padx=2)
        ttk.Button(char_btn_frame, text="編集", command=self.edit_character).pack(side=tk.LEFT, padx=2)
        ttk.Button(char_btn_frame, text="削除", command=self.delete_character).pack(side=tk.LEFT, padx=2)
        
        self.character_listbox = tk.Listbox(character_tab, height=20)
        self.character_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 世界設定タブ
        setting_tab = ttk.Frame(tab_control)
        tab_control.add(setting_tab, text="世界設定")
        
        ttk.Label(setting_tab, text="設定資料", font=("", 11, "bold")).pack(pady=5)
        
        set_btn_frame = ttk.Frame(setting_tab)
        set_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(set_btn_frame, text="+ 追加", command=self.add_setting).pack(side=tk.LEFT, padx=2)
        ttk.Button(set_btn_frame, text="編集", command=self.edit_setting).pack(side=tk.LEFT, padx=2)
        ttk.Button(set_btn_frame, text="削除", command=self.delete_setting).pack(side=tk.LEFT, padx=2)
        
        self.setting_listbox = tk.Listbox(setting_tab, height=20)
        self.setting_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def setup_editor(self):
        editor_frame = ttk.Frame(self.main_container)
        self.main_container.add(editor_frame, weight=1)
        
        # ツールバー
        toolbar = ttk.Frame(editor_frame)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(toolbar, text="タイトル:").pack(side=tk.LEFT, padx=5)
        self.title_entry = ttk.Entry(toolbar, width=40)
        self.title_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(toolbar, text="保存", command=self.save_current_episode).pack(side=tk.LEFT, padx=5)
        
        # バージョン表示
        self.version_label = ttk.Label(toolbar, text="[作業中]", foreground="blue")
        self.version_label.pack(side=tk.LEFT, padx=20)
        
        # フォント設定
        ttk.Label(toolbar, text="サイズ:").pack(side=tk.LEFT, padx=5)
        self.font_size = ttk.Combobox(toolbar, values=[10, 12, 14, 16, 18, 20], width=5)
        self.font_size.set(12)
        self.font_size.pack(side=tk.LEFT, padx=5)
        self.font_size.bind("<<ComboboxSelected>>", lambda e: self.update_font())
        
        # テキストエディタ
        editor_container = ttk.Frame(editor_frame)
        editor_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.text_editor = scrolledtext.ScrolledText(
            editor_container,
            wrap=tk.WORD,
            font=("游明朝", 12),
            undo=True
        )
        self.text_editor.pack(fill=tk.BOTH, expand=True)
        self.text_editor.bind("<KeyRelease>", self.update_word_count)
        
        # 文字数表示
        count_frame = ttk.Frame(editor_frame)
        count_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.word_count_label = ttk.Label(count_frame, text="文字数: 0", font=("", 10))
        self.word_count_label.pack(side=tk.LEFT, padx=10)
        
        self.goal_label = ttk.Label(count_frame, text="目標: 2000", font=("", 10))
        self.goal_label.pack(side=tk.LEFT, padx=10)
        
        self.chapter_count_label = ttk.Label(count_frame, text="章合計: 0", font=("", 10))
        self.chapter_count_label.pack(side=tk.LEFT, padx=10)
    
    def setup_tools(self):
        tools_frame = ttk.Frame(self.main_container, width=300)
        self.main_container.add(tools_frame, weight=0)
        
        tool_tabs = ttk.Notebook(tools_frame)
        tool_tabs.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 文章チェックタブ
        check_tab = ttk.Frame(tool_tabs)
        tool_tabs.add(check_tab, text="文章チェック")
        
        ttk.Label(check_tab, text="文章分析", font=("", 11, "bold")).pack(pady=10)
        ttk.Button(check_tab, text="繰り返し語句検出", command=self.check_repetition).pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(check_tab, text="文体バランス分析", command=self.check_balance).pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(check_tab, text="句読点チェック", command=self.check_punctuation).pack(fill=tk.X, padx=10, pady=5)
        
        self.check_result = scrolledtext.ScrolledText(check_tab, height=15, wrap=tk.WORD)
        self.check_result.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 進捗管理タブ
        progress_tab = ttk.Frame(tool_tabs)
        tool_tabs.add(progress_tab, text="進捗")
        
        ttk.Label(progress_tab, text="執筆進捗", font=("", 11, "bold")).pack(pady=10)
        
        goal_frame = ttk.Frame(progress_tab)
        goal_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(goal_frame, text="1日の目標:").pack(side=tk.LEFT)
        self.goal_entry = ttk.Entry(goal_frame, width=10)
        self.goal_entry.insert(0, "2000")
        self.goal_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(goal_frame, text="設定", command=self.set_goal).pack(side=tk.LEFT)
        
        self.progress_text = scrolledtext.ScrolledText(progress_tab, height=20, wrap=tk.WORD)
        self.progress_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.update_progress()
    
    def setup_statusbar(self):
        statusbar = ttk.Frame(self.root)
        statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_label = ttk.Label(statusbar, text="準備完了", relief=tk.SUNKEN)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2, pady=2)
        
        self.time_label = ttk.Label(statusbar, text="", relief=tk.SUNKEN)
        self.time_label.pack(side=tk.RIGHT, padx=2, pady=2)
        self.update_time()
    
    # プロジェクト管理
    def new_project(self):
        name = tk.simpledialog.askstring("新規プロジェクト", "プロジェクト名を入力:")
        if name:
            project = {
                "name": name,
                "chapters": [],
                "characters": [],
                "settings": [],
                "current_chapter": None,
                "current_episode": None,
                "writing_goal": 2000,
                "versions": {}
            }
            self.projects.append(project)
            self.current_project = project
            self.current_project_idx = len(self.projects) - 1
            self.refresh_project_list()
            self.refresh_ui()
            self.status_label.config(text=f"プロジェクト '{name}' を作成")
    
    def switch_project(self):
        if not self.projects:
            messagebox.showinfo("情報", "プロジェクトがありません")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("プロジェクト選択")
        dialog.geometry("400x300")
        
        listbox = tk.Listbox(dialog)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        for proj in self.projects:
            listbox.insert(tk.END, proj["name"])
        
        def select_project():
            selection = listbox.curselection()
            if selection:
                self.current_project_idx = selection[0]
                self.current_project = self.projects[self.current_project_idx]
                self.refresh_ui()
                self.status_label.config(text=f"切替: {self.current_project['name']}")
                dialog.destroy()
        
        ttk.Button(dialog, text="選択", command=select_project).pack(pady=10)
    
    def on_project_select(self, event):
        idx = self.project_combo.current()
        if idx >= 0:
            self.current_project_idx = idx
            self.current_project = self.projects[idx]
            self.refresh_ui()
    
    def refresh_project_list(self):
        names = [p["name"] for p in self.projects]
        self.project_combo['values'] = names
        if self.current_project_idx is not None:
            self.project_combo.current(self.current_project_idx)
    
    # 章管理
    def add_chapter(self):
        chapter_num = len(self.current_project["chapters"]) + 1
        chapter = {
            "id": datetime.now().strftime("%Y%m%d%H%M%S"),
            "title": f"第{chapter_num}章",
            "episodes": []
        }
        self.current_project["chapters"].append(chapter)
        self.refresh_chapters()
        self.status_label.config(text=f"{chapter['title']}を追加")
    
    def delete_chapter(self):
        selection = self.chapter_listbox.curselection()
        if selection:
            if messagebox.askyesno("確認", "選択した章とその話をすべて削除しますか?"):
                idx = selection[0]
                del self.current_project["chapters"][idx]
                self.refresh_chapters()
                self.episode_listbox.delete(0, tk.END)
    
    def on_chapter_select(self, event):
        selection = self.chapter_listbox.curselection()
        if selection:
            self.current_project["current_chapter"] = selection[0]
            self.refresh_episodes()
    
    def refresh_chapters(self):
        self.chapter_listbox.delete(0, tk.END)
        for ch in self.current_project["chapters"]:
            episode_count = len(ch["episodes"])
            self.chapter_listbox.insert(tk.END, f"{ch['title']} ({episode_count}話)")
    
    # 話管理
    def add_episode(self):
        if self.current_project["current_chapter"] is None:
            messagebox.showwarning("警告", "先に章を選択してください")
            return
        
        ch_idx = self.current_project["current_chapter"]
        chapter = self.current_project["chapters"][ch_idx]
        
        episode_num = len(chapter["episodes"]) + 1
        episode = {
            "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
            "title": f"{chapter['title']} - 第{episode_num}話",
            "content": "",
            "memo": "",
            "word_count": 0
        }
        chapter["episodes"].append(episode)
        self.refresh_episodes()
        self.refresh_chapters()
        self.status_label.config(text=f"{episode['title']}を追加")
    
    def delete_episode(self):
        selection = self.episode_listbox.curselection()
        if selection and self.current_project["current_chapter"] is not None:
            if messagebox.askyesno("確認", "選択した話を削除しますか?"):
                ch_idx = self.current_project["current_chapter"]
                ep_idx = selection[0]
                del self.current_project["chapters"][ch_idx]["episodes"][ep_idx]
                self.refresh_episodes()
                self.refresh_chapters()
                self.text_editor.delete(1.0, tk.END)
                self.title_entry.delete(0, tk.END)
    
    def load_episode(self, event):
        selection = self.episode_listbox.curselection()
        if selection and self.current_project["current_chapter"] is not None:
            ch_idx = self.current_project["current_chapter"]
            ep_idx = selection[0]
            
            self.save_current_episode()
            
            self.current_project["current_episode"] = ep_idx
            chapter = self.current_project["chapters"][ch_idx]
            episode = chapter["episodes"][ep_idx]
            
            self.title_entry.delete(0, tk.END)
            self.title_entry.insert(0, episode["title"])
            
            self.text_editor.delete(1.0, tk.END)
            self.text_editor.insert(1.0, episode["content"])
            
            self.update_word_count()
            self.update_chapter_count()
    
    def save_current_episode(self):
        if (self.current_project["current_chapter"] is not None and 
            self.current_project["current_episode"] is not None):
            ch_idx = self.current_project["current_chapter"]
            ep_idx = self.current_project["current_episode"]
            
            if ch_idx < len(self.current_project["chapters"]):
                chapter = self.current_project["chapters"][ch_idx]
                if ep_idx < len(chapter["episodes"]):
                    episode = chapter["episodes"][ep_idx]
                    episode["title"] = self.title_entry.get()
                    episode["content"] = self.text_editor.get(1.0, tk.END).strip()
                    episode["word_count"] = len(episode["content"])
                    
                    self.refresh_episodes()
                    self.refresh_chapters()
    
    def refresh_episodes(self):
        self.episode_listbox.delete(0, tk.END)
        if self.current_project["current_chapter"] is not None:
            ch_idx = self.current_project["current_chapter"]
            chapter = self.current_project["chapters"][ch_idx]
            for ep in chapter["episodes"]:
                self.episode_listbox.insert(tk.END, f"{ep['title']} ({ep['word_count']}字)")
    
    # バージョン管理
    def save_version(self, version_type):
        if (self.current_project["current_chapter"] is None or 
            self.current_project["current_episode"] is None):
            messagebox.showwarning("警告", "話を選択してください")
            return
        
        self.save_current_episode()
        
        ch_idx = self.current_project["current_chapter"]
        ep_idx = self.current_project["current_episode"]
        chapter = self.current_project["chapters"][ch_idx]
        episode = chapter["episodes"][ep_idx]
        
        version_key = f"{ch_idx}_{ep_idx}"
        if version_key not in self.current_project["versions"]:
            self.current_project["versions"][version_key] = []
        
        version_name = "下書き" if version_type == "draft" else "校正版"
        timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        
        version = {
            "type": version_type,
            "name": f"{version_name} - {timestamp}",
            "content": episode["content"],
            "timestamp": timestamp
        }
        
        self.current_project["versions"][version_key].append(version)
        self.save_all()
        
        messagebox.showinfo("保存完了", f"{version_name}として保存しました")
        self.status_label.config(text=f"{version_name}保存: {episode['title']}")
    
    def show_version_history(self):
        if (self.current_project["current_chapter"] is None or 
            self.current_project["current_episode"] is None):
            messagebox.showwarning("警告", "話を選択してください")
            return
        
        ch_idx = self.current_project["current_chapter"]
        ep_idx = self.current_project["current_episode"]
        version_key = f"{ch_idx}_{ep_idx}"
        
        versions = self.current_project["versions"].get(version_key, [])
        
        if not versions:
            messagebox.showinfo("情報", "バージョン履歴がありません")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("バージョン履歴")
        dialog.geometry("600x400")
        
        listbox = tk.Listbox(dialog)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        for ver in versions:
            listbox.insert(tk.END, ver["name"])
        
        def load_version():
            selection = listbox.curselection()
            if selection:
                ver = versions[selection[0]]
                if messagebox.askyesno("確認", "このバージョンを読み込みますか?\n現在の内容は上書きされます。"):
                    self.text_editor.delete(1.0, tk.END)
                    self.text_editor.insert(1.0, ver["content"])
                    self.version_label.config(text=f"[{ver['name']}]")
                    dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="読み込み", command=load_version).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="閉じる", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    # キャラクター管理
    def add_character(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("キャラクター追加")
        dialog.geometry("450x350")
        
        ttk.Label(dialog, text="名前:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        name_entry = ttk.Entry(dialog, width=30)
        name_entry.grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="年齢:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        age_entry = ttk.Entry(dialog, width=30)
        age_entry.grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="性格:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        personality_entry = ttk.Entry(dialog, width=30)
        personality_entry.grid(row=2, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="背景:").grid(row=3, column=0, padx=10, pady=10, sticky=tk.NW)
        background_text = scrolledtext.ScrolledText(dialog, width=30, height=6)
        background_text.grid(row=3, column=1, padx=10, pady=10)
        
        def save_character():
            character = {
                "name": name_entry.get(),
                "age": age_entry.get(),
                "personality": personality_entry.get(),
                "background": background_text.get(1.0, tk.END).strip()
            }
            self.current_project["characters"].append(character)
            self.refresh_characters()
            dialog.destroy()
        
        ttk.Button(dialog, text="保存", command=save_character).grid(row=4, column=0, columnspan=2, pady=20)
    
    def edit_character(self):
        selection = self.character_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "キャラクターを選択してください")
            return
        
        idx = selection[0]
        char = self.current_project["characters"][idx]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("キャラクター情報")
        dialog.geometry("500x400")
        
        info_text = scrolledtext.ScrolledText(dialog, wrap=tk.WORD)
        info_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        info = f"""名前: {char['name']}
年齢: {char['age']}
性格: {char['personality']}

背景:
{char['background']}
"""
        info_text.insert(1.0, info)
    
    def delete_character(self):
        selection = self.character_listbox.curselection()
        if selection:
            if messagebox.askyesno("確認", "選択したキャラクターを削除しますか?"):
                del self.current_project["characters"][selection[0]]
                self.refresh_characters()
    
    def refresh_characters(self):
        self.character_listbox.delete(0, tk.END)
        for char in self.current_project["characters"]:
            self.character_listbox.insert(tk.END, char["name"])
    
    # 世界設定管理
    def add_setting(self):
        name = tk.simpledialog.askstring("世界設定", "設定名を入力:")
        if name:
            detail = tk.simpledialog.askstring("世界設定", "詳細を入力:")
            setting = {"name": name, "detail": detail or ""}
            self.current_project["settings"].append(setting)
            self.refresh_settings()
    
    def edit_setting(self):
        selection = self.setting_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "設定を選択してください")
            return
        
        idx = selection[0]
        setting = self.current_project["settings"][idx]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("設定詳細")
        dialog.geometry("500x300")
        
        ttk.Label(dialog, text=setting["name"], font=("", 12, "bold")).pack(pady=10)
        
        detail_text = scrolledtext.ScrolledText(dialog, wrap=tk.WORD)
        detail_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        detail_text.insert(1.0, setting["detail"])
    
    def delete_setting(self):
        selection = self.setting_listbox.curselection()
        if selection:
            if messagebox.askyesno("確認", "選択した設定を削除しますか?"):
                del self.current_project["settings"][selection[0]]
                self.refresh_settings()
    
    def refresh_settings(self):
        self.setting_listbox.delete(0, tk.END)
        for setting in self.current_project["settings"]:
            self.setting_listbox.insert(tk.END, setting["name"])
    
    # 文章チェック機能
    def check_repetition(self):
        text = self.text_editor.get(1.0, tk.END)
        words = {}
        
        # 3文字以上の繰り返しをチェック
        for i in range(len(text) - 2):
            word = text[i:i+3]
            if word.strip() and not word.isspace():
                words[word] = words.get(word, 0) + 1
        
        repeated = [(w, c) for w, c in words.items() if c > 3]
        repeated.sort(key=lambda x: x[1], reverse=True)
        
        self.check_result.delete(1.0, tk.END)
        if repeated:
            self.check_result.insert(tk.END, "繰り返し語句検出:\n\n")
            for word, count in repeated[:15]:
                self.check_result.insert(tk.END, f"「{word}」: {count}回\n")
        else:
            self.check_result.insert(tk.END, "目立った繰り返しは検出されませんでした。")
    
    def check_balance(self):
        text = self.text_editor.get(1.0, tk.END).strip()
        
        if not text:
            self.check_result.delete(1.0, tk.END)
            self.check_result.insert(tk.END, "テキストがありません。")
            return
        
        kanji_count = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
        hiragana_count = sum(1 for c in text if '\u3040' <= c <= '\u309f')
        katakana_count = sum(1 for c in text if '\u30a0' <= c <= '\u30ff')
        
        total = len(text)
        
        self.check_result.delete(1.0, tk.END)
        result = f"""文字種バランス分析:

漢字: {kanji_count}文字 ({kanji_count/total*100:.1f}%)
ひらがな: {hiragana_count}文字 ({hiragana_count/total*100:.1f}%)
カタカナ: {katakana_count}文字 ({katakana_count/total*100:.1f}%)

総文字数: {total}

【推奨バランス】
漢字: 20-30%
ひらがな: 60-70%
カタカナ: 5-10%

【評価】
"""
        
        # 簡易評価
        if 20 <= kanji_count/total*100 <= 30:
            result += "✓ 漢字バランス良好\n"
        else:
            result += "△ 漢字の割合を調整してみてください\n"
        
        if 60 <= hiragana_count/total*100 <= 70:
            result += "✓ ひらがなバランス良好\n"
        else:
            result += "△ ひらがなの割合を調整してみてください\n"
        
        self.check_result.insert(tk.END, result)
    
    def check_punctuation(self):
        text = self.text_editor.get(1.0, tk.END)
        
        # 句読点の連続をチェック
        issues = []
        
        for i in range(len(text) - 1):
            if text[i] == '、' and text[i+1] == '、':
                issues.append(f"行{self.get_line_number(text, i)}: 読点が連続")
            if text[i] == '。' and text[i+1] == '。':
                issues.append(f"行{self.get_line_number(text, i)}: 句点が連続")
        
        comma_count = text.count('、')
        period_count = text.count('。')
        
        self.check_result.delete(1.0, tk.END)
        self.check_result.insert(tk.END, f"句読点チェック:\n\n")
        self.check_result.insert(tk.END, f"読点（、）: {comma_count}個\n")
        self.check_result.insert(tk.END, f"句点（。）: {period_count}個\n\n")
        
        if issues:
            self.check_result.insert(tk.END, "【検出された問題】\n")
            for issue in issues:
                self.check_result.insert(tk.END, f"- {issue}\n")
        else:
            self.check_result.insert(tk.END, "問題は検出されませんでした。")
    
    def get_line_number(self, text, pos):
        return text[:pos].count('\n') + 1
    
    # 進捗管理
    def update_word_count(self, event=None):
        text = self.text_editor.get(1.0, tk.END).strip()
        count = len(text)
        self.word_count_label.config(text=f"文字数: {count}")
        
        goal = self.current_project["writing_goal"]
        progress = min(100, (count / goal * 100))
        self.goal_label.config(text=f"目標: {goal} ({progress:.1f}%)")
    
    def update_chapter_count(self):
        if self.current_project["current_chapter"] is not None:
            ch_idx = self.current_project["current_chapter"]
            chapter = self.current_project["chapters"][ch_idx]
            total = sum(ep.get("word_count", 0) for ep in chapter["episodes"])
            self.chapter_count_label.config(text=f"章合計: {total}")
    
    def set_goal(self):
        try:
            goal = int(self.goal_entry.get())
            self.current_project["writing_goal"] = goal
            self.update_progress()
            messagebox.showinfo("設定完了", f"目標文字数を{goal}に設定しました")
        except ValueError:
            messagebox.showerror("エラー", "数値を入力してください")
    
    def update_progress(self):
        total_words = 0
        chapter_stats = []
        
        for ch in self.current_project["chapters"]:
            ch_total = sum(ep.get("word_count", 0) for ep in ch["episodes"])
            total_words += ch_total
            chapter_stats.append((ch["title"], ch_total, len(ch["episodes"])))
        
        goal = self.current_project["writing_goal"]
        
        self.progress_text.delete(1.0, tk.END)
        self.progress_text.insert(tk.END, f"【プロジェクト進捗】\n")
        self.progress_text.insert(tk.END, f"プロジェクト名: {self.current_project['name']}\n\n")
        self.progress_text.insert(tk.END, f"総文字数: {total_words:,} 文字\n")
        self.progress_text.insert(tk.END, f"今日の目標: {goal:,} 文字\n")
        self.progress_text.insert(tk.END, f"進捗率: {min(100, total_words/goal*100):.1f}%\n\n")
        
        self.progress_text.insert(tk.END, "【章別統計】\n")
        for title, words, ep_count in chapter_stats:
            self.progress_text.insert(tk.END, f"\n{title}\n")
            self.progress_text.insert(tk.END, f"  文字数: {words:,} 文字\n")
            self.progress_text.insert(tk.END, f"  話数: {ep_count}話\n")
    
    # UI機能
    def update_font(self):
        size = int(self.font_size.get())
        self.text_editor.config(font=("游明朝", size))
    
    def focus_mode(self):
        # 集中モード：サイドバーを非表示
        for i in range(2):
            try:
                self.main_container.forget(0)
            except:
                pass
        self.status_label.config(text="集中モード - ESCキーで通常モードに戻ります")
        self.root.bind("<Escape>", lambda e: self.normal_mode())
    
    def normal_mode(self):
        # 通常モードに戻す
        self.main_container.forget(0)
        self.setup_sidebar()
        self.setup_tools()
        self.refresh_ui()
        self.status_label.config(text="通常モード")
        self.root.unbind("<Escape>")
    
    def export_project(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text File", "*.txt"), ("All Files", "*.*")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(f"{self.current_project['name']}\n")
                f.write("="*80 + "\n\n")
                
                for ch in self.current_project["chapters"]:
                    f.write(f"\n{'='*80}\n")
                    f.write(f"{ch['title']}\n")
                    f.write(f"{'='*80}\n\n")
                    
                    for ep in ch["episodes"]:
                        f.write(f"\n{'-'*60}\n")
                        f.write(f"{ep['title']}\n")
                        f.write(f"{'-'*60}\n\n")
                        f.write(ep["content"])
                        f.write("\n\n")
            
            messagebox.showinfo("エクスポート完了", f"{filename}に保存しました")
            self.status_label.config(text=f"エクスポート完了: {os.path.basename(filename)}")
    
    # データ保存・読込
    def save_all(self):
        self.save_current_episode()
        
        data = {
            "projects": self.projects,
            "current_project_idx": self.current_project_idx
        }
        
        with open("novels_data.json", 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        self.status_label.config(text="自動保存完了")
    
    def load_projects(self):
        if os.path.exists("novels_data.json"):
            try:
                with open("novels_data.json", 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.projects = data.get("projects", [])
                    self.current_project_idx = data.get("current_project_idx")
                    
                    if self.projects and self.current_project_idx is not None:
                        if self.current_project_idx < len(self.projects):
                            self.current_project = self.projects[self.current_project_idx]
                    
                    self.refresh_project_list()
                    self.refresh_ui()
            except Exception as e:
                print(f"読込エラー: {e}")
    
    def open_project_dialog(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Novel Data", "*.json"), ("All Files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.projects = data.get("projects", [])
                    self.current_project_idx = data.get("current_project_idx")
                    
                    if self.projects and self.current_project_idx is not None:
                        if self.current_project_idx < len(self.projects):
                            self.current_project = self.projects[self.current_project_idx]
                    
                    self.refresh_project_list()
                    self.refresh_ui()
                    self.status_label.config(text=f"読込完了: {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました:\n{e}")
    
    def start_auto_save(self):
        self.save_all()
        self.auto_save_timer = self.root.after(60000, self.start_auto_save)  # 1分ごと
    
    def update_time(self):
        now = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        self.time_label.config(text=now)
        self.root.after(1000, self.update_time)
    
    def refresh_ui(self):
        # 章リスト更新
        self.refresh_chapters()
        
        # 話リスト更新
        if self.current_project["current_chapter"] is not None:
            self.refresh_episodes()
        
        # キャラクターリスト更新
        self.refresh_characters()
        
        # 世界設定リスト更新
        self.refresh_settings()
        
        # 進捗更新
        self.update_progress()
        
        # プロジェクトリスト更新
        self.refresh_project_list()

if __name__ == "__main__":
    root = tk.Tk()
    app = NovelWriterPro(root)

    root.mainloop()
