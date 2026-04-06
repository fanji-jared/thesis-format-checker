#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import threading
from pathlib import Path
from typing import Optional, Callable, Any, Dict

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .ui_styles import AppStyles, CardFrame, StatCard
from .docx_parser import DocxParser
from .config_manager import ConfigManager, ConfigValidationError
from .ai_config_generator import AIConfigGenerator, GeneratorProgress, GeneratorState, OllamaConfig
from .format_checker import FormatChecker, CheckReport
from .models import DocumentInfo, ParseResult


class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("论文格式检测工具")
        self.root.geometry("900x750")
        self.root.minsize(800, 600)
        
        AppStyles.apply_styles(self.root)
        
        self.docx_parser = DocxParser()
        self.config_manager = ConfigManager()
        self.ai_generator: Optional[AIConfigGenerator] = None
        
        self.current_document: Optional[DocumentInfo] = None
        self.current_config: Optional[Dict] = None
        self.current_config_path: Optional[str] = None
        self.check_report: Optional[CheckReport] = None
        
        self._create_menu()
        self._create_widgets()
        self._load_config_list()
        self._check_ollama_status()
    
    def _create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="打开文档", command=self._select_document)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        
        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="配置", menu=config_menu)
        config_menu.add_command(label="刷新配置列表", command=self._load_config_list)
        config_menu.add_command(label="导入配置", command=self._import_config)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self._show_about)
    
    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, style='TFrame')
        main_frame.pack(fill='both', expand=True, padx=AppStyles.PADDING_LARGE, 
                       pady=AppStyles.PADDING_LARGE)
        
        title_label = ttk.Label(main_frame, text="论文格式检测工具", style='Title.TLabel')
        title_label.pack(pady=(0, AppStyles.PADDING_LARGE))
        
        self._create_file_section(main_frame)
        self._create_config_section(main_frame)
        self._create_ai_section(main_frame)
        self._create_check_section(main_frame)
        self._create_status_bar()
    
    def _create_file_section(self, parent):
        file_card = CardFrame(parent, title="文档选择")
        file_card.pack(fill='x', pady=(0, AppStyles.PADDING_NORMAL))
        
        select_frame = ttk.Frame(file_card.content_frame, style='Card.TFrame')
        select_frame.pack(fill='x')
        
        self.file_path_var = tk.StringVar(value="未选择文档")
        file_entry = ttk.Entry(select_frame, textvariable=self.file_path_var, 
                              state='readonly', width=60)
        file_entry.pack(side='left', fill='x', expand=True, padx=(0, AppStyles.PADDING_NORMAL))
        
        select_btn = ttk.Button(select_frame, text="选择文档", command=self._select_document)
        select_btn.pack(side='right')
        
        self.doc_info_frame = ttk.Frame(file_card.content_frame, style='Card.TFrame')
        self.doc_info_frame.pack(fill='x', pady=(AppStyles.PADDING_NORMAL, 0))
        
        self.doc_info_labels = {}
        info_items = [
            ('file_name', '文件名'),
            ('page_count', '页数'),
            ('paragraph_count', '段落数'),
            ('character_count', '字符数')
        ]
        
        for i, (key, label) in enumerate(info_items):
            ttk.Label(self.doc_info_frame, text=f"{label}:", style='Surface.TLabel',
                     background=AppStyles.COLOR_SURFACE).grid(row=0, column=i*2, padx=(0, 5), sticky='e')
            self.doc_info_labels[key] = ttk.Label(self.doc_info_frame, text="-", 
                                                   style='Surface.TLabel',
                                                   background=AppStyles.COLOR_SURFACE)
            self.doc_info_labels[key].grid(row=0, column=i*2+1, padx=(0, 15), sticky='w')
    
    def _create_config_section(self, parent):
        config_card = CardFrame(parent, title="配置文件选择")
        config_card.pack(fill='x', pady=(0, AppStyles.PADDING_NORMAL))
        
        select_frame = ttk.Frame(config_card.content_frame, style='Card.TFrame')
        select_frame.pack(fill='x')
        
        ttk.Label(select_frame, text="选择配置:", style='Surface.TLabel',
                 background=AppStyles.COLOR_SURFACE).pack(side='left', padx=(0, AppStyles.PADDING_SMALL))
        
        self.config_var = tk.StringVar()
        self.config_combo = ttk.Combobox(select_frame, textvariable=self.config_var,
                                         state='readonly', width=50)
        self.config_combo.pack(side='left', fill='x', expand=True, padx=(0, AppStyles.PADDING_NORMAL))
        self.config_combo.bind('<<ComboboxSelected>>', self._on_config_selected)
        
        refresh_btn = ttk.Button(select_frame, text="刷新", command=self._load_config_list, width=6)
        refresh_btn.pack(side='right')
        
        self.config_info_frame = ttk.Frame(config_card.content_frame, style='Card.TFrame')
        self.config_info_frame.pack(fill='x', pady=(AppStyles.PADDING_NORMAL, 0))
        
        self.config_info_labels = {}
        info_items = [
            ('name', '配置名称'),
            ('version', '版本'),
            ('description', '描述')
        ]
        
        for i, (key, label) in enumerate(info_items):
            row = i // 2
            col = (i % 2) * 2
            ttk.Label(self.config_info_frame, text=f"{label}:", style='Surface.TLabel',
                     background=AppStyles.COLOR_SURFACE).grid(row=row, column=col, padx=(0, 5), sticky='e')
            self.config_info_labels[key] = ttk.Label(self.config_info_frame, text="-",
                                                      style='Surface.TLabel',
                                                      background=AppStyles.COLOR_SURFACE)
            self.config_info_labels[key].grid(row=row, column=col+1, padx=(0, 20), sticky='w')
    
    def _create_ai_section(self, parent):
        ai_card = CardFrame(parent, title="AI配置生成")
        ai_card.pack(fill='x', pady=(0, AppStyles.PADDING_NORMAL))
        
        rules_frame = ttk.Frame(ai_card.content_frame, style='Card.TFrame')
        rules_frame.pack(fill='x')
        
        ttk.Label(rules_frame, text="规则文档:", style='Surface.TLabel',
                 background=AppStyles.COLOR_SURFACE).pack(side='left', padx=(0, AppStyles.PADDING_SMALL))
        
        self.rules_path_var = tk.StringVar(value="未选择")
        rules_entry = ttk.Entry(rules_frame, textvariable=self.rules_path_var,
                               state='readonly', width=40)
        rules_entry.pack(side='left', fill='x', expand=True, padx=(0, AppStyles.PADDING_SMALL))
        
        rules_btn = ttk.Button(rules_frame, text="选择", command=self._select_rules_file, width=6)
        rules_btn.pack(side='left', padx=(0, AppStyles.PADDING_SMALL))
        
        model_frame = ttk.Frame(ai_card.content_frame, style='Card.TFrame')
        model_frame.pack(fill='x', pady=(AppStyles.PADDING_NORMAL, 0))
        
        ttk.Label(model_frame, text="AI模型:", style='Surface.TLabel',
                 background=AppStyles.COLOR_SURFACE).pack(side='left', padx=(0, AppStyles.PADDING_SMALL))
        
        self.model_var = tk.StringVar()
        self.model_combo = ttk.Combobox(model_frame, textvariable=self.model_var,
                                        state='readonly', width=30)
        self.model_combo.pack(side='left', padx=(0, AppStyles.PADDING_NORMAL))
        
        self.ollama_status_label = ttk.Label(model_frame, text="●", style='Surface.TLabel',
                                             background=AppStyles.COLOR_SURFACE)
        self.ollama_status_label.pack(side='left')
        
        generate_frame = ttk.Frame(ai_card.content_frame, style='Card.TFrame')
        generate_frame.pack(fill='x', pady=(AppStyles.PADDING_NORMAL, 0))
        
        self.generate_btn = ttk.Button(generate_frame, text="生成配置", 
                                       command=self._generate_config, style='Primary.TButton')
        self.generate_btn.pack(side='left')
        
        self.ai_progress_var = tk.DoubleVar(value=0)
        self.ai_progress = ttk.Progressbar(generate_frame, variable=self.ai_progress_var,
                                           maximum=100, length=300)
        self.ai_progress.pack(side='left', padx=(AppStyles.PADDING_NORMAL, 0))
        
        self.ai_status_var = tk.StringVar(value="")
        ttk.Label(generate_frame, textvariable=self.ai_status_var, style='Small.TLabel',
                 background=AppStyles.COLOR_SURFACE).pack(side='left', padx=(AppStyles.PADDING_SMALL, 0))
    
    def _create_check_section(self, parent):
        check_card = CardFrame(parent, title="检测控制")
        check_card.pack(fill='x', pady=(0, AppStyles.PADDING_NORMAL))
        
        btn_frame = ttk.Frame(check_card.content_frame, style='Card.TFrame')
        btn_frame.pack(fill='x')
        
        self.check_btn = ttk.Button(btn_frame, text="开始检测", 
                                    command=self._start_check, style='Success.TButton')
        self.check_btn.pack(side='left', padx=(0, AppStyles.PADDING_NORMAL))
        
        self.check_progress_var = tk.DoubleVar(value=0)
        self.check_progress = ttk.Progressbar(btn_frame, variable=self.check_progress_var,
                                              maximum=100, length=400)
        self.check_progress.pack(side='left', fill='x', expand=True)
        
        self.check_status_var = tk.StringVar(value="就绪")
        ttk.Label(btn_frame, textvariable=self.check_status_var, style='Surface.TLabel',
                 background=AppStyles.COLOR_SURFACE).pack(side='left', padx=(AppStyles.PADDING_NORMAL, 0))
    
    def _create_status_bar(self):
        status_frame = ttk.Frame(self.root, style='TFrame')
        status_frame.pack(fill='x', side='bottom')
        
        ttk.Separator(status_frame, orient='horizontal').pack(fill='x')
        
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, style='Small.TLabel')
        status_label.pack(side='left', padx=AppStyles.PADDING_NORMAL, pady=AppStyles.PADDING_SMALL)
    
    def _select_document(self):
        file_path = filedialog.askopenfilename(
            title="选择Word文档",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self._load_document(file_path)
    
    def _load_document(self, file_path: str):
        self.status_var.set("正在加载文档...")
        self.file_path_var.set(file_path)
        
        def load_task():
            try:
                result: ParseResult = self.docx_parser.parse(file_path)
                
                if result.success and result.document_info:
                    self.current_document = result.document_info
                    self.root.after(0, lambda: self._update_doc_info(result.document_info))
                    self.root.after(0, lambda: self.status_var.set("文档加载完成"))
                else:
                    error_msg = result.error_message or "未知错误"
                    self.root.after(0, lambda: messagebox.showerror("错误", f"文档解析失败: {error_msg}"))
                    self.root.after(0, lambda: self.status_var.set("文档加载失败"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"文档解析异常: {str(e)}"))
                self.root.after(0, lambda: self.status_var.set("文档加载失败"))
        
        threading.Thread(target=load_task, daemon=True).start()
    
    def _update_doc_info(self, doc_info: DocumentInfo):
        self.doc_info_labels['file_name'].configure(text=doc_info.file_name[:30] + "..." if len(doc_info.file_name) > 30 else doc_info.file_name)
        self.doc_info_labels['page_count'].configure(text=str(doc_info.page_count))
        self.doc_info_labels['paragraph_count'].configure(text=str(doc_info.paragraph_count))
        self.doc_info_labels['character_count'].configure(text=str(doc_info.character_count))
    
    def _load_config_list(self):
        configs = self.config_manager.get_config_list()
        
        config_names = [f"{c['name']} ({c['version']})" for c in configs]
        self.config_combo['values'] = config_names
        
        self._config_list = configs
        
        if configs:
            self.config_combo.current(0)
            self._on_config_selected(None)
    
    def _on_config_selected(self, event):
        idx = self.config_combo.current()
        if idx >= 0 and idx < len(self._config_list):
            config_info = self._config_list[idx]
            
            self.config_info_labels['name'].configure(text=config_info.get('name', '-'))
            self.config_info_labels['version'].configure(text=config_info.get('version', '-'))
            self.config_info_labels['description'].configure(text=config_info.get('description', '-')[:50])
            
            try:
                self.current_config = self.config_manager.load_config(config_info['path'])
                self.current_config_path = config_info['path']
            except ConfigValidationError as e:
                messagebox.showwarning("配置验证警告", str(e))
                self.current_config = self.config_manager.get_current_config()
            except Exception as e:
                messagebox.showerror("错误", f"加载配置失败: {str(e)}")
    
    def _import_config(self):
        file_path = filedialog.askopenfilename(
            title="导入配置文件",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                self.config_manager.import_config(file_path)
                self._load_config_list()
                messagebox.showinfo("成功", "配置导入成功")
            except Exception as e:
                messagebox.showerror("错误", f"配置导入失败: {str(e)}")
    
    def _check_ollama_status(self):
        def check():
            try:
                import requests
                response = requests.get("http://localhost:11434/api/tags", timeout=5)
                if response.status_code == 200:
                    models = response.json().get("models", [])
                    model_names = [m.get("name", "") for m in models]
                    self.root.after(0, lambda: self._update_ollama_status(True, model_names))
                else:
                    self.root.after(0, lambda: self._update_ollama_status(False, []))
            except:
                self.root.after(0, lambda: self._update_ollama_status(False, []))
        
        threading.Thread(target=check, daemon=True).start()
    
    def _update_ollama_status(self, connected: bool, models: list):
        if connected:
            self.ollama_status_label.configure(foreground=AppStyles.COLOR_SUCCESS)
            self.model_combo['values'] = models
            if models:
                self.model_combo.current(0)
        else:
            self.ollama_status_label.configure(foreground=AppStyles.COLOR_ERROR)
            self.model_combo['values'] = ["Ollama未连接"]
            self.model_combo.current(0)
    
    def _select_rules_file(self):
        file_path = filedialog.askopenfilename(
            title="选择规则文档",
            filetypes=[("Markdown文件", "*.md"), ("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.rules_path_var.set(file_path)
    
    def _generate_config(self):
        rules_path = self.rules_path_var.get()
        if rules_path == "未选择":
            messagebox.showwarning("提示", "请先选择规则文档")
            return
        
        model = self.model_var.get()
        if not model or model == "Ollama未连接":
            messagebox.showwarning("提示", "请确保Ollama服务正在运行")
            return
        
        self.generate_btn.configure(state='disabled')
        self.ai_progress_var.set(0)
        self.ai_status_var.set("正在生成...")
        
        def progress_callback(progress: GeneratorProgress):
            self.root.after(0, lambda: self._update_ai_progress(progress))
        
        ollama_config = OllamaConfig(model=model)
        self.ai_generator = AIConfigGenerator(
            ollama_config=ollama_config,
            config_manager=self.config_manager,
            progress_callback=progress_callback
        )
        
        def generate_task():
            try:
                result = self.ai_generator.generate_config(rules_path)
                self.root.after(0, lambda: self._on_config_generated(result))
            except Exception as e:
                self.root.after(0, lambda: self._on_config_generate_error(str(e)))
        
        threading.Thread(target=generate_task, daemon=True).start()
    
    def _update_ai_progress(self, progress: GeneratorProgress):
        self.ai_progress_var.set(progress.progress * 100)
        self.ai_status_var.set(progress.message)
    
    def _on_config_generated(self, result):
        self.generate_btn.configure(state='normal')
        
        if result.success:
            self.ai_status_var.set("生成成功")
            self._load_config_list()
            messagebox.showinfo("成功", f"配置文件已生成并保存:\n{result.config_path}")
        else:
            self.ai_status_var.set("生成失败")
            messagebox.showerror("错误", f"配置生成失败:\n{result.error_message}")
    
    def _on_config_generate_error(self, error: str):
        self.generate_btn.configure(state='normal')
        self.ai_status_var.set("生成失败")
        messagebox.showerror("错误", f"配置生成异常:\n{error}")
    
    def _start_check(self):
        if not self.current_document:
            messagebox.showwarning("提示", "请先选择要检测的文档")
            return
        
        if not self.current_config:
            messagebox.showwarning("提示", "请先选择配置文件")
            return
        
        self.check_btn.configure(state='disabled')
        self.check_progress_var.set(0)
        self.check_status_var.set("正在检测...")
        
        def check_task():
            try:
                checker = FormatChecker(self.current_config)
                
                steps = [
                    (10, "正在检测页面设置..."),
                    (25, "正在检测目录格式..."),
                    (40, "正在检测章节标题..."),
                    (55, "正在检测正文格式..."),
                    (70, "正在检测图片格式..."),
                    (80, "正在检测表格格式..."),
                    (90, "正在检测公式格式..."),
                    (95, "正在检测参考文献..."),
                ]
                
                for progress, status in steps:
                    self.root.after(0, lambda p=progress, s=status: (
                        self.check_progress_var.set(p),
                        self.check_status_var.set(s)
                    ))
                
                self.check_report = checker.check(self.current_document)
                
                self.root.after(0, lambda: self._on_check_completed())
            except Exception as e:
                self.root.after(0, lambda: self._on_check_error(str(e)))
        
        threading.Thread(target=check_task, daemon=True).start()
    
    def _on_check_completed(self):
        self.check_progress_var.set(100)
        self.check_status_var.set("检测完成")
        self.check_btn.configure(state='normal')
        
        if self.check_report:
            self._show_result_window()
    
    def _on_check_error(self, error: str):
        self.check_progress_var.set(0)
        self.check_status_var.set("检测失败")
        self.check_btn.configure(state='normal')
        messagebox.showerror("错误", f"检测过程出错:\n{error}")
    
    def _show_result_window(self):
        from .ui_result import ResultWindow
        
        result_window = tk.Toplevel(self.root)
        result_window.title("检测结果")
        result_window.geometry("1000x700")
        result_window.minsize(800, 600)
        
        ResultWindow(result_window, self.check_report, self.current_document)
    
    def _show_about(self):
        messagebox.showinfo("关于", 
            "论文格式检测工具 v1.0.0\n\n"
            "基于Python和AI技术的论文格式自动检测工具\n\n"
            "功能:\n"
            "- 自动解析Word文档结构\n"
            "- AI辅助生成检测配置\n"
            "- 多维度格式检测\n"
            "- 详细问题报告导出"
        )
