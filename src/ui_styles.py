#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk
from typing import Dict, Any


class AppStyles:
    COLOR_ERROR = "#E74C3C"
    COLOR_ERROR_LIGHT = "#FADBD8"
    COLOR_WARNING = "#F39C12"
    COLOR_WARNING_LIGHT = "#FCF3CF"
    COLOR_INFO = "#3498DB"
    COLOR_INFO_LIGHT = "#D4E6F1"
    COLOR_SUCCESS = "#27AE60"
    COLOR_SUCCESS_LIGHT = "#D5F5E3"
    COLOR_PRIMARY = "#2C3E50"
    COLOR_PRIMARY_LIGHT = "#5D6D7E"
    COLOR_SECONDARY = "#7F8C8D"
    COLOR_BACKGROUND = "#F8F9FA"
    COLOR_SURFACE = "#FFFFFF"
    COLOR_TEXT = "#2C3E50"
    COLOR_TEXT_LIGHT = "#7F8C8D"
    COLOR_BORDER = "#DEE2E6"
    
    FONT_FAMILY = "Microsoft YaHei UI"
    FONT_FAMILY_MONO = "Consolas"
    FONT_SIZE_TITLE = 14
    FONT_SIZE_HEADING = 12
    FONT_SIZE_NORMAL = 10
    FONT_SIZE_SMALL = 9
    
    PADDING_SMALL = 5
    PADDING_NORMAL = 10
    PADDING_LARGE = 15
    
    @classmethod
    def apply_styles(cls, root: tk.Tk):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('TFrame', background=cls.COLOR_BACKGROUND)
        style.configure('Surface.TFrame', background=cls.COLOR_SURFACE)
        style.configure('Card.TFrame', background=cls.COLOR_SURFACE)
        
        style.configure('TLabel',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_TEXT,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        style.configure('Surface.TLabel',
                       background=cls.COLOR_SURFACE,
                       foreground=cls.COLOR_TEXT,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        style.configure('Title.TLabel',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_TITLE, 'bold'))
        style.configure('Heading.TLabel',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_HEADING, 'bold'))
        style.configure('Small.TLabel',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_TEXT_LIGHT,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_SMALL))
        style.configure('Error.TLabel',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_ERROR,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        style.configure('Warning.TLabel',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_WARNING,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        style.configure('Info.TLabel',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_INFO,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        style.configure('Success.TLabel',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_SUCCESS,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL))
        
        style.configure('TButton',
                       background=cls.COLOR_PRIMARY,
                       foreground=cls.COLOR_SURFACE,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL),
                       padding=(cls.PADDING_NORMAL, cls.PADDING_SMALL))
        style.map('TButton',
                 background=[('active', cls.COLOR_PRIMARY_LIGHT),
                            ('pressed', cls.COLOR_PRIMARY)],
                 foreground=[('active', cls.COLOR_SURFACE)])
        
        style.configure('Primary.TButton',
                       background=cls.COLOR_PRIMARY,
                       foreground=cls.COLOR_SURFACE,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL, 'bold'),
                       padding=(cls.PADDING_LARGE, cls.PADDING_NORMAL))
        style.map('Primary.TButton',
                 background=[('active', cls.COLOR_PRIMARY_LIGHT),
                            ('pressed', cls.COLOR_PRIMARY)])
        
        style.configure('Success.TButton',
                       background=cls.COLOR_SUCCESS,
                       foreground=cls.COLOR_SURFACE,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL, 'bold'),
                       padding=(cls.PADDING_LARGE, cls.PADDING_NORMAL))
        style.map('Success.TButton',
                 background=[('active', '#219A52'),
                            ('pressed', cls.COLOR_SUCCESS)])
        
        style.configure('Danger.TButton',
                       background=cls.COLOR_ERROR,
                       foreground=cls.COLOR_SURFACE,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL, 'bold'),
                       padding=(cls.PADDING_LARGE, cls.PADDING_NORMAL))
        style.map('Danger.TButton',
                 background=[('active', '#C0392B'),
                            ('pressed', cls.COLOR_ERROR)])
        
        style.configure('TEntry',
                       fieldbackground=cls.COLOR_SURFACE,
                       foreground=cls.COLOR_TEXT,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL),
                       padding=cls.PADDING_SMALL)
        
        style.configure('TCombobox',
                       fieldbackground=cls.COLOR_SURFACE,
                       foreground=cls.COLOR_TEXT,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL),
                       padding=cls.PADDING_SMALL)
        
        style.configure('Treeview',
                       background=cls.COLOR_SURFACE,
                       foreground=cls.COLOR_TEXT,
                       fieldbackground=cls.COLOR_SURFACE,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL),
                       rowheight=28)
        style.configure('Treeview.Heading',
                       background=cls.COLOR_PRIMARY,
                       foreground=cls.COLOR_SURFACE,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL, 'bold'))
        style.map('Treeview',
                 background=[('selected', cls.COLOR_INFO_LIGHT)],
                 foreground=[('selected', cls.COLOR_TEXT)])
        
        style.configure('Horizontal.TProgressbar',
                       background=cls.COLOR_PRIMARY,
                       troughcolor=cls.COLOR_BORDER,
                       thickness=20)
        style.configure('Success.Horizontal.TProgressbar',
                       background=cls.COLOR_SUCCESS,
                       troughcolor=cls.COLOR_BORDER)
        style.configure('Warning.Horizontal.TProgressbar',
                       background=cls.COLOR_WARNING,
                       troughcolor=cls.COLOR_BORDER)
        style.configure('Error.Horizontal.TProgressbar',
                       background=cls.COLOR_ERROR,
                       troughcolor=cls.COLOR_BORDER)
        
        style.configure('TLabelframe',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_TEXT)
        style.configure('TLabelframe.Label',
                       background=cls.COLOR_BACKGROUND,
                       foreground=cls.COLOR_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_HEADING, 'bold'))
        
        style.configure('TNotebook',
                       background=cls.COLOR_BACKGROUND)
        style.configure('TNotebook.Tab',
                       background=cls.COLOR_SURFACE,
                       foreground=cls.COLOR_TEXT,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_NORMAL),
                       padding=(cls.PADDING_NORMAL, cls.PADDING_SMALL))
        style.map('TNotebook.Tab',
                 background=[('selected', cls.COLOR_PRIMARY)],
                 foreground=[('selected', cls.COLOR_SURFACE)])
        
        style.configure('TScrollbar',
                       background=cls.COLOR_BORDER,
                       troughcolor=cls.COLOR_BACKGROUND)
        
        style.configure('Card.TLabelframe',
                       background=cls.COLOR_SURFACE)
        style.configure('Card.TLabelframe.Label',
                       background=cls.COLOR_SURFACE,
                       foreground=cls.COLOR_PRIMARY,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE_HEADING, 'bold'))
    
    @classmethod
    def get_severity_color(cls, severity: str) -> str:
        severity_colors = {
            'error': cls.COLOR_ERROR,
            'warning': cls.COLOR_WARNING,
            'info': cls.COLOR_INFO
        }
        return severity_colors.get(severity.lower(), cls.COLOR_TEXT)
    
    @classmethod
    def get_severity_light_color(cls, severity: str) -> str:
        severity_colors = {
            'error': cls.COLOR_ERROR_LIGHT,
            'warning': cls.COLOR_WARNING_LIGHT,
            'info': cls.COLOR_INFO_LIGHT
        }
        return severity_colors.get(severity.lower(), cls.COLOR_BACKGROUND)
    
    @classmethod
    def get_category_display_name(cls, category: str) -> str:
        category_names = {
            'page_settings': '页面设置',
            'toc': '目录格式',
            'chapter_title': '章节标题',
            'body_text': '正文格式',
            'figure': '图片格式',
            'table': '表格格式',
            'formula': '公式格式',
            'reference': '参考文献'
        }
        return category_names.get(category, category)
    
    @classmethod
    def get_severity_display_name(cls, severity: str) -> str:
        severity_names = {
            'error': '错误',
            'warning': '警告',
            'info': '提示'
        }
        return severity_names.get(severity.lower(), severity)


class CardFrame(ttk.Frame):
    def __init__(self, parent, title: str = "", **kwargs):
        super().__init__(parent, style='Card.TFrame', **kwargs)
        
        self._create_widgets(title)
    
    def _create_widgets(self, title: str):
        if title:
            title_frame = ttk.Frame(self, style='Card.TFrame')
            title_frame.pack(fill='x', padx=AppStyles.PADDING_NORMAL, 
                           pady=(AppStyles.PADDING_NORMAL, 0))
            
            ttk.Label(title_frame, text=title, style='Heading.TLabel',
                     background=AppStyles.COLOR_SURFACE).pack(side='left')
            
            ttk.Separator(self, orient='horizontal').pack(fill='x', 
                         padx=AppStyles.PADDING_NORMAL, pady=AppStyles.PADDING_SMALL)
        
        self.content_frame = ttk.Frame(self, style='Card.TFrame')
        self.content_frame.pack(fill='both', expand=True, 
                               padx=AppStyles.PADDING_NORMAL,
                               pady=(0, AppStyles.PADDING_NORMAL))


class StatCard(ttk.Frame):
    def __init__(self, parent, title: str, value: str = "0", 
                 color: str = AppStyles.COLOR_PRIMARY, **kwargs):
        super().__init__(parent, style='Card.TFrame', **kwargs)
        
        self._create_widgets(title, value, color)
    
    def _create_widgets(self, title: str, value: str, color: str):
        self.configure(style='Card.TFrame')
        
        value_label = tk.Label(self, text=value, 
                              font=(AppStyles.FONT_FAMILY, 24, 'bold'),
                              fg=color, bg=AppStyles.COLOR_SURFACE)
        value_label.pack(pady=(AppStyles.PADDING_NORMAL, 0))
        
        title_label = tk.Label(self, text=title,
                              font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL),
                              fg=AppStyles.COLOR_TEXT_LIGHT, bg=AppStyles.COLOR_SURFACE)
        title_label.pack(pady=(0, AppStyles.PADDING_NORMAL))
    
    def update_value(self, value: str):
        for widget in self.winfo_children():
            if isinstance(widget, tk.Label):
                if widget.cget('font')[2] == 'bold':
                    widget.configure(text=value)
                    break


class SeverityIcon:
    ICONS = {
        'error': '✖',
        'warning': '⚠',
        'info': 'ℹ',
        'success': '✓'
    }
    
    @classmethod
    def get_icon(cls, severity: str) -> str:
        return cls.ICONS.get(severity.lower(), '•')


class ProgressBar(ttk.Frame):
    def __init__(self, parent, value: float = 0, max_value: float = 100,
                 color: str = AppStyles.COLOR_PRIMARY, show_text: bool = True, **kwargs):
        super().__init__(parent, style='TFrame', **kwargs)
        
        self._value = value
        self._max_value = max_value
        self._color = color
        self._show_text = show_text
        
        self._create_widgets()
    
    def _create_widgets(self):
        self._bar_frame = ttk.Frame(self, style='TFrame')
        self._bar_frame.pack(fill='x', expand=True)
        
        self._background = tk.Frame(self._bar_frame, bg=AppStyles.COLOR_BORDER, height=20)
        self._background.pack(fill='x', expand=True)
        
        self._fill = tk.Frame(self._background, bg=self._color, height=20)
        self._fill.place(x=0, y=0, relheight=1.0)
        
        if self._show_text:
            self._label = tk.Label(self._bar_frame, text='',
                                  font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL),
                                  fg=AppStyles.COLOR_TEXT, bg=AppStyles.COLOR_BACKGROUND)
            self._label.pack(side='right', padx=(AppStyles.PADDING_SMALL, 0))
        
        self._update_display()
    
    def _update_display(self):
        if self._max_value > 0:
            percentage = min(100, (self._value / self._max_value) * 100)
        else:
            percentage = 0
        
        self._fill.place(x=0, y=0, relwidth=percentage / 100, relheight=1.0)
        
        if self._show_text:
            self._label.configure(text=f'{percentage:.1f}%')
    
    def set_value(self, value: float):
        self._value = max(0, min(value, self._max_value))
        self._update_display()
    
    def set_color(self, color: str):
        self._color = color
        self._fill.configure(bg=color)


class DistributionChart(ttk.Frame):
    def __init__(self, parent, data: dict = None, **kwargs):
        super().__init__(parent, style='Card.TFrame', **kwargs)
        
        self._data = data or {}
        self._create_widgets()
    
    def _create_widgets(self):
        if not self._data:
            ttk.Label(self, text="暂无数据", style='Surface.TLabel',
                     background=AppStyles.COLOR_SURFACE).pack(pady=AppStyles.PADDING_NORMAL)
            return
        
        total = sum(self._data.values())
        if total == 0:
            total = 1
        
        chart_frame = ttk.Frame(self, style='Card.TFrame')
        chart_frame.pack(fill='x', padx=AppStyles.PADDING_NORMAL, pady=AppStyles.PADDING_NORMAL)
        
        colors = {
            'error': AppStyles.COLOR_ERROR,
            'warning': AppStyles.COLOR_WARNING,
            'info': AppStyles.COLOR_INFO,
            'success': AppStyles.COLOR_SUCCESS
        }
        
        labels = {
            'error': '错误',
            'warning': '警告',
            'info': '提示',
            'success': '通过'
        }
        
        bar_frame = tk.Frame(chart_frame, bg=AppStyles.COLOR_SURFACE, height=30)
        bar_frame.pack(fill='x', pady=(0, AppStyles.PADDING_SMALL))
        bar_frame.pack_propagate(False)
        
        for key, value in self._data.items():
            if value > 0:
                percentage = (value / total) * 100
                color = colors.get(key, AppStyles.COLOR_TEXT)
                segment = tk.Frame(bar_frame, bg=color, width=int(percentage * 3))
                segment.pack(side='left', fill='y')
        
        legend_frame = ttk.Frame(chart_frame, style='Card.TFrame')
        legend_frame.pack(fill='x')
        
        for key, value in self._data.items():
            if value > 0:
                item_frame = ttk.Frame(legend_frame, style='Card.TFrame')
                item_frame.pack(side='left', padx=(0, AppStyles.PADDING_NORMAL))
                
                color = colors.get(key, AppStyles.COLOR_TEXT)
                indicator = tk.Frame(item_frame, bg=color, width=12, height=12)
                indicator.pack(side='left', padx=(0, AppStyles.PADDING_SMALL))
                
                label = labels.get(key, key)
                ttk.Label(item_frame, text=f"{label}: {value}", style='Surface.TLabel',
                         background=AppStyles.COLOR_SURFACE).pack(side='left')


class IssueCard(ttk.Frame):
    def __init__(self, parent, issue_data: dict, **kwargs):
        super().__init__(parent, style='Card.TFrame', **kwargs)
        
        self._issue = issue_data
        self._create_widgets()
    
    def _create_widgets(self):
        severity = self._issue.get('severity', 'info')
        color = AppStyles.get_severity_color(severity)
        light_color = AppStyles.get_severity_light_color(severity)
        icon = SeverityIcon.get_icon(severity)
        severity_text = AppStyles.get_severity_display_name(severity)
        
        header_frame = tk.Frame(self, bg=light_color)
        header_frame.pack(fill='x')
        
        icon_label = tk.Label(header_frame, text=icon,
                             font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_HEADING),
                             fg=color, bg=light_color)
        icon_label.pack(side='left', padx=(AppStyles.PADDING_NORMAL, AppStyles.PADDING_SMALL))
        
        severity_label = tk.Label(header_frame, text=f"【{severity_text}】",
                                 font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_NORMAL, 'bold'),
                                 fg=color, bg=light_color)
        severity_label.pack(side='left')
        
        category = AppStyles.get_category_display_name(self._issue.get('category', ''))
        category_label = tk.Label(header_frame, text=category,
                                 font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL),
                                 fg=AppStyles.COLOR_TEXT_LIGHT, bg=light_color)
        category_label.pack(side='right', padx=AppStyles.PADDING_NORMAL)
        
        content_frame = ttk.Frame(self, style='Card.TFrame')
        content_frame.pack(fill='both', expand=True, padx=AppStyles.PADDING_NORMAL,
                          pady=AppStyles.PADDING_SMALL)
        
        location = self._issue.get('location', '-')
        if location:
            location_frame = ttk.Frame(content_frame, style='Card.TFrame')
            location_frame.pack(fill='x', pady=(0, AppStyles.PADDING_SMALL))
            
            ttk.Label(location_frame, text="📍 位置:", style='Surface.TLabel',
                     background=AppStyles.COLOR_SURFACE,
                     font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL, 'bold')).pack(side='left')
            ttk.Label(location_frame, text=location, style='Surface.TLabel',
                     background=AppStyles.COLOR_SURFACE).pack(side='left', padx=(AppStyles.PADDING_SMALL, 0))
        
        message = self._issue.get('message', '')
        if message:
            message_label = tk.Label(content_frame, text=message,
                                    font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_NORMAL),
                                    fg=AppStyles.COLOR_TEXT, bg=AppStyles.COLOR_SURFACE,
                                    wraplength=400, justify='left')
            message_label.pack(fill='x', pady=(0, AppStyles.PADDING_SMALL))
        
        suggestion = self._issue.get('suggestion', '')
        if suggestion:
            suggestion_frame = tk.Frame(content_frame, bg=AppStyles.COLOR_INFO_LIGHT)
            suggestion_frame.pack(fill='x', pady=(0, AppStyles.PADDING_SMALL))
            
            tk.Label(suggestion_frame, text="💡 修改建议:",
                    font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL, 'bold'),
                    fg=AppStyles.COLOR_INFO, bg=AppStyles.COLOR_INFO_LIGHT).pack(anchor='w',
                    padx=AppStyles.PADDING_SMALL, pady=(AppStyles.PADDING_SMALL, 0))
            
            tk.Label(suggestion_frame, text=suggestion,
                    font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL),
                    fg=AppStyles.COLOR_TEXT, bg=AppStyles.COLOR_INFO_LIGHT,
                    wraplength=380, justify='left').pack(anchor='w',
                    padx=AppStyles.PADDING_SMALL, pady=(0, AppStyles.PADDING_SMALL))
        
        details_frame = ttk.Frame(content_frame, style='Card.TFrame')
        details_frame.pack(fill='x')
        
        actual = self._issue.get('actual_value')
        expected = self._issue.get('expected_value')
        
        if actual or expected:
            if actual:
                actual_frame = ttk.Frame(details_frame, style='Card.TFrame')
                actual_frame.pack(side='left', padx=(0, AppStyles.PADDING_NORMAL))
                ttk.Label(actual_frame, text="当前值:", style='Surface.TLabel',
                         background=AppStyles.COLOR_SURFACE,
                         font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL)).pack(side='left')
                ttk.Label(actual_frame, text=str(actual), style='Error.TLabel',
                         font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL)).pack(side='left')
            
            if expected:
                expected_frame = ttk.Frame(details_frame, style='Card.TFrame')
                expected_frame.pack(side='left')
                ttk.Label(expected_frame, text="期望值:", style='Surface.TLabel',
                         background=AppStyles.COLOR_SURFACE,
                         font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL)).pack(side='left')
                ttk.Label(expected_frame, text=str(expected), style='Success.TLabel',
                         font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL)).pack(side='left')


class EnhancedStatCard(ttk.Frame):
    def __init__(self, parent, title: str, value: str = "0",
                 icon: str = "", color: str = AppStyles.COLOR_PRIMARY,
                 subtitle: str = "", **kwargs):
        super().__init__(parent, style='Card.TFrame', **kwargs)
        
        self._create_widgets(title, value, icon, color, subtitle)
    
    def _create_widgets(self, title: str, value: str, icon: str, color: str, subtitle: str):
        self.configure(style='Card.TFrame')
        
        header_frame = ttk.Frame(self, style='Card.TFrame')
        header_frame.pack(fill='x', padx=AppStyles.PADDING_NORMAL, pady=(AppStyles.PADDING_NORMAL, 0))
        
        if icon:
            icon_label = tk.Label(header_frame, text=icon,
                                 font=(AppStyles.FONT_FAMILY, 20),
                                 fg=color, bg=AppStyles.COLOR_SURFACE)
            icon_label.pack(side='left')
        
        title_label = tk.Label(header_frame, text=title,
                              font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL),
                              fg=AppStyles.COLOR_TEXT_LIGHT, bg=AppStyles.COLOR_SURFACE)
        title_label.pack(side='left', padx=(AppStyles.PADDING_SMALL, 0))
        
        value_label = tk.Label(self, text=value,
                              font=(AppStyles.FONT_FAMILY, 28, 'bold'),
                              fg=color, bg=AppStyles.COLOR_SURFACE)
        value_label.pack(pady=(AppStyles.PADDING_SMALL, 0))
        
        if subtitle:
            subtitle_label = tk.Label(self, text=subtitle,
                                     font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL),
                                     fg=AppStyles.COLOR_TEXT_LIGHT, bg=AppStyles.COLOR_SURFACE)
            subtitle_label.pack(pady=(0, AppStyles.PADDING_NORMAL))
        else:
            value_label.pack(pady=(0, AppStyles.PADDING_NORMAL))
    
    def update_value(self, value: str, subtitle: str = ""):
        for widget in self.winfo_children():
            if isinstance(widget, tk.Label):
                font = widget.cget('font')
                if isinstance(font, tuple) and len(font) >= 3 and font[2] == 'bold':
                    widget.configure(text=value)
                elif subtitle and widget.cget('fg') == AppStyles.COLOR_TEXT_LIGHT:
                    widget.configure(text=subtitle)
