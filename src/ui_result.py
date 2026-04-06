#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
from datetime import datetime
from typing import Optional, List, Dict, Any

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .ui_styles import (
    AppStyles, StatCard, SeverityIcon, ProgressBar, 
    DistributionChart, IssueCard, EnhancedStatCard
)
from .format_checker import CheckReport, CheckResult, Issue, Severity, CheckCategory
from .models import DocumentInfo
from .report_template import HTMLReportTemplate


class ResultWindow:
    def __init__(self, root: tk.Toplevel, report: CheckReport, document: Optional[DocumentInfo] = None):
        self.root = root
        self.report = report
        self.document = document
        
        self._current_filter = "all"
        self._issues_data: List[Dict] = []
        
        self._create_widgets()
        self._load_data()
    
    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, style='TFrame')
        main_frame.pack(fill='both', expand=True, padx=AppStyles.PADDING_LARGE,
                       pady=AppStyles.PADDING_LARGE)
        
        self._create_header(main_frame)
        self._create_stats_section(main_frame)
        self._create_chart_section(main_frame)
        self._create_filter_section(main_frame)
        self._create_issues_section(main_frame)
        self._create_export_section(main_frame)
    
    def _create_header(self, parent):
        header_frame = ttk.Frame(parent, style='TFrame')
        header_frame.pack(fill='x', pady=(0, AppStyles.PADDING_NORMAL))
        
        title_frame = ttk.Frame(header_frame, style='TFrame')
        title_frame.pack(side='left')
        
        ttk.Label(title_frame, text="📄 检测结果报告", style='Title.TLabel').pack(side='left')
        
        pass_rate = self.report.statistics.pass_rate
        if pass_rate >= 80:
            status_text = "✓ 格式良好"
            status_color = AppStyles.COLOR_SUCCESS
        elif pass_rate >= 60:
            status_text = "⚠ 需要改进"
            status_color = AppStyles.COLOR_WARNING
        else:
            status_text = "✗ 问题较多"
            status_color = AppStyles.COLOR_ERROR
        
        status_label = tk.Label(title_frame, text=status_text,
                               font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_HEADING, 'bold'),
                               fg=status_color, bg=AppStyles.COLOR_BACKGROUND)
        status_label.pack(side='left', padx=(AppStyles.PADDING_NORMAL, 0))
        
        info_text = f"📁 {self.report.document_name}  |  🕐 {self.report.check_time[:19]}"
        ttk.Label(header_frame, text=info_text, style='Small.TLabel').pack(side='right')
    
    def _create_stats_section(self, parent):
        stats_frame = ttk.Frame(parent, style='TFrame')
        stats_frame.pack(fill='x', pady=(0, AppStyles.PADDING_NORMAL))
        
        total_card = EnhancedStatCard(
            stats_frame, "总问题数", str(self.report.statistics.total_issues),
            "📊", AppStyles.COLOR_PRIMARY
        )
        total_card.pack(side='left', padx=(0, AppStyles.PADDING_SMALL), ipadx=15)
        
        error_card = EnhancedStatCard(
            stats_frame, "错误", str(self.report.statistics.error_count),
            "✖", AppStyles.COLOR_ERROR
        )
        error_card.pack(side='left', padx=(0, AppStyles.PADDING_SMALL), ipadx=15)
        
        warning_card = EnhancedStatCard(
            stats_frame, "警告", str(self.report.statistics.warning_count),
            "⚠", AppStyles.COLOR_WARNING
        )
        warning_card.pack(side='left', padx=(0, AppStyles.PADDING_SMALL), ipadx=15)
        
        info_card = EnhancedStatCard(
            stats_frame, "提示", str(self.report.statistics.info_count),
            "ℹ", AppStyles.COLOR_INFO
        )
        info_card.pack(side='left', padx=(0, AppStyles.PADDING_SMALL), ipadx=15)
        
        pass_rate = f"{self.report.statistics.pass_rate:.1f}%"
        rate_card = EnhancedStatCard(
            stats_frame, "通过率", pass_rate,
            "✓", AppStyles.COLOR_SUCCESS,
            f"{self.report.statistics.passed_checks}/{self.report.statistics.total_checks}项通过"
        )
        rate_card.pack(side='left', padx=(0, AppStyles.PADDING_SMALL), ipadx=15)
    
    def _create_chart_section(self, parent):
        chart_frame = ttk.Frame(parent, style='TFrame')
        chart_frame.pack(fill='x', pady=(0, AppStyles.PADDING_NORMAL))
        
        chart_data = {
            'error': self.report.statistics.error_count,
            'warning': self.report.statistics.warning_count,
            'info': self.report.statistics.info_count,
            'success': self.report.statistics.passed_checks
        }
        
        chart = DistributionChart(chart_frame, data=chart_data)
        chart.pack(fill='x', ipady=5)
        
        progress_frame = ttk.Frame(parent, style='TFrame')
        progress_frame.pack(fill='x', pady=(0, AppStyles.PADDING_NORMAL))
        
        ttk.Label(progress_frame, text="通过率进度:", style='TLabel').pack(side='left', padx=(0, AppStyles.PADDING_SMALL))
        
        pass_rate = self.report.statistics.pass_rate
        if pass_rate >= 80:
            bar_color = AppStyles.COLOR_SUCCESS
        elif pass_rate >= 60:
            bar_color = AppStyles.COLOR_WARNING
        else:
            bar_color = AppStyles.COLOR_ERROR
        
        progress_bar = ProgressBar(progress_frame, value=pass_rate, max_value=100, color=bar_color)
        progress_bar.pack(side='left', fill='x', expand=True, padx=(AppStyles.PADDING_SMALL, 0))
    
    def _create_filter_section(self, parent):
        filter_frame = ttk.Frame(parent, style='TFrame')
        filter_frame.pack(fill='x', pady=(0, AppStyles.PADDING_NORMAL))
        
        ttk.Label(filter_frame, text="🔍 筛选:", style='TLabel').pack(side='left', padx=(0, AppStyles.PADDING_SMALL))
        
        self.filter_var = tk.StringVar(value="all")
        
        filters = [
            ("全部", "all"),
            ("✖ 错误", "error"),
            ("⚠ 警告", "warning"),
            ("ℹ 提示", "info"),
        ]
        
        for text, value in filters:
            rb = ttk.Radiobutton(filter_frame, text=text, value=value,
                                variable=self.filter_var, command=self._apply_filter)
            rb.pack(side='left', padx=(0, AppStyles.PADDING_NORMAL))
        
        ttk.Separator(filter_frame, orient='vertical').pack(side='left', fill='y', padx=AppStyles.PADDING_NORMAL)
        
        ttk.Label(filter_frame, text="📂 分类:", style='TLabel').pack(side='left', padx=(0, AppStyles.PADDING_SMALL))
        
        self.category_var = tk.StringVar(value="全部")
        
        categories = [
            "全部",
            "页面设置",
            "目录格式",
            "章节标题",
            "正文格式",
            "图片格式",
            "表格格式",
            "公式格式",
            "参考文献",
        ]
        
        self.category_combo = ttk.Combobox(filter_frame, textvariable=self.category_var,
                                           state='readonly', width=12)
        self.category_combo['values'] = categories
        self.category_combo.current(0)
        self.category_combo.pack(side='left')
        self.category_combo.bind('<<ComboboxSelected>>', self._apply_filter)
        
        self._category_map = {text: value for text, value in [
            ("全部", "all"),
            ("页面设置", "page_settings"),
            ("目录格式", "toc"),
            ("章节标题", "chapter_title"),
            ("正文格式", "body_text"),
            ("图片格式", "figure"),
            ("表格格式", "table"),
            ("公式格式", "formula"),
            ("参考文献", "reference"),
        ]}
    
    def _create_issues_section(self, parent):
        issues_container = ttk.Frame(parent, style='Surface.TFrame')
        issues_container.pack(fill='both', expand=True, pady=(0, AppStyles.PADDING_NORMAL))
        
        list_frame = ttk.Frame(issues_container, style='Surface.TFrame')
        list_frame.pack(side='left', fill='both', expand=True)
        
        columns = ('severity', 'category', 'location', 'message')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=12)
        
        self.tree.heading('severity', text='严重程度')
        self.tree.heading('category', text='类别')
        self.tree.heading('location', text='位置')
        self.tree.heading('message', text='问题描述')
        
        self.tree.column('severity', width=90, anchor='center')
        self.tree.column('category', width=100, anchor='center')
        self.tree.column('location', width=200, anchor='w')
        self.tree.column('message', width=400, anchor='w')
        
        self.tree.tag_configure('error', foreground=AppStyles.COLOR_ERROR)
        self.tree.tag_configure('warning', foreground=AppStyles.COLOR_WARNING)
        self.tree.tag_configure('info', foreground=AppStyles.COLOR_INFO)
        
        scrollbar_y = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar_y.pack(side='right', fill='y')
        
        self.tree.bind('<<TreeviewSelect>>', self._on_issue_selected)
        self.tree.bind('<Double-1>', self._show_issue_detail)
        
        detail_frame = ttk.LabelFrame(issues_container, text="问题详情预览", style='Card.TLabelframe')
        detail_frame.pack(side='right', fill='both', padx=(AppStyles.PADDING_NORMAL, 0), ipadx=10, ipady=5)
        
        self.detail_widget = IssueDetailWidget(detail_frame)
        self.detail_widget.pack(fill='both', expand=True, padx=AppStyles.PADDING_SMALL, pady=AppStyles.PADDING_SMALL)
    
    def _create_export_section(self, parent):
        export_frame = ttk.Frame(parent, style='TFrame')
        export_frame.pack(fill='x')
        
        count_text = f"共 {self.report.statistics.total_issues} 个问题"
        if self.report.statistics.error_count > 0:
            count_text += f"（{self.report.statistics.error_count} 个错误需要修复）"
        ttk.Label(export_frame, text=count_text, style='TLabel').pack(side='left')
        
        btn_frame = ttk.Frame(export_frame, style='TFrame')
        btn_frame.pack(side='right')
        
        ttk.Button(btn_frame, text="📄 导出HTML报告",
                  command=self._export_html).pack(side='right', padx=(AppStyles.PADDING_SMALL, 0))
        
        ttk.Button(btn_frame, text="📋 导出JSON报告",
                  command=self._export_json).pack(side='right', padx=(AppStyles.PADDING_SMALL, 0))
        
        ttk.Button(btn_frame, text="🔍 查看详情",
                  command=self._show_selected_detail).pack(side='right', padx=(AppStyles.PADDING_SMALL, 0))
    
    def _load_data(self):
        self._issues_data = []
        
        for result in self.report.results:
            for issue in result.issues:
                self._issues_data.append({
                    'severity': issue.severity.value,
                    'category': result.category.value,
                    'location': issue.location or '-',
                    'message': issue.message,
                    'actual_value': issue.actual_value,
                    'expected_value': issue.expected_value,
                    'suggestion': issue.suggestion,
                    'paragraph_index': issue.paragraph_index,
                    'check_name': result.check_name
                })
        
        self._refresh_tree()
    
    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        severity_filter = self.filter_var.get()
        category_text = self.category_var.get()
        category_filter = self._category_map.get(category_text, 'all')
        
        for issue in self._issues_data:
            if severity_filter != 'all' and issue['severity'] != severity_filter:
                continue
            if category_filter != 'all' and issue['category'] != category_filter:
                continue
            
            icon = SeverityIcon.get_icon(issue['severity'])
            severity_display = f"{icon} {AppStyles.get_severity_display_name(issue['severity'])}"
            category_display = AppStyles.get_category_display_name(issue['category'])
            
            location_display = issue['location'][:35] if issue['location'] and len(issue['location']) > 35 else (issue['location'] or '-')
            message_display = issue['message'][:55] + "..." if len(issue['message']) > 55 else issue['message']
            
            self.tree.insert('', 'end', values=(
                severity_display,
                category_display,
                location_display,
                message_display
            ), tags=(issue['severity'],))
    
    def _apply_filter(self):
        self._refresh_tree()
    
    def _on_issue_selected(self, event):
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        
        for issue in self._issues_data:
            severity_display = AppStyles.get_severity_display_name(issue['severity'])
            category_display = AppStyles.get_category_display_name(issue['category'])
            
            tree_severity = values[0].replace(SeverityIcon.get_icon(issue['severity']), '').strip()
            if tree_severity == severity_display and values[1] == category_display:
                self.detail_widget.update_issue(issue)
                break
    
    def _show_selected_detail(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个问题")
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        
        for issue in self._issues_data:
            severity_display = AppStyles.get_severity_display_name(issue['severity'])
            category_display = AppStyles.get_category_display_name(issue['category'])
            
            tree_severity = values[0].replace(SeverityIcon.get_icon(issue['severity']), '').strip()
            if tree_severity == severity_display and values[1] == category_display:
                self._show_issue_detail_dialog(issue)
                break
    
    def _show_issue_detail(self, event):
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        
        for issue in self._issues_data:
            severity_display = AppStyles.get_severity_display_name(issue['severity'])
            category_display = AppStyles.get_category_display_name(issue['category'])
            
            tree_severity = values[0].replace(SeverityIcon.get_icon(issue['severity']), '').strip()
            if tree_severity == severity_display and values[1] == category_display:
                self._show_issue_detail_dialog(issue)
                break
    
    def _show_issue_detail_dialog(self, issue: Dict):
        dialog = tk.Toplevel(self.root)
        dialog.title("问题详情")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, style='TFrame')
        main_frame.pack(fill='both', expand=True, padx=AppStyles.PADDING_NORMAL,
                       pady=AppStyles.PADDING_NORMAL)
        
        card = IssueCard(main_frame, issue)
        card.pack(fill='both', expand=True)
        
        btn_frame = ttk.Frame(main_frame, style='TFrame')
        btn_frame.pack(fill='x', pady=(AppStyles.PADDING_NORMAL, 0))
        
        ttk.Button(btn_frame, text="关闭", command=dialog.destroy).pack(side='right')
    
    def _export_html(self):
        file_path = filedialog.asksaveasfilename(
            title="保存HTML报告",
            defaultextension=".html",
            filetypes=[("HTML文件", "*.html")],
            initialfile=f"检测报告_{self.report.document_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        )
        
        if file_path:
            try:
                report_data = self.report.to_dict()
                html_content = HTMLReportTemplate.generate_report(report_data)
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                messagebox.showinfo("成功", f"报告已导出至:\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败:\n{str(e)}")
    
    def _export_json(self):
        file_path = filedialog.asksaveasfilename(
            title="保存JSON报告",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json")],
            initialfile=f"检测报告_{self.report.document_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        
        if file_path:
            try:
                import json
                report_data = self.report.to_dict()
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(report_data, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", f"报告已导出至:\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败:\n{str(e)}")


class IssueDetailWidget(ttk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, style='Card.TFrame', **kwargs)
        
        self._current_issue = None
        self._create_widgets()
    
    def _create_widgets(self):
        self._placeholder = ttk.Label(self, text="请选择一个问题查看详情",
                                      style='Surface.TLabel',
                                      background=AppStyles.COLOR_SURFACE)
        self._placeholder.pack(expand=True)
    
    def update_issue(self, issue: Dict):
        self._current_issue = issue
        
        for widget in self.winfo_children():
            widget.destroy()
        
        severity = issue.get('severity', 'info')
        color = AppStyles.get_severity_color(severity)
        light_color = AppStyles.get_severity_light_color(severity)
        icon = SeverityIcon.get_icon(severity)
        severity_text = AppStyles.get_severity_display_name(severity)
        
        header_frame = tk.Frame(self, bg=light_color)
        header_frame.pack(fill='x')
        
        tk.Label(header_frame, text=f"{icon} {severity_text}",
                font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_HEADING, 'bold'),
                fg=color, bg=light_color).pack(side='left', padx=AppStyles.PADDING_SMALL, pady=AppStyles.PADDING_SMALL)
        
        category = AppStyles.get_category_display_name(issue.get('category', ''))
        tk.Label(header_frame, text=category,
                font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL),
                fg=AppStyles.COLOR_TEXT_LIGHT, bg=light_color).pack(side='right', padx=AppStyles.PADDING_SMALL, pady=AppStyles.PADDING_SMALL)
        
        content_frame = ttk.Frame(self, style='Card.TFrame')
        content_frame.pack(fill='both', expand=True, padx=AppStyles.PADDING_SMALL, pady=AppStyles.PADDING_SMALL)
        
        location = issue.get('location', '')
        if location and location != '-':
            loc_frame = ttk.Frame(content_frame, style='Card.TFrame')
            loc_frame.pack(fill='x', pady=(0, AppStyles.PADDING_SMALL))
            
            ttk.Label(loc_frame, text="📍 位置:", style='Surface.TLabel',
                     background=AppStyles.COLOR_SURFACE,
                     font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL, 'bold')).pack(side='left')
            ttk.Label(loc_frame, text=location[:40], style='Surface.TLabel',
                     background=AppStyles.COLOR_SURFACE).pack(side='left', padx=(AppStyles.PADDING_SMALL, 0))
        
        message = issue.get('message', '')
        if message:
            msg_label = tk.Label(content_frame, text=message,
                                font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL),
                                fg=AppStyles.COLOR_TEXT, bg=AppStyles.COLOR_SURFACE,
                                wraplength=280, justify='left')
            msg_label.pack(fill='x', pady=(0, AppStyles.PADDING_SMALL))
        
        suggestion = issue.get('suggestion', '')
        if suggestion:
            sug_frame = tk.Frame(content_frame, bg=AppStyles.COLOR_INFO_LIGHT)
            sug_frame.pack(fill='x', pady=(0, AppStyles.PADDING_SMALL))
            
            tk.Label(sug_frame, text="💡 修改建议:",
                    font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL, 'bold'),
                    fg=AppStyles.COLOR_INFO, bg=AppStyles.COLOR_INFO_LIGHT).pack(anchor='w',
                    padx=AppStyles.PADDING_SMALL, pady=(AppStyles.PADDING_SMALL, 0))
            
            tk.Label(sug_frame, text=suggestion[:100] + "..." if len(suggestion) > 100 else suggestion,
                    font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL),
                    fg=AppStyles.COLOR_TEXT, bg=AppStyles.COLOR_INFO_LIGHT,
                    wraplength=260, justify='left').pack(anchor='w',
                    padx=AppStyles.PADDING_SMALL, pady=(0, AppStyles.PADDING_SMALL))
        
        actual = issue.get('actual_value')
        expected = issue.get('expected_value')
        
        if actual or expected:
            details_frame = ttk.Frame(content_frame, style='Card.TFrame')
            details_frame.pack(fill='x')
            
            if actual:
                actual_frame = ttk.Frame(details_frame, style='Card.TFrame')
                actual_frame.pack(fill='x', pady=(0, AppStyles.PADDING_SMALL))
                ttk.Label(actual_frame, text="当前值:", style='Surface.TLabel',
                         background=AppStyles.COLOR_SURFACE,
                         font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL)).pack(side='left')
                ttk.Label(actual_frame, text=str(actual)[:30], style='Error.TLabel',
                         font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL)).pack(side='left', padx=(AppStyles.PADDING_SMALL, 0))
            
            if expected:
                expected_frame = ttk.Frame(details_frame, style='Card.TFrame')
                expected_frame.pack(fill='x')
                ttk.Label(expected_frame, text="期望值:", style='Surface.TLabel',
                         background=AppStyles.COLOR_SURFACE,
                         font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL)).pack(side='left')
                ttk.Label(expected_frame, text=str(expected)[:30], style='Success.TLabel',
                         font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE_SMALL)).pack(side='left', padx=(AppStyles.PADDING_SMALL, 0))
