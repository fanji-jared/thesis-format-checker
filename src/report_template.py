#!/usr/bin/env python
# -*- coding: utf-8 -*-

from datetime import datetime
from typing import List, Dict, Any, Optional


class HTMLReportTemplate:
    CSS_STYLES = """
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: "Microsoft YaHei", "Segoe UI", Tahoma, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
            line-height: 1.6;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2C3E50 0%, #3498DB 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 32px;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.2);
        }
        
        .header .subtitle {
            font-size: 16px;
            opacity: 0.9;
        }
        
        .meta-info {
            background: #f8f9fa;
            padding: 20px 40px;
            border-bottom: 1px solid #e9ecef;
            display: flex;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 15px;
        }
        
        .meta-item {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .meta-item .icon {
            font-size: 18px;
        }
        
        .meta-item .label {
            color: #7F8C8D;
            font-size: 14px;
        }
        
        .meta-item .value {
            color: #2C3E50;
            font-weight: 600;
        }
        
        .content {
            padding: 40px;
        }
        
        .stats-section {
            margin-bottom: 40px;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .stat-card {
            background: white;
            border-radius: 12px;
            padding: 25px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
            border: 1px solid #e9ecef;
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }
        
        .stat-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0, 0, 0, 0.12);
        }
        
        .stat-card .icon {
            font-size: 28px;
            margin-bottom: 10px;
        }
        
        .stat-card .value {
            font-size: 36px;
            font-weight: bold;
            margin-bottom: 5px;
        }
        
        .stat-card .label {
            color: #7F8C8D;
            font-size: 14px;
        }
        
        .stat-card.total {
            border-left: 4px solid #2C3E50;
        }
        
        .stat-card.total .value {
            color: #2C3E50;
        }
        
        .stat-card.error {
            border-left: 4px solid #E74C3C;
        }
        
        .stat-card.error .value {
            color: #E74C3C;
        }
        
        .stat-card.warning {
            border-left: 4px solid #F39C12;
        }
        
        .stat-card.warning .value {
            color: #F39C12;
        }
        
        .stat-card.info {
            border-left: 4px solid #3498DB;
        }
        
        .stat-card.info .value {
            color: #3498DB;
        }
        
        .stat-card.success {
            border-left: 4px solid #27AE60;
        }
        
        .stat-card.success .value {
            color: #27AE60;
        }
        
        .chart-section {
            background: #f8f9fa;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 40px;
        }
        
        .chart-title {
            font-size: 16px;
            font-weight: 600;
            color: #2C3E50;
            margin-bottom: 20px;
        }
        
        .distribution-bar {
            height: 30px;
            border-radius: 15px;
            overflow: hidden;
            display: flex;
            background: #e9ecef;
            margin-bottom: 15px;
        }
        
        .distribution-segment {
            height: 100%;
            transition: width 0.5s ease;
        }
        
        .distribution-segment.error {
            background: linear-gradient(90deg, #E74C3C, #C0392B);
        }
        
        .distribution-segment.warning {
            background: linear-gradient(90deg, #F39C12, #D68910);
        }
        
        .distribution-segment.info {
            background: linear-gradient(90deg, #3498DB, #2980B9);
        }
        
        .distribution-segment.success {
            background: linear-gradient(90deg, #27AE60, #1E8449);
        }
        
        .legend {
            display: flex;
            justify-content: center;
            flex-wrap: wrap;
            gap: 20px;
        }
        
        .legend-item {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .legend-color {
            width: 16px;
            height: 16px;
            border-radius: 4px;
        }
        
        .legend-text {
            font-size: 14px;
            color: #7F8C8D;
        }
        
        .section-title {
            font-size: 20px;
            font-weight: 600;
            color: #2C3E50;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #3498DB;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .section-title .icon {
            font-size: 24px;
        }
        
        .overview-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 40px;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
        }
        
        .overview-table th {
            background: linear-gradient(135deg, #2C3E50 0%, #34495E 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }
        
        .overview-table td {
            padding: 15px;
            border-bottom: 1px solid #e9ecef;
        }
        
        .overview-table tr:last-child td {
            border-bottom: none;
        }
        
        .overview-table tr:hover {
            background: #f8f9fa;
        }
        
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 13px;
            font-weight: 600;
        }
        
        .status-badge.passed {
            background: #D5F5E3;
            color: #27AE60;
        }
        
        .status-badge.failed {
            background: #FADBD8;
            color: #E74C3C;
        }
        
        .issues-section {
            margin-bottom: 40px;
        }
        
        .issue-card {
            background: white;
            border-radius: 12px;
            margin-bottom: 15px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
            overflow: hidden;
            border: 1px solid #e9ecef;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        
        .issue-card:hover {
            transform: translateX(5px);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.12);
        }
        
        .issue-header {
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .issue-header.error {
            background: linear-gradient(90deg, #FADBD8, #F5B7B1);
        }
        
        .issue-header.warning {
            background: linear-gradient(90deg, #FCF3CF, #F9E79F);
        }
        
        .issue-header.info {
            background: linear-gradient(90deg, #D4E6F1, #AED6F1);
        }
        
        .issue-severity {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .issue-severity .icon {
            font-size: 20px;
        }
        
        .issue-severity .text {
            font-weight: 600;
        }
        
        .issue-severity.error .icon,
        .issue-severity.error .text {
            color: #E74C3C;
        }
        
        .issue-severity.warning .icon,
        .issue-severity.warning .text {
            color: #F39C12;
        }
        
        .issue-severity.info .icon,
        .issue-severity.info .text {
            color: #3498DB;
        }
        
        .issue-category {
            background: rgba(255, 255, 255, 0.7);
            padding: 4px 12px;
            border-radius: 15px;
            font-size: 13px;
            color: #7F8C8D;
        }
        
        .issue-body {
            padding: 20px;
        }
        
        .issue-location {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 12px;
            padding: 10px 15px;
            background: #f8f9fa;
            border-radius: 8px;
        }
        
        .issue-location .icon {
            font-size: 16px;
        }
        
        .issue-location .label {
            font-weight: 600;
            color: #7F8C8D;
        }
        
        .issue-location .value {
            color: #2C3E50;
        }
        
        .issue-message {
            font-size: 15px;
            color: #2C3E50;
            margin-bottom: 15px;
            line-height: 1.6;
        }
        
        .issue-suggestion {
            background: linear-gradient(90deg, #D4E6F1, #EBF5FB);
            border-left: 4px solid #3498DB;
            padding: 15px;
            border-radius: 0 8px 8px 0;
            margin-bottom: 15px;
        }
        
        .issue-suggestion .suggestion-title {
            display: flex;
            align-items: center;
            gap: 8px;
            font-weight: 600;
            color: #3498DB;
            margin-bottom: 8px;
        }
        
        .issue-suggestion .suggestion-text {
            color: #2C3E50;
            font-size: 14px;
        }
        
        .issue-details {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }
        
        .detail-item {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .detail-item .label {
            color: #7F8C8D;
            font-size: 14px;
        }
        
        .detail-item .value {
            font-weight: 600;
            padding: 3px 10px;
            border-radius: 4px;
        }
        
        .detail-item .value.actual {
            background: #FADBD8;
            color: #E74C3C;
        }
        
        .detail-item .value.expected {
            background: #D5F5E3;
            color: #27AE60;
        }
        
        .footer {
            background: #2C3E50;
            color: white;
            padding: 30px 40px;
            text-align: center;
        }
        
        .footer .generator {
            font-size: 14px;
            opacity: 0.8;
            margin-bottom: 10px;
        }
        
        .footer .timestamp {
            font-size: 12px;
            opacity: 0.6;
        }
        
        .no-issues {
            text-align: center;
            padding: 60px 20px;
            background: #D5F5E3;
            border-radius: 12px;
            margin-bottom: 40px;
        }
        
        .no-issues .icon {
            font-size: 60px;
            margin-bottom: 20px;
        }
        
        .no-issues .title {
            font-size: 24px;
            font-weight: 600;
            color: #27AE60;
            margin-bottom: 10px;
        }
        
        .no-issues .text {
            color: #7F8C8D;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .container {
                box-shadow: none;
                border-radius: 0;
            }
            
            .stat-card:hover,
            .issue-card:hover {
                transform: none;
                box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
            }
            
            .header {
                background: #2C3E50 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            
            .overview-table th {
                background: #2C3E50 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            
            .issue-header.error {
                background: #FADBD8 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            
            .issue-header.warning {
                background: #FCF3CF !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            
            .issue-header.info {
                background: #D4E6F1 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            
            .page-break {
                page-break-before: always;
            }
        }
        
        @media screen and (max-width: 768px) {
            .header {
                padding: 25px;
            }
            
            .header h1 {
                font-size: 24px;
            }
            
            .content {
                padding: 20px;
            }
            
            .stats-grid {
                grid-template-columns: repeat(2, 1fr);
            }
            
            .meta-info {
                flex-direction: column;
                padding: 15px 20px;
            }
        }
    """
    
    SEVERITY_ICONS = {
        'error': '✖',
        'warning': '⚠',
        'info': 'ℹ',
        'success': '✓'
    }
    
    CATEGORY_NAMES = {
        'page_settings': '页面设置',
        'toc': '目录格式',
        'chapter_title': '章节标题',
        'body_text': '正文格式',
        'figure': '图片格式',
        'table': '表格格式',
        'formula': '公式格式',
        'reference': '参考文献'
    }
    
    SEVERITY_NAMES = {
        'error': '错误',
        'warning': '警告',
        'info': '提示'
    }
    
    @classmethod
    def generate_report(cls, report_data: Dict[str, Any]) -> str:
        html = cls._get_html_start(report_data)
        html += cls._get_header_section(report_data)
        html += cls._get_meta_section(report_data)
        html += cls._get_stats_section(report_data)
        html += cls._get_chart_section(report_data)
        html += cls._get_overview_section(report_data)
        html += cls._get_issues_section(report_data)
        html += cls._get_footer_section(report_data)
        html += cls._get_html_end()
        return html
    
    @classmethod
    def _get_html_start(cls, report_data: Dict[str, Any]) -> str:
        return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>论文格式检测报告 - {report_data.get('document_name', '未知文档')}</title>
    <style>
        {cls.CSS_STYLES}
    </style>
</head>
<body>
    <div class="container">
"""
    
    @classmethod
    def _get_header_section(cls, report_data: Dict[str, Any]) -> str:
        return f"""
        <div class="header">
            <h1>📄 论文格式检测报告</h1>
            <div class="subtitle">专业格式检测 · 精准问题定位 · 智能修改建议</div>
        </div>
"""
    
    @classmethod
    def _get_meta_section(cls, report_data: Dict[str, Any]) -> str:
        doc_name = report_data.get('document_name', '未知文档')
        check_time = report_data.get('check_time', '')[:19]
        config_name = report_data.get('config_name', '默认配置')
        config_version = report_data.get('config_version', '-')
        
        return f"""
        <div class="meta-info">
            <div class="meta-item">
                <span class="icon">📁</span>
                <span class="label">文档名称:</span>
                <span class="value">{doc_name}</span>
            </div>
            <div class="meta-item">
                <span class="icon">🕐</span>
                <span class="label">检测时间:</span>
                <span class="value">{check_time}</span>
            </div>
            <div class="meta-item">
                <span class="icon">⚙️</span>
                <span class="label">配置方案:</span>
                <span class="value">{config_name} v{config_version}</span>
            </div>
        </div>
"""
    
    @classmethod
    def _get_stats_section(cls, report_data: Dict[str, Any]) -> str:
        stats = report_data.get('statistics', {})
        total = stats.get('total_issues', 0)
        errors = stats.get('error_count', 0)
        warnings = stats.get('warning_count', 0)
        infos = stats.get('info_count', 0)
        pass_rate = stats.get('pass_rate', 100)
        
        return f"""
        <div class="content">
            <div class="stats-section">
                <div class="stats-grid">
                    <div class="stat-card total">
                        <div class="icon">📊</div>
                        <div class="value">{total}</div>
                        <div class="label">总问题数</div>
                    </div>
                    <div class="stat-card error">
                        <div class="icon">✖</div>
                        <div class="value">{errors}</div>
                        <div class="label">错误</div>
                    </div>
                    <div class="stat-card warning">
                        <div class="icon">⚠</div>
                        <div class="value">{warnings}</div>
                        <div class="label">警告</div>
                    </div>
                    <div class="stat-card info">
                        <div class="icon">ℹ</div>
                        <div class="value">{infos}</div>
                        <div class="label">提示</div>
                    </div>
                    <div class="stat-card success">
                        <div class="icon">✓</div>
                        <div class="value">{pass_rate:.1f}%</div>
                        <div class="label">通过率</div>
                    </div>
                </div>
            </div>
"""
    
    @classmethod
    def _get_chart_section(cls, report_data: Dict[str, Any]) -> str:
        stats = report_data.get('statistics', {})
        errors = stats.get('error_count', 0)
        warnings = stats.get('warning_count', 0)
        infos = stats.get('info_count', 0)
        passed = stats.get('passed_checks', 0)
        
        total = errors + warnings + infos + passed
        if total == 0:
            total = 1
        
        error_width = (errors / total) * 100 if errors > 0 else 0
        warning_width = (warnings / total) * 100 if warnings > 0 else 0
        info_width = (infos / total) * 100 if infos > 0 else 0
        success_width = (passed / total) * 100 if passed > 0 else 0
        
        return f"""
            <div class="chart-section">
                <div class="chart-title">📈 问题分布统计</div>
                <div class="distribution-bar">
                    <div class="distribution-segment error" style="width: {error_width}%"></div>
                    <div class="distribution-segment warning" style="width: {warning_width}%"></div>
                    <div class="distribution-segment info" style="width: {info_width}%"></div>
                    <div class="distribution-segment success" style="width: {success_width}%"></div>
                </div>
                <div class="legend">
                    <div class="legend-item">
                        <div class="legend-color" style="background: #E74C3C"></div>
                        <span class="legend-text">错误: {errors}</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background: #F39C12"></div>
                        <span class="legend-text">警告: {warnings}</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background: #3498DB"></div>
                        <span class="legend-text">提示: {infos}</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background: #27AE60"></div>
                        <span class="legend-text">通过: {passed}</span>
                    </div>
                </div>
            </div>
"""
    
    @classmethod
    def _get_overview_section(cls, report_data: Dict[str, Any]) -> str:
        results = report_data.get('results', [])
        
        html = """
            <h2 class="section-title"><span class="icon">📋</span> 检测概览</h2>
            <table class="overview-table">
                <thead>
                    <tr>
                        <th>检测项</th>
                        <th>状态</th>
                        <th>检测数</th>
                        <th>通过数</th>
                        <th>问题数</th>
                    </tr>
                </thead>
                <tbody>
"""
        
        for result in results:
            check_name = result.get('check_name', '-')
            passed = result.get('passed', True)
            checked = result.get('checked_count', 0)
            passed_count = result.get('passed_count', 0)
            failed = result.get('failed_count', 0)
            
            if passed:
                status_html = '<span class="status-badge passed">✓ 通过</span>'
            else:
                status_html = '<span class="status-badge failed">✗ 未通过</span>'
            
            html += f"""
                    <tr>
                        <td>{check_name}</td>
                        <td>{status_html}</td>
                        <td>{checked}</td>
                        <td>{passed_count}</td>
                        <td>{failed}</td>
                    </tr>
"""
        
        html += """
                </tbody>
            </table>
"""
        return html
    
    @classmethod
    def _get_issues_section(cls, report_data: Dict[str, Any]) -> str:
        results = report_data.get('results', [])
        issues_data = []
        
        for result in results:
            category = result.get('category', '')
            check_name = result.get('check_name', '')
            for issue in result.get('issues', []):
                issue['category'] = category
                issue['check_name'] = check_name
                issues_data.append(issue)
        
        if not issues_data:
            return """
            <div class="no-issues">
                <div class="icon">🎉</div>
                <div class="title">恭喜！未发现问题</div>
                <div class="text">您的论文格式完全符合要求</div>
            </div>
"""
        
        html = """
            <div class="page-break"></div>
            <h2 class="section-title"><span class="icon">🔍</span> 问题详情列表</h2>
            <div class="issues-section">
"""
        
        for idx, issue in enumerate(issues_data, 1):
            html += cls._generate_issue_card(issue, idx)
        
        html += """
            </div>
"""
        return html
    
    @classmethod
    def _generate_issue_card(cls, issue: Dict[str, Any], index: int) -> str:
        severity = issue.get('severity', 'info')
        category = issue.get('category', '')
        location = issue.get('location', '-')
        message = issue.get('message', '')
        actual = issue.get('actual_value', '')
        expected = issue.get('expected_value', '')
        suggestion = issue.get('suggestion', '')
        
        icon = cls.SEVERITY_ICONS.get(severity, '•')
        severity_text = cls.SEVERITY_NAMES.get(severity, severity)
        category_text = cls.CATEGORY_NAMES.get(category, category)
        
        suggestion_html = ""
        if suggestion:
            suggestion_html = f"""
                <div class="issue-suggestion">
                    <div class="suggestion-title">
                        <span>💡</span>
                        <span>修改建议</span>
                    </div>
                    <div class="suggestion-text">{suggestion}</div>
                </div>
"""
        
        details_html = ""
        if actual or expected:
            details_html = '<div class="issue-details">'
            if actual:
                details_html += f"""
                <div class="detail-item">
                    <span class="label">当前值:</span>
                    <span class="value actual">{actual}</span>
                </div>
"""
            if expected:
                details_html += f"""
                <div class="detail-item">
                    <span class="label">期望值:</span>
                    <span class="value expected">{expected}</span>
                </div>
"""
            details_html += '</div>'
        
        return f"""
                <div class="issue-card">
                    <div class="issue-header {severity}">
                        <div class="issue-severity {severity}">
                            <span class="icon">{icon}</span>
                            <span class="text">【{severity_text}】</span>
                        </div>
                        <span class="issue-category">{category_text}</span>
                    </div>
                    <div class="issue-body">
                        <div class="issue-location">
                            <span class="icon">📍</span>
                            <span class="label">位置:</span>
                            <span class="value">{location}</span>
                        </div>
                        <div class="issue-message">{message}</div>
                        {suggestion_html}
                        {details_html}
                    </div>
                </div>
"""
    
    @classmethod
    def _get_footer_section(cls, report_data: Dict[str, Any]) -> str:
        return f"""
        </div>
        <div class="footer">
            <div class="generator">📄 论文格式检测工具</div>
            <div class="timestamp">报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
        </div>
"""
    
    @classmethod
    def _get_html_end(cls) -> str:
        return """
    </div>
</body>
</html>
"""
