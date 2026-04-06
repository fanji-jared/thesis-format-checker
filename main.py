#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import argparse
import os

def main():
    parser = argparse.ArgumentParser(description='论文格式检测工具')
    parser.add_argument('-i', '--input', type=str, help='输入文档路径')
    parser.add_argument('-o', '--output', type=str, default='output', help='输出目录')
    parser.add_argument('-c', '--config', type=str, help='配置文件路径')
    parser.add_argument('--cli', action='store_true', help='使用命令行模式')
    parser.add_argument('--no-gui', action='store_true', help='禁用图形界面')
    
    args = parser.parse_args()
    
    if args.cli or args.no_gui:
        run_cli_mode(args)
    else:
        run_gui_mode(args)


def run_gui_mode(args):
    import tkinter as tk
    from tkinter import messagebox
    
    try:
        from src.ui_main import MainWindow
    except ImportError:
        from ui_main import MainWindow
    
    root = tk.Tk()
    
    try:
        app = MainWindow(root)
        
        if args.input and os.path.exists(args.input):
            root.after(100, lambda: app._load_document(args.input))
        
        if args.config and os.path.exists(args.config):
            try:
                from src.config_manager import ConfigManager
            except ImportError:
                from config_manager import ConfigManager
            
            cm = ConfigManager()
            app.current_config = cm.load_config(args.config)
            app.current_config_path = args.config
        
        root.mainloop()
        
    except Exception as e:
        messagebox.showerror("启动错误", f"程序启动失败:\n{str(e)}")
        sys.exit(1)


def run_cli_mode(args):
    if not args.input:
        print("错误: 命令行模式需要指定输入文档路径 (-i/--input)")
        sys.exit(1)
    
    if not os.path.exists(args.input):
        print(f"错误: 文件不存在: {args.input}")
        sys.exit(1)
    
    print("=" * 60)
    print("论文格式检测工具 - 命令行模式")
    print("=" * 60)
    print(f"输入文件: {args.input}")
    print(f"输出目录: {args.output}")
    print(f"配置文件: {args.config or '默认配置'}")
    print()
    
    try:
        from src.docx_parser import DocxParser
        from src.config_manager import ConfigManager
        from src.format_checker import FormatChecker
        from src.models import ParseResult
    except ImportError:
        from docx_parser import DocxParser
        from config_manager import ConfigManager
        from format_checker import FormatChecker
        from models import ParseResult
    
    print("[1/3] 正在解析文档...")
    parser = DocxParser()
    result: ParseResult = parser.parse(args.input)
    
    if not result.success:
        print(f"错误: 文档解析失败 - {result.error_message}")
        sys.exit(1)
    
    doc_info = result.document_info
    print(f"  - 文件名: {doc_info.file_name}")
    print(f"  - 页数: {doc_info.page_count}")
    print(f"  - 段落数: {doc_info.paragraph_count}")
    print(f"  - 字符数: {doc_info.character_count}")
    print()
    
    print("[2/3] 正在加载配置...")
    config_manager = ConfigManager()
    
    if args.config:
        config = config_manager.load_config(args.config)
    else:
        configs = config_manager.get_config_list()
        if configs:
            config = config_manager.load_config(configs[0]['path'])
            print(f"  - 使用配置: {configs[0]['name']}")
        else:
            print("错误: 未找到配置文件")
            sys.exit(1)
    print()
    
    print("[3/3] 正在执行检测...")
    checker = FormatChecker(config)
    report = checker.check(doc_info)
    
    print()
    print("=" * 60)
    print("检测结果摘要")
    print("=" * 60)
    print(f"总问题数: {report.statistics.total_issues}")
    print(f"  - 错误: {report.statistics.error_count}")
    print(f"  - 警告: {report.statistics.warning_count}")
    print(f"  - 提示: {report.statistics.info_count}")
    print(f"通过率: {report.statistics.pass_rate:.1f}%")
    print()
    
    print("检测项详情:")
    for result in report.results:
        status = "✓" if result.passed else "✗"
        print(f"  [{status}] {result.check_name}: {result.passed_count}/{result.checked_count} 通过")
    print()
    
    os.makedirs(args.output, exist_ok=True)
    
    import json
    from datetime import datetime
    
    output_file = os.path.join(
        args.output,
        f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    )
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(report.to_dict(), f, ensure_ascii=False, indent=2)
    
    print(f"报告已保存至: {output_file}")


if __name__ == '__main__':
    main()
