#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import time
import unittest
from pathlib import Path
from typing import Dict, Any, Optional
from unittest.mock import MagicMock, patch

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.docx_parser import DocxParser
from src.config_manager import ConfigManager, ConfigValidationError
from src.format_checker import FormatChecker, CheckReport, Severity, CheckCategory
from src.models import DocumentInfo, ParseResult, ParagraphType, HeadingLevel
from src.cache import DocumentCache
from src.logger import get_logger

logger = get_logger()

SAMPLE_DOC_PATH = os.path.join(
    project_root, 
    ".document", 
    "基于Spring Boot+检索增强生成（RAG）的智能医疗网盘系统设计与实现.docx"
)
DEFAULT_CONFIG_PATH = os.path.join(project_root, "config", "config_default.json")


class TestDocumentParser(unittest.TestCase):
    """文档解析器测试"""
    
    @classmethod
    def setUpClass(cls):
        cls.parser = DocxParser
        cls.sample_doc_exists = os.path.exists(SAMPLE_DOC_PATH)
        if not cls.sample_doc_exists:
            logger.warning(f"示例文档不存在: {SAMPLE_DOC_PATH}")
    
    def test_parser_initialization(self):
        """测试解析器初始化"""
        parser = self.parser(SAMPLE_DOC_PATH if self.sample_doc_exists else "test.docx")
        self.assertIsNotNone(parser)
        self.assertEqual(parser.file_path, SAMPLE_DOC_PATH if self.sample_doc_exists else "test.docx")
    
    def test_parse_sample_document(self):
        """测试解析示例文档"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = self.parser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success, f"解析失败: {result.error_message}")
        self.assertIsNotNone(result.document_info)
        
        doc_info = result.document_info
        self.assertEqual(doc_info.file_name, "基于Spring Boot+检索增强生成（RAG）的智能医疗网盘系统设计与实现.docx")
        self.assertGreater(doc_info.paragraph_count, 0)
        self.assertGreater(doc_info.character_count, 0)
        logger.info(f"文档解析成功: {doc_info.paragraph_count}段落, {doc_info.character_count}字符")
    
    def test_extract_page_settings(self):
        """测试页面设置提取"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = self.parser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        doc_info = result.document_info
        
        self.assertGreater(len(doc_info.page_settings), 0)
        settings = doc_info.page_settings[0]
        
        self.assertGreater(settings.page_width, 0)
        self.assertGreater(settings.page_height, 0)
        logger.info(f"页面设置: {settings.page_width}x{settings.page_height}cm")
    
    def test_extract_chapters(self):
        """测试章节提取"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = self.parser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        doc_info = result.document_info
        
        self.assertGreater(len(doc_info.chapters), 0, "应该检测到章节")
        logger.info(f"检测到 {len(doc_info.chapters)} 个章节")
        
        for chapter in doc_info.chapters[:3]:
            logger.info(f"  - 第{chapter.chapter_number}章: {chapter.chapter_title}")
    
    def test_extract_figures(self):
        """测试图片提取"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = self.parser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        doc_info = result.document_info
        
        logger.info(f"检测到 {len(doc_info.figures)} 个图片")
        for fig in doc_info.figures[:3]:
            logger.info(f"  - {fig.figure_id}: {fig.caption}")
    
    def test_extract_tables(self):
        """测试表格提取"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = self.parser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        doc_info = result.document_info
        
        logger.info(f"检测到 {len(doc_info.tables)} 个表格")
    
    def test_extract_references(self):
        """测试参考文献提取"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = self.parser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        doc_info = result.document_info
        
        self.assertGreater(len(doc_info.references), 0, "应该检测到参考文献")
        logger.info(f"检测到 {len(doc_info.references)} 条参考文献")
        logger.info(f"检测到 {len(doc_info.citations)} 处引用")


class TestConfigManager(unittest.TestCase):
    """配置管理器测试"""
    
    @classmethod
    def setUpClass(cls):
        cls.config_manager = ConfigManager()
    
    def test_get_config_list(self):
        """测试获取配置列表"""
        configs = self.config_manager.get_config_list()
        self.assertIsInstance(configs, list)
        logger.info(f"找到 {len(configs)} 个配置文件")
        
        for config in configs:
            logger.info(f"  - {config['name']} v{config['version']}")
    
    def test_load_default_config(self):
        """测试加载默认配置"""
        configs = self.config_manager.get_config_list()
        
        if not configs:
            self.skipTest("没有可用的配置文件")
        
        config = self.config_manager.load_config(configs[0]['path'])
        self.assertIsNotNone(config)
        self.assertIn('meta', config)
        self.assertIn('page_settings', config)
        logger.info(f"成功加载配置: {config['meta']['name']}")
    
    def test_validate_config(self):
        """测试配置验证"""
        valid_config = self.config_manager.create_empty_config("测试配置")
        errors = self.config_manager.validate_config(valid_config)
        
        self.assertEqual(len(errors), 0, f"配置验证应该通过: {errors}")
        logger.info("配置验证通过")
    
    def test_invalid_config_validation(self):
        """测试无效配置验证"""
        invalid_config = {"meta": {"name": "测试"}}
        errors = self.config_manager.validate_config(invalid_config)
        
        self.assertGreater(len(errors), 0, "无效配置应该有错误")
        logger.info(f"检测到 {len(errors)} 个配置错误")
    
    def test_create_empty_config(self):
        """测试创建空配置"""
        config = self.config_manager.create_empty_config("测试配置", "这是一个测试配置")
        
        self.assertEqual(config['meta']['name'], "测试配置")
        self.assertEqual(config['meta']['description'], "这是一个测试配置")
        self.assertIn('page_settings', config)
        self.assertIn('body_text', config)
        logger.info("空配置创建成功")


class TestFormatChecker(unittest.TestCase):
    """格式检测引擎测试"""
    
    @classmethod
    def setUpClass(cls):
        cls.config_manager = ConfigManager()
        configs = cls.config_manager.get_config_list()
        
        if configs:
            cls.config = cls.config_manager.load_config(configs[0]['path'])
        else:
            cls.config = cls.config_manager.create_empty_config("测试配置")
        
        cls.sample_doc_exists = os.path.exists(SAMPLE_DOC_PATH)
    
    def test_checker_initialization(self):
        """测试检测器初始化"""
        checker = FormatChecker(self.config)
        self.assertIsNotNone(checker)
        self.assertEqual(checker.config, self.config)
    
    def test_check_page_settings(self):
        """测试页面设置检测"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        checker = FormatChecker(self.config)
        check_result = checker.check_page_settings(result.document_info)
        
        self.assertIsNotNone(check_result)
        self.assertEqual(check_result.category, CheckCategory.PAGE_SETTINGS)
        logger.info(f"页面设置检测: 通过 {check_result.passed_count}/{check_result.checked_count}")
    
    def test_check_toc(self):
        """测试目录检测"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        checker = FormatChecker(self.config)
        check_result = checker.check_toc(result.document_info)
        
        self.assertIsNotNone(check_result)
        logger.info(f"目录检测: 通过 {check_result.passed_count}/{check_result.checked_count}")
    
    def test_check_chapter_titles(self):
        """测试章节标题检测"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        checker = FormatChecker(self.config)
        check_result = checker.check_chapter_titles(result.document_info)
        
        self.assertIsNotNone(check_result)
        logger.info(f"章节标题检测: 通过 {check_result.passed_count}/{check_result.checked_count}")
    
    def test_check_body_text(self):
        """测试正文格式检测"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        checker = FormatChecker(self.config)
        check_result = checker.check_body_text(result.document_info)
        
        self.assertIsNotNone(check_result)
        logger.info(f"正文格式检测: 通过 {check_result.passed_count}/{check_result.checked_count}")
    
    def test_check_figures(self):
        """测试图片格式检测"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        checker = FormatChecker(self.config)
        check_result = checker.check_figures(result.document_info)
        
        self.assertIsNotNone(check_result)
        logger.info(f"图片格式检测: 通过 {check_result.passed_count}/{check_result.checked_count}")
    
    def test_check_tables(self):
        """测试表格格式检测"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        checker = FormatChecker(self.config)
        check_result = checker.check_tables(result.document_info)
        
        self.assertIsNotNone(check_result)
        logger.info(f"表格格式检测: 通过 {check_result.passed_count}/{check_result.checked_count}")
    
    def test_check_references(self):
        """测试参考文献检测"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        checker = FormatChecker(self.config)
        check_result = checker.check_references(result.document_info)
        
        self.assertIsNotNone(check_result)
        logger.info(f"参考文献检测: 通过 {check_result.passed_count}/{check_result.checked_count}")
    
    def test_full_check(self):
        """测试完整检测流程"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        checker = FormatChecker(self.config)
        report = checker.check(result.document_info)
        
        self.assertIsNotNone(report)
        self.assertEqual(len(report.results), 8)
        
        logger.info("=" * 50)
        logger.info("完整检测结果:")
        logger.info(f"  总问题数: {report.statistics.total_issues}")
        logger.info(f"  错误: {report.statistics.error_count}")
        logger.info(f"  警告: {report.statistics.warning_count}")
        logger.info(f"  提示: {report.statistics.info_count}")
        logger.info(f"  通过率: {report.statistics.pass_rate:.1f}%")
        logger.info("=" * 50)


class TestDocumentCache(unittest.TestCase):
    """文档缓存测试"""
    
    @classmethod
    def setUpClass(cls):
        cls.cache = DocumentCache(ttl=300)
        cls.sample_doc_exists = os.path.exists(SAMPLE_DOC_PATH)
    
    def test_cache_initialization(self):
        """测试缓存初始化"""
        self.assertIsNotNone(self.cache)
        self.assertTrue(self.cache.cache_dir.exists())
    
    def test_cache_set_and_get(self):
        """测试缓存存取"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        self.cache.set(SAMPLE_DOC_PATH, result.document_info)
        
        cached_data = self.cache.get(SAMPLE_DOC_PATH)
        self.assertIsNotNone(cached_data)
        self.assertEqual(cached_data.file_name, result.document_info.file_name)
        logger.info("缓存存取测试通过")
    
    def test_cache_invalidation(self):
        """测试缓存失效"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        parser = DocxParser(SAMPLE_DOC_PATH)
        result = parser.parse()
        
        self.assertTrue(result.success)
        
        self.cache.set(SAMPLE_DOC_PATH, result.document_info)
        
        self.cache.invalidate(SAMPLE_DOC_PATH)
        
        cached_data = self.cache.get(SAMPLE_DOC_PATH)
        self.assertIsNone(cached_data)
        logger.info("缓存失效测试通过")


class TestAIConfigGenerator(unittest.TestCase):
    """AI配置生成器测试"""
    
    def test_generator_initialization(self):
        """测试生成器初始化"""
        try:
            from src.ai_config_generator import AIConfigGenerator, OllamaConfig
            
            ollama_config = OllamaConfig()
            generator = AIConfigGenerator(ollama_config=ollama_config)
            
            self.assertIsNotNone(generator)
            logger.info("AI配置生成器初始化成功")
        except ImportError as e:
            self.skipTest(f"导入失败: {e}")
    
    def test_check_ollama_connection(self):
        """测试Ollama连接检查"""
        try:
            from src.ai_config_generator import AIConfigGenerator, OllamaConfig
            
            ollama_config = OllamaConfig()
            generator = AIConfigGenerator(ollama_config=ollama_config)
            
            connected, message = generator.check_ollama_connection()
            
            logger.info(f"Ollama连接状态: {message}")
            
            if connected:
                logger.info("Ollama服务可用")
            else:
                logger.warning(f"Ollama服务不可用: {message}")
        except ImportError as e:
            self.skipTest(f"导入失败: {e}")
    
    def test_parse_markdown_rules(self):
        """测试Markdown规则解析"""
        try:
            from src.ai_config_generator import AIConfigGenerator, OllamaConfig
            
            rules_path = os.path.join(project_root, ".document", "最终版完整论文格式规则.md")
            
            if not os.path.exists(rules_path):
                self.skipTest("规则文档不存在")
            
            ollama_config = OllamaConfig()
            generator = AIConfigGenerator(ollama_config=ollama_config)
            
            sections = generator.parse_markdown_rules(rules_path)
            
            self.assertIn('_full_content', sections)
            self.assertIn('_meta', sections)
            
            logger.info(f"规则解析成功: {sections['_meta']['section_count']} 个章节")
        except ImportError as e:
            self.skipTest(f"导入失败: {e}")


class TestIntegration(unittest.TestCase):
    """集成测试"""
    
    @classmethod
    def setUpClass(cls):
        cls.config_manager = ConfigManager()
        configs = cls.config_manager.get_config_list()
        
        if configs:
            cls.config = cls.config_manager.load_config(configs[0]['path'])
        else:
            cls.config = cls.config_manager.create_empty_config("测试配置")
        
        cls.sample_doc_exists = os.path.exists(SAMPLE_DOC_PATH)
        cls.cache = DocumentCache(ttl=300)
    
    def test_full_workflow(self):
        """测试完整工作流程"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        logger.info("=" * 60)
        logger.info("开始完整工作流程测试")
        logger.info("=" * 60)
        
        start_time = time.time()
        
        logger.info("[1/4] 解析文档...")
        parser = DocxParser(SAMPLE_DOC_PATH)
        parse_result = parser.parse()
        
        self.assertTrue(parse_result.success, f"文档解析失败: {parse_result.error_message}")
        doc_info = parse_result.document_info
        logger.info(f"  文档: {doc_info.file_name}")
        logger.info(f"  段落数: {doc_info.paragraph_count}")
        logger.info(f"  字符数: {doc_info.character_count}")
        
        logger.info("[2/4] 缓存文档信息...")
        self.cache.set(SAMPLE_DOC_PATH, doc_info)
        cached = self.cache.get(SAMPLE_DOC_PATH)
        self.assertIsNotNone(cached)
        logger.info("  缓存成功")
        
        logger.info("[3/4] 执行格式检测...")
        checker = FormatChecker(self.config)
        report = checker.check(doc_info)
        
        logger.info(f"  检测完成: {len(report.results)} 项")
        logger.info(f"  总问题: {report.statistics.total_issues}")
        logger.info(f"  通过率: {report.statistics.pass_rate:.1f}%")
        
        logger.info("[4/4] 生成报告...")
        report_dict = report.to_dict()
        self.assertIn('statistics', report_dict)
        self.assertIn('results', report_dict)
        
        output_dir = project_root / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        
        import json
        report_file = output_dir / "test_report.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report_dict, f, ensure_ascii=False, indent=2)
        
        logger.info(f"  报告已保存: {report_file}")
        
        elapsed_time = time.time() - start_time
        logger.info("=" * 60)
        logger.info(f"完整工作流程测试完成，耗时: {elapsed_time:.2f}秒")
        logger.info("=" * 60)
    
    def test_performance_large_document(self):
        """测试大文档处理性能"""
        if not self.sample_doc_exists:
            self.skipTest("示例文档不存在")
        
        logger.info("=" * 60)
        logger.info("开始性能测试")
        logger.info("=" * 60)
        
        times = []
        
        for i in range(3):
            start_time = time.time()
            
            parser = DocxParser(SAMPLE_DOC_PATH)
            result = parser.parse()
            
            if result.success:
                checker = FormatChecker(self.config)
                report = checker.check(result.document_info)
            
            elapsed = time.time() - start_time
            times.append(elapsed)
            logger.info(f"  第{i+1}次: {elapsed:.2f}秒")
        
        avg_time = sum(times) / len(times)
        logger.info(f"平均耗时: {avg_time:.2f}秒")
        
        self.assertLess(avg_time, 30, "大文档处理应该在30秒内完成")
        logger.info("性能测试通过")


class TestErrorHandling(unittest.TestCase):
    """错误处理测试"""
    
    def test_invalid_file_path(self):
        """测试无效文件路径"""
        parser = DocxParser("nonexistent.docx")
        result = parser.parse()
        
        self.assertFalse(result.success)
        self.assertIsNotNone(result.error_message)
        logger.info(f"正确处理无效文件: {result.error_message}")
    
    def test_invalid_config(self):
        """测试无效配置"""
        checker = FormatChecker({})
        mock_doc_info = MagicMock()
        mock_doc_info.page_settings = []
        mock_doc_info.file_name = "test.docx"
        
        result = checker.check_page_settings(mock_doc_info)
        self.assertIsNotNone(result)
    
    def test_corrupted_document(self):
        """测试损坏文档处理"""
        fake_doc_path = os.path.join(project_root, "test_fake.docx")
        
        with open(fake_doc_path, 'wb') as f:
            f.write(b"fake content")
        
        try:
            parser = DocxParser(fake_doc_path)
            result = parser.parse()
            
            self.assertFalse(result.success)
            logger.info(f"正确处理损坏文档: {result.error_message}")
        finally:
            if os.path.exists(fake_doc_path):
                os.remove(fake_doc_path)


def run_tests():
    """运行所有测试"""
    logger.info("=" * 60)
    logger.info("开始运行集成测试")
    logger.info("=" * 60)
    
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    suite.addTests(loader.loadTestsFromTestCase(TestDocumentParser))
    suite.addTests(loader.loadTestsFromTestCase(TestConfigManager))
    suite.addTests(loader.loadTestsFromTestCase(TestFormatChecker))
    suite.addTests(loader.loadTestsFromTestCase(TestDocumentCache))
    suite.addTests(loader.loadTestsFromTestCase(TestAIConfigGenerator))
    suite.addTests(loader.loadTestsFromTestCase(TestIntegration))
    suite.addTests(loader.loadTestsFromTestCase(TestErrorHandling))
    
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    logger.info("=" * 60)
    logger.info("测试结果摘要")
    logger.info("=" * 60)
    logger.info(f"运行测试: {result.testsRun}")
    logger.info(f"成功: {result.testsRun - len(result.failures) - len(result.errors)}")
    logger.info(f"失败: {len(result.failures)}")
    logger.info(f"错误: {len(result.errors)}")
    logger.info(f"跳过: {len(result.skipped)}")
    
    return result


if __name__ == '__main__':
    result = run_tests()
    sys.exit(0 if result.wasSuccessful() else 1)
