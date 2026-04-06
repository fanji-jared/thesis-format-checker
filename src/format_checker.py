from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any, Callable
from enum import Enum
from datetime import datetime

from .models import (
    DocumentInfo, ParagraphInfo, ParagraphType, ParagraphFormat,
    FontInfo, RunInfo, PageSettings, SectionInfo, TOCInfo, TOCEntry,
    FigureInfo, TableInfo, FormulaInfo, ReferenceInfo, CitationInfo,
    ChapterInfo, SubSectionInfo, HeadingLevel
)
from .logger import get_logger

logger = get_logger()


class Severity(Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


class CheckCategory(Enum):
    PAGE_SETTINGS = "page_settings"
    TOC = "toc"
    CHAPTER_TITLE = "chapter_title"
    BODY_TEXT = "body_text"
    FIGURE = "figure"
    TABLE = "table"
    FORMULA = "formula"
    REFERENCE = "reference"


@dataclass
class Issue:
    severity: Severity
    message: str
    location: Optional[str] = None
    actual_value: Optional[Any] = None
    expected_value: Optional[Any] = None
    suggestion: Optional[str] = None
    paragraph_index: Optional[int] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            'severity': self.severity.value,
            'message': self.message,
            'location': self.location,
            'actual_value': self.actual_value,
            'expected_value': self.expected_value,
            'suggestion': self.suggestion,
            'paragraph_index': self.paragraph_index
        }


@dataclass
class CheckResult:
    category: CheckCategory
    check_name: str
    passed: bool
    issues: List[Issue] = field(default_factory=list)
    checked_count: int = 0
    passed_count: int = 0
    failed_count: int = 0

    def add_issue(self, issue: Issue):
        self.issues.append(issue)
        self.failed_count += 1
        self.passed = False

    def add_issues(self, issues: List[Issue]):
        for issue in issues:
            self.add_issue(issue)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'category': self.category.value,
            'check_name': self.check_name,
            'passed': self.passed,
            'issues': [issue.to_dict() for issue in self.issues],
            'checked_count': self.checked_count,
            'passed_count': self.passed_count,
            'failed_count': self.failed_count
        }


@dataclass
class CheckStatistics:
    total_checks: int = 0
    total_issues: int = 0
    error_count: int = 0
    warning_count: int = 0
    info_count: int = 0
    passed_checks: int = 0
    failed_checks: int = 0

    @property
    def pass_rate(self) -> float:
        if self.total_checks == 0:
            return 100.0
        return (self.passed_checks / self.total_checks) * 100

    def to_dict(self) -> Dict[str, Any]:
        return {
            'total_checks': self.total_checks,
            'total_issues': self.total_issues,
            'error_count': self.error_count,
            'warning_count': self.warning_count,
            'info_count': self.info_count,
            'passed_checks': self.passed_checks,
            'failed_checks': self.failed_checks,
            'pass_rate': round(self.pass_rate, 2)
        }


@dataclass
class CheckReport:
    document_name: str
    check_time: str
    results: List[CheckResult] = field(default_factory=list)
    statistics: CheckStatistics = field(default_factory=CheckStatistics)
    config_name: Optional[str] = None
    config_version: Optional[str] = None

    def add_result(self, result: CheckResult):
        self.results.append(result)
        self._update_statistics(result)

    def _update_statistics(self, result: CheckResult):
        self.statistics.total_checks += 1
        if result.passed:
            self.statistics.passed_checks += 1
        else:
            self.statistics.failed_checks += 1

        for issue in result.issues:
            self.statistics.total_issues += 1
            if issue.severity == Severity.ERROR:
                self.statistics.error_count += 1
            elif issue.severity == Severity.WARNING:
                self.statistics.warning_count += 1
            else:
                self.statistics.info_count += 1

    def get_issues_by_severity(self, severity: Severity) -> List[Issue]:
        issues = []
        for result in self.results:
            for issue in result.issues:
                if issue.severity == severity:
                    issues.append(issue)
        return issues

    def get_issues_by_category(self, category: CheckCategory) -> List[Issue]:
        issues = []
        for result in self.results:
            if result.category == category:
                issues.extend(result.issues)
        return issues

    def to_dict(self) -> Dict[str, Any]:
        return {
            'document_name': self.document_name,
            'check_time': self.check_time,
            'config_name': self.config_name,
            'config_version': self.config_version,
            'statistics': self.statistics.to_dict(),
            'results': [result.to_dict() for result in self.results]
        }


class CheckProgress:
    def __init__(self, stage: str, progress: float, message: str, details: Optional[Dict[str, Any]] = None):
        self.stage = stage
        self.progress = progress
        self.message = message
        self.details = details or {}
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'stage': self.stage,
            'progress': self.progress,
            'message': self.message,
            'details': self.details
        }


class FormatChecker:
    SIZE_MAP = {
        '初号': 42, '小初': 36, '一号': 26, '小一': 24,
        '二号': 22, '小二': 18, '三号': 16, '小三': 15,
        '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
        '六号': 7.5, '小六': 6.5, '七号': 5.5, '八号': 5
    }

    TOLERANCE_MARGIN = 0.3
    TOLERANCE_FONT_SIZE = 1.0

    def __init__(self, config: Dict[str, Any], progress_callback: Optional[Callable[[CheckProgress], None]] = None):
        self.config = config
        self._size_map = self.SIZE_MAP.copy()
        self._progress_callback = progress_callback

    def _update_progress(self, stage: str, progress: float, message: str, details: Optional[Dict[str, Any]] = None):
        if self._progress_callback:
            progress_info = CheckProgress(stage, progress, message, details)
            self._progress_callback(progress_info)
        logger.debug(f"[{stage}] {progress*100:.0f}% - {message}")

    def check(self, document_info: DocumentInfo) -> CheckReport:
        logger.info(f"开始格式检测: {document_info.file_name}")
        
        report = CheckReport(
            document_name=document_info.file_name,
            check_time=datetime.now().isoformat(),
            config_name=self.config.get('meta', {}).get('name'),
            config_version=self.config.get('meta', {}).get('version')
        )

        check_steps = [
            (0.05, "页面设置检测", self.check_page_settings),
            (0.15, "目录格式检测", self.check_toc),
            (0.30, "章节标题检测", self.check_chapter_titles),
            (0.45, "正文格式检测", self.check_body_text),
            (0.60, "图片格式检测", self.check_figures),
            (0.75, "表格格式检测", self.check_tables),
            (0.85, "公式格式检测", self.check_formulas),
            (0.95, "参考文献检测", self.check_references),
        ]

        for progress, stage_name, check_func in check_steps:
            self._update_progress("checking", progress, f"正在{stage_name}...")
            result = check_func(document_info)
            report.add_result(result)
            logger.info(f"{stage_name}: 通过 {result.passed_count}/{result.checked_count}")

        self._update_progress("completed", 1.0, "检测完成")
        logger.info(f"检测完成: 总问题 {report.statistics.total_issues}, 通过率 {report.statistics.pass_rate:.1f}%")

        return report

    def check_page_settings(self, document_info: DocumentInfo) -> CheckResult:
        result = CheckResult(
            category=CheckCategory.PAGE_SETTINGS,
            check_name="页面设置检测",
            passed=True
        )

        if not document_info.page_settings:
            result.add_issue(Issue(
                severity=Severity.ERROR,
                message="无法获取页面设置信息",
                suggestion="请检查文档是否损坏"
            ))
            return result

        page_config = self.config.get('page_settings', {})
        margins_config = page_config.get('margins', {})
        binding_config = page_config.get('binding', {})
        header_footer_config = page_config.get('header_footer', {})

        for idx, settings in enumerate(document_info.page_settings):
            section_name = f"第{idx + 1}节" if len(document_info.page_settings) > 1 else "文档"

            if margins_config:
                result.checked_count += 4

                if self._check_margin(settings.top_margin, margins_config.get('top')):
                    result.passed_count += 1
                else:
                    result.add_issue(Issue(
                        severity=Severity.ERROR,
                        message=f"{section_name}上边距不符合要求",
                        location=f"{section_name}页面设置",
                        actual_value=f"{settings.top_margin:.2f}cm",
                        expected_value=f"{margins_config.get('top')}cm",
                        suggestion=f"请将上边距设置为{margins_config.get('top')}cm"
                    ))

                if self._check_margin(settings.bottom_margin, margins_config.get('bottom')):
                    result.passed_count += 1
                else:
                    result.add_issue(Issue(
                        severity=Severity.ERROR,
                        message=f"{section_name}下边距不符合要求",
                        location=f"{section_name}页面设置",
                        actual_value=f"{settings.bottom_margin:.2f}cm",
                        expected_value=f"{margins_config.get('bottom')}cm",
                        suggestion=f"请将下边距设置为{margins_config.get('bottom')}cm"
                    ))

                if self._check_margin(settings.left_margin, margins_config.get('left')):
                    result.passed_count += 1
                else:
                    result.add_issue(Issue(
                        severity=Severity.ERROR,
                        message=f"{section_name}左边距不符合要求",
                        location=f"{section_name}页面设置",
                        actual_value=f"{settings.left_margin:.2f}cm",
                        expected_value=f"{margins_config.get('left')}cm",
                        suggestion=f"请将左边距设置为{margins_config.get('left')}cm"
                    ))

                if self._check_margin(settings.right_margin, margins_config.get('right')):
                    result.passed_count += 1
                else:
                    result.add_issue(Issue(
                        severity=Severity.ERROR,
                        message=f"{section_name}右边距不符合要求",
                        location=f"{section_name}页面设置",
                        actual_value=f"{settings.right_margin:.2f}cm",
                        expected_value=f"{margins_config.get('right')}cm",
                        suggestion=f"请将右边距设置为{margins_config.get('right')}cm"
                    ))

            if binding_config:
                result.checked_count += 1
                expected_gutter = binding_config.get('width', 0)
                if self._check_margin(settings.gutter, expected_gutter):
                    result.passed_count += 1
                else:
                    result.add_issue(Issue(
                        severity=Severity.WARNING,
                        message=f"{section_name}装订线宽度不符合要求",
                        location=f"{section_name}页面设置",
                        actual_value=f"{settings.gutter:.2f}cm",
                        expected_value=f"{expected_gutter}cm",
                        suggestion=f"请将装订线宽度设置为{expected_gutter}cm"
                    ))

            if header_footer_config:
                result.checked_count += 2

                expected_header = header_footer_config.get('header_margin')
                if expected_header is not None:
                    if self._check_margin(settings.header_distance, expected_header):
                        result.passed_count += 1
                    else:
                        result.add_issue(Issue(
                            severity=Severity.WARNING,
                            message=f"{section_name}页眉边距不符合要求",
                            location=f"{section_name}页面设置",
                            actual_value=f"{settings.header_distance:.2f}cm",
                            expected_value=f"{expected_header}cm",
                            suggestion=f"请将页眉边距设置为{expected_header}cm"
                        ))

                expected_footer = header_footer_config.get('footer_margin')
                if expected_footer is not None:
                    if self._check_margin(settings.footer_distance, expected_footer):
                        result.passed_count += 1
                    else:
                        result.add_issue(Issue(
                            severity=Severity.WARNING,
                            message=f"{section_name}页脚边距不符合要求",
                            location=f"{section_name}页面设置",
                            actual_value=f"{settings.footer_distance:.2f}cm",
                            expected_value=f"{expected_footer}cm",
                            suggestion=f"请将页脚边距设置为{expected_footer}cm"
                        ))

        if result.checked_count > 0:
            result.passed = result.failed_count == 0

        return result

    def check_toc(self, document_info: DocumentInfo) -> CheckResult:
        result = CheckResult(
            category=CheckCategory.TOC,
            check_name="目录格式检测",
            passed=True
        )

        toc_config = self.config.get('toc_settings', {})

        result.checked_count += 1
        if not document_info.has_toc or not document_info.toc_info:
            result.add_issue(Issue(
                severity=Severity.ERROR,
                message="文档缺少目录",
                suggestion="请在文档开头添加目录"
            ))
            return result
        result.passed_count += 1

        if toc_config.get('auto_generated', True):
            result.checked_count += 1
            if document_info.toc_info.is_auto_generated:
                result.passed_count += 1
            else:
                result.add_issue(Issue(
                    severity=Severity.WARNING,
                    message="目录未使用Word自动生成功能",
                    actual_value="手动创建的目录",
                    expected_value="自动生成的目录",
                    suggestion="建议使用Word的目录自动生成功能，以便自动更新页码"
                ))

        display_levels = toc_config.get('display_levels', 3)
        result.checked_count += 1
        max_level = max((entry.level for entry in document_info.toc_info.entries), default=0)
        if max_level <= display_levels:
            result.passed_count += 1
        else:
            result.add_issue(Issue(
                severity=Severity.INFO,
                message=f"目录显示级别超过配置要求",
                actual_value=f"{max_level}级",
                expected_value=f"{display_levels}级",
                suggestion=f"目录应只显示{display_levels}级标题"
            ))

        result.checked_count += 1
        if document_info.toc_info.entries:
            result.passed_count += 1
            title_config = toc_config.get('title', {})
            if title_config:
                title_issues = self._check_toc_title_format(document_info, title_config)
                result.add_issues(title_issues)
        else:
            result.add_issue(Issue(
                severity=Severity.ERROR,
                message="目录内容为空",
                suggestion="请确保文档中有标题并正确生成目录"
            ))

        return result

    def check_chapter_titles(self, document_info: DocumentInfo) -> CheckResult:
        result = CheckResult(
            category=CheckCategory.CHAPTER_TITLE,
            check_name="章节标题格式检测",
            passed=True
        )

        chapter_config = self.config.get('chapter_titles', {})

        level1_config = chapter_config.get('level1', {})
        level2_config = chapter_config.get('level2', {})
        level3_config = chapter_config.get('level3', {})

        for para in document_info.all_paragraphs:
            if para.heading_level == HeadingLevel.LEVEL_1:
                result.checked_count += 1
                issues = self._check_heading_format(para, level1_config, "一级标题")
                if not issues:
                    result.passed_count += 1
                else:
                    result.add_issues(issues)

            elif para.heading_level == HeadingLevel.LEVEL_2:
                result.checked_count += 1
                issues = self._check_heading_format(para, level2_config, "二级标题")
                if not issues:
                    result.passed_count += 1
                else:
                    result.add_issues(issues)

            elif para.heading_level == HeadingLevel.LEVEL_3:
                result.checked_count += 1
                issues = self._check_heading_format(para, level3_config, "三级标题")
                if not issues:
                    result.passed_count += 1
                else:
                    result.add_issues(issues)

        if result.checked_count == 0:
            result.passed = True
        else:
            result.passed = result.failed_count == 0

        return result

    def check_body_text(self, document_info: DocumentInfo) -> CheckResult:
        result = CheckResult(
            category=CheckCategory.BODY_TEXT,
            check_name="正文格式检测",
            passed=True
        )

        body_config = self.config.get('body_text', {})

        body_paragraphs = [
            para for para in document_info.all_paragraphs
            if para.paragraph_type == ParagraphType.BODY_TEXT and para.text.strip()
        ]

        if not body_paragraphs:
            result.passed = True
            return result

        sample_size = min(len(body_paragraphs), 50)
        sample_paragraphs = body_paragraphs[:sample_size]

        for para in sample_paragraphs:
            result.checked_count += 1
            issues = self._check_body_paragraph_format(para, body_config)
            if not issues:
                result.passed_count += 1
            else:
                result.add_issues(issues)

        if result.checked_count > 0:
            result.passed = result.failed_count == 0

        return result

    def check_figures(self, document_info: DocumentInfo) -> CheckResult:
        result = CheckResult(
            category=CheckCategory.FIGURE,
            check_name="图片格式检测",
            passed=True
        )

        figure_config = self.config.get('figures', {})

        if not document_info.figures:
            result.passed = True
            return result

        for fig in document_info.figures:
            result.checked_count += 1
            issues = self._check_figure_format(fig, figure_config)
            if not issues:
                result.passed_count += 1
            else:
                result.add_issues(issues)

        numbering_config = figure_config.get('numbering', {})
        if numbering_config.get('by_chapter', True):
            result.checked_count += 1
            numbering_issues = self._check_figure_numbering(document_info.figures)
            if not numbering_issues:
                result.passed_count += 1
            else:
                result.add_issues(numbering_issues)

        if result.checked_count > 0:
            result.passed = result.failed_count == 0

        return result

    def check_tables(self, document_info: DocumentInfo) -> CheckResult:
        result = CheckResult(
            category=CheckCategory.TABLE,
            check_name="表格格式检测",
            passed=True
        )

        table_config = self.config.get('tables', {})

        if not document_info.tables:
            result.passed = True
            return result

        for table in document_info.tables:
            result.checked_count += 1
            issues = self._check_table_format(table, table_config)
            if not issues:
                result.passed_count += 1
            else:
                result.add_issues(issues)

        three_line_config = table_config.get('three_line_table', {})
        if three_line_config.get('enabled', True):
            result.checked_count += 1
            three_line_issues = self._check_three_line_table(document_info.tables, three_line_config)
            if not three_line_issues:
                result.passed_count += 1
            else:
                result.add_issues(three_line_issues)

        numbering_config = table_config.get('numbering', {})
        if numbering_config.get('by_chapter', True):
            result.checked_count += 1
            numbering_issues = self._check_table_numbering(document_info.tables)
            if not numbering_issues:
                result.passed_count += 1
            else:
                result.add_issues(numbering_issues)

        if result.checked_count > 0:
            result.passed = result.failed_count == 0

        return result

    def check_formulas(self, document_info: DocumentInfo) -> CheckResult:
        result = CheckResult(
            category=CheckCategory.FORMULA,
            check_name="公式格式检测",
            passed=True
        )

        formula_config = self.config.get('formulas', {})

        if not document_info.formulas:
            result.passed = True
            return result

        for formula in document_info.formulas:
            result.checked_count += 1
            issues = self._check_formula_format(formula, formula_config)
            if not issues:
                result.passed_count += 1
            else:
                result.add_issues(issues)

        numbering_config = formula_config.get('numbering', {})
        if numbering_config.get('by_chapter', True):
            result.checked_count += 1
            numbering_issues = self._check_formula_numbering(document_info.formulas)
            if not numbering_issues:
                result.passed_count += 1
            else:
                result.add_issues(numbering_issues)

        if result.checked_count > 0:
            result.passed = result.failed_count == 0

        return result

    def check_references(self, document_info: DocumentInfo) -> CheckResult:
        result = CheckResult(
            category=CheckCategory.REFERENCE,
            check_name="参考文献格式检测",
            passed=True
        )

        ref_config = self.config.get('references', {})

        result.checked_count += 1
        if not document_info.references:
            result.add_issue(Issue(
                severity=Severity.ERROR,
                message="文档缺少参考文献",
                suggestion="请在文档末尾添加参考文献列表"
            ))
            return result
        result.passed_count += 1

        citation_config = ref_config.get('in_text_citation', {})
        if citation_config:
            result.checked_count += 1
            citation_issues = self._check_citations(document_info.citations, citation_config)
            if not citation_issues:
                result.passed_count += 1
            else:
                result.add_issues(citation_issues)

        cross_ref_required = ref_config.get('cross_reference_required', True)
        if cross_ref_required:
            result.checked_count += 1
            cross_ref_issues = self._check_cross_references(document_info)
            if not cross_ref_issues:
                result.passed_count += 1
            else:
                result.add_issues(cross_ref_issues)

        list_config = ref_config.get('list_format', {})
        if list_config:
            for ref in document_info.references[:min(len(document_info.references), 30)]:
                result.checked_count += 1
                issues = self._check_reference_format(ref, list_config)
                if not issues:
                    result.passed_count += 1
                else:
                    result.add_issues(issues)

        if result.checked_count > 0:
            result.passed = result.failed_count == 0

        return result

    def _check_margin(self, actual: float, expected: float) -> bool:
        if expected is None:
            return True
        return abs(actual - expected) <= self.TOLERANCE_MARGIN

    def _check_font_size(self, actual: Optional[float], expected: str) -> bool:
        if actual is None or not expected:
            return True
        expected_pt = self._size_map.get(expected)
        if expected_pt is None:
            try:
                expected_pt = float(expected)
            except ValueError:
                return True
        return abs(actual - expected_pt) <= self.TOLERANCE_FONT_SIZE

    def _check_font_name(self, actual: Optional[str], expected: str) -> bool:
        if not actual or not expected:
            return True
        return expected in actual or actual in expected

    def _get_paragraph_font_info(self, para: ParagraphInfo) -> Dict[str, Any]:
        font_info = {
            'chinese_font': None,
            'english_font': None,
            'size': None,
            'bold': False
        }

        if para.runs:
            first_run = para.runs[0]
            if first_run.font:
                font_info['chinese_font'] = first_run.font.name_east_asia or first_run.font.name
                font_info['english_font'] = first_run.font.name_ascii or first_run.font.name
                font_info['size'] = first_run.font.size
                font_info['bold'] = first_run.font.bold

        return font_info

    def _check_toc_title_format(self, document_info: DocumentInfo, title_config: Dict) -> List[Issue]:
        issues = []
        return issues

    def _check_heading_format(self, para: ParagraphInfo, config: Dict, level_name: str) -> List[Issue]:
        issues = []

        if not config:
            return issues

        font_info = self._get_paragraph_font_info(para)
        location = f"{level_name}: {para.text[:30]}..."

        expected_font = config.get('font')
        if expected_font:
            actual_font = font_info.get('chinese_font')
            if not self._check_font_name(actual_font, expected_font):
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message=f"{level_name}字体不符合要求",
                    location=location,
                    actual_value=actual_font,
                    expected_value=expected_font,
                    suggestion=f"请将字体设置为{expected_font}",
                    paragraph_index=para.index
                ))

        expected_size = config.get('size')
        if expected_size:
            actual_size = font_info.get('size')
            if not self._check_font_size(actual_size, expected_size):
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message=f"{level_name}字号不符合要求",
                    location=location,
                    actual_value=f"{actual_size}磅" if actual_size else "未设置",
                    expected_value=expected_size,
                    suggestion=f"请将字号设置为{expected_size}",
                    paragraph_index=para.index
                ))

        expected_bold = config.get('bold')
        if expected_bold is not None:
            actual_bold = font_info.get('bold', False)
            if actual_bold != expected_bold:
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message=f"{level_name}加粗设置不符合要求",
                    location=location,
                    actual_value="已加粗" if actual_bold else "未加粗",
                    expected_value="已加粗" if expected_bold else "未加粗",
                    suggestion=f"请{'设置' if expected_bold else '取消'}加粗",
                    paragraph_index=para.index
                ))

        expected_alignment = config.get('alignment')
        if expected_alignment and para.format:
            actual_alignment = para.format.alignment
            if actual_alignment != expected_alignment:
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message=f"{level_name}对齐方式不符合要求",
                    location=location,
                    actual_value=actual_alignment or "未设置",
                    expected_value=expected_alignment,
                    suggestion=f"请将对齐方式设置为{expected_alignment}",
                    paragraph_index=para.index
                ))

        expected_line_spacing = config.get('line_spacing')
        if expected_line_spacing is not None and para.format:
            actual_line_spacing = para.format.line_spacing
            if actual_line_spacing is not None:
                if abs(actual_line_spacing - expected_line_spacing) > 0.1:
                    issues.append(Issue(
                        severity=Severity.WARNING,
                        message=f"{level_name}行距不符合要求",
                        location=location,
                        actual_value=f"{actual_line_spacing}倍",
                        expected_value=f"{expected_line_spacing}倍",
                        suggestion=f"请将行距设置为{expected_line_spacing}倍",
                        paragraph_index=para.index
                    ))

        return issues

    def _check_body_paragraph_format(self, para: ParagraphInfo, config: Dict) -> List[Issue]:
        issues = []

        if not config:
            return issues

        font_info = self._get_paragraph_font_info(para)
        location = f"正文段落: {para.text[:30]}..."

        expected_font = config.get('font')
        if expected_font:
            actual_font = font_info.get('chinese_font')
            if not self._check_font_name(actual_font, expected_font):
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message="正文字体不符合要求",
                    location=location,
                    actual_value=actual_font,
                    expected_value=expected_font,
                    suggestion=f"请将正文字体设置为{expected_font}",
                    paragraph_index=para.index
                ))

        expected_size = config.get('size')
        if expected_size:
            actual_size = font_info.get('size')
            if not self._check_font_size(actual_size, expected_size):
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message="正文字号不符合要求",
                    location=location,
                    actual_value=f"{actual_size}磅" if actual_size else "未设置",
                    expected_value=expected_size,
                    suggestion=f"请将正文字号设置为{expected_size}",
                    paragraph_index=para.index
                ))

        expected_indent = config.get('first_line_indent')
        if expected_indent is not None and para.format:
            actual_indent = para.format.first_line_indent
            expected_cm = expected_indent * 0.35
            if actual_indent is None or abs(actual_indent - expected_cm) > 0.1:
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message="正文首行缩进不符合要求",
                    location=location,
                    actual_value=f"{actual_indent:.2f}cm" if actual_indent else "未设置",
                    expected_value=f"{expected_indent}字符(约{expected_cm:.2f}cm)",
                    suggestion=f"请将首行缩进设置为{expected_indent}字符",
                    paragraph_index=para.index
                ))

        expected_alignment = config.get('alignment')
        if expected_alignment and para.format:
            actual_alignment = para.format.alignment
            if actual_alignment != expected_alignment:
                issues.append(Issue(
                    severity=Severity.WARNING,
                    message="正文对齐方式不符合要求",
                    location=location,
                    actual_value=actual_alignment or "未设置",
                    expected_value=expected_alignment,
                    suggestion=f"请将对齐方式设置为{expected_alignment}",
                    paragraph_index=para.index
                ))

        expected_line_spacing = config.get('line_spacing')
        if expected_line_spacing is not None and para.format:
            actual_line_spacing = para.format.line_spacing
            if actual_line_spacing is not None:
                if abs(actual_line_spacing - expected_line_spacing) > 0.1:
                    issues.append(Issue(
                        severity=Severity.WARNING,
                        message="正文行距不符合要求",
                        location=location,
                        actual_value=f"{actual_line_spacing}倍",
                        expected_value=f"{expected_line_spacing}倍",
                        suggestion=f"请将行距设置为{expected_line_spacing}倍",
                        paragraph_index=para.index
                    ))

        return issues

    def _check_figure_format(self, fig: FigureInfo, config: Dict) -> List[Issue]:
        issues = []

        if not config:
            return issues

        caption_config = config.get('caption', {})
        if caption_config and fig.caption_paragraph:
            font_info = self._get_paragraph_font_info(fig.caption_paragraph)
            location = f"图片: {fig.figure_id}"

            expected_font = caption_config.get('font')
            if expected_font:
                actual_font = font_info.get('chinese_font')
                if not self._check_font_name(actual_font, expected_font):
                    issues.append(Issue(
                        severity=Severity.ERROR,
                        message="图题字体不符合要求",
                        location=location,
                        actual_value=actual_font,
                        expected_value=expected_font,
                        suggestion=f"请将图题字体设置为{expected_font}",
                        paragraph_index=fig.caption_paragraph.index
                    ))

            expected_size = caption_config.get('size')
            if expected_size:
                actual_size = font_info.get('size')
                if not self._check_font_size(actual_size, expected_size):
                    issues.append(Issue(
                        severity=Severity.ERROR,
                        message="图题字号不符合要求",
                        location=location,
                        actual_value=f"{actual_size}磅" if actual_size else "未设置",
                        expected_value=expected_size,
                        suggestion=f"请将图题字号设置为{expected_size}",
                        paragraph_index=fig.caption_paragraph.index
                    ))

            expected_bold = caption_config.get('bold')
            if expected_bold is not None:
                actual_bold = font_info.get('bold', False)
                if actual_bold != expected_bold:
                    issues.append(Issue(
                        severity=Severity.ERROR,
                        message="图题加粗设置不符合要求",
                        location=location,
                        actual_value="已加粗" if actual_bold else "未加粗",
                        expected_value="已加粗" if expected_bold else "未加粗",
                        suggestion=f"请{'设置' if expected_bold else '取消'}图题加粗",
                        paragraph_index=fig.caption_paragraph.index
                    ))

            expected_alignment = caption_config.get('alignment')
            if expected_alignment and fig.caption_paragraph.format:
                actual_alignment = fig.caption_paragraph.format.alignment
                if actual_alignment != expected_alignment:
                    issues.append(Issue(
                        severity=Severity.ERROR,
                        message="图题对齐方式不符合要求",
                        location=location,
                        actual_value=actual_alignment or "未设置",
                        expected_value=expected_alignment,
                        suggestion=f"请将图题对齐方式设置为{expected_alignment}",
                        paragraph_index=fig.caption_paragraph.index
                    ))

        return issues

    def _check_figure_numbering(self, figures: List[FigureInfo]) -> List[Issue]:
        issues = []

        chapter_figures = {}
        for fig in figures:
            if fig.chapter_number is not None:
                if fig.chapter_number not in chapter_figures:
                    chapter_figures[fig.chapter_number] = []
                chapter_figures[fig.chapter_number].append(fig)

        for chapter_num, figs in chapter_figures.items():
            expected_nums = list(range(1, len(figs) + 1))
            actual_nums = [f.figure_number for f in figs if f.figure_number is not None]

            if sorted(actual_nums) != expected_nums[:len(actual_nums)]:
                issues.append(Issue(
                    severity=Severity.WARNING,
                    message=f"第{chapter_num}章图片编号不连续",
                    actual_value=f"现有编号: {sorted(actual_nums)}",
                    expected_value=f"应为: 1到{len(figs)}",
                    suggestion="请检查图片编号是否连续"
                ))

        return issues

    def _check_table_format(self, table: TableInfo, config: Dict) -> List[Issue]:
        issues = []

        if not config:
            return issues

        caption_config = config.get('caption', {})
        if caption_config and table.caption_paragraph:
            font_info = self._get_paragraph_font_info(table.caption_paragraph)
            location = f"表格: {table.table_id}"

            expected_font = caption_config.get('font')
            if expected_font:
                actual_font = font_info.get('chinese_font')
                if not self._check_font_name(actual_font, expected_font):
                    issues.append(Issue(
                        severity=Severity.ERROR,
                        message="表题字体不符合要求",
                        location=location,
                        actual_value=actual_font,
                        expected_value=expected_font,
                        suggestion=f"请将表题字体设置为{expected_font}",
                        paragraph_index=table.caption_paragraph.index
                    ))

            expected_size = caption_config.get('size')
            if expected_size:
                actual_size = font_info.get('size')
                if not self._check_font_size(actual_size, expected_size):
                    issues.append(Issue(
                        severity=Severity.ERROR,
                        message="表题字号不符合要求",
                        location=location,
                        actual_value=f"{actual_size}磅" if actual_size else "未设置",
                        expected_value=expected_size,
                        suggestion=f"请将表题字号设置为{expected_size}",
                        paragraph_index=table.caption_paragraph.index
                    ))

            expected_bold = caption_config.get('bold')
            if expected_bold is not None:
                actual_bold = font_info.get('bold', False)
                if actual_bold != expected_bold:
                    issues.append(Issue(
                        severity=Severity.ERROR,
                        message="表题加粗设置不符合要求",
                        location=location,
                        actual_value="已加粗" if actual_bold else "未加粗",
                        expected_value="已加粗" if expected_bold else "未加粗",
                        suggestion=f"请{'设置' if expected_bold else '取消'}表题加粗",
                        paragraph_index=table.caption_paragraph.index
                    ))

        return issues

    def _check_three_line_table(self, tables: List[TableInfo], config: Dict) -> List[Issue]:
        issues = []

        no_vertical = config.get('no_vertical_borders', True)
        no_left_right = config.get('no_left_right_borders', True)

        for table in tables:
            location = f"表格: {table.table_id or f'表格{table.index + 1}'}"
            borders = table.has_borders

            if no_vertical and borders.get('inside_v', False):
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message="三线表不应有垂直边线",
                    location=location,
                    actual_value="存在垂直边线",
                    expected_value="无垂直边线",
                    suggestion="请删除表格的垂直边线，使用三线表格式"
                ))

            if no_left_right and (borders.get('left', False) or borders.get('right', False)):
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message="三线表不应有左右边线",
                    location=location,
                    actual_value="存在左右边线",
                    expected_value="无左右边线",
                    suggestion="请删除表格的左右边线，使用三线表格式"
                ))

        return issues

    def _check_table_numbering(self, tables: List[TableInfo]) -> List[Issue]:
        issues = []

        chapter_tables = {}
        for table in tables:
            if table.chapter_number is not None:
                if table.chapter_number not in chapter_tables:
                    chapter_tables[table.chapter_number] = []
                chapter_tables[table.chapter_number].append(table)

        for chapter_num, tbls in chapter_tables.items():
            expected_nums = list(range(1, len(tbls) + 1))
            actual_nums = [t.table_number for t in tbls if t.table_number is not None]

            if sorted(actual_nums) != expected_nums[:len(actual_nums)]:
                issues.append(Issue(
                    severity=Severity.WARNING,
                    message=f"第{chapter_num}章表格编号不连续",
                    actual_value=f"现有编号: {sorted(actual_nums)}",
                    expected_value=f"应为: 1到{len(tbls)}",
                    suggestion="请检查表格编号是否连续"
                ))

        return issues

    def _check_formula_format(self, formula: FormulaInfo, config: Dict) -> List[Issue]:
        issues = []

        if not config:
            return issues

        location = f"公式: {formula.formula_id}"

        if config.get('formula_centered', True) and formula.formula_paragraph:
            para_format = formula.formula_paragraph.format
            if para_format and para_format.alignment != 'center':
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message="公式应居中对齐",
                    location=location,
                    actual_value=para_format.alignment or "未设置",
                    expected_value="center",
                    suggestion="请将公式设置为居中对齐",
                    paragraph_index=formula.formula_paragraph.index
                ))

        if config.get('use_equation_editor', True) and not formula.is_omath:
            issues.append(Issue(
                severity=Severity.WARNING,
                message="建议使用公式编辑器输入公式",
                location=location,
                actual_value="普通文本",
                expected_value="公式编辑器",
                suggestion="建议使用Word公式编辑器输入公式以获得更好的显示效果",
                paragraph_index=formula.position_paragraph_index
            ))

        return issues

    def _check_formula_numbering(self, formulas: List[FormulaInfo]) -> List[Issue]:
        issues = []

        chapter_formulas = {}
        for formula in formulas:
            if formula.chapter_number is not None:
                if formula.chapter_number not in chapter_formulas:
                    chapter_formulas[formula.chapter_number] = []
                chapter_formulas[formula.chapter_number].append(formula)

        for chapter_num, forms in chapter_formulas.items():
            expected_nums = list(range(1, len(forms) + 1))
            actual_nums = [f.formula_number for f in forms if f.formula_number is not None]

            if sorted(actual_nums) != expected_nums[:len(actual_nums)]:
                issues.append(Issue(
                    severity=Severity.WARNING,
                    message=f"第{chapter_num}章公式编号不连续",
                    actual_value=f"现有编号: {sorted(actual_nums)}",
                    expected_value=f"应为: 1到{len(forms)}",
                    suggestion="请检查公式编号是否连续"
                ))

        return issues

    def _check_citations(self, citations: List[CitationInfo], config: Dict) -> List[Issue]:
        issues = []

        if not citations:
            issues.append(Issue(
                severity=Severity.WARNING,
                message="正文中未发现引用标注",
                suggestion="请在正文中添加参考文献引用标注"
            ))
            return issues

        expected_superscript = config.get('superscript', True)
        if expected_superscript:
            non_superscript = [c for c in citations if not c.is_superscript]
            if non_superscript:
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message=f"发现{len(non_superscript)}处引用标注未设置为上角标",
                    actual_value="普通文本格式",
                    expected_value="上角标格式",
                    suggestion="请将引用标注设置为上角标格式"
                ))

        return issues

    def _check_cross_references(self, document_info: DocumentInfo) -> List[Issue]:
        issues = []

        if not document_info.references:
            return issues

        cited_numbers = set(c.citation_number for c in document_info.citations)
        ref_count = len(document_info.references)

        for i in range(1, ref_count + 1):
            if i not in cited_numbers:
                issues.append(Issue(
                    severity=Severity.WARNING,
                    message=f"参考文献[{i}]未在正文中引用",
                    suggestion="请确保所有参考文献都在正文中被引用，或删除未引用的参考文献"
                ))

        for num in cited_numbers:
            if num < 1 or num > ref_count:
                issues.append(Issue(
                    severity=Severity.ERROR,
                    message=f"引用标注[{num}]超出参考文献范围",
                    actual_value=f"引用编号: {num}",
                    expected_value=f"应为: 1到{ref_count}",
                    suggestion=f"请检查引用编号是否正确，当前共有{ref_count}条参考文献"
                ))

        return issues

    def _check_reference_format(self, ref: ReferenceInfo, config: Dict) -> List[Issue]:
        issues = []

        if not config:
            return issues

        location = f"参考文献[{ref.index + 1}]"

        document_types = config.get('document_types', [])
        if document_types and ref.reference_type:
            if ref.reference_type not in ['J', 'M', 'D', 'P', 'S', 'R', 'N', 'EB/OL']:
                issues.append(Issue(
                    severity=Severity.WARNING,
                    message="参考文献类型标识可能不正确",
                    location=location,
                    actual_value=f"[{ref.reference_type}]",
                    expected_value=f"应为: {', '.join(document_types)}",
                    suggestion="请检查文献类型标识是否正确",
                    paragraph_index=ref.paragraph.index if ref.paragraph else None
                ))

        if not ref.year:
            issues.append(Issue(
                severity=Severity.INFO,
                message="参考文献缺少年份信息",
                location=location,
                suggestion="建议补充文献发表年份",
                paragraph_index=ref.paragraph.index if ref.paragraph else None
            ))

        return issues
