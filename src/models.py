from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from enum import Enum


class HeadingLevel(Enum):
    LEVEL_1 = 1
    LEVEL_2 = 2
    LEVEL_3 = 3
    NONE = 0


class ParagraphType(Enum):
    HEADING_1 = "heading_1"
    HEADING_2 = "heading_2"
    HEADING_3 = "heading_3"
    BODY_TEXT = "body_text"
    FIGURE_CAPTION = "figure_caption"
    TABLE_CAPTION = "table_caption"
    FORMULA = "formula"
    REFERENCE = "reference"
    TOC_ENTRY = "toc_entry"
    UNKNOWN = "unknown"


@dataclass
class FontInfo:
    name: str
    name_ascii: Optional[str] = None
    name_east_asia: Optional[str] = None
    size: Optional[float] = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Optional[str] = None


@dataclass
class ParagraphFormat:
    alignment: Optional[str] = None
    first_line_indent: Optional[float] = None
    left_indent: Optional[float] = None
    right_indent: Optional[float] = None
    hanging_indent: Optional[float] = None
    line_spacing: Optional[float] = None
    line_spacing_rule: Optional[str] = None
    space_before: Optional[float] = None
    space_after: Optional[float] = None
    outline_level: Optional[int] = None


@dataclass
class RunInfo:
    text: str
    font: FontInfo
    is_formula: bool = False
    is_superscript: bool = False
    is_subscript: bool = False


@dataclass
class ParagraphInfo:
    index: int
    text: str
    paragraph_type: ParagraphType
    runs: List[RunInfo] = field(default_factory=list)
    format: Optional[ParagraphFormat] = None
    style_name: Optional[str] = None
    heading_level: HeadingLevel = HeadingLevel.NONE
    heading_number: Optional[str] = None
    heading_title: Optional[str] = None
    chapter_number: Optional[int] = None
    section_number: Optional[str] = None


@dataclass
class PageInfo:
    page_settings: 'PageSettings'
    sections: List['SectionInfo'] = field(default_factory=list)


@dataclass
class PageSettings:
    top_margin: float = 0.0
    bottom_margin: float = 0.0
    left_margin: float = 0.0
    right_margin: float = 0.0
    gutter: float = 0.0
    gutter_position: str = "left"
    header_distance: float = 0.0
    footer_distance: float = 0.0
    page_width: float = 0.0
    page_height: float = 0.0
    orientation: str = "portrait"


@dataclass
class SectionInfo:
    section_index: int
    start_page: int
    page_settings: PageSettings
    has_different_first_page: bool = False
    has_different_odd_even: bool = False


@dataclass
class TOCEntry:
    level: int
    text: str
    page_number: Optional[int] = None
    heading_number: Optional[str] = None
    heading_title: Optional[str] = None


@dataclass
class TOCInfo:
    entries: List[TOCEntry] = field(default_factory=list)
    is_auto_generated: bool = False
    title: str = "目录"


@dataclass
class FigureInfo:
    index: int
    figure_id: Optional[str] = None
    caption: Optional[str] = None
    caption_paragraph: Optional[ParagraphInfo] = None
    image_data: Optional[bytes] = None
    image_width: Optional[float] = None
    image_height: Optional[float] = None
    position_paragraph_index: Optional[int] = None
    chapter_number: Optional[int] = None
    figure_number: Optional[int] = None


@dataclass
class TableInfo:
    index: int
    table_id: Optional[str] = None
    caption: Optional[str] = None
    caption_paragraph: Optional[ParagraphInfo] = None
    rows: int = 0
    columns: int = 0
    is_three_line_table: bool = False
    has_borders: Dict[str, bool] = field(default_factory=dict)
    cell_contents: List[List[str]] = field(default_factory=list)
    position_paragraph_index: Optional[int] = None
    chapter_number: Optional[int] = None
    table_number: Optional[int] = None


@dataclass
class FormulaInfo:
    index: int
    formula_id: Optional[str] = None
    formula_text: Optional[str] = None
    formula_paragraph: Optional[ParagraphInfo] = None
    is_omath: bool = False
    position_paragraph_index: Optional[int] = None
    chapter_number: Optional[int] = None
    formula_number: Optional[int] = None


@dataclass
class ReferenceInfo:
    index: int
    reference_id: Optional[str] = None
    full_text: str = ""
    reference_type: Optional[str] = None
    authors: Optional[str] = None
    title: Optional[str] = None
    source: Optional[str] = None
    year: Optional[str] = None
    volume: Optional[str] = None
    issue: Optional[str] = None
    pages: Optional[str] = None
    doi: Optional[str] = None
    url: Optional[str] = None
    paragraph: Optional[ParagraphInfo] = None


@dataclass
class CitationInfo:
    citation_number: int
    citation_text: str
    paragraph_index: int
    position_in_paragraph: int
    is_superscript: bool = False


@dataclass
class ChapterInfo:
    chapter_number: int
    chapter_title: str
    heading_paragraph: ParagraphInfo
    start_paragraph_index: int
    end_paragraph_index: Optional[int] = None
    sections: List['SubSectionInfo'] = field(default_factory=list)
    figures: List[FigureInfo] = field(default_factory=list)
    tables: List[TableInfo] = field(default_factory=list)
    formulas: List[FormulaInfo] = field(default_factory=list)


@dataclass
class SubSectionInfo:
    section_number: str
    section_title: str
    heading_paragraph: ParagraphInfo
    level: int
    start_paragraph_index: int
    end_paragraph_index: Optional[int] = None
    sub_sections: List['SubSectionInfo'] = field(default_factory=list)


@dataclass
class DocumentInfo:
    file_path: str
    file_name: str
    page_count: int = 0
    paragraph_count: int = 0
    section_count: int = 0
    character_count: int = 0
    word_count: int = 0
    has_toc: bool = False
    toc_info: Optional[TOCInfo] = None
    page_settings: List[PageSettings] = field(default_factory=list)
    sections: List[SectionInfo] = field(default_factory=list)
    chapters: List[ChapterInfo] = field(default_factory=list)
    all_paragraphs: List[ParagraphInfo] = field(default_factory=list)
    figures: List[FigureInfo] = field(default_factory=list)
    tables: List[TableInfo] = field(default_factory=list)
    formulas: List[FormulaInfo] = field(default_factory=list)
    references: List[ReferenceInfo] = field(default_factory=list)
    citations: List[CitationInfo] = field(default_factory=list)
    reference_section_start: Optional[int] = None
    reference_section_end: Optional[int] = None


@dataclass
class ParseResult:
    success: bool
    document_info: Optional[DocumentInfo] = None
    error_message: Optional[str] = None
    warnings: List[str] = field(default_factory=list)
