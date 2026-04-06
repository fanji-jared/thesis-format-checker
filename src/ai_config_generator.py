import json
import re
import time
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import requests
from requests.exceptions import ConnectionError, Timeout, RequestException

from .config_manager import ConfigManager, ConfigValidationError


class GeneratorState(Enum):
    IDLE = "idle"
    CONNECTING = "connecting"
    PARSING = "parsing"
    GENERATING = "generating"
    VALIDATING = "validating"
    SAVING = "saving"
    COMPLETED = "completed"
    FAILED = "failed"


@dataclass
class GeneratorProgress:
    state: GeneratorState
    progress: float
    message: str
    details: Optional[Dict[str, Any]] = None


@dataclass
class OllamaConfig:
    base_url: str = "http://localhost:11434"
    model: str = "qwen2.5:7b"
    timeout: int = 300
    max_retries: int = 3
    retry_delay: float = 2.0
    temperature: float = 0.1
    top_p: float = 0.9


@dataclass
class GenerationResult:
    success: bool
    config_data: Optional[Dict[str, Any]] = None
    config_path: Optional[str] = None
    error_message: Optional[str] = None
    raw_response: Optional[str] = None
    validation_errors: List[Dict[str, Any]] = field(default_factory=list)


class AIConfigGeneratorError(Exception):
    def __init__(self, message: str, error_type: str = "unknown", details: Optional[Dict] = None):
        super().__init__(message)
        self.error_type = error_type
        self.details = details or {}


class AIConfigGenerator:
    DEFAULT_SCHEMA_PATH = "config/config_schema.json"
    DEFAULT_CONFIG_DIR = "config"
    
    PROMPT_TEMPLATE = """你是一个专业的论文格式配置生成助手。你的任务是根据用户提供的论文格式规则文档，生成符合指定JSON Schema的配置文件。

## 输入信息

### 论文格式规则文档：
{rules_content}

### JSON Schema定义：
{schema_content}

## 输出要求

1. 你必须严格按照JSON Schema的定义生成配置文件
2. 所有必填字段（required字段）都必须包含
3. 数值类型必须符合范围限制（minimum/maximum）
4. 字符串枚举值必须在允许的范围内
5. 输出必须是纯JSON格式，不要包含任何Markdown代码块标记或其他说明文字
6. 配置文件应该准确反映规则文档中描述的所有格式要求

## 字段映射指南

以下是规则文档中常见内容与配置字段的对应关系：

### 页面设置 (page_settings)
- 页边距：上/下/左/右边距 -> margins.top/bottom/left/right (单位: cm)
- 装订线 -> binding.position 和 binding.width
- 页眉页脚边距 -> header_footer.header_margin/footer_margin

### 目录设置 (toc_settings)
- 目录标题格式 -> title
- 目录各级标题格式 -> level1/level2/level3
- 显示级别 -> display_levels

### 章节标题 (chapter_titles)
- 一级标题(章标题) -> level1
- 二级标题(节标题) -> level2
- 三级标题(小节标题) -> level3
- 大纲级别 -> outline_level
- 是否另起一页 -> page_break_before

### 正文格式 (body_text)
- 字体字号 -> font, size, english_font
- 对齐方式 -> alignment
- 行距 -> line_spacing
- 首行缩进 -> first_line_indent

### 图片格式 (figures)
- 编号格式 -> numbering.format
- 图题位置 -> caption.position ("above"或"below")
- 图题字体字号 -> caption.font, caption.size

### 表格格式 (tables)
- 编号格式 -> numbering.format
- 表题位置 -> caption.position
- 三线表设置 -> three_line_table
- 续表格式 -> continuation

### 公式格式 (formulas)
- 编号格式 -> numbering.format
- 对齐方式 -> alignment, formula_centered, number_right_aligned

### 参考文献 (references)
- 文中引用格式 -> in_text_citation
- 文末列表格式 -> list_format
- 著录标准 -> list_format.standard

### 附录格式 (appendix)
- 字体字号 -> font, size
- 对齐方式 -> alignment

### 检测阈值 (thresholds)
- 重复率阈值 -> plagiarism_rate
- AIGC检测阈值 -> aigc_rate
- 格式得分阈值 -> format_score

## 注意事项

1. 字号转换参考：
   - 初号 = 42pt, 小初 = 36pt
   - 一号 = 26pt, 小一 = 24pt
   - 二号 = 22pt, 小二 = 18pt
   - 三号 = 16pt, 小三 = 15pt
   - 四号 = 14pt, 小四 = 12pt
   - 五号 = 10.5pt, 小五 = 9pt

2. 对齐方式枚举值：left, center, right, justify

3. 布尔值使用 true/false

4. 如果规则文档中某些信息缺失，使用合理的默认值

请直接输出JSON配置文件，不要有任何其他内容："""

    RULE_SECTION_PROMPT = """请根据以下论文格式规则片段，提取相关的配置信息。

## 规则片段：
{section_content}

## 需要提取的配置项：
{config_items}

请以JSON格式输出提取的配置信息，只输出JSON，不要有其他内容："""

    def __init__(
        self,
        ollama_config: Optional[OllamaConfig] = None,
        config_manager: Optional[ConfigManager] = None,
        progress_callback: Optional[Callable[[GeneratorProgress], None]] = None
    ):
        self.ollama_config = ollama_config or OllamaConfig()
        self.config_manager = config_manager
        self.progress_callback = progress_callback
        self._state = GeneratorState.IDLE
        self._schema_cache: Optional[Dict] = None
        
        if self.config_manager is None:
            project_root = Path(__file__).parent.parent
            self.config_manager = ConfigManager(str(project_root / self.DEFAULT_CONFIG_DIR))
    
    @property
    def state(self) -> GeneratorState:
        return self._state
    
    def _update_progress(self, state: GeneratorState, progress: float, message: str, details: Optional[Dict] = None):
        self._state = state
        if self.progress_callback:
            progress_info = GeneratorProgress(
                state=state,
                progress=progress,
                message=message,
                details=details
            )
            self.progress_callback(progress_info)
    
    def _load_schema(self) -> Dict[str, Any]:
        if self._schema_cache is None:
            self._schema_cache = self.config_manager.schema
        return self._schema_cache
    
    def check_ollama_connection(self) -> Tuple[bool, str]:
        try:
            response = requests.get(
                f"{self.ollama_config.base_url}/api/tags",
                timeout=10
            )
            if response.status_code == 200:
                models = response.json().get("models", [])
                model_names = [m.get("name", "") for m in models]
                if self.ollama_config.model in model_names:
                    return True, f"Ollama连接成功，模型 {self.ollama_config.model} 可用"
                else:
                    return False, f"模型 {self.ollama_config.model} 不可用，可用模型: {', '.join(model_names)}"
            else:
                return False, f"Ollama服务响应异常: HTTP {response.status_code}"
        except ConnectionError:
            return False, "无法连接到Ollama服务，请确保Ollama正在运行"
        except Timeout:
            return False, "连接Ollama服务超时"
        except Exception as e:
            return False, f"连接检查失败: {str(e)}"
    
    def _call_ollama_api(
        self,
        prompt: str,
        stream: bool = False
    ) -> Tuple[bool, str]:
        url = f"{self.ollama_config.base_url}/api/generate"
        payload = {
            "model": self.ollama_config.model,
            "prompt": prompt,
            "stream": stream,
            "options": {
                "temperature": self.ollama_config.temperature,
                "top_p": self.ollama_config.top_p
            }
        }
        
        last_error = None
        for attempt in range(self.ollama_config.max_retries):
            try:
                response = requests.post(
                    url,
                    json=payload,
                    timeout=self.ollama_config.timeout
                )
                
                if response.status_code == 200:
                    if stream:
                        full_response = ""
                        for line in response.iter_lines():
                            if line:
                                data = json.loads(line)
                                full_response += data.get("response", "")
                                if data.get("done", False):
                                    break
                        return True, full_response
                    else:
                        result = response.json()
                        return True, result.get("response", "")
                else:
                    last_error = f"API返回错误: HTTP {response.status_code}"
                    if attempt < self.ollama_config.max_retries - 1:
                        time.sleep(self.ollama_config.retry_delay * (attempt + 1))
                        
            except ConnectionError as e:
                last_error = f"连接错误: {str(e)}"
                if attempt < self.ollama_config.max_retries - 1:
                    time.sleep(self.ollama_config.retry_delay * (attempt + 1))
                    
            except Timeout as e:
                last_error = f"请求超时: {str(e)}"
                if attempt < self.ollama_config.max_retries - 1:
                    time.sleep(self.ollama_config.retry_delay * (attempt + 1))
                    
            except RequestException as e:
                last_error = f"请求异常: {str(e)}"
                if attempt < self.ollama_config.max_retries - 1:
                    time.sleep(self.ollama_config.retry_delay * (attempt + 1))
                    
            except json.JSONDecodeError as e:
                last_error = f"响应解析错误: {str(e)}"
                break
        
        return False, last_error or "未知错误"
    
    def parse_markdown_rules(self, rules_path: str) -> Dict[str, str]:
        path = Path(rules_path)
        if not path.exists():
            raise AIConfigGeneratorError(
                f"规则文档不存在: {rules_path}",
                error_type="file_not_found"
            )
        
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            raise AIConfigGeneratorError(
                f"读取规则文档失败: {str(e)}",
                error_type="file_read_error"
            )
        
        sections = {}
        section_pattern = re.compile(
            r'^(#{1,3})\s+(.+?)(?:\n|$)(.*?)(?=^#{1,3}\s+|\Z)',
            re.MULTILINE | re.DOTALL
        )
        
        for match in section_pattern.finditer(content):
            level = len(match.group(1))
            title = match.group(2).strip()
            section_content = match.group(3).strip()
            
            section_key = f"level{level}_{re.sub(r'[^\w\u4e00-\u9fff]', '_', title)}"
            sections[section_key] = {
                'title': title,
                'level': level,
                'content': section_content
            }
        
        sections['_full_content'] = content
        sections['_meta'] = {
            'file_path': str(path),
            'total_length': len(content),
            'section_count': len([k for k in sections.keys() if not k.startswith('_')])
        }
        
        return sections
    
    def _extract_json_from_response(self, response: str) -> Optional[Dict[str, Any]]:
        json_patterns = [
            r'```json\s*([\s\S]*?)\s*```',
            r'```\s*([\s\S]*?)\s*```',
            r'(\{[\s\S]*\})',
        ]
        
        for pattern in json_patterns:
            matches = re.findall(pattern, response)
            for match in matches:
                try:
                    config = json.loads(match.strip())
                    if isinstance(config, dict):
                        return config
                except json.JSONDecodeError:
                    continue
        
        return None
    
    def _validate_and_fix_config(
        self,
        config: Dict[str, Any],
        schema: Dict[str, Any]
    ) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        errors = []
        fixed_config = config.copy()
        
        required_fields = schema.get('required', [])
        properties = schema.get('properties', {})
        
        for field in required_fields:
            if field not in fixed_config:
                if field == 'meta':
                    fixed_config['meta'] = {
                        'name': 'AI生成的配置',
                        'version': '1.0.0',
                        'description': '由AI根据规则文档自动生成',
                        'created_at': '',
                        'updated_at': '',
                        'author': 'AI生成'
                    }
                elif field == 'page_settings':
                    fixed_config['page_settings'] = self._get_default_page_settings()
                elif field == 'toc_settings':
                    fixed_config['toc_settings'] = self._get_default_toc_settings()
                elif field == 'chapter_titles':
                    fixed_config['chapter_titles'] = self._get_default_chapter_titles()
                elif field == 'body_text':
                    fixed_config['body_text'] = self._get_default_body_text()
                elif field == 'figures':
                    fixed_config['figures'] = self._get_default_figures()
                elif field == 'tables':
                    fixed_config['tables'] = self._get_default_tables()
                elif field == 'formulas':
                    fixed_config['formulas'] = self._get_default_formulas()
                elif field == 'references':
                    fixed_config['references'] = self._get_default_references()
                elif field == 'appendix':
                    fixed_config['appendix'] = self._get_default_appendix()
                else:
                    errors.append({
                        'type': 'missing_field',
                        'field': field,
                        'message': f'缺少必填字段: {field}'
                    })
        
        return fixed_config, errors
    
    def _get_default_page_settings(self) -> Dict[str, Any]:
        return {
            "margins": {"top": 2.0, "bottom": 1.5, "left": 2.0, "right": 2.0},
            "binding": {"position": "left", "width": 1.0},
            "header_footer": {"header_margin": 1.5, "footer_margin": 1.0},
            "font_rules": {"english_font": "Times New Roman"}
        }
    
    def _get_default_toc_settings(self) -> Dict[str, Any]:
        return {
            "title": {"text": "目录", "font": "黑体", "size": "三号", "alignment": "center"},
            "level1": {"font": "黑体", "size": "小四", "alignment": "justify", "line_spacing": 1.5, "space_before": 0.25, "space_after": 0.25},
            "level2": {"font": "宋体", "size": "小四", "alignment": "justify", "line_spacing": 1.5, "space_before": 0, "space_after": 0, "left_indent": 2},
            "level3": {"font": "宋体", "size": "小四", "alignment": "justify", "line_spacing": 1.5, "space_before": 0, "space_after": 0},
            "auto_generated": True,
            "display_levels": 3
        }
    
    def _get_default_chapter_titles(self) -> Dict[str, Any]:
        return {
            "level1": {
                "font": "黑体", "size": "三号", "alignment": "center", "bold": False,
                "line_spacing": 1.5, "space_before": 1, "space_after": 1,
                "outline_level": 1, "page_break_before": True,
                "numbering_format": "第X章", "space_between_number_title": 1
            },
            "level2": {
                "font": "黑体", "size": "四号", "alignment": "left", "bold": True,
                "line_spacing": 1.5, "space_before": 0.5, "space_after": 0.5,
                "outline_level": 2, "numbering_format": "X.X"
            },
            "level3": {
                "font": "黑体", "size": "小四", "alignment": "left", "bold": True,
                "line_spacing": 1.5, "space_before": 0.5, "space_after": 0.5,
                "outline_level": 3, "numbering_format": "X.X.X"
            }
        }
    
    def _get_default_body_text(self) -> Dict[str, Any]:
        return {
            "font": "宋体", "english_font": "Times New Roman", "size": "小四",
            "alignment": "justify", "line_spacing": 1.5,
            "space_before": 0, "space_after": 0, "first_line_indent": 2
        }
    
    def _get_default_figures(self) -> Dict[str, Any]:
        return {
            "numbering": {"format": "图X-X", "by_chapter": True},
            "caption": {
                "position": "below", "font": "宋体", "english_font": "Times New Roman",
                "size": "五号", "bold": True, "alignment": "center",
                "line_spacing": 1.5, "space_after_pt": 6
            },
            "alignment": "center", "line_spacing": 1.5, "space_before": 0, "space_after": 0
        }
    
    def _get_default_tables(self) -> Dict[str, Any]:
        return {
            "numbering": {"format": "表X-X", "by_chapter": True},
            "caption": {
                "position": "above", "font": "宋体", "english_font": "Times New Roman",
                "size": "五号", "bold": True, "alignment": "center",
                "line_spacing": 1.5, "space_before_pt": 6
            },
            "three_line_table": {"enabled": True, "no_vertical_borders": True, "no_left_right_borders": True},
            "content": {
                "font": "宋体", "english_font": "Times New Roman", "size": "五号",
                "alignment": "center", "line_spacing": 1.5, "header_bold": True
            },
            "continuation": {
                "format": "续表X-X", "font": "宋体", "size": "五号",
                "alignment": "right", "line_spacing_pt": 22,
                "space_before": 0, "space_after": 0
            }
        }
    
    def _get_default_formulas(self) -> Dict[str, Any]:
        return {
            "numbering": {"format": "(X-X)", "by_chapter": True},
            "alignment": "center", "formula_centered": True, "number_right_aligned": True,
            "use_equation_editor": True, "separate_line": True
        }
    
    def _get_default_references(self) -> Dict[str, Any]:
        return {
            "in_text_citation": {
                "format": "[X]", "superscript": True, "brackets": True,
                "position": "after_sentence"
            },
            "list_format": {
                "standard": "GB/T 7714", "font": "宋体", "english_font": "Times New Roman",
                "size": "五号", "alignment": "justify", "line_spacing": 1.5,
                "space_before": 0, "space_after": 0, "hanging_indent": 2,
                "document_types": ["[J]", "[M]", "[D]", "[P]", "[S]", "[R]", "[N]", "[EB/OL]"]
            },
            "cross_reference_required": True
        }
    
    def _get_default_appendix(self) -> Dict[str, Any]:
        return {
            "font": "宋体", "english_font": "Times New Roman", "size": "小四",
            "alignment": "left", "line_spacing": 1.5,
            "space_before": 0, "space_after": 0, "first_line_indent": 2
        }
    
    def generate_config(
        self,
        rules_path: str,
        config_name: Optional[str] = None,
        description: Optional[str] = None,
        save_config: bool = True,
        overwrite: bool = False
    ) -> GenerationResult:
        self._update_progress(GeneratorState.CONNECTING, 0.0, "正在连接Ollama服务...")
        
        connected, message = self.check_ollama_connection()
        if not connected:
            self._update_progress(GeneratorState.FAILED, 0.0, message)
            return GenerationResult(
                success=False,
                error_message=message
            )
        
        self._update_progress(GeneratorState.PARSING, 0.1, "正在解析规则文档...")
        
        try:
            rules_sections = self.parse_markdown_rules(rules_path)
        except AIConfigGeneratorError as e:
            self._update_progress(GeneratorState.FAILED, 0.1, str(e))
            return GenerationResult(
                success=False,
                error_message=str(e)
            )
        
        self._update_progress(GeneratorState.GENERATING, 0.2, "正在加载配置Schema...")
        
        try:
            schema = self._load_schema()
        except Exception as e:
            self._update_progress(GeneratorState.FAILED, 0.2, f"加载Schema失败: {str(e)}")
            return GenerationResult(
                success=False,
                error_message=f"加载Schema失败: {str(e)}"
            )
        
        self._update_progress(GeneratorState.GENERATING, 0.3, "正在生成配置文件...")
        
        rules_content = rules_sections.get('_full_content', '')
        schema_str = json.dumps(schema, ensure_ascii=False, indent=2)
        
        prompt = self.PROMPT_TEMPLATE.format(
            rules_content=rules_content,
            schema_content=schema_str
        )
        
        success, response = self._call_ollama_api(prompt)
        
        if not success:
            self._update_progress(GeneratorState.FAILED, 0.3, f"AI生成失败: {response}")
            return GenerationResult(
                success=False,
                error_message=f"AI生成失败: {response}",
                raw_response=response
            )
        
        self._update_progress(GeneratorState.VALIDATING, 0.7, "正在解析AI响应...")
        
        config_data = self._extract_json_from_response(response)
        
        if config_data is None:
            self._update_progress(GeneratorState.FAILED, 0.7, "无法从AI响应中提取有效的JSON配置")
            return GenerationResult(
                success=False,
                error_message="无法从AI响应中提取有效的JSON配置",
                raw_response=response
            )
        
        self._update_progress(GeneratorState.VALIDATING, 0.8, "正在验证和修复配置...")
        
        config_data, fix_errors = self._validate_and_fix_config(config_data, schema)
        
        if config_name:
            if 'meta' not in config_data:
                config_data['meta'] = {}
            config_data['meta']['name'] = config_name
        if description:
            if 'meta' not in config_data:
                config_data['meta'] = {}
            config_data['meta']['description'] = description
        
        validation_errors = self.config_manager.validate_config(config_data)
        
        if validation_errors:
            self._update_progress(GeneratorState.FAILED, 0.8, f"配置验证失败，共{len(validation_errors)}个错误")
            return GenerationResult(
                success=False,
                error_message=f"配置验证失败，共{len(validation_errors)}个错误",
                raw_response=response,
                config_data=config_data,
                validation_errors=validation_errors
            )
        
        if not save_config:
            self._update_progress(GeneratorState.COMPLETED, 1.0, "配置生成完成")
            return GenerationResult(
                success=True,
                config_data=config_data,
                raw_response=response
            )
        
        self._update_progress(GeneratorState.SAVING, 0.9, "正在保存配置文件...")
        
        try:
            config_path = self.config_manager.save_config(
                config_data,
                overwrite=overwrite
            )
            self._update_progress(GeneratorState.COMPLETED, 1.0, "配置文件保存成功")
            return GenerationResult(
                success=True,
                config_data=config_data,
                config_path=config_path,
                raw_response=response
            )
        except FileExistsError as e:
            self._update_progress(GeneratorState.FAILED, 0.9, str(e))
            return GenerationResult(
                success=False,
                error_message=str(e),
                raw_response=response,
                config_data=config_data
            )
        except Exception as e:
            self._update_progress(GeneratorState.FAILED, 0.9, f"保存配置失败: {str(e)}")
            return GenerationResult(
                success=False,
                error_message=f"保存配置失败: {str(e)}",
                raw_response=response,
                config_data=config_data
            )
    
    def generate_config_incremental(
        self,
        rules_path: str,
        config_name: Optional[str] = None,
        description: Optional[str] = None,
        save_config: bool = True,
        overwrite: bool = False
    ) -> GenerationResult:
        self._update_progress(GeneratorState.CONNECTING, 0.0, "正在连接Ollama服务...")
        
        connected, message = self.check_ollama_connection()
        if not connected:
            self._update_progress(GeneratorState.FAILED, 0.0, message)
            return GenerationResult(
                success=False,
                error_message=message
            )
        
        self._update_progress(GeneratorState.PARSING, 0.1, "正在解析规则文档...")
        
        try:
            rules_sections = self.parse_markdown_rules(rules_path)
        except AIConfigGeneratorError as e:
            self._update_progress(GeneratorState.FAILED, 0.1, str(e))
            return GenerationResult(
                success=False,
                error_message=str(e)
            )
        
        self._update_progress(GeneratorState.GENERATING, 0.2, "正在加载配置Schema...")
        
        try:
            schema = self._load_schema()
        except Exception as e:
            self._update_progress(GeneratorState.FAILED, 0.2, f"加载Schema失败: {str(e)}")
            return GenerationResult(
                success=False,
                error_message=f"加载Schema失败: {str(e)}"
            )
        
        section_configs = {}
        total_sections = len([k for k in rules_sections.keys() if not k.startswith('_')])
        processed_sections = 0
        
        section_mapping = {
            '页面': 'page_settings',
            '目录': 'toc_settings',
            '章节': 'chapter_titles',
            '正文': 'body_text',
            '图片': 'figures',
            '表格': 'tables',
            '公式': 'formulas',
            '参考文献': 'references',
            '附录': 'appendix',
            '检测': 'thresholds'
        }
        
        for section_key, section_data in rules_sections.items():
            if section_key.startswith('_'):
                continue
            
            section_title = section_data.get('title', '')
            section_content = section_data.get('content', '')
            
            if not section_content.strip():
                continue
            
            matched_config_key = None
            for keyword, config_key in section_mapping.items():
                if keyword in section_title:
                    matched_config_key = config_key
                    break
            
            if matched_config_key is None:
                continue
            
            progress = 0.2 + (processed_sections / total_sections) * 0.5
            self._update_progress(
                GeneratorState.GENERATING,
                progress,
                f"正在处理章节: {section_title}"
            )
            
            config_items = self._get_config_items_description(matched_config_key, schema)
            
            prompt = self.RULE_SECTION_PROMPT.format(
                section_content=section_content,
                config_items=config_items
            )
            
            success, response = self._call_ollama_api(prompt)
            
            if success:
                section_config = self._extract_json_from_response(response)
                if section_config:
                    section_configs[matched_config_key] = section_config
            
            processed_sections += 1
        
        self._update_progress(GeneratorState.VALIDATING, 0.7, "正在合并配置...")
        
        config_data = {
            "meta": {
                "name": config_name or "AI生成的配置",
                "version": "1.0.0",
                "description": description or "由AI根据规则文档自动生成",
                "created_at": "",
                "updated_at": "",
                "author": "AI生成"
            }
        }
        
        for config_key in ['page_settings', 'toc_settings', 'chapter_titles', 'body_text',
                          'figures', 'tables', 'formulas', 'references', 'appendix', 'thresholds']:
            if config_key in section_configs:
                config_data[config_key] = section_configs[config_key]
            else:
                default_method = getattr(self, f'_get_default_{config_key}', None)
                if default_method:
                    config_data[config_key] = default_method()
        
        self._update_progress(GeneratorState.VALIDATING, 0.8, "正在验证配置...")
        
        validation_errors = self.config_manager.validate_config(config_data)
        
        if validation_errors:
            config_data, fix_errors = self._validate_and_fix_config(config_data, schema)
            validation_errors = self.config_manager.validate_config(config_data)
            
            if validation_errors:
                self._update_progress(GeneratorState.FAILED, 0.8, f"配置验证失败，共{len(validation_errors)}个错误")
                return GenerationResult(
                    success=False,
                    error_message=f"配置验证失败，共{len(validation_errors)}个错误",
                    config_data=config_data,
                    validation_errors=validation_errors
                )
        
        if not save_config:
            self._update_progress(GeneratorState.COMPLETED, 1.0, "配置生成完成")
            return GenerationResult(
                success=True,
                config_data=config_data
            )
        
        self._update_progress(GeneratorState.SAVING, 0.9, "正在保存配置文件...")
        
        try:
            config_path = self.config_manager.save_config(
                config_data,
                overwrite=overwrite
            )
            self._update_progress(GeneratorState.COMPLETED, 1.0, "配置文件保存成功")
            return GenerationResult(
                success=True,
                config_data=config_data,
                config_path=config_path
            )
        except Exception as e:
            self._update_progress(GeneratorState.FAILED, 0.9, f"保存配置失败: {str(e)}")
            return GenerationResult(
                success=False,
                error_message=f"保存配置失败: {str(e)}",
                config_data=config_data
            )
    
    def _get_config_items_description(self, config_key: str, schema: Dict) -> str:
        properties = schema.get('properties', {})
        if config_key not in properties:
            return ""
        
        config_schema = properties[config_key]
        required = config_schema.get('required', [])
        props = config_schema.get('properties', {})
        
        descriptions = []
        for field in required:
            if field in props:
                field_info = props[field]
                desc = field_info.get('description', field)
                field_type = field_info.get('type', 'unknown')
                
                if field_type == 'object':
                    nested_required = field_info.get('required', [])
                    nested_props = field_info.get('properties', {})
                    nested_desc = []
                    for nf in nested_required:
                        if nf in nested_props:
                            nested_desc.append(f"  - {nf}: {nested_props[nf].get('description', nf)}")
                    descriptions.append(f"{field} ({desc}):\n" + "\n".join(nested_desc))
                elif field_type == 'string':
                    enum_values = field_info.get('enum', [])
                    if enum_values:
                        descriptions.append(f"{field}: {desc} (可选值: {', '.join(enum_values)})")
                    else:
                        descriptions.append(f"{field}: {desc}")
                elif field_type == 'number':
                    minimum = field_info.get('minimum')
                    maximum = field_info.get('maximum')
                    if minimum is not None and maximum is not None:
                        descriptions.append(f"{field}: {desc} (范围: {minimum}-{maximum})")
                    else:
                        descriptions.append(f"{field}: {desc}")
                elif field_type == 'boolean':
                    descriptions.append(f"{field}: {desc} (true/false)")
                elif field_type == 'integer':
                    minimum = field_info.get('minimum')
                    maximum = field_info.get('maximum')
                    if minimum is not None and maximum is not None:
                        descriptions.append(f"{field}: {desc} (范围: {minimum}-{maximum})")
                    else:
                        descriptions.append(f"{field}: {desc}")
                else:
                    descriptions.append(f"{field}: {desc}")
        
        return "\n".join(descriptions)
    
    def get_available_models(self) -> Tuple[bool, List[str], str]:
        try:
            response = requests.get(
                f"{self.ollama_config.base_url}/api/tags",
                timeout=10
            )
            if response.status_code == 200:
                models = response.json().get("models", [])
                model_names = [m.get("name", "") for m in models]
                return True, model_names, f"找到 {len(model_names)} 个可用模型"
            else:
                return False, [], f"获取模型列表失败: HTTP {response.status_code}"
        except Exception as e:
            return False, [], f"获取模型列表失败: {str(e)}"
    
    def set_model(self, model_name: str):
        self.ollama_config.model = model_name
    
    def set_base_url(self, base_url: str):
        self.ollama_config.base_url = base_url
    
    def set_timeout(self, timeout: int):
        self.ollama_config.timeout = timeout
    
    def set_retry_config(self, max_retries: int, retry_delay: float):
        self.ollama_config.max_retries = max_retries
        self.ollama_config.retry_delay = retry_delay
    
    def set_temperature(self, temperature: float):
        self.ollama_config.temperature = temperature
    
    def set_top_p(self, top_p: float):
        self.ollama_config.top_p = top_p
