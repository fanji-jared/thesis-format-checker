import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import jsonschema
from jsonschema import validate, ValidationError as JsonSchemaValidationError


class ConfigValidationError(Exception):
    """配置文件验证错误"""
    def __init__(self, message: str, errors: List[Dict[str, Any]] = None):
        super().__init__(message)
        self.errors = errors or []


class ConfigManager:
    """配置文件管理器"""
    
    CONFIG_FILE_PREFIX = "config_"
    CONFIG_FILE_EXTENSION = ".json"
    SCHEMA_FILE = "config_schema.json"
    
    def __init__(self, config_dir: str = None):
        if config_dir is None:
            config_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config")
        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self._schema: Optional[Dict] = None
        self._current_config: Optional[Dict] = None
        self._current_config_path: Optional[Path] = None
    
    @property
    def schema(self) -> Dict:
        """获取配置文件Schema"""
        if self._schema is None:
            schema_path = self.config_dir / self.SCHEMA_FILE
            if schema_path.exists():
                with open(schema_path, 'r', encoding='utf-8') as f:
                    self._schema = json.load(f)
            else:
                raise FileNotFoundError(f"Schema文件不存在: {schema_path}")
        return self._schema
    
    def get_config_list(self) -> List[Dict[str, Any]]:
        """获取所有配置文件列表"""
        configs = []
        for file_path in self.config_dir.glob(f"{self.CONFIG_FILE_PREFIX}*{self.CONFIG_FILE_EXTENSION}"):
            if file_path.name == self.SCHEMA_FILE:
                continue
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                meta = config_data.get('meta', {})
                configs.append({
                    'path': str(file_path),
                    'filename': file_path.name,
                    'name': meta.get('name', file_path.stem),
                    'description': meta.get('description', ''),
                    'version': meta.get('version', '1.0.0'),
                    'created_at': meta.get('created_at', ''),
                    'updated_at': meta.get('updated_at', ''),
                    'author': meta.get('author', '')
                })
            except (json.JSONDecodeError, KeyError) as e:
                configs.append({
                    'path': str(file_path),
                    'filename': file_path.name,
                    'name': file_path.stem,
                    'description': f'读取失败: {str(e)}',
                    'version': '',
                    'created_at': '',
                    'updated_at': '',
                    'author': '',
                    'error': str(e)
                })
        configs.sort(key=lambda x: x.get('updated_at', '') or x.get('created_at', ''), reverse=True)
        return configs
    
    def load_config(self, config_path: str) -> Dict[str, Any]:
        """加载配置文件"""
        path = Path(config_path)
        if not path.is_absolute():
            path = self.config_dir / path
        
        if not path.exists():
            raise FileNotFoundError(f"配置文件不存在: {path}")
        
        with open(path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        
        validation_errors = self.validate_config(config_data)
        if validation_errors:
            raise ConfigValidationError(
                f"配置文件验证失败，共{len(validation_errors)}个错误",
                validation_errors
            )
        
        self._current_config = config_data
        self._current_config_path = path
        return config_data
    
    def save_config(self, config_data: Dict[str, Any], filename: str = None, overwrite: bool = False) -> str:
        """保存配置文件"""
        validation_errors = self.validate_config(config_data)
        if validation_errors:
            raise ConfigValidationError(
                f"配置数据验证失败，共{len(validation_errors)}个错误",
                validation_errors
            )
        
        if filename is None:
            if 'meta' in config_data and 'name' in config_data['meta']:
                name = config_data['meta']['name']
                safe_name = "".join(c for c in name if c.isalnum() or c in ('-', '_'))
                filename = f"{self.CONFIG_FILE_PREFIX}{safe_name}{self.CONFIG_FILE_EXTENSION}"
            else:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"{self.CONFIG_FILE_PREFIX}{timestamp}{self.CONFIG_FILE_EXTENSION}"
        
        if not filename.startswith(self.CONFIG_FILE_PREFIX):
            filename = f"{self.CONFIG_FILE_PREFIX}{filename}"
        if not filename.endswith(self.CONFIG_FILE_EXTENSION):
            filename = f"{filename}{self.CONFIG_FILE_EXTENSION}"
        
        file_path = self.config_dir / filename
        
        if file_path.exists() and not overwrite:
            raise FileExistsError(f"配置文件已存在: {file_path}")
        
        if 'meta' not in config_data:
            config_data['meta'] = {}
        
        now = datetime.now().isoformat()
        if 'created_at' not in config_data['meta'] or not config_data['meta']['created_at']:
            config_data['meta']['created_at'] = now
        config_data['meta']['updated_at'] = now
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, ensure_ascii=False, indent=2)
        
        self._current_config = config_data
        self._current_config_path = file_path
        
        return str(file_path)
    
    def delete_config(self, config_path: str) -> bool:
        """删除配置文件"""
        path = Path(config_path)
        if not path.is_absolute():
            path = self.config_dir / path
        
        if not path.exists():
            raise FileNotFoundError(f"配置文件不存在: {path}")
        
        if path.name == self.SCHEMA_FILE:
            raise ValueError("不能删除Schema文件")
        
        path.unlink()
        
        if self._current_config_path == path:
            self._current_config = None
            self._current_config_path = None
        
        return True
    
    def validate_config(self, config_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """验证配置文件"""
        errors = []
        
        try:
            validate(instance=config_data, schema=self.schema)
        except JsonSchemaValidationError as e:
            errors.append({
                'type': 'schema_validation',
                'path': list(e.absolute_path) if e.absolute_path else [],
                'message': e.message,
                'validator': e.validator,
                'value': e.instance
            })
            return errors
        
        errors.extend(self._validate_required_fields(config_data))
        errors.extend(self._validate_data_types(config_data))
        errors.extend(self._validate_value_ranges(config_data))
        errors.extend(self._validate_consistency(config_data))
        
        return errors
    
    def _validate_required_fields(self, config_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """验证必填字段"""
        errors = []
        
        required_top_level = ['meta', 'page_settings', 'toc_settings', 'chapter_titles', 
                            'body_text', 'figures', 'tables', 'formulas', 'references', 'appendix']
        
        for field in required_top_level:
            if field not in config_data:
                errors.append({
                    'type': 'missing_field',
                    'path': [field],
                    'message': f"缺少必填字段: {field}"
                })
        
        if 'meta' in config_data:
            meta = config_data['meta']
            required_meta = ['name', 'version', 'description']
            for field in required_meta:
                if field not in meta:
                    errors.append({
                        'type': 'missing_field',
                        'path': ['meta', field],
                        'message': f"meta中缺少必填字段: {field}"
                    })
        
        return errors
    
    def _validate_data_types(self, config_data: Dict[str, Any], path: List[str] = None) -> List[Dict[str, Any]]:
        """验证数据类型"""
        errors = []
        if path is None:
            path = []
        
        def check_type(value, expected_type, current_path):
            if expected_type == 'number' and not isinstance(value, (int, float)):
                return {
                    'type': 'type_error',
                    'path': current_path,
                    'message': f"期望类型: {expected_type}, 实际类型: {type(value).__name__}",
                    'expected': expected_type,
                    'actual': type(value).__name__
                }
            elif expected_type == 'string' and not isinstance(value, str):
                return {
                    'type': 'type_error',
                    'path': current_path,
                    'message': f"期望类型: {expected_type}, 实际类型: {type(value).__name__}",
                    'expected': expected_type,
                    'actual': type(value).__name__
                }
            elif expected_type == 'boolean' and not isinstance(value, bool):
                return {
                    'type': 'type_error',
                    'path': current_path,
                    'message': f"期望类型: {expected_type}, 实际类型: {type(value).__name__}",
                    'expected': expected_type,
                    'actual': type(value).__name__
                }
            elif expected_type == 'integer' and not isinstance(value, int):
                return {
                    'type': 'type_error',
                    'path': current_path,
                    'message': f"期望类型: {expected_type}, 实际类型: {type(value).__name__}",
                    'expected': expected_type,
                    'actual': type(value).__name__
                }
            return None
        
        type_checks = [
            (['page_settings', 'margins', 'top'], 'number'),
            (['page_settings', 'margins', 'bottom'], 'number'),
            (['page_settings', 'margins', 'left'], 'number'),
            (['page_settings', 'margins', 'right'], 'number'),
            (['page_settings', 'binding', 'width'], 'number'),
            (['page_settings', 'header_footer', 'header_margin'], 'number'),
            (['page_settings', 'header_footer', 'footer_margin'], 'number'),
            (['toc_settings', 'display_levels'], 'integer'),
            (['toc_settings', 'auto_generated'], 'boolean'),
            (['chapter_titles', 'level1', 'outline_level'], 'integer'),
            (['chapter_titles', 'level1', 'bold'], 'boolean'),
            (['chapter_titles', 'level1', 'page_break_before'], 'boolean'),
            (['chapter_titles', 'level2', 'outline_level'], 'integer'),
            (['chapter_titles', 'level2', 'bold'], 'boolean'),
            (['chapter_titles', 'level3', 'outline_level'], 'integer'),
            (['chapter_titles', 'level3', 'bold'], 'boolean'),
            (['body_text', 'first_line_indent'], 'number'),
            (['body_text', 'line_spacing'], 'number'),
            (['figures', 'caption', 'bold'], 'boolean'),
            (['tables', 'three_line_table', 'enabled'], 'boolean'),
            (['tables', 'content', 'header_bold'], 'boolean'),
            (['formulas', 'formula_centered'], 'boolean'),
            (['formulas', 'number_right_aligned'], 'boolean'),
            (['references', 'in_text_citation', 'superscript'], 'boolean'),
            (['references', 'in_text_citation', 'brackets'], 'boolean'),
            (['references', 'cross_reference_required'], 'boolean'),
        ]
        
        for check_path, expected_type in type_checks:
            try:
                value = config_data
                for key in check_path:
                    value = value[key]
                error = check_type(value, expected_type, check_path)
                if error:
                    errors.append(error)
            except (KeyError, TypeError):
                pass
        
        return errors
    
    def _validate_value_ranges(self, config_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """验证数值范围"""
        errors = []
        
        range_checks = [
            (['page_settings', 'margins', 'top'], 0, 10, 'cm'),
            (['page_settings', 'margins', 'bottom'], 0, 10, 'cm'),
            (['page_settings', 'margins', 'left'], 0, 10, 'cm'),
            (['page_settings', 'margins', 'right'], 0, 10, 'cm'),
            (['page_settings', 'binding', 'width'], 0, 5, 'cm'),
            (['toc_settings', 'display_levels'], 1, 9, '级'),
            (['chapter_titles', 'level1', 'outline_level'], 1, 9, '级'),
            (['chapter_titles', 'level2', 'outline_level'], 1, 9, '级'),
            (['chapter_titles', 'level3', 'outline_level'], 1, 9, '级'),
            (['body_text', 'line_spacing'], 0.5, 3.0, '倍'),
            (['body_text', 'first_line_indent'], 0, 10, '字符'),
        ]
        
        for check_path, min_val, max_val, unit in range_checks:
            try:
                value = config_data
                for key in check_path:
                    value = value[key]
                if isinstance(value, (int, float)):
                    if value < min_val or value > max_val:
                        errors.append({
                            'type': 'range_error',
                            'path': check_path,
                            'message': f"值 {value} 超出有效范围 [{min_val}, {max_val}] {unit}",
                            'value': value,
                            'min': min_val,
                            'max': max_val
                        })
            except (KeyError, TypeError):
                pass
        
        if 'thresholds' in config_data:
            thresholds = config_data['thresholds']
            for field in ['plagiarism_rate', 'aigc_rate']:
                if field in thresholds:
                    value = thresholds[field]
                    if isinstance(value, (int, float)):
                        if value < 0 or value > 100:
                            errors.append({
                                'type': 'range_error',
                                'path': ['thresholds', field],
                                'message': f"阈值 {value} 必须在 0-100 之间",
                                'value': value,
                                'min': 0,
                                'max': 100
                            })
        
        return errors
    
    def _validate_consistency(self, config_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """验证一致性"""
        errors = []
        
        if 'chapter_titles' in config_data:
            levels = config_data['chapter_titles']
            if 'level1' in levels and 'level2' in levels:
                if levels['level1'].get('outline_level', 1) >= levels['level2'].get('outline_level', 2):
                    errors.append({
                        'type': 'consistency_error',
                        'path': ['chapter_titles'],
                        'message': "一级标题的大纲级别应该小于二级标题"
                    })
            
            if 'level2' in levels and 'level3' in levels:
                if levels['level2'].get('outline_level', 2) >= levels['level3'].get('outline_level', 3):
                    errors.append({
                        'type': 'consistency_error',
                        'path': ['chapter_titles'],
                        'message': "二级标题的大纲级别应该小于三级标题"
                    })
        
        return errors
    
    def get_current_config(self) -> Optional[Dict[str, Any]]:
        """获取当前加载的配置"""
        return self._current_config
    
    def get_current_config_path(self) -> Optional[str]:
        """获取当前配置文件路径"""
        return str(self._current_config_path) if self._current_config_path else None
    
    def create_empty_config(self, name: str, description: str = "") -> Dict[str, Any]:
        """创建空配置模板"""
        return {
            "meta": {
                "name": name,
                "version": "1.0.0",
                "description": description,
                "created_at": datetime.now().isoformat(),
                "updated_at": datetime.now().isoformat(),
                "author": ""
            },
            "page_settings": {
                "margins": {"top": 2.0, "bottom": 1.5, "left": 2.0, "right": 2.0},
                "binding": {"position": "left", "width": 1.0},
                "header_footer": {"header_margin": 1.5, "footer_margin": 1.0},
                "font_rules": {"english_font": "Times New Roman"}
            },
            "toc_settings": {
                "title": {"text": "目录", "font": "黑体", "size": "三号", "alignment": "center"},
                "level1": {"font": "黑体", "size": "小四", "alignment": "justify", "line_spacing": 1.5, "space_before": 0.25, "space_after": 0.25},
                "level2": {"font": "宋体", "size": "小四", "alignment": "justify", "line_spacing": 1.5, "space_before": 0, "space_after": 0, "left_indent": 2},
                "level3": {"font": "宋体", "size": "小四", "alignment": "justify", "line_spacing": 1.5, "space_before": 0, "space_after": 0},
                "auto_generated": True,
                "display_levels": 3
            },
            "chapter_titles": {
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
            },
            "body_text": {
                "font": "宋体", "english_font": "Times New Roman", "size": "小四",
                "alignment": "justify", "line_spacing": 1.5,
                "space_before": 0, "space_after": 0, "first_line_indent": 2
            },
            "figures": {
                "numbering": {"format": "图X-X", "by_chapter": True},
                "caption": {
                    "position": "below", "font": "宋体", "english_font": "Times New Roman",
                    "size": "五号", "bold": True, "alignment": "center",
                    "line_spacing": 1.5, "space_after_pt": 6
                },
                "alignment": "center", "line_spacing": 1.5, "space_before": 0, "space_after": 0
            },
            "tables": {
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
            },
            "formulas": {
                "numbering": {"format": "(X-X)", "by_chapter": True},
                "alignment": "center", "formula_centered": True, "number_right_aligned": True,
                "use_equation_editor": True, "separate_line": True
            },
            "references": {
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
            },
            "appendix": {
                "font": "宋体", "english_font": "Times New Roman", "size": "小四",
                "alignment": "left", "line_spacing": 1.5,
                "space_before": 0, "space_after": 0, "first_line_indent": 2
            },
            "thresholds": {
                "plagiarism_rate": 30,
                "aigc_rate": 40,
                "format_score": 45
            }
        }
    
    def export_config(self, config_data: Dict[str, Any], export_path: str) -> str:
        """导出配置文件到指定路径"""
        path = Path(export_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, ensure_ascii=False, indent=2)
        
        return str(path)
    
    def import_config(self, import_path: str, save_as: str = None) -> str:
        """从外部导入配置文件"""
        path = Path(import_path)
        if not path.exists():
            raise FileNotFoundError(f"导入文件不存在: {path}")
        
        with open(path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        
        validation_errors = self.validate_config(config_data)
        if validation_errors:
            raise ConfigValidationError(
                f"导入的配置文件验证失败，共{len(validation_errors)}个错误",
                validation_errors
            )
        
        return self.save_config(config_data, filename=save_as, overwrite=True)
