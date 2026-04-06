# 论文格式检测工具

## 项目简介

论文格式检测工具是一个用于自动检测Word文档格式是否符合学术规范的Python工具。该工具能够检测论文的字体、字号、段落格式、页边距等多种格式要素，帮助用户快速发现并修正格式问题。

## 功能特性

- 支持检测Word文档（.docx格式）
- 可配置的格式规则
- 生成详细的检测报告
- 支持批量处理

## 项目结构

```
论文格式检测/
├── src/              # 源代码目录
├── config/           # 配置文件目录
├── tests/            # 测试文件目录
├── output/           # 输出目录
├── main.py           # 主程序入口
├── requirements.txt  # 项目依赖
└── README.md         # 项目说明文档
```

## 安装说明

### 环境要求

- Python 3.8+
- pip

### 安装步骤

1. 克隆或下载项目到本地

2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

```bash
python main.py -i <输入文档路径> -o <输出目录> -c <配置文件路径>
```

### 参数说明

- `-i, --input`: 输入文档路径（必需）
- `-o, --output`: 输出目录（默认：output）
- `-c, --config`: 配置文件路径（默认：config/rules.yaml）

### 示例

```bash
python main.py -i thesis.docx -o output -c config/rules.yaml
```

## 开发计划

- [ ] 实现文档解析模块
- [ ] 实现格式检测规则引擎
- [ ] 实现报告生成模块
- [ ] 添加单元测试
- [ ] 完善配置文件系统

## 许可证

MIT License
