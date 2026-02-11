# 中文文本词频统计分析工具 v3.0

科研级面板数据词频统计工具。输出「公司×年份」面板数据格式，适合学术研究和数据分析。

## 核心功能

- 支持 `.xlsx` / `.xls` / `.csv` 数据文件，自动递归扫描子文件夹
- 支持分类词典管理（多分类、多关键词）
- 双模式匹配：**正则匹配**（推荐）和 **jieba 中文分词**
- 大文件自动分块处理（CSV > 100MB 自动分块读取）
- 输出 Excel 面板数据（多 Sheet），含分类统计和关键词明细
- 支持 TF 词频标准化、停用词过滤、命中句子导出
- 自动识别股票代码和日期列
- 超大数据集自动截断保护 + 完整 CSV 备份

## 系统要求

- macOS 10.15 (Catalina) 或更高版本
- Windows 10 / 11（64 位）
- Python 3.10+（从源码运行时）
- 建议内存 8GB 以上（处理大文件时建议 16GB）

## 快速开始

### 方式一：下载预编译程序（推荐）

前往 [Releases](../../releases) 页面下载对应系统的安装包：

- **macOS**：下载 `词频统计分析工具_macOS.zip`，解压后运行 `.app` 文件
- **Windows**：下载 `词频统计分析工具.exe`（需在 Windows 上构建，详见下方说明）

> macOS 首次打开可能提示「无法验证开发者」，右键点击 → 选择「打开」→ 点击「打开」即可。

### 方式二：从源码运行

```bash
# 克隆仓库
git clone https://github.com/chjchen77/word-freq-analyzer.git
cd word-freq-analyzer

# 安装依赖
pip install -r requirements.txt

# 运行程序
python word_freq_analyzer.py
```

## 从源码构建

### macOS

```bash
chmod +x build_app.sh
./build_app.sh
```

### Windows

双击 `build_windows.bat` 或在命令提示符中运行：

```cmd
build_windows.bat
```

构建完成后在 `dist/` 文件夹找到可执行文件。

## 使用说明

程序界面分为 4 个步骤：

1. **数据选择** — 添加数据文件夹，扫描文件，配置列映射（公司代码列、年份列、文本列）
2. **词典管理** — 导入或手动创建分类词典（支持 Excel 和 TXT 格式）
3. **分析设置** — 选择匹配模式、停用词、输出选项
4. **运行分析** — 开始统计，查看实时进度和日志

### 词典格式

**Excel 词典**（`.xlsx` / `.xls`）：两列表格，第一列为分类名，第二列为关键词。

| 分类 | 关键词 |
|------|--------|
| 低碳战略 | 碳中和 |
| 低碳战略 | 碳达峰 |
| 绿色创新 | 绿色技术 |

**文本词典**（`.txt`）：每行一个分类，冒号分隔。

```
低碳战略：碳中和,碳达峰,低碳转型
绿色创新：绿色技术,清洁能源,节能减排
```

### 输出文件

| Sheet | 内容 |
|-------|------|
| Sheet1 | 公司×年份分类统计（含可选 TF 标准化） |
| Sheet2 | 关键词明细（关键词、次数、分类） |
| Sheet3 | 分类汇总（总次数、占比） |
| Sheet4 | 命中句子（可选，含原文句子） |

## 依赖

- [jieba](https://github.com/fxsjy/jieba) >= 0.42.1 — 中文分词
- [pandas](https://pandas.pydata.org/) >= 2.0.0 — 数据处理
- [openpyxl](https://openpyxl.readthedocs.io/) >= 3.0.0 — Excel 读写
- [xlrd](https://xlrd.readthedocs.io/) >= 2.0.1 — .xls 文件读取

## 作者

**陈浩杰** — 澳门城市大学金融学院

## 许可证

MIT License
