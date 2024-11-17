# word_frequency_statistics
# 词频统计程序使用说明手册

## 1. 程序概述

本程序用于处理文本文件（包括 `.txt`, `.docx`, `.pptx`, `.xlsx`, `.xls`, `.pdf` 等格式），并生成词频统计报告。程序支持自定义黑名单、保留词频最高的词汇数量、权重系数和最终截断数量。

## 2. 程序功能

- **文件类型支持**：支持多种文本文件格式，包括 `.txt`, `.docx`, `.pptx`, `.xlsx`, `.xls`, `.pdf`。
- **支持语言**：中文、英文
- **词频统计**：统计文本文件中的词汇频率，并过滤掉停用词和黑名单词。
- **自定义参数**：支持自定义黑名单文件、保留词频最高的词汇数量、权重系数和最终截断数量。
- **日志记录**：记录程序运行的关键信息，包括处理的文件、不支持的文件、错误信息等。
- **输出报告**：生成词频统计报告并保存为 CSV 文件。

## 3. 程序安装

### 3.1 下载程序

1. 从指定的下载链接或路径下载 `utils_jieba_nltk.exe` 文件。
2. 将下载的 `utils_jieba_nltk.exe` 文件解压到一个合适的目录，例如 `C:\Program Files\WordCountTool`。

### 3.2 验证安装

1. 导航到安装目录：

   ```sh
   cd C:\Program Files\WordCountTool
   ```

2. 运行程序以验证安装是否成功：

   ```sh
   utils_jieba_nltk.exe --help
   ```

   如果程序成功运行并显示帮助信息，则安装成功。

## 4. 使用说明

### 4.1 命令行参数

程序支持以下命令行参数：

- `--directory`：指定要处理的目录路径。
- `--blacklist`：指定包含黑名单词的文件路径，每行一个词。
- `--top-n`：指定每个文件保留词频最高的词汇数量（默认值为 10）。
- `--weight-coefficient`：指定权重系数（默认值为 1.0）。
- `--cut-n`：指定权重字典最终截断数量（默认值为 10）。

### 4.2 运行示例

假设你的程序安装在 `C:\Program Files\WordCountTool` 目录下，要处理的文件位于 `C:\Data` 目录，黑名单文件为 `C:\Data\blacklist.txt`，可以使用以下命令运行程序：

```sh
cd C:\Program Files\WordCountTool
utils_jieba_nltk.exe C:\Data --blacklist C:\Data\blacklist.txt --top-n 10 --weight-coefficient 0.5 --cut-n 5
```

### 4.3 参数说明

- `C:\Data`：指定要处理的目录路径。
- `--blacklist C:\Data\blacklist.txt`：指定包含黑名单词的文件路径。
- `--top-n 10`：保留每个文件词频最高的 10 个词汇。
- `--weight-coefficient 0.5`：设置权重系数为 0.5。（系数越小最终输出的权重字典的值越低）
- `--cut-n 5`：权重字典最终保留 5 个词汇。

## 5. 输出文件

### 5.1 词频统计报告

程序会生成一个 CSV 文件，包含词频统计结果。文件名格式为 `{目录名}_{当前时间}.csv`，例如 `Data_20241112_220000.csv`。

### 5.2 日志文件

程序会生成一个日志文件 `word_count.log`，记录程序运行的关键信息，包括处理的文件、不支持的文件、错误信息等。

## 6. 常见问题

### 6.1 权限问题

如果在运行程序时遇到权限问题，尝试以管理员身份运行命令行或文件资源管理器。

### 6.2 路径问题

确保所有路径都是正确的，特别是在 `--add-data` 参数中指定的路径。

### 7.0 如何将python脚本打包为exe文件
 ```sh
pyinstaller --onefile --add-data "nltk_data;nltk_data" utils_jieba_nltk.exe --help
```

# 更新记录

2024/11/15   版本号 2.0（MD5：361229f250a75af5ec27324d63a72e2a）

- 增加了对于 doc、xlsm、ppt、pptm文件格式的支持，但要求客户环境中安装Microsoft Office
- 支持对指定目录下子目录的递归分析



version 2.0  update : 2024/11/15 
