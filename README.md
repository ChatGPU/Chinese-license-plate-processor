# Excel License Plate Location Processor | Excel车牌归属地处理工具

A simple yet powerful Python script that batch-processes Excel files to add province and city information based on Chinese license plate numbers. Created with Google Gemini.

一个简洁而强大的Python脚本，可批量处理Excel文件，根据中国车牌号码自动添加省份和城市归属地信息。

---

### ✨ Key Features | 主要特性

* **Batch Processing / 批量处理**: Processes `.xlsx` and `.xls` files from configurable paths, with optional recursive search. / 可配置多个路径并可递归处理 `.xlsx` 与 `.xls` 文件。
* **Easy to Configure / 易于配置**: All settings (like column names) are in a clear `CONFIG` section at the top of the script. No need to dig through code. / 所有重要设置（如列名）都集中在脚本顶部的 `CONFIG` 配置区，无需深入代码即可修改。
* **Auto Column Detection / 自动识别列名**: Supports aliases and keyword matching when the exact column name differs. / 支持别名与关键字匹配，列名不一致也能自动识别。
* **Robust Input Cleaning / 输入清洗**: Trims spaces, separators, and full-width characters; normalizes letter case. / 自动清理空格、分隔符、全角字符，并统一字母大小写。
* **Multi-Sheet Safe / 多表安全**: Preserves other sheets and can process all or specified sheets. / 可处理多工作表并保留未处理的表。
* **Safe / 安全可靠**: Never modifies your original files. It saves the results in a new subfolder named `处理后表格` (or a custom name you set). / 绝不修改原始文件。脚本会将处理后的结果保存到一个新的子文件夹中（默认为 `处理后表格`），确保您的源数据安全。
* **Easy to Maintain / 易于维护**: The license plate prefix data is stored in a simple Python dictionary, making it easy to update or correct. / 所有的车牌前缀与地区对应数据都储存在一个独立的Python字典中，更新和修正数据非常方便。
* **Intelligent Column Placement / 智能列排序**: Inserts the new Province and City columns directly before the original license plate column for easy comparison. / 自动将新添加的“省份”和“城市”列放置在原始车牌号列的前面，方便数据核对与比较。

---

### ⚙️ Installation & Requirements | 安装与环境要求

You need to have Python 3 installed on your system. Then, install the required libraries using pip:

您需要在系统中安装 Python 3。然后，使用 pip 安装所需的第三方库：

```bash
pip install pandas openpyxl xlrd
```

Optional (only needed if you want to keep `.xls` output without conversion):

可选（仅当需要保留 `.xls` 输出时安装）：

```bash
pip install xlwt
```

---

### 🧩 Configuration Notes | 配置说明

* **`input_paths`**: A list of files, folders, or glob patterns to process. / 可配置文件、文件夹或通配符路径列表。
* **`recursive_search`**: Set to `True` to search subfolders. / 设为 `True` 时可递归查找子目录。
* **`process_all_sheets` & `sheet_names`**: Process all sheets or only specific ones. / 可处理全部工作表或仅指定工作表。
* **`preserve_other_sheets`**: Keep unprocessed sheets in the output. / 输出文件保留未处理的工作表。
* **`input_column_aliases` & `input_column_keywords`**: Used for auto-detecting the plate column. / 用于自动识别车牌列。
* **`overwrite_existing_output_columns`**: Set to `False` to avoid overwriting existing columns. / 设为 `False` 可避免覆盖已有列。
* **`.xls` handling**: Without `xlwt`, `.xls` files will be saved as `.xlsx`. / 未安装 `xlwt` 时，`.xls` 会保存为 `.xlsx`。
