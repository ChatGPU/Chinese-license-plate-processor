# Excel License Plate Location Processor | Excel车牌归属地处理工具

A simple yet powerful Python script that batch-processes Excel files to add province and city information based on Chinese license plate numbers. Created with Google Gemini.

一个简洁而强大的Python脚本，可批量处理Excel文件，根据中国车牌号码自动添加省份和城市归属地信息。

---

### ✨ Key Features | 主要特性

* **Batch Processing / 批量处理**: Automatically finds and processes all `.xlsx` and `.xls` files in its folder. / 自动查找并处理当前文件夹下的所有 `.xlsx` 和 `.xls` 格式的Excel文件。
* **Easy to Configure / 易于配置**: All settings (like column names) are in a clear `CONFIG` section at the top of the script. No need to dig through code. / 所有重要设置（如列名）都集中在脚本顶部的 `CONFIG` 配置区，无需深入代码即可修改。
* **Safe / 安全可靠**: Never modifies your original files. It saves the results in a new subfolder named `处理后表格` (or a custom name you set). / 绝不修改原始文件。脚本会将处理后的结果保存到一个新的子文件夹中（默认为 `处理后表格`），确保您的源数据安全。
* **Robust / 稳定运行**: Gracefully skips files that don't have the required license plate column and handles potential errors. / 当在文件中找不到指定的车牌号列时，会自动跳过该文件，避免程序因错误而中断。
* **Easy to Maintain / 易于维护**: The license plate prefix data is stored in a simple Python dictionary, making it easy to update or correct. / 所有的车牌前缀与地区对应数据都储存在一个独立的Python字典中，更新和修正数据非常方便。
* **Intelligent Column Placement / 智能列排序**: Inserts the new Province and City columns directly before the original license plate column for easy comparison. / 自动将新添加的“省份”和“城市”列放置在原始车牌号列的前面，方便数据核对与比较。

---

### ⚙️ Installation & Requirements | 安装与环境要求

You need to have Python 3 installed on your system. Then, install the required libraries using pip:

您需要在系统中安装 Python 3。然后，使用 pip 安装所需的第三方库：

```bash
pip install pandas openpyxl xlrd
