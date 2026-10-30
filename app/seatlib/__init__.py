"""座位轮转助手 · Excel 引擎（纯 Python 标准库，零第三方依赖）。

子模块：
- xlsx_io     : 读取 .xlsx -> ARCHITECTURE §3.1 的 sheet JSON
- xlsx_export : 在原始 xlsx 上原位替换单元格内容的高保真导出

import 本包不产生任何副作用，也不会读取文件。
"""

__all__ = ["xlsx_io", "xlsx_export"]
