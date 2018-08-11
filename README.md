# create_sql
Use python to generate sql for insert, update, delete.Intent to save time.

目的：
在工作中需要根据excel文件对数据库进行 insert,update,delete操作，每次重新写不合理。
这个小工具是为了方便下次任务来时一键生成。

所需工具：
安装python 
依赖包：
argparse
xlwt
在安装pip的情况下：
pip install argparse
pip install xlwt

使用：
下载 read_xlsx_base.py 到本地

执行：
python read_xlsx_base.py --path file_path_name.xlsx --sheet_name Sheet_name --sql_type insert --table_name db_table name

不明白命令行可以使用

