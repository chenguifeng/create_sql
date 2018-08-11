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
python read_xlsx_base.py --h 查看：

python read_xlsx_base.py --h
usage: read_xlsx_base.py [-h] [--path PATH] [--sheet_name SHEET_NAME]
                         [--sql_type SQL_TYPE] [--table_name TABLE_NAME]

optional arguments:
  -h, --help            show this help message and exit
  --path PATH           give the xlsx path to load
  --sheet_name SHEET_NAME
                        give specific sheet which belong to xlsx file
  --sql_type SQL_TYPE   there offer there types of sql statement: insert,
                        update, delete
  --table_name TABLE_NAME
                        sql table name
