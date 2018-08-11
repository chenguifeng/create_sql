import sys
import os
import argparse
from pyexcel_xls import get_data


class ReadXlsxBase(object):

    def __init__(self, path):
        self.xlsx_path = path
        self.xlsx_data = get_data(self.xlsx_path)

    def check_sheet_name_is_exist(self, sheet_name):
        # check sheet_name is exist
        if sheet_name not in self.xlsx_data.keys():
            return False

        return True

    def get_all_sheet_key(self):
        return [sheet_name for sheet_name in self.xlsx_data.keys()] #

    def get_all_field_name_by_sheet_name(self, sheet_name):
        # the first line of the sheet must be the field
        return self.xlsx_data[sheet_name][0]

    def get_all_data_by_sheet_name(self, sheet_name):
        return self.xlsx_data[sheet_name][1:]

    def get_default_sheet_name(self):
        return self.xlsx_data.keys()[0]


class GenerateSql(object):

    def __init__(self, path, table_name, sheet_name):
        self.xlsx = ReadXlsxBase(path)
        self.table_name = table_name
        if not self.xlsx.check_sheet_name_is_exist(sheet_name):
            raise Exception("file:%s does not contain sheet:%s, please check!" % (path, sheet_name))

        self.sheet_name = sheet_name

        # name -> function
        self.support_sql_type = {'insert': self.create_insert_sql,
                                 'update': self.create_update_sql,
                                 'delete': self.create_delete_sql
                                 }

    insert_head = 'insert into '
    update_head = 'update '

    def create_sql_by_type(self, sql_type, sheet_name):
        print sql_type, sheet_name
        func = self.support_sql_type.get(sql_type)
        if func is not None:
            return func()

    def create_insert_sql(self):
        '''
        insert into tablename (xx,xx,xx) values (xx,xx,xx);
        :param sheet_name:
        :return:
        '''
        sql_statement = []
        sheet_name = self.sheet_name
        value_list = self.get_value_sql()
        for value in value_list:
            sql = self.insert_head + self.table_name + ' ' + self.get_field_sql() + \
                  ' values ' + value + ";"
            sql_statement.append(sql)

        return sql_statement

    def create_update_sql(self):
        '''
        update tablename set xx = xx, where id = xx;
        and we assume the id was in the first position
        :param sheet_name:
        :return:
        '''
        sheet_name = self.sheet_name
        sql_statement = []
        data = self.xlsx.get_all_data_by_sheet_name(sheet_name)
        field = self.xlsx.get_all_field_name_by_sheet_name(sheet_name)

        # get index and value in the same time
        for values in data:
            sql = self.update_head + self.table_name + " set "
            set_value = ""
            where_sql = ""
            for index, value in enumerate(values):
                if index == 0:
                    where_sql = " where " + str(field[index]) + "=" + self.get_value_str(value)
                    where_sql = where_sql[:len(where_sql) - 1]

                set_value += (str(field[index]) + " = " + self.get_value_str(value))

            set_value = set_value[:len(set_value) - 1]
            sql = sql + set_value + where_sql + ";"
            sql_statement.append(sql)

        return sql_statement

    def create_delete_sql(self):
        '''
        delete from table_name where id = xx;
        :param sheet_name:
        :return:
        '''
        sheet_name = self.sheet_name
        sql_statement = []
        data = self.xlsx.get_all_data_by_sheet_name(sheet_name)
        filed = self.xlsx.get_all_field_name_by_sheet_name(sheet_name)

        for values in data:
            value_str = self.get_value_str(values[0])
            value_str = value_str[: len(value_str) - 1]

            delete_sql = "delete from " + self.table_name + " where " + \
            str(filed[0]) + " = " + value_str + ";"
            print delete_sql
            sql_statement.append(delete_sql)

        return sql_statement

    def get_field_sql(self):
        sheet_name = self.sheet_name
        field = self.xlsx.get_all_field_name_by_sheet_name(sheet_name)
        field_start = '('
        field_end = ')'
        field_content = ''
        for field_name in field:
            field_content += (str(field_name) + ',')

        # drop the last ','
        field_content = field_content[:len(field_content) - 1]

        return field_start + field_content + field_end

    def get_value_sql(self):
        sheet_name = self.sheet_name
        value_start = '('
        value_end = ')'
        value_list = []
        data = self.xlsx.get_all_data_by_sheet_name(sheet_name)

        # get the value of the values
        for values in data:
            value_content = ''
            for value in values:
                value_content += self.get_value_str(value)

            value_content = value_content[:len(value_content) - 1]

            value_list.append(value_start + value_content + value_end)

        return value_list

    @staticmethod
    def get_value_str(value):
        # if values was int or float, it don't need " "
        if isinstance(value, (int, float)):
            return str(value) + ","

        return '"' + value + '",'


def check_path(path):
    if os.access(path, os.F_OK) and os.access(path, os.R_OK):
        return True
    return False


if __name__ == "__main__":
    '''
        use 
        argv[1]: xlsx path
        argv[2]: xlsx sheet_name
        argv[3]: xlsx sql_type
    '''
    parser = argparse.ArgumentParser()
    parser.add_argument('--path', help="give the xlsx path to load", type=str)
    parser.add_argument('--sheet_name', help="give specific sheet which belong to xlsx file", type=str)
    parser.add_argument('--sql_type', help='there offer there types of sql statement: insert, update, delete', type=str)
    parser.add_argument('--table_name', help='sql table name', type=str)

    args = parser.parse_args()

    print ('your argument path:%s sheet_name:%s sql_type:%s table_name:%s ' %(args.path, args.sheet_name, args.sql_type, args.table_name))

    # check the path is exist
    if not check_path(args.path):
        print ('path:%s does not exist or have not read permission, please check!' % args.path)
        sys.exit(0)

    try:
        generate_sql = GenerateSql(args.path, args.table_name, args.sheet_name)
        sql_list = generate_sql.create_sql_by_type(args.sql_type, args.sheet_name)
        print ('total len of the sql:%d' % len(sql_list))

        for sql in sql_list:
            print sql

    except Exception as e:
        print e

