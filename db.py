from sqlite3 import connect

import pandas as pd


class Database:

    """
    This class allows you to work with the database
    -----
    * method create() allows you to create tables in the database,
    based on the information received in the form of a list

    * method filing() populating the database

    * method takeinfo() allows you to extract information from the database

    * method close() closes connection to db
    """

    def __init__(self, name: str) -> None:
        self.connect = connect(name)
        self.cur = self.connect.cursor()


    def create(self, *data: list):
        """Create_database."""

        name, values, type_col = data
        param_all = ''
        param_key = ''
        insert = ''
        table_list: list = []
        for item in type_col:
            value = item
            typ = type_col[item]
            try:
                table_list.append(tuple(values[0][item]))
                insert += f'"{value}", '
                param_all += f'"{value}" {typ}'
            except:
                param_key += f'{typ}'

        x = insert.rfind(' ')
        insert = insert[:x-1]
        final_str = f'''CREATE TABLE IF NOT EXISTS {name}({param_all + param_key})'''
        print(final_str)
        self.cur.execute(final_str)

        p = []
        for i in range(len(table_list[0])):
            b = []

            for k in range(len(table_list)):
                b.append(table_list[k][i])

            p.append(tuple(b))

        for key in p:
            into = f'''INSERT INTO {name}({insert}) VALUES {key};'''
            self.cur.execute(into)

        self.connect.commit()

    def filing(self, data, root_dir, excel: 'Excel'):
        """Insert_data_in_database/table."""

        book = excel.load(root_dir)
        list_keys: list = []
        sheetname = book.worksheets

        for sheet in sheetname:
            keys: dict = {}

            for COL in sheet.iter_cols(1, sheet.max_column + 1):
                
                main_key: list = []
                list_item: list = []
                
                if type(COL[0].value) is not str:
                    continue

                value = COL[0].value.upper()
                
                if value in data:
                        
                    if value not in keys:
                        for key in COL[1:]:
                            filter_key = excel.filterkey(key.value, value)
                            if 'TEXT' in data[value]:
                                if value in excel.FILTER_COL:
                                    main_key.append(str(key.value).replace(' ', '').upper())
                                list_item.append(str(filter_key))
                            else:
                                list_item.append(filter_key)

                        if value in excel.FILTER_COL:
                            keys['MAIN_KEY'] = main_key
                        
                        keys[value] = list_item

            long = len(keys)
            if long != 0:
                list_keys.append(keys)

        return list_keys

    def takeinfo(self, sql: str):
        """Update_database."""

        df_read = pd.read_sql(sql, self.connect).to_dict(orient='records')
        return df_read

    def close(self):
        """Close_connection."""
        self.connect.close()
