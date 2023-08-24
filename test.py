# # from sqlite3 import connect
# # import pandas as pd
# # sql =  'SELECT Закупка2."PN" AS "ЗИП", Закупка2."СУММА СОВМЕСТНОЙ ЗАКУПКИ" AS "$, СТОИМОСТЬ ЗАКУПКИ ЗИП", Закупка2."ОЦЕНОЧНАЯ СТОИМОСТЬ", Закупка2."МАГАЗИН" AS "НАЗНАЧЕНИЕ", Закупка2."ЗАКУПАЕМ ПОД ЗАКАЗЧИКА", Закупка2."КЛИЕНТЫ" FROM Закупка2 WHERE Закупка2."PN" = "FPR9KFAN" GROUP BY Закупка2."PN" HAVING MIN(Закупка2."ID");'
# # # output_dict = pd.DataFrame(sql).value_counts().to_dict()# .to_dict()['FieldType']

# # connectt = connect('data.db')
# # cur = connectt.cursor()
# # df_read = pd.read_sql(sql, connectt).to_dict('records')
# # print(*df_read)

a = 'aA56'

print(a.upper())
