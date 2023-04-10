import jinja2
from docxtpl import DocxTemplate
import pymysql

#编写获取元数据字典的脚本
db1="testdb1"
schemasql = """select schema_name ,
            DEFAULT_CHARACTER_SET_NAME
from information_schema.schemata 
where schema_name not in('information_schema','mysql','performance_schema','sys');
"""

tablesql = f"""
SELECT row_number() over(ORDER BY table_name) as table_no,
       table_name as table_english_name,
       TABLE_COMMENT as table_chinese_name,
       table_catalog as table_catalog,
       table_schema as table_schema,
       TABLE_ROWS as pg_size_pretty
  FROM information_schema.tables 
 WHERE table_schema = "{db1}"
 ORDER BY 2"""

tablecolumnsql = f"""
SELECT TABLE_NAME as table_english_name,
          column_name column_english_name,
          COLUMN_COMMENT column_chinese_name,
          DATA_TYPE column_data_type,
          ORDINAL_POSITION as column_no,
          COALESCE(character_maximum_length, CHARACTER_OCTET_LENGTH) column_length,
          (case when IS_NULLABLE = 'YES' then 'Y' else 'N' end) as column_null_flag,
          (case when COLUMN_KEY = 'PRI' then 'Y' else 'N' end) as column_primarykey_flag                 
  FROM information_schema.COLUMNS 
 WHERE table_schema="{db1}"
order by 1     
"""

#构造数据库链接，获取数据
def gettablecolumninfo():
    schemalist=[]
    tablelist = []
    tablecolumnlist = []

    host = "10.30.128.187"
    port = 8306
    username = "dba_admin"
    password = "DpzYAuwW9M13db0J"
    database = "testdb1"
    charset = "utf8"

    connection = pymysql.connect(host=host, port=port, user=username, password=password, database=database,charset=charset)
    cur = connection.cursor()

    cur.execute(schemasql)
    results = cur.fetchall()
    for result in results:
        schemalist.append(result)

    cur.execute(tablesql)
    results = cur.fetchall()
    for result in results:
        tablelist.append(result)

    cur.execute(tablecolumnsql)
    results = cur.fetchall()
    for result in results:
        tablecolumnlist.append(result)
    cur.close()

    connection.close()
    return schemalist, tablelist, tablecolumnlist

#将数据结构转换为嵌套方式，以便前端容易实现
def trandataformat(tablelist, tablecolumnlist):
    tablecolumnlist_out = []
    for table in tablelist:
        tabledict = {}
        tabledict['table_no'] = table[0]
        tabledict['table_english_name'] = table[1]
        tabledict['table_chinese_name'] = table[2]
        tabledict['table_schema'] = table[4]
        columnlist = []
        for column in tablecolumnlist:
            if tabledict['table_english_name'] == column[0]:
                columndict = {}
                columndict['table_english_name'] = column[0]
                columndict['column_english_name'] = column[1]
                columndict['column_chinese_name'] = column[2]
                columndict['column_data_type'] = column[3]
                columndict['column_no'] = column[4]
                columndict['column_length'] = column[5]
                columndict['column_pk_flag'] = column[7]
                columnlist.append(columndict)
        tabledict['columns'] = columnlist
        tablecolumnlist_out.append(tabledict)
    return tablecolumnlist_out

#进行转换，将模板和数据输出到文件中
schemalist, tablelist, tablecolumnlist = gettablecolumninfo()

tablecolumnlist_out = trandataformat(tablelist, tablecolumnlist)


project_name = '系统支撑中心'
project_unit_name = '资源工作组群'
project_description = '数据库一组'
db_type = 'MySQL'
db_name = db1

templatefile = r'E:\docxtpl单库版\(单库)数据表汇总列表模板.docx'
tpl_table = DocxTemplate(templatefile)
context = {"project_name": project_name,
           "project_unit_name": project_unit_name,
           "project_description": project_description,
           "db_type": db_type,
           "db_name": db_name,
           "tablecolumnlist": tablecolumnlist_out
           }
jinja_env = jinja2.Environment()
tpl_table.render(context, jinja_env)
tpl_table.save("E:\docxtpl单库版\(单库)数据表汇总列表.docx")

templatefile2 = r'E:\docxtpl单库版\(单库)数据表结构模板.docx'
tpl_table_structure = DocxTemplate(templatefile2)
context2 = {
           "tablecolumnlist": tablecolumnlist_out
           }
tpl_table_structure.render(context2, jinja_env)
tpl_table_structure.save("E:\docxtpl单库版\(单库)数据表结构列表.docx")



