import jinja2
from docxtpl import DocxTemplate
import pymysql

#编写获取元数据字典的脚本
schemasql = """select schema_name ,
            DEFAULT_CHARACTER_SET_NAME
from information_schema.SCHEMATA
where schema_name not in('information_schema','mysql','performance_schema','sys');
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
    database = "testdb2"
    charset = "utf8"

    connection = pymysql.connect(host=host, port=port, user=username, password=password, database=database, charset=charset)
    cur = connection.cursor()

    cur.execute(schemasql)
    results = cur.fetchall()
    for result in results:
        schemalist.append(result)


    for db_info in schemalist:
        tablelist_tmp = []
        tablecolumnlist_tmp = []
        db=db_info[0]

        # 编写获取元数据字典的脚本
        tablesql = f"""
        SELECT row_number() over(ORDER BY table_name) as table_no,
               table_name as table_english_name,
               TABLE_COMMENT as table_chinese_name,
               table_catalog as table_catalog,
               table_schema as table_schema,
               TABLE_ROWS as pg_size_pretty
          FROM information_schema.tables 
         WHERE table_schema = "{db}"
         and table_schema !=''
         ORDER BY 2"""

        tablecolumnsql = f"""
        SELECT TABLE_NAME as table_english_name,
                  column_name column_english_name,
                  COLUMN_COMMENT column_chinese_name,
                  DATA_TYPE column_data_type,
                  ORDINAL_POSITION as column_no,
                  COALESCE(character_maximum_length, CHARACTER_OCTET_LENGTH) column_length,
                  (case when IS_NULLABLE = 'YES' then 'Y' else 'N' end) as column_null_flag,
                  (case when COLUMN_KEY = 'PRI' then 'Y' else 'N' end) as column_primarykey_flag,
                  TABLE_SCHEMA as table_schema                 
          FROM information_schema.COLUMNS 
         WHERE table_schema="{db}"
        order by ORDINAL_POSITION    
        """

        cur.execute(tablesql)
        results = cur.fetchall()
        if results:
            for result in results:
                tablelist_tmp.append(result)
            tablelist.append(tablelist_tmp)

        cur.execute(tablecolumnsql)
        results = cur.fetchall()
        if results:
            for result in results:
                tablecolumnlist_tmp.append(result)
            tablecolumnlist.append(tablecolumnlist_tmp)

    cur.close()
    connection.close()
    return schemalist, tablelist, tablecolumnlist

#将数据结构转换为嵌套方式，以便前端容易实现
def trandataformat(tablelist, tablecolumnlist):
    tablecolumnlist_out = []
    for table_tmp in tablelist:
        for table in table_tmp:
            tabledict = {}

            tabledict['table_no'] = table[0]
            tabledict['table_english_name'] = table[1]
            tabledict['table_chinese_name'] = table[2]
            tabledict['table_schema'] = table[4]
            tabledict['table_rows'] = table[5]

            columnlist = []
            for column_tmp in tablecolumnlist:
                for column in column_tmp:

                    if tabledict['table_english_name'] == column[0] and tabledict['table_schema'] == column[8]:
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

def trandataformat_summary_list(tablecolumnlist_out):

    summary_out_list = []

    for table_tmp in tablelist:
        summary_out_dict = {}
        summary_out_dict['table_schema'] = table_tmp[0][4]
        summary_list = []
        for list_out in tablecolumnlist_out:
            if summary_out_dict['table_schema'] == list_out['table_schema']:
                summary_dict = {}
                summary_dict['table_no'] = list_out['table_no']
                summary_dict['table_english_name'] = list_out['table_english_name']
                summary_dict['table_chinese_name'] = list_out['table_chinese_name']
                summary_dict['table_rows'] = list_out['table_rows']
                summary_list.append(summary_dict)
        summary_out_dict["table_info"] = summary_list
        summary_out_list.append(summary_out_dict)

    return summary_out_list


#进行转换，将模板和数据输出到文件中
schemalist, tablelist, tablecolumnlist = gettablecolumninfo()

tablecolumnlist_out = trandataformat(tablelist, tablecolumnlist)

summary_out_list = trandataformat_summary_list(tablecolumnlist_out)


project_name = '系统支撑中心'
project_unit_name = '资源工作组群'
project_description = '数据库一组'
db_type = 'MySQL'


templatefile = r'E:\docxtpl全库版\数据表汇总列表模板.docx'
tpl_table = DocxTemplate(templatefile)
context = {"project_name": project_name,
           "project_unit_name": project_unit_name,
           "project_description": project_description,
           "db_type": db_type,
           "tablecolumnlist": summary_out_list
           }
jinja_env = jinja2.Environment()
tpl_table.render(context, jinja_env)
tpl_table.save("E:\docxtpl全库版\数据表汇总列表.docx")

templatefile2 = r'E:\docxtpl全库版\数据表结构模板.docx'
tpl_table_structure = DocxTemplate(templatefile2)
context2 = {
           "tablecolumnlist": tablecolumnlist_out
           }
tpl_table_structure.render(context2, jinja_env)
tpl_table_structure.save("E:\docxtpl全库版\数据表结构列表.docx")

