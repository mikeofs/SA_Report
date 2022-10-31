import pypyodbc
from data_struct  import SqlTag, DataTypeOPC, excel_row_map

vil_equipment = (
    [1,"Extruder SB"],
   # [2, "Extruder SD"],
   #  [3, "Extruder SF"],
    [4, "Extruder CC"])

def MSSqlInsertEquipment (production_unit,equipment_list):
    connection_string = 'Driver={{SQL Server Native Client 11.0}};' \
                    'Server={server_name};Database={database_name};' \
                    'UID=DataLog;PWD=123456;' \
                    'Trusted_Connection=no;' \
                    .format(server_name='VOLGCC2W63\SA_REPORTS', database_name='Reports')
    sql_text = 'SET NOCOUNT ON;  DECLARE @return_ident int;' \
               'EXEC  @return_ident = EQUIPMENT_INSERT \'{0}\',{1};' \
               'SELECT @return_ident'

    with pypyodbc.connect(connection_string) as con:
        for equip in equipment_list:
            cur = con.cursor()
            cur.execute(sql_text.format(equip[1],production_unit))
            equip[0]=cur.fetchall()[0][0]

def MSSqlInsertEquipmentParameter (parameters_list:SqlTag):
    connection_string = 'Driver={{SQL Server Native Client 11.0}};' \
                    'Server={server_name};Database={database_name};' \
                    'UID=DataLog;PWD=123456;' \
                    'Trusted_Connection=no;' \
                    .format(server_name='VOLGCC2W63\SA_REPORTS', database_name='Reports')
    sql_text = 'SET NOCOUNT ON;  DECLARE @return_ident int;' \
               'EXEC  @return_ident = EQUIPMENT_PARAMETER_INSERT \'{0}\',{1},{2},\'{3}\';' \
               'SELECT @return_ident'

    with pypyodbc.connect(connection_string) as con:
        for parameter in parameters_list:
            cur = con.cursor()
            cur.execute(sql_text.format(parameter.name,
                                        DataTypeOPC[parameter.type],
                                        parameter.equipment[0],
                                        parameter.adress
                                        ))
            # parameter[0]=cur.fetchall()[0][0]
def GetParametersFromSQL():
    connection_string = 'Driver={{SQL Server Native Client 11.0}};' \
                        'Server={server_name};Database={database_name};' \
                        'UID=AlehinS;PWD=900800Als;' \
                        'Trusted_Connection=no;' \
                        .format(server_name='VOLGCC2W63\SA_REPORTS', database_name='Reports')
    sql_text = "Select P.[Name] ,P.[Id] " \
               "From  dbo.EquipmentParameter as P INNER JOIN dbo.Equipment as E ON (E.Id = P.Equipment) where E.ProductionUnit = 1"
    Result=[]
    with pypyodbc.connect(connection_string) as con:
        cur = con.cursor()
        cur.execute(sql_text)
        Result=cur.fetchall()
    return Result

def InsertDescriptionSQL(paramList, tagList):

    connection_string = 'Driver={{SQL Server Native Client 11.0}};' \
                        'Server={server_name};Database={database_name};' \
                        'UID=DataLog;PWD=123456;' \
                        'Trusted_Connection=no;' \
                        .format(server_name='VOLGCC2W63\SA_REPORTS', database_name='Reports')
    sql_text = "SET NOCOUNT ON;  UPDATE [dbo].[EquipmentParameter] SET [Description]=\'{0}\' WHERE [Id] = {1}"
    with pypyodbc.connect(connection_string) as con:
        for paramId, tag in enumerate(tagList):
                cur = con.cursor()
                cur.execute(sql_text.format(tag.discr,paramList[paramId][1]))
            # print(sql_text.format(tag.discr,paramList[paramId][1]))

def testReport():
    connection_string = 'Driver={{SQL Server Native Client 11.0}};' \
                        'Server={server_name};Database={database_name};' \
                        'UID=AlehinS;PWD=900800Als;' \
                        'Trusted_Connection=no;' \
        .format(server_name='VOLGCC2W63\SA_REPORTS', database_name='Reports')
    sql_text = "Select P.[Name] ,CONCAT(V.ValueInt, V.ValueSTR, V.ValueREAL) "\
                "From dbo.ParametersValue as V INNER JOIN dbo.EquipmentParameter as P ON (V.Parameter = p.Id) "\
                    "where V.Report = 7"
    Data_sent = []
    with pypyodbc.connect(connection_string) as con:
        cur = con.cursor()
        cur.execute(sql_text)
        Data_sent=cur.fetchall()

    for ddf in Data_sent:
        print(ddf[0],ddf[1])
# MSSqlInsertEquipment(1,vil_equipment)
# inn=[SqlTag("SUB_DF",'aa',"Ether.EXTR.moreno12","Short","122")]
# inn[0].equipment=[1,"Extruder SB"]
# MSSqlInsertEquipmentParameter(inn)
# print(vil_equipment)
#test()
#itList = GetParametersFromSQL()
#itList.sort(key=lambda namee : namee[0])
#list(map(print,itList))
