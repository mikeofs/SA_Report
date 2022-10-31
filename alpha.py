from data_struct  import SqlTag, DataTypeOPC, excel_row_map
import csv
import re
import time
from  datetime import  datetime
from collections import namedtuple


from DB_work_MS_SQL import  MSSqlInsertEquipment, MSSqlInsertEquipmentParameter, GetParametersFromSQL, InsertDescriptionSQL

start = datetime.now()

query = "Insert INTO  EquipmentParameter ([Name],[Equipment],[DataType],[OPC_Connection])  values ('{0}',@MultiIDENT,{1},'{2}');"


def GetIFixScriptTags(script_path):
    script_file = open(script_path, "r")
    result_tag_list = {}
    constNode = "MOC" # for Vil  --   THISNODE
    for dat in script_file:

        temps = re.search(".*Worksheets\((\d+)\)\.Cells\((\d+), *(\d+)\).*{}\.([\w%]+)\.(A_CV|F_CV).*".format(constNode),dat)
        #".*Worksheets.(\d+).Cells.(\d+).*THISNODE.(\w+).F_CV|A_CV|A_DESC.*"
        if not temps is None:
            result_tag_list[temps.group(4)] = (temps.group(1), temps.group(2), temps.group(3))
    return  result_tag_list


def GetIgsAllenBradlleyTags(address, data_type, *args):

    result_tags_list = {}
    for FilePath in args:
        file_name = FilePath[FilePath.rindex('\\')+1:FilePath.rindex('.')]
        print(file_name)
        with open(FilePath,'r') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=",")
            header = next(csv_reader)
            if address in header and data_type in header:
                result_tags_list.update({"Ethernet.{}.{}".format(file_name, csv_row[header.index(address)]).upper():
                                          csv_row[header.index(data_type)]
                                          for csv_row in csv_reader})

    return result_tags_list

def GetDescriptions (Path):
    result_list = []
    with open(Path, "r") as CFileDescr:
        csv_reader = csv.reader(CFileDescr, delimiter=";")
        for csv_Row in csv_reader:
            result_list.append((csv_Row[1].strip(),csv_Row[2].strip(),csv_Row[3].strip()))
    return result_list

ifix_db_types_markup ={
                        "AI":(3,9),
                        "AA":(3,9),
                        "A0":(3,6),
                        "DI":(3,6),
                        "DO":(3,6),
                        "AR":(2,6),
                        "DR":(2,6),
                        "TX":(1,7)
                      }
equipment_tuple=namedtuple("Equipment", "DB_ID Name Filter") #Coat Sub Line

vil_equipment = (
    [1,"Extruder SB", ".*([._]SB|SB[._]).*"],
    [2, "Extruder SD", ".*([._]SD|SD[._]).*"],
    [3, "Extruder SF", ".*([._]SF|SF[._]).*"],
    [4, "Extruder CC", ".*([._]CC|CC[._]).*"],
    [5, "Extruder CD", ".*([._]CD|CD[._]).*"],
    [6, "Extruder CF", ".*([._]CF|CF[._]).*"],
    [7, "Extruder CG", ".*([._]CG|CG[._]).*"],
    [8, "Coating", ".*([._][Cc][Oo][Aa][Tt]|[Cc][Oo][Aa][Tt][._]).*"],
    [9, "Substrate", ".*([._][Ss][Uu][Bb]|[Ss][Uu][Bb][._]).*"],
    [10,"Line",".*"]
)



def EquipmentSearch(source_str, equipment_list):
     for equipment in equipment_list:
        if re.search(equipment[2], source_str):
            return equipment
     # return (equ for equ in equipment_list if re.search(equ[2], source_str))[0]



def GetIfixDatabaseTags(db_file_path, source_tags, markup):
    result_tag_list = []
    with open(db_file_path, "r") as CFile:
        db_reader = csv.reader(CFile, delimiter=",")
        for db_Row in db_reader:
            if db_Row.__len__() > 1:
                if  db_Row[1] in source_tags:
                        result_tag_list.append(SqlTag(db_Row[1],
                                                      db_Row[markup[db_Row[0].strip()][0]],
                                                      db_Row[markup[db_Row[0].strip()][1]],
                                                      "!",
                                                      source_tags[db_Row[1]]))
                        # print(source_tags[db_Row[1]])

    return result_tag_list


# ifix_script_tags = GetIFixScriptTags("d:\\Xtest\\rep.txt")
# PLC_tags = GetIgsAllenBradlleyTags('Tag Name', 'Data Type', "d:\\Xtest\\EBR1EXTR.csv", "d:\\Xtest\\EBR1ERCS.csv")
# ifix_DB_tags = GetIfixDatabaseTags("d:\\Xtest\\EBR.csv", ifix_script_tags, ifix_db_types_markup)


# for k, tag_sql in enumerate(ifix_DB_tags):
#      tag_sql.type = PLC_tags[tag_sql.adress.upper()]
#      tag_sql.equipment = EquipmentSearch(tag_sql.adress, vil_equipment)
     # print( tag_sql.type, f" -{k}-- ", tag_sql.adress, tag_sql.report_position, tag_sql.name, tag_sql.equipment[1], DataTypeOPC[tag_sql.type], tag_sql.discr )
     #print( tag_sql.equipment[1],',',tag_sql.name,',',
 #            excel_row_map[tag_sql.report_position[2]]+tag_sql.report_position[1], ' ')
#print(vil_equipment)
#MSSqlInsertEquipment(1,vil_equipment)
#MSSqlInsertEquipmentParameter(ifix_DB_tags)

# ifix_DB_tags.sort(key = lambda taggg : taggg.name)
#list(map(lambda tempp : print(tempp.name),ifix_DB_tags))
# listOfDescr=GetDescriptions("D:\Xtest\ZaecDescr.csv")
# listOfDescr.sort(key= lambda kname: kname[0])

# for k, tag in enumerate(ifix_DB_tags):
#    if listOfDescr[k][0] == tag.name:
#         tag.discr = listOfDescr[k][2]
        #print(listOfDescr[k][0], tag.name, tag.discr)

# itList = GetParametersFromSQL()
# itList.sort(key=lambda namee : namee[0])

#for id, tag in enumerate(ifix_DB_tags):
#    print (itList[id][0],itList[id][1],tag.name, tag.discr)
# InsertDescriptionSQL(itList,ifix_DB_tags)
ifix_script_tags = GetIFixScriptTags("D:\\WORK\\FlatCast\\Reports\\FCL_Recipe.txt")

flat_tagrow = "writevalue CDec(Excelapp.Worksheets({}).Cells({}, {}).Value), \"Fix32.MOC.{}.F_CV\""
for fg in ifix_script_tags.items():
    print(flat_tagrow.format(fg[1][0],fg[1][1],fg[1][2],fg[0]))

print(datetime.now()-start)











