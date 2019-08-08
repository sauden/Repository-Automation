__author__ = 'i20764'
import os
import re
import config.config
import dbconnect.dbconnection
from mantis.mantis_detail import createdir,getMantisDetail,get_dir
from excel.excel_work import copy_file,write_file,copyData
from excel.check_mismatch import get_header,get_file_count,get_column,check_valid

db = dbconnect.dbconnection.dbconnect()

def main():
    note = getMantisDetail()
    x=1
    while x==1:
        for summary in note:
            if summary[1] is not None:
                result_single = summary[1].split('.')
                if len(result_single) <= 4:
                    table_name = result_single[1]
                    operation = result_single[2].split()
                    id = str(summary[0])
                    print(id)
                    if ('INSERT' or 'Insert' or 'insert') in operation:
                        ind_operation = 'Insert'
                    elif ('UPDATE' or 'Update' or 'update') in operation:
                        ind_operation = 'Update'
                    else:
                        ind_operation = 'Delete'

                    wd = createdir(id)

                    if wd==1:
                        x=note
                        continue

                    try:
                        cur = db.cursor()
                        cur.execute('select content from mantis.mantis_bug_file_table where bug_id = '+id+' and id in (select max(id) from mantis.mantis_bug_file_table where bug_id = '+id+')')
                        res = cur.fetchone()[0]
                        write_file(res,config.config.repo_excel+'\\'+id+'.xlsx')
                    except Exception as e:
                        print(str(e))
                    finally:
                        cur.close()

                    list_of_dir = get_dir()

                    if table_name in list_of_dir:
                        listOfFile = os.listdir(config.config.template_path+table_name+'\\'+ind_operation+'\\')

                        for file in listOfFile:
                            with open(config.config.template_path+table_name+'\\'+ind_operation+'\\'+file,'r',errors = 'ignore') as f:
                                filedata = f.read()
                            file = re.sub('xxxxx',id,file,flags=re.IGNORECASE)
                            filedata = re.sub('xxxxx',id,filedata,flags=re.IGNORECASE)
                            wf = open(config.config.path+id+'\\'+file,'w')
                            wf.write(filedata)
                            wf.close()


                        file_xlsx = id+'.xlsx'
                        print(file_xlsx)

                        if os.path.isfile(config.config.repo_excel+file_xlsx):
                            check = check_valid(config.config.repo_excel+'\\'+id+'.xlsx',config.config.path+id+'\\')
                            if check == 'Pass':
                                        (config.config.repo_excel+'\\'+id+'.xlsx',config.config.path+id+'\\')
                            else:
                                print('Erro:', check)
                                continue
                        else:
                            print('***Please confirm spreadsheets are attached or not in mantis -'+id+' and also look into mantis description.***')

                    else:
                        print('*****Please check note of manits - '+id+'*****')
                        continue
                else: #changed here
                    result_double = [x.upper() for x in summary[1].split()]
                    id = str(summary[0])
                    print(id)
                    for i in range(len(result_double)):
                        if result_double[i][-1] == '.':
                            result_double[i] = result_double[i][:-1]

                    search = ['INSERT','UPDATE','DELETE']
                    update_and_insert = ['UPDATE','INSERT']
                    update_and_delete = ['UPDATE','DELETE']
                    tablename_check = ['REPOSITORY.PCI_EDIT_NONCOV_CODES','REPOSITORY.PCI_EDIT_NONCOV_RULE']

                    for_table_check = list(set(result_double)&set(tablename_check))


                    for_1_operation = list(set(result_double)&set(search))
                    if len(for_1_operation)==1:
                        for x in for_1_operation:
                            if x == 'INSERT':
                                table_name = 'PCI_EDIT_NONCOV_RULE'
                                operation = 'Insert_Rule_Code'
                            if x == 'UPDATE':
                                table_name = 'PCI_EDIT_NONCOV_RULE'
                                operation = 'Update_Rule_Code'
                            if x == 'DELETE':
                                for_table_check = list(set(tablename_check)&set(result_double))
                                if len(for_table_check) == 2:
                                    table_name = 'PCI_EDIT_NONCOV_RULE_ONLY'
                                    operation = 'Delete'
                    else:
                        for_2_operation = list(set(for_1_operation)&set(update_and_insert))
                        if len(for_2_operation) == 2:
                            table_name = 'PCI_EDIT_NONCOV_RULE'
                            operation = 'Update_Rule_Ins_Code'
                        else:
                            table_name = 'PCI_EDIT_NONCOV_RULE'
                            operation = 'Update_Rule_Del_codes'

                    wd = createdir(id)

                    if wd==1:
                        x = note
                        continue

                    try:
                        cur = db.cursor()
                        cur.execute('select content from mantis.mantis_bug_file_table where bug_id = '+id+' and id in (select max(id) from mantis.mantis_bug_file_table where bug_id = '+id+')')
                        res = cur.fetchone()[0]
                        write_file(res,config.config.repo_excel+'\\'+id+'.xlsx')
                    except Exception as e:
                        print(str(e))
                    finally:
                        cur.close()

                    list_of_dir = get_dir()

                    if table_name in list_of_dir:
                        listOfFile = os.listdir(config.config.template_path+table_name+'\\'+operation+'\\')

                        for file in listOfFile:
                            with open(config.config.template_path+table_name+'\\'+operation+'\\'+file,'r') as f:
                                filedata = f.read()

                            file = re.sub('xxxxx',id,file,flags=re.IGNORECASE)
                            filedata = re.sub('xxxxx',id,filedata,flags=re.IGNORECASE)
                            wf = open(config.config.path+id+'\\'+file,'w')
                            wf.write(filedata)
                            wf.close()

                        file_xlsx = id+'.xlsx'
                        print(file_xlsx)

                        if os.path.isfile(config.config.repo_excel+file_xlsx):
                            check = check_valid(config.config.repo_excel+'\\'+id+'.xlsx',config.config.path+id+'\\')
                            if check == 'Pass':
                                        copyData(config.config.repo_excel+'\\'+id+'.xlsx',config.config.path+id+'\\')
                            else:
                                print('Erro:', check)
                                continue
                        else:
                            print('***Please confirm spreadsheets are attached or not in mantis -'+id+' and also look into mantis description.***')

                    else:
                        print('*****Please check note of manits - '+id+'*****')
                        continue
                copy_file(config.config.create_sqlplus,config.config.path+id, symlinks=False, ignore=None)
            else:
                print('***Please look into mantis -'+str(summary[0])+':(may be no added note in this mantis***')
                continue
        x=note
    print("Successfully files are created")

if __name__ == '__main__':
    main()






