#coding: utf-8
import  xdrlib ,sys
import xlrd
import re
#import chardet

reload(sys)
sys.setdefaultencoding('utf-8')

Op_Col_Name = ['������','��Ʒ����','����״̬','���id','�Ӷ������','����ǳ�','��Ʒ����','��Ʒ���','��Ʒ����','��Ʒ����','��Ʒ�ܼ�','�˷�','�����Ż���Ϣ','�ܽ��','��ҹ�����','�ջ�������','�ջ���ַ-ʡ��','�ջ���ַ-�ֵ���ַ','�ʱ�','�ջ����ֻ�','�ջ��˵绰','���ѡ�����ͷ�ʽ','���ұ�������','��������ʱ��','����ʱ��','������˾','��������','��������','��Ʊ̧ͷ','�����ʼ�']
Im_Col_Name = ['�������','�ͻ�����','��Ʒ���','�̼���Ʒ���','����������','����','��Ʒ��Ϣ','����ʱ��','֧��ʱ��','����Ӧ�����(Ԫ)','�������(Ԫ)','����״̬','��ϵ��','���֤','��ϵ�绰','��ϵ��ַ','�ʱ�','��ע']
#Col_Name_Mapping = {Op_Col_Name[0]:Im_Col_Name[0],'����״̬':'����Ѹ���','���id':'��ϵ��','��Ʒ����':'��Ʒ��Ϣ','��Ʒ�ܼ�':'�������','�ܽ��':'�������','�ջ�������':'��ϵ��','�ջ���ַ-�ֵ���ַ':'��ϵ��ַ','�ʱ�':'�ʱ�','�ջ����ֻ�':'��ϵ�绰','��������ʱ��':'����ʱ��','����ʱ��':'֧��ʱ��','��Ʊ̧ͷ':'��ע'}
Col_Name_Mapping = {'������':'�������','���id':'��ϵ��','��Ʒ����':'��Ʒ��Ϣ','��Ʒ�ܼ�':'�������(Ԫ)','�ܽ��':'�������(Ԫ)','�ջ�������':'��ϵ��','�ջ���ַ-�ֵ���ַ':'��ϵ��ַ','�ʱ�':'�ʱ�','�ջ����ֻ�':'��ϵ�绰','��������ʱ��':'����ʱ��','����ʱ��':'֧��ʱ��','��Ʊ̧ͷ':'��ע'}

def open_excel(file= '8-10����.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)
#����������ȡExcel����е�����   ����:file��Excel�ļ�·��     colnameindex����ͷ���������е�����  ��by_index���������
#def excel_table_byindex(file= '8-10����.xls',colnameindex=0,by_index=0):
#    data = open_excel(file)
#    table = data.sheets()[by_index]
#    nrows = table.nrows #����
#    ncols = table.ncols #����
#    colnames =  table.row_values(colnameindex) #ĳһ������ 
#    list =[]
#    for rownum in range(1,nrows):
#
#         row = table.row_values(rownum)
#         if row:
#             app = {}
#             for i in range(len(colnames)):
#                app[colnames[i]] = row[i] 
#             list.append(app)
#    return list

def excel_table_byindex(file= 'input.xls',colnameindex=0,by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #����
    ncols = table.ncols #����
#   colnames =  table.row_values(colnameindex) #ĳһ������ 
    list =[]
    for rownum in range(1,nrows):

         row = table.row_values(rownum)
         if row:
             app = {}
             op_app = {}
             for i in range(len(Im_Col_Name)):
                app[Im_Col_Name[i]] = row[i]

             for j in range(len(Op_Col_Name)):
                if Op_Col_Name[j] in Col_Name_Mapping:
                    op_app[Op_Col_Name[j]] = app[Col_Name_Mapping[Op_Col_Name[j]]]

#'��Ʒ����':'','��Ʒ����':'' '����״̬':'����Ѹ���',
                 
                     
             list.append(op_app)
    return list
#�������ƻ�ȡExcel����е�����   ����:file��Excel�ļ�·��     colnameindex����ͷ���������е�����  ��by_name��Sheet1����
def excel_table_byname(file= 'file.xls',colnameindex=0,by_name=u'Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows #���� 
    colnames =  table.row_values(colnameindex) #ĳһ������ 
    list =[]
    for rownum in range(1,nrows):
         row = table.row_values(rownum)
         if row:
             app = {}
             for i in range(len(colnames)):
                app[colnames[i]] = row[i]
             list.append(app)
    return list

def main():
   source_file = raw_input('�����ļ���:')
   target_file = raw_input('����ļ���:')

   tables = excel_table_byindex(source_file)
   f = open(target_file,'w')
#��ӡ����
   for key in Op_Col_Name:
       f.write('%s,' % key)
   f.write('\n')
#��ӡ����
   for row in tables:
#��ȡ�������ٶ���Ʒ�����и�����ɫ֮��ľ�����Ʒ����������֮��û��������Ϣ
       tmpStr = row['��Ʒ����']
       number = ''
       xx = u'[\u4e00-\u9fa5]+'
       pattern = re.compile(xx)
       match = pattern.split(tmpStr)
       pnum = re.compile(r'[\d]+')
       pre_number = pnum.findall(match[-1])
       if len(pre_number):
          number = pre_number[0]
#���㵥��
       pprice = re.compile(r'[\d]+')
       tmpStr = row['��Ʒ�ܼ�']
       price = int(row['��Ʒ�ܼ�']) / int(number)

#������ļ�
       for key in Op_Col_Name:
       	  if key in row.keys():
            print '%s,' % row[key]
            f.write((u'%s,' % row[key]).encode('gbk'))
          else:
            if key == '����״̬':
                f.write('����Ѹ���,')
            elif key == '��Ʒ����':
                f.write('%s,' % number)
            elif key == '��Ʒ����':
                f.write('%d,' % price)
            else:
          	    f.write(',')

       f.write('\n')
   f.flush()
   f.close()

#   tables = excel_table_byname()
#   for row in tables:
#       print row

if __name__=="__main__":
    main()