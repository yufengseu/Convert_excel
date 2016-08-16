#coding: utf-8
from Tkinter import *
import tkMessageBox, tkFileDialog
import  xdrlib ,sys
import xlrd
import re
reload(sys)
sys.setdefaultencoding('utf-8')


Op_Col_Name = ['������','��Ʒ����','����״̬','���id','�Ӷ������','����ǳ�','��Ʒ����','��Ʒ���','��Ʒ����','��Ʒ����','��Ʒ�ܼ�','�˷�','�����Ż���Ϣ','�ܽ��','��ҹ�����','�ջ�������','�ջ���ַ-ʡ��','�ջ���ַ-�ֵ���ַ','�ʱ�','�ջ����ֻ�','�ջ��˵绰','���ѡ�����ͷ�ʽ','���ұ�������','��������ʱ��','����ʱ��','������˾','��������','��������','��Ʊ̧ͷ','�����ʼ�']
Im_Col_Name = ['�������','�ͻ�����','��Ʒ���','�̼���Ʒ���','����������','����','��Ʒ��Ϣ','����ʱ��','֧��ʱ��','����Ӧ�����(Ԫ)','�������(Ԫ)','����״̬','��ϵ��','���֤','��ϵ�绰','��ϵ��ַ','�ʱ�','��ע']
#Col_Name_Mapping = {Op_Col_Name[0]:Im_Col_Name[0],'����״̬':'����Ѹ���','���id':'��ϵ��','��Ʒ����':'��Ʒ��Ϣ','��Ʒ�ܼ�':'�������','�ܽ��':'�������','�ջ�������':'��ϵ��','�ջ���ַ-�ֵ���ַ':'��ϵ��ַ','�ʱ�':'�ʱ�','�ջ����ֻ�':'��ϵ�绰','��������ʱ��':'����ʱ��','����ʱ��':'֧��ʱ��','��Ʊ̧ͷ':'��ע'}
Col_Name_Mapping = {'������':'�������','���id':'��ϵ��','��Ʒ����':'��Ʒ��Ϣ','��Ʒ�ܼ�':'�������(Ԫ)','�ܽ��':'�������(Ԫ)','�ջ�������':'��ϵ��','�ջ���ַ-�ֵ���ַ':'��ϵ��ַ','�ʱ�':'�ʱ�','�ջ����ֻ�':'��ϵ�绰','��������ʱ��':'����ʱ��','����ʱ��':'֧��ʱ��','��Ʊ̧ͷ':'��ע'}


class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        self.inputFrame = Frame(self)
        self.inputlabel = Label(self.inputFrame, text='�����ļ���:'.decode('gbk'),width=10, height=2)
        self.inputlabel.pack(side=LEFT)
        self.openInputButton = Button(self.inputFrame, text='open', command=self.openInputFile)
        self.openInputButton.pack(side=RIGHT)
        self.nameInput = Entry(self.inputFrame)
        self.nameInput.pack(side=RIGHT)
        self.inputFrame.pack(side=TOP)
        self.outputFrame = Frame(self)
        self.outputlabel = Label(self.outputFrame, text='����ļ���:'.decode('gbk'),width=10, height=2)
        self.outputlabel.pack(side=LEFT)
        self.openOutputButton = Button(self.outputFrame, text='open', command=self.openOutputFile)
        self.openOutputButton.pack(side=RIGHT)
        self.nameOutput = Entry(self.outputFrame)
        self.nameOutput.pack(side=RIGHT)
        self.outputFrame.pack(side=TOP)
        self.alertButton = Button(self, text='��ʼ����Ǩ��'.decode('gbk'), command=self.convert)
        self.alertButton.pack()

#check file


    def openInputFile(self):
        #name = self.nameInput.get() or 'world'
        name = tkFileDialog.askopenfilename()
        self.nameInput.delete(0,END)
        self.nameInput.insert(END, name)

    def openOutputFile(self):
        #name = self.nameInput.get() or 'world'
        name = tkFileDialog.asksaveasfilename()
        self.nameOutput.delete(0,END)
        self.nameOutput.insert(END, name)

    def open_excel(self,file= '8-10����.xls'):
        try:
            data = xlrd.open_workbook(file)
            return data
        except Exception,e:
            print str(e)

    def excel_table_byindex(self,file= 'input.xls',colnameindex=0,by_index=0):
        data = self.open_excel(file)
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

    def convert(self):
       source_file = self.nameInput.get() or 'input.xls'
       target_file = self.nameOutput.get() or 'test.csv'

    
       tables = self.excel_table_byindex(source_file)
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
           tmpStr = ''
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
       tkMessageBox.showinfo('Message', ('�ɹ��������'.decode('gbk')+': %s' % target_file))   

app = Application()
app.master.title('���ת��'.decode('gbk'))
app.master.geometry('300x200')
# ����Ϣѭ��:
app.mainloop()