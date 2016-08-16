#coding: utf-8
from Tkinter import *
import tkMessageBox, tkFileDialog
import  xdrlib ,sys
import xlrd
import re
reload(sys)
sys.setdefaultencoding('utf-8')


Op_Col_Name = ['订单号','产品条码','订单状态','买家id','子订单编号','买家昵称','商品名称','产品规格','商品单价','商品数量','商品总价','运费','购买优惠信息','总金额','买家购买附言','收货人姓名','收货地址-省市','收货地址-街道地址','邮编','收货人手机','收货人电话','买家选择运送方式','卖家备忘内容','订单创建时间','付款时间','物流公司','物流单号','发货附言','发票抬头','电子邮件']
Im_Col_Name = ['订单编号','客户单号','商品编号','商家商品编号','主机订单号','渠道','商品信息','创建时间','支付时间','当期应付金额(元)','订单金额(元)','订单状态','联系人','身份证','联系电话','联系地址','邮编','备注']
#Col_Name_Mapping = {Op_Col_Name[0]:Im_Col_Name[0],'订单状态':'买家已付款','买家id':'联系人','商品名称':'商品信息','商品总价':'订单金额','总金额':'订单金额','收货人姓名':'联系人','收货地址-街道地址':'联系地址','邮编':'邮编','收货人手机':'联系电话','订单创建时间':'创建时间','付款时间':'支付时间','发票抬头':'备注'}
Col_Name_Mapping = {'订单号':'订单编号','买家id':'联系人','商品名称':'商品信息','商品总价':'订单金额(元)','总金额':'订单金额(元)','收货人姓名':'联系人','收货地址-街道地址':'联系地址','邮编':'邮编','收货人手机':'联系电话','订单创建时间':'创建时间','付款时间':'支付时间','发票抬头':'备注'}


class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        self.inputFrame = Frame(self)
        self.inputlabel = Label(self.inputFrame, text='输入文件名:'.decode('gbk'),width=10, height=2)
        self.inputlabel.pack(side=LEFT)
        self.openInputButton = Button(self.inputFrame, text='open', command=self.openInputFile)
        self.openInputButton.pack(side=RIGHT)
        self.nameInput = Entry(self.inputFrame)
        self.nameInput.pack(side=RIGHT)
        self.inputFrame.pack(side=TOP)
        self.outputFrame = Frame(self)
        self.outputlabel = Label(self.outputFrame, text='输出文件名:'.decode('gbk'),width=10, height=2)
        self.outputlabel.pack(side=LEFT)
        self.openOutputButton = Button(self.outputFrame, text='open', command=self.openOutputFile)
        self.openOutputButton.pack(side=RIGHT)
        self.nameOutput = Entry(self.outputFrame)
        self.nameOutput.pack(side=RIGHT)
        self.outputFrame.pack(side=TOP)
        self.alertButton = Button(self, text='开始数据迁移'.decode('gbk'), command=self.convert)
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

    def open_excel(self,file= '8-10上午.xls'):
        try:
            data = xlrd.open_workbook(file)
            return data
        except Exception,e:
            print str(e)

    def excel_table_byindex(self,file= 'input.xls',colnameindex=0,by_index=0):
        data = self.open_excel(file)
        table = data.sheets()[by_index]
        nrows = table.nrows #行数
        ncols = table.ncols #列数
    #   colnames =  table.row_values(colnameindex) #某一行数据 
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
    
    #'商品单价':'','商品数量':'' '订单状态':'买家已付款',
                     
                         
                 list.append(op_app)
        return list

    def convert(self):
       source_file = self.nameInput.get() or 'input.xls'
       target_file = self.nameOutput.get() or 'test.csv'

    
       tables = self.excel_table_byindex(source_file)
       f = open(target_file,'w')
    #打印列名
       for key in Op_Col_Name:
           f.write('%s,' % key)
       f.write('\n')
    #打印数据
       for row in tables:
    #提取数量，假定商品名称中跟在颜色之后的就是商品数量，而且之后没有其他信息
           tmpStr = row['商品名称']
           number = ''
           xx = u'[\u4e00-\u9fa5]+'
           pattern = re.compile(xx)
           match = pattern.split(tmpStr)
           pnum = re.compile(r'[\d]+')
           pre_number = pnum.findall(match[-1])
           if len(pre_number):
              number = pre_number[0]
    #计算单价
           pprice = re.compile(r'[\d]+')
           tmpStr = row['商品总价']
           price = int(row['商品总价']) / int(number)
    
    #输出到文件
           tmpStr = ''
           for key in Op_Col_Name:
              if key in row.keys():
                print '%s,' % row[key]
                f.write((u'%s,' % row[key]).encode('gbk'))
              else:
                if key == '订单状态':
                    f.write('买家已付款,')
                elif key == '商品数量':
                    f.write('%s,' % number)
                elif key == '商品单价':
                    f.write('%d,' % price)
                else:
          	        f.write(',')
    
           f.write('\n')
       f.flush()
       f.close()
       tkMessageBox.showinfo('Message', ('成功！输出到'.decode('gbk')+': %s' % target_file))   

app = Application()
app.master.title('表格转换'.decode('gbk'))
app.master.geometry('300x200')
# 主消息循环:
app.mainloop()