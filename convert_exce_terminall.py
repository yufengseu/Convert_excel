#coding: utf-8
import  xdrlib ,sys
import xlrd
import re
#import chardet

reload(sys)
sys.setdefaultencoding('utf-8')

Op_Col_Name = ['订单号','产品条码','订单状态','买家id','子订单编号','买家昵称','商品名称','产品规格','商品单价','商品数量','商品总价','运费','购买优惠信息','总金额','买家购买附言','收货人姓名','收货地址-省市','收货地址-街道地址','邮编','收货人手机','收货人电话','买家选择运送方式','卖家备忘内容','订单创建时间','付款时间','物流公司','物流单号','发货附言','发票抬头','电子邮件']
Im_Col_Name = ['订单编号','客户单号','商品编号','商家商品编号','主机订单号','渠道','商品信息','创建时间','支付时间','当期应付金额(元)','订单金额(元)','订单状态','联系人','身份证','联系电话','联系地址','邮编','备注']
#Col_Name_Mapping = {Op_Col_Name[0]:Im_Col_Name[0],'订单状态':'买家已付款','买家id':'联系人','商品名称':'商品信息','商品总价':'订单金额','总金额':'订单金额','收货人姓名':'联系人','收货地址-街道地址':'联系地址','邮编':'邮编','收货人手机':'联系电话','订单创建时间':'创建时间','付款时间':'支付时间','发票抬头':'备注'}
Col_Name_Mapping = {'订单号':'订单编号','买家id':'联系人','商品名称':'商品信息','商品总价':'订单金额(元)','总金额':'订单金额(元)','收货人姓名':'联系人','收货地址-街道地址':'联系地址','邮编':'邮编','收货人手机':'联系电话','订单创建时间':'创建时间','付款时间':'支付时间','发票抬头':'备注'}

def open_excel(file= '8-10上午.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)
#根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
#def excel_table_byindex(file= '8-10上午.xls',colnameindex=0,by_index=0):
#    data = open_excel(file)
#    table = data.sheets()[by_index]
#    nrows = table.nrows #行数
#    ncols = table.ncols #列数
#    colnames =  table.row_values(colnameindex) #某一行数据 
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
#根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_name：Sheet1名称
def excel_table_byname(file= 'file.xls',colnameindex=0,by_name=u'Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows #行数 
    colnames =  table.row_values(colnameindex) #某一行数据 
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
   source_file = raw_input('输入文件名:')
   target_file = raw_input('输出文件名:')

   tables = excel_table_byindex(source_file)
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

#   tables = excel_table_byname()
#   for row in tables:
#       print row

if __name__=="__main__":
    main()