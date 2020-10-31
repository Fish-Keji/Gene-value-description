#匹配基因名与差异表达倍数及p值

#将两个列表组成为键值对，构建字典
def creat_dic_from_2list(list1, list2):
    new_dict = dict(map(lambda x,y:[x,y],list1,list2))
    return new_dict
   
import xlrd
import xlwt
#读取被查找表格
book1 = xlrd.open_workbook('/Users/apky/Desktop/待查表格.xlsx')
sheet1 = book1.sheet_by_name('sheet名称')
gene_ID = sheet1.col_values(0)
FoldChange = sheet1.col_values(9)
Pvalue = sheet1.col_values(10)
padj = sheet1.col_values(11)
abrev = sheet1.col_values(12)
descript = sheet1.col_values(19)
#读取索引
book2 = xlrd.open_workbook('/Users/apky/Desktop/包含索引表格.xlsx')
sheet2 = book2.sheet_by_name('UP-intersection')
list_to_find = sheet2.col_values(0)
#通过列表构建字典，引用函数
FoldChange_dict = creat_dic_from_2list(gene_ID, FoldChange)
Pvalue_dict = creat_dic_from_2list(gene_ID, Pvalue)
padj_dict = creat_dic_from_2list(gene_ID, padj)
abrev_dict = creat_dic_from_2list(gene_ID, abrev)
descript_dict = creat_dic_from_2list(gene_ID, descript)
#新建excel文件，原文件只读
new_excel = xlwt.Workbook()
sheet_match = new_excel.add_sheet('sheet_match')

#初始行值为n = 0
n = 0
#历遍索引
for name in list_to_find:
    a = FoldChange_dict[name]
    b = Pvalue_dict[name]
    c = padj_dict[name]
    d = abrev_dict[name]
    e = descript_dict[name]
    sheet_match.write(n, 1, a)#写入Foldchange
    sheet_match.write(n, 2, b) #写入p
    sheet_match.write(n, 3, c) #写入padj
    sheet_match.write(n, 4, d) #写入简名
    sheet_match.write(n, 5, e) #写入基因长注释
    sheet_match.write(n, 0, name)#写入基因索引
    n = n+1 #迭代
new_excel.save('/Users/apky/Desktop/新文件.xls')#存储为新文件
