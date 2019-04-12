import docx
import numpy as np
import pandas as pd
file_name_1=r'测试要求.docx'
file_name_2=r'引脚定义_01.docx'
file1=docx.Document(file_name_1)
file2=docx.Document(file_name_2)
print(file2.paragraphs[12].text)
table=file1.tables[1]
table_list = []
#读取测试要求的表格
for row in table.rows:  # 读每行
	row_content = []
	for cell in row.cells[1:5]:  # 读一行中的所有单元格
		c = cell.text
		row_content.append(c)
	# print(row_content)
	table_list.append(row_content)
print(table_list[2][0])
print(len(file2.tables))
table2=file2.tables[0]
table_list2 = []
#读取引脚定义的表格
for row in table2.rows[1:]:  # 读每行
	  # 读一行中的第三个单元格
	c = row.cells[2].text

	table_list2.append(c)


var_seq=[]
for item in table_list2:
	temp=[]
	temp=item.split('\n')
	var_seq+=temp

print(var_seq)
print(len(var_seq))
#统计行数和配置参数
num_rows=len(table_list)
mm=[]
nn=[]
num_mn=[]

#将统计的一些参数分别放入列表中
for i in range(1,num_rows):
	j=table_list[i][3][0]
	k=table_list[i][3][2]
	j=int(j)
	k=int(k)
	mm.append(j)
	nn.append(k)
	num_mn.append(j*k)

num_sum=sum(num_mn)#计算增加后的总行数
print(num_sum)

var_config=[]
#插空行填补
for i in range(num_rows-1):
	
	for j in range(mm[i]):
		for k in range(nn[i]):
			
			var_t=table_list[i+1][1]+'_%d'%(j+1)+'%d'%(k+1)
			var_config.append(var_t)



#将空的行先配置到列表中
bb=[None]*4

n=2

for i in range(num_rows-1):
	
	if num_mn[i]>1:
		for j in range(num_mn[i]-1):
			table_list.insert(n,bb)
	n+=num_mn[i]

#生成dataframe结构数据类型
df1=pd.DataFrame(table_list[1:num_sum+1],columns=table_list[0])



#增加dataframe中的列
df1['通道配置参数']=var_config
df1['编号']=var_seq
df1.to_excel(r'config.xlsx', index=None, columns=None)
#print(df1)
