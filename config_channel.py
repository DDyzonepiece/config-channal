import docx
import numpy as np
import pandas as pd

def psi_out(x):
	psi=(int(x)-101.25)/6.895
	print(psi)
	if psi<5:
		return 5
	elif 5<=psi<15:
		return 15
	elif 15<=psi<30:
		return 30
	elif 30<=psi<100:
		return 100
	elif 100<=psi<250:
		return 250
	elif 250<=psi<500:
		return 500
	elif 500 <= psi < 750:
		return 750
	else:
		return 0


file_name_1=r'测试要求.docx'
file_name_2=r'引脚定义_01.docx'
file1=docx.Document(file_name_1)
file2=docx.Document(file_name_2)
print(file2.paragraphs[12].text)
table=file1.tables[1]
table_list_1 = []
range_list=[]
#读取测试要求的表格
for row in table.rows:  # 读每行
	row_content = []
	range_list.append(row.cells[5].text)
	for cell in row.cells[1:7]:  # 读一行中的所有单元格
		c = cell.text
		row_content.append(c)
	# 用二维列表去表示读取的表格
	table_list_1.append(row_content)
print(table_list_1[2][0])
print(len(file2.tables))
table2=file2.tables[0]
table_list_2 = []
#读取引脚定义的表格
for row in table2.rows[1:]:  # 读每行
	  # 读一行中的第三个单元格
	c = row.cells[2].text

	table_list_2.append(c)


var_seq=[]
for item in table_list_2:
	temp=[]
	temp=item.split('\n')
	var_seq+=temp

#print(var_seq)
print(len(var_seq))
#统计行数和配置参数
num_rows=len(table_list_1)
mm=[]
nn=[]
num_mn=[]

#将统计的一些参数分别放入列表中
for i in range(1,num_rows):
	j=table_list_1[i][3][0]
	k=table_list_1[i][3][2]
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
			
			var_t=table_list_1[i+1][1]+'_%d'%(j+1)+'%d'%(k+1)
			var_config.append(var_t)


#table_list_1[0].append('量程（表压）')
max_range_list=[]
#分割范围
for item in range_list[1:]:
	temp=[]
	temp=item.split('~')
	print(temp)
	max_range_list.append(temp[1])

print(max_range_list)

#psi_list

psi_list=[psi_out(x) for x in max_range_list]
print(psi_list)
#将非压力的参数改为空格
for i,item in enumerate(psi_list):
	if table_list_1[i+1][1][0]!='P':
		psi_list[i]=' '



# psi_list_extend=[]
# for i in range(len(psi_list)):
# 	item=[psi_list[i]]*num_mn[i]
# 	psi_list_extend.extend(item)
#
# print(psi_list_extend)
# print(len(psi_list_extend))


#将空的行先配置到列表中
bb=[None]*6

n=2

for i in range(num_rows-1):
	
	if num_mn[i]>1:
		for j in range(num_mn[i]-1):
			table_list_1.insert(n,bb)
			psi_list.insert(n-1,None)
	n+=num_mn[i]

#生成dataframe结构数据类型
df1=pd.DataFrame(table_list_1[1:num_sum+1],columns=table_list_1[0])


print(psi_list)

#增加dataframe中的列
df1['量程（表压）（psi）']=psi_list
df1['通道配置参数']=var_config
df1['编号']=var_seq

df1.to_excel(r'config.xlsx', index=None, columns=None)
#print(df1)
