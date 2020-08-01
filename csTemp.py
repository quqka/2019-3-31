import os
import csv
import xlrd

# 写入csv文件
def csvWriter(name, datas):
	# 打开csv文件
	with open("csTemp/{}.csv".format(name),"a", encoding="utf-8", newline='') as c:
		csv_writer = csv.writer(c)
		csv_writer.writerow([datas])
		c.close()

# 新建csv文件
def csvName(name):
	csvfilename = "csTemp/{}.csv".format(name)
	# if not os.path.exists(csvfilename):
	# 	c=open(csvfilename,"w")
	# 	c.close()
	c=open(csvfilename, "w", encoding="utf-8")
	c.close()


def isComment(data, li):
	for i in csvNameList:
		# 分类excel内容到csv中
		if i.split('_')[-1] == tw:
			# 判断是否有注释(中文符号)
			if '*注：' in data:
				csvWriter(i, data.split('*注：',1)[0])
			# 判断是否有对话(中文符号)
			if "：“" in data and "”" in data:
				ds=data.split('：“',1)[0]
				ds_list.append(ds)
				dsa = data.split('：“',1)[-1].split('”',-1)[0]
				# if data_dict.__contains__(ds):
				# 	pass
				# else:
				# 	data_dict[ds] = dsa
				csvWriter(i, dsa)
			li.append(data)
			csvWriter(i, data)


# 判断是否存在csvTemp文件，如果不存在则新建
if not os.path.exists("csTemp"):
	os.mkdir("csTemp")
	os.chdir("csTemp")
	# if not os.path.exists("csTemp.csv"):
	# 	j=open("csTemp.csv","w", encoding="utf-8")
	# 	j.close()

# with open('csTemp/csTemp.csv', 'r', encoding="utf-8") as f:
# 	reader = csv.reader(f)
# 	if f:
# 		count = 4
		# count = int(list(reader)[0][0])
		
ds_list = []
wb = xlrd.open_workbook(r'sc.xlsx')
wbname = wb.sheet_names()
wbnum = []
for i in range(len(wbname)):
	wbnum.append(i)

csvNameList = []
for num,plotName in zip(wbnum, wbname):
	csvName('000{}_{}'.format(num,plotName))
	csvNameList.append("000{}_{}".format(num,plotName))


wbtlist = []
for sheetname in wbname:
	locals()["table"+str(sheetname)] = wb.sheet_by_name(sheetname)
	wbtlist.append(locals()["table"+str(sheetname)])


data_dict = {}
tw_list = []
for tablename, tw in zip(wbtlist, wbname):
	locals()[tw] = []
	tablename = wb.sheet_by_name(tw)
	for rowNum in range(tablename.nrows):
		rowVale = tablename.row_values(rowNum)
		for colNum in range(tablename.ncols):
			if rowNum > 0 and colNum == 0:
				isComment(rowVale[0],locals()[tw])
			else:
				isComment(rowVale[colNum],locals()[tw])
	tw_list.append(locals()[tw])
	data_dict[tw] = locals()[tw]
	print(len(locals()[tw]))

print(data_dict)
list_ds_list = list({}.fromkeys(ds_list).keys())
