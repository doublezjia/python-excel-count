# 考勤表格筛选统计
# 环境 python3.5
# 需要模块 xlrd xlwt os sys
# xlrd 和 xlwt 是第三方的库，可以通过pip进行安装，分别用来读取Excel表格和生成Excel表格
# 使用时候，把源表格放到sourceXls中，然后运行脚本，就会自动筛选出来并生成表格放在destinationXls目录中。
# 注意：请保证sourceXls存在Excel文件，且表格格式与考勤表格相同。表格格式可参看templates中的Excel模板。
#       暂时脚本只能筛选一个表格的原因，所以不管在sourceXls放多少个文件，也只能筛选最新的表格。
#       所以最好保持sourceXls中只存在一个Excel文件。这个以后进行修改
# 版本：ver 1.0
# 最后更新时间 2017.08.21
#       

import xlrd,xlwt,os,sys





# 读取源表格
def open_Source_Xls(Sourfname):
	filename=Sourfname
	try:
		book = xlrd.open_workbook(filename)
	except PermissionError:
		print ('表格已经打开，请先关闭表格，然后再运行脚本')
		sys.exit()
	sheet1 = book.sheet_by_index(0)
	return sheet1


# 获取人名列表
def namelist(Sourfname):
	global name
	Sourfname = Sourfname
	sheet1=open_Source_Xls(Sourfname)
	for n in range(1,sheet1.nrows-1):
		m = sheet1.row_values(n)
		name.append(m[1])
	# 去除重复人名
	name = sorted(set(name),key=name.index)



# 目标表格写入
def Des_Xls_write(Sourfname,Desfname):
	
	# 声明变量
	rownum=1
	late = 0
	wtime = 0
	firtime = 0
	sedtime = 0
	thitime = 0
	leavenum = 0

	# 源表格的目录和名称
	Sourfname = os.path.join(sourcedir,Sourfname)
	# 生成的目标表格存放的目录和名称
	Desfname = os.path.join(destinationdir,Desfname)

	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	# 表格标题内容
	headlist= ["姓名","部门","迟到时长（小时/月）","工作时长（小时/月）",
	"19-21点下班次数","21-23点下班次数","23点后下班次数","备注"]
	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on');
	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)


	sheet1=open_Source_Xls(Sourfname)
	for i in range(1,sheet1.nrows-1):
		sourSheet = sheet1.row_values(i)
		# 如果行号跟name中的这个人名的下标不一样就把数字归零
		num = name.index(sourSheet[1])+1
		if rownum < num:
			late = 0
			wtime = 0
			firtime = 0
			sedtime = 0
			thitime = 0
			leavenum = 0

		# 让行号等于这个人名的name.index
		rownum = name.index(sourSheet[1])+1

		uname = sourSheet[1]
		# 如果数值为空则让其等于零，使得能够进行相加，下面同理		
		if sourSheet[8] == '':
			sourSheet[8] = 0
		late = late+sourSheet[8]
		if sourSheet[6] == '':
			sourSheet[6] = 0
		wtime = wtime+sourSheet[6]
		if sourSheet[10] == '':
			sourSheet[10] = 0
		firtime = firtime+sourSheet[10]
		if sourSheet[11] == '':
			sourSheet[11] = 0
		sedtime = sedtime+sourSheet[11]
		if sourSheet[12] == '':
			sourSheet[12] = 0
		thitime = thitime+sourSheet[12]


		#判断请假的次数
		if sourSheet[3] == '' and sourSheet[4] == '':
			leavenum = leavenum+1

		# 把值写入表格
		write_sheet.write(rownum,0,uname)
		write_sheet.write(rownum,2,late)
		write_sheet.write(rownum,3,round(wtime,1))
		write_sheet.write(rownum,4,firtime)
		write_sheet.write(rownum,5,sedtime)
		write_sheet.write(rownum,6,thitime)
		# 如果请假次数大于三的，进行登记
		if leavenum >= 3 :
			remark = '黄色空白'+str(leavenum)
			write_sheet.write(rownum,7,remark)

	# 保存表格
	desxls.save(Desfname)


# 查询源表格中的Excel文件，并获取文件名
def Des_filename(sourcedir):
	Sdfname = []
	os.chdir(sourcedir)
	for i in os.listdir('.'):
		if os.path.isfile(i):
			if os.path.splitext(i)[1] == '.xlsx':
				Sourfname = os.path.splitext(i)[0]+os.path.splitext(i)[1]
				desname=os.path.splitext(i)[0]+'统计.xls'
				Sdfname.append(Sourfname)
				Sdfname.append(desname)
				return Sdfname
			elif os.path.splitext(i)[1] == '.xls':
				Sourfname = os.path.splitext(i)[0]+os.path.splitext(i)[1]
				desname=os.path.splitext(i)[0]+'统计.xls'
				Sdfname.append(Sourfname)
				Sdfname.append(desname)
				return Sdfname
			else:
				print ('Excel的后缀应为 xls 或者 xlsx 请检查一下 sourceXls 文件夹中的文件后缀是否为正确.')
				sys.exit()
		elif not os.path.isfile(i):
			print ('请把要统计的考勤数据的Excel表格放进 sourceXls 的文件夹中。然后重新运行脚本.')
			sys.exit()




if __name__ == '__main__':
	# 存放源文件和目标文件的目录
	# 脚本根目录
	rootdir = os.path.abspath('.')
	# 源表格的目录
	sourcedir = os.path.join(rootdir,'sourceXls')
	# 生成的目标表格的目录
	destinationdir = os.path.join(rootdir,'destinationXls')
	# 用于存储人名列表
	name  = []
	tip = '''
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
考勤表格筛选统计 
说明：
使用前先把考勤表格放到 sourceXls 文件夹中，
然后再运行程序，避免出错。
运行后生成的文件放在 destinationXls 文件夹中。

注意：
请保证 sourceXls 文件夹中存在Excel文件，且表格格式与考勤表格相同。
表格格式可参看 templates 文件夹中的Excel模板。

因为暂时只能读取一个表格,
所以请保持 sourceXls 文件夹中只存在一个你要统计的Excel表格。
如果 sourceXls 文件夹中存在多个表格，也只会读取最新的表格。

因为脚本只统计Excel表格中的第一个表,
所以请把统计的数据放在Excel表格中的第一个表格。
要先把表格按名称排好顺序,否则统计有问题。

支持的Excel格式为 .xls 和 .xlsx

最后更新：ver1.0 2017-08-21
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'''
print (tip)
while True:
	Yoq = input('请输入 y 运行程序，退出请输入 q ：')
	if Yoq == 'y' or Yoq == 'Y' :
		try:
			Sdfnamelist = Des_filename(sourcedir)
			Sourfname = Sdfnamelist[0]
			Desfname = Sdfnamelist[1]
		except TypeError:
			print ('请把要统计的考勤数据的Excel表格放进 sourceXls 的文件夹中。然后重新运行脚本.')
			sys.exit()
		# 获取人名列表
		namelist(Sourfname)
		# 表格数据写入
		Des_Xls_write(Sourfname,Desfname)
		print ("表格统计成功，统计的文件存放在 destinationXls 文件夹中.")
		sys.exit()
	elif Yoq == 'q' or Yoq == 'Q':
		print ('退出脚本。')
		sys.exit()
	else:
		print ('请输入 y 执行脚本 或者输入 q 退出脚本。')
