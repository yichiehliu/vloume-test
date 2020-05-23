import csv
from pythonds.basic import Stack
import xlrd
import pandas as pd
import numpy as np

if __name__ == '__main__':
	s=Stack()
	filename_p = r"C:\Users\wii76_000\Desktop\VolumeTest\201904-202004.xlsx"
	pbook = xlrd.open_workbook(filename_p)
	price = pbook.sheet_by_index(0)
	print('sheel_1.name:',price.name,'sheel_1.nrows:',price.nrows,'sheel_1.ncols:',price.ncols)
	filename_v = r"C:\Users\wii76_000\Desktop\VolumeTest\201904-202004.xlsx"
	vbook = xlrd.open_workbook(filename_v)
	volume =  vbook.sheet_by_index(1)
	print('sheel_1.name:',volume.name,'sheel_1.nrows:',volume.nrows,'sheel_1.ncols:',volume.ncols)


	result = open(r'C:\Users\wii76_000\Desktop\VolumeTest\2020_invest_20200423.csv', 'a', newline = '\n',encoding = 'utf-8')

	for i in range(1, price.ncols):
		allinfo_p = []
		allinfo_v = []
		allresult = []
		tensumv = []
		for j in range(2,price.nrows):
			allinfo_p.append(price.cell_value(j,i))

		# print(allinfo_p)
		
		# for j in range(2,volume.nrows):
			allinfo_v.append(volume.cell_value(j,i))


		# print(allinfo_v)


		# for k in range(len(allinfo_v)-1, 2, -1):
			k = j-2
			# print("kkk",k)
			if (len(tensumv) == 10):
				if (allinfo_v[k] != '' and (float(allinfo_v[k]) > 2*sum(tensumv))):
					allresult.append([price.cell_value(j,0),price.cell_value(1,i),i,allinfo_p[k],allinfo_v[k]]) 
					print(1,i,j,allinfo_p[k],allinfo_v[k])
					x = pd.DataFrame()
					x['result'] = allresult
					x.to_csv(result,encoding = 'utf-8-sig')	
					allresult = []
				# 
				tensumv.pop()
			# print(allinfo_v[k])
			try:
				cvt = float(allinfo_v[k])
				tensumv.insert(0,cvt)
			except:
				pass
			# 	print(allinfo_v[k],tensumv)
				# tensumv.insert(0,"Blank")
				

			# cnt += 1
			# print("11m",tensumv)
			# print(allresult)
			
			# tensumv.append(allinfo_v[i])
		# print("Price",allinfo_p[i])
		# print("Volume",allinfo_v[i])


