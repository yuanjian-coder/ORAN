import xlrd
from collections import OrderedDict
import json
import codecs
from time import *
import os
import json
import tablib
import sys

def file_to_json(file_name,num):
	wb = xlrd.open_workbook(file_name+'.xlsx')
	convert_list = []
	sh = wb.sheet_by_index(0)
	title = sh.row_values(0)	
	for rownum in range(1, sh.nrows):
		rowvalue = sh.row_values(rownum)
		single = OrderedDict()
		if(rowvalue[2]=='int'):
			single[rowvalue[0]] =int(rowvalue[1])
		else:
			single[rowvalue[0]] =rowvalue[1]
		convert_list.append(single)
	if(num==1):
		print('Size before encoding:'+str(sys.getsizeof(convert_list))+'B')
	begin_time = time()
	j = json.dumps(convert_list)
	end_time=time() 
	if(num==1):
		print('Size after encoding:'+str(sys.getsizeof(j))+'B')
	with codecs.open(file_name+'.json',"w","utf-8") as f:
	    f.write(j)
	run_time = end_time-begin_time
	return run_time

def file_to_excel(file_name,num):
	begin_time = time()
	with open(file_name, 'r',encoding='utf-8',errors='ignore') as f:
	   rows = json.load(f)
	end_time=time()
	head=[]
	head.append('name')
	head.append('values')
	header=tuple(head)
	data = []
	for row in rows:
	    body = []
	    for v in row.keys():
	        body.append(v)
	    for k in row.values():
	    	body.append(k)
	    data.append(tuple(body))
	data = tablib.Dataset(*data,headers=header)
	run_time = end_time-begin_time	   
	return run_time

time1=0
time2=0
for i in range(10000):
	time1+=file_to_json('Table1. test message content-20 fields',i)
for i in range(10000):
	time2+=file_to_excel('Table1. test message content-20 fields.json',i)
print('20 fields JSON encoding time：'+str(time1/10000*1000)+'ms')
print('20 fields JSON decoding time：'+str(time2/10000*1000)+'ms')
print('----------')

time3=0
time4=0
for i in range(10000):
	time3+=file_to_json('Table2. test message content-100 fields',i)
for i in range(10000):
	time4+=file_to_excel('Table2. test message content-100 fields.json',i)
print('100 fields JSON encoding time：'+str(time3/10000*1000)+'ms')
print('100 fields JSON encoding time：'+str(time4/10000*1000)+'ms')
print('----------')

time5=0
time6=0
for i in range(10000):
	time5+=file_to_json('Table3. test message content-500 fields',i)
for i in range(10000):
	time6+=file_to_excel('Table3. test message content-500 fields.json',i)
print('500 fields JSON encoding time：'+str(time5/10000*1000)+'ms')
print('500 fields JSON encoding time：'+str(time6/10000*1000)+'ms')
