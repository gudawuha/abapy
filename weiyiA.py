# coding=utf-8
from abaqus import *
from abaqusConstants import *
from caeModules import *
from driverUtils import executeOnCaeStartup
from odbAccess import *
from  abaqus import session
import numpy as np
import sys
import os
import xlsxwriter
path = 'D:/abaqus/sxzdzfenduan2xaaVUSD'
opath = path + '.odb'
odb = session.openOdb(name=opath, readOnly=True)
session.viewports['Viewport: 1'].setValues(displayedObject=odb)
stepname=odb.steps.keys()[-1]
mystep=odb.steps['Accel']
myframeNO=len(mystep.frames) #frame数，等于二维数组行数
nodeset=odb.rootAssembly.instances['PART-1-1'].nodeSets.keys()
nodeset1=odb.rootAssembly.instances['PART-1-1'].nodeSets['NNA1']	#选择NNA，需要在part里建立点集
nodeNO1=len(nodeset1.nodes) #node数，等于二维数组列数
NNANO=0
cenggao=[1,6,4.5,3.27,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3]	#层高向量,第一个为1，总数和nodeNO1相同
for vv in nodeset:	#统计有多少包含NNA的结点集
	if vv[:3]=='NNA':
		NNANO=NNANO+1
matrix_NNAx = np.zeros((NNANO, myframeNO, nodeNO1))	#x向，每个时间的位移差，三维数组np.zeros(((二维数组个数, 二维数行数, 二维数列数)))
matrix_NNAy = np.zeros((NNANO, myframeNO, nodeNO1))	#y向
matrix_disp = np.zeros((2, NNANO, nodeNO1))	#三维数组,每个NNA的最大值，np.zeros(((二维数组个数, 二维数行数, 二维数列数)))，转置，方便计算
max_disp = np.zeros((nodeNO1, 2))	#最大位移角
for t in range(NNANO):
	NNFLOOR='PART-1-1.NNA'+str(t+1)
	xyListx = xyPlot.xyDataListFromField(odb=odb, outputPosition=NODAL, variable=((	#选择x向数据
		'U', NODAL, ((COMPONENT, 'U1'), )), ), nodeSets=(NNFLOOR, ))
	xyListy = xyPlot.xyDataListFromField(odb=odb, outputPosition=NODAL, variable=((	#选择y向数据
		'U', NODAL, ((COMPONENT, 'U2'), )), ), nodeSets=(NNFLOOR, ))
	my_arrayx = np.array(xyListx)	#x向，len(xyList1[1]行数，len(xyList1)列数,NNANO表示NNA的个数
	my_arrayy = np.array(xyListy)	#y向
	matrix_NNAx[t][:, 0] = np.abs(my_arrayx[0][:, 0])	#第一列时间，取第一个点的数据x向
	matrix_NNAy[t][:, 0] = np.abs(my_arrayy[0][:, 0])	#第一列时间，取第一个点的数据y向
	for i in range(1,nodeNO1):
		matrix_NNAx[t][:, i] = abs(my_arrayx[i][:, 1] - my_arrayx[i-1][:, 1])/cenggao[i]	#计算位移差x向
		matrix_NNAy[t][:, i] = abs(my_arrayy[i][:, 1] - my_arrayy[i-1][:, 1])/cenggao[i]	#y向
	for x in range(1, nodeNO1):	#计算所有frames下的最大位移角
		matrix_disp[0][t][x] = max(matrix_NNAx[t][:, x])	#[0]x向
		matrix_disp[1][t][x] = max(matrix_NNAy[t][:, x])	#[1]y向
for a in range(nodeNO1):	#计算NNA中最大的位移角
	max_disp[a][0] = max(matrix_disp[0][:, a])	#[0]x向位移角
	max_disp[a][1] = max(matrix_disp[1][:, a])	#[1]y向位移角
epath = path + 'A.xls'
wb = xlsxwriter.Workbook(epath) #生成excel
ws1 = wb.add_worksheet('tongji')	#生成sheet
ws2 = wb.add_worksheet('WEIYI')	#生成sheet
for ii in range(nodeNO1):	#写入excel
	ws1.write(ii, 0, ii)
	for jj in range(2):
		ws1.write(ii, jj+1, max_disp[ii][jj])
for ff in range(nodeNO1):	#写入excel:WEIYI
	ws2.write(ff, 0, ff)
for xx in range(2):
	for yy in range(NNANO):
		for zz in range(nodeNO1):
			ws2.write(zz, yy+xx*NNANO+1, matrix_disp[xx][yy][zz])	#转置输出
chart = wb.add_chart({'type': 'scatter','subtype': 'smooth'})
chart.set_x_axis({'name': u'Disp'})
chart.set_y_axis({'name': u'Floor'})
chart.set_legend({'none': True})
chart.add_series({		
	'categories': '=tongji!$a$1:$a$100',
	'values':     '=tongji!$b$1:$b$100',
})
chart.add_series({		
	'categories': '=tongji!$a$1:$a$100',
	'values':     '=tongji!$C$1:$C$100',
})
ws1.insert_chart('E8',chart,{'x_offset':20,'y_offset':5})
wb.close()	#保存