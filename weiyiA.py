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
myframeNO=len(mystep.frames) #frame�������ڶ�ά��������
nodeset=odb.rootAssembly.instances['PART-1-1'].nodeSets.keys()
nodeset1=odb.rootAssembly.instances['PART-1-1'].nodeSets['NNA1']	#ѡ��NNA����Ҫ��part�ｨ���㼯
nodeNO1=len(nodeset1.nodes) #node�������ڶ�ά��������
NNANO=0
cenggao=[1,6,4.5,3.27,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3]	#�������,��һ��Ϊ1��������nodeNO1��ͬ
for vv in nodeset:	#ͳ���ж��ٰ���NNA�Ľ�㼯
	if vv[:3]=='NNA':
		NNANO=NNANO+1
matrix_NNAx = np.zeros((NNANO, myframeNO, nodeNO1))	#x��ÿ��ʱ���λ�Ʋ��ά����np.zeros(((��ά�������, ��ά������, ��ά������)))
matrix_NNAy = np.zeros((NNANO, myframeNO, nodeNO1))	#y��
matrix_disp = np.zeros((2, NNANO, nodeNO1))	#��ά����,ÿ��NNA�����ֵ��np.zeros(((��ά�������, ��ά������, ��ά������)))��ת�ã��������
max_disp = np.zeros((nodeNO1, 2))	#���λ�ƽ�
for t in range(NNANO):
	NNFLOOR='PART-1-1.NNA'+str(t+1)
	xyListx = xyPlot.xyDataListFromField(odb=odb, outputPosition=NODAL, variable=((	#ѡ��x������
		'U', NODAL, ((COMPONENT, 'U1'), )), ), nodeSets=(NNFLOOR, ))
	xyListy = xyPlot.xyDataListFromField(odb=odb, outputPosition=NODAL, variable=((	#ѡ��y������
		'U', NODAL, ((COMPONENT, 'U2'), )), ), nodeSets=(NNFLOOR, ))
	my_arrayx = np.array(xyListx)	#x��len(xyList1[1]������len(xyList1)����,NNANO��ʾNNA�ĸ���
	my_arrayy = np.array(xyListy)	#y��
	matrix_NNAx[t][:, 0] = np.abs(my_arrayx[0][:, 0])	#��һ��ʱ�䣬ȡ��һ���������x��
	matrix_NNAy[t][:, 0] = np.abs(my_arrayy[0][:, 0])	#��һ��ʱ�䣬ȡ��һ���������y��
	for i in range(1,nodeNO1):
		matrix_NNAx[t][:, i] = abs(my_arrayx[i][:, 1] - my_arrayx[i-1][:, 1])/cenggao[i]	#����λ�Ʋ�x��
		matrix_NNAy[t][:, i] = abs(my_arrayy[i][:, 1] - my_arrayy[i-1][:, 1])/cenggao[i]	#y��
	for x in range(1, nodeNO1):	#��������frames�µ����λ�ƽ�
		matrix_disp[0][t][x] = max(matrix_NNAx[t][:, x])	#[0]x��
		matrix_disp[1][t][x] = max(matrix_NNAy[t][:, x])	#[1]y��
for a in range(nodeNO1):	#����NNA������λ�ƽ�
	max_disp[a][0] = max(matrix_disp[0][:, a])	#[0]x��λ�ƽ�
	max_disp[a][1] = max(matrix_disp[1][:, a])	#[1]y��λ�ƽ�
epath = path + 'A.xls'
wb = xlsxwriter.Workbook(epath) #����excel
ws1 = wb.add_worksheet('tongji')	#����sheet
ws2 = wb.add_worksheet('WEIYI')	#����sheet
for ii in range(nodeNO1):	#д��excel
	ws1.write(ii, 0, ii)
	for jj in range(2):
		ws1.write(ii, jj+1, max_disp[ii][jj])
for ff in range(nodeNO1):	#д��excel:WEIYI
	ws2.write(ff, 0, ff)
for xx in range(2):
	for yy in range(NNANO):
		for zz in range(nodeNO1):
			ws2.write(zz, yy+xx*NNANO+1, matrix_disp[xx][yy][zz])	#ת�����
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
wb.close()	#����