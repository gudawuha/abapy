#***coding:UTF-8***
from abaqus import *
from odbAccess import *
from abaqusConstants import *
from viewerModules import *
import __main__
import section
import regionToolset
import displayGroupMdbToolset as dgm
import part
import material
import assembly
import step
import interaction
import load
import mesh
import optimization
import job
import sketch
import visualization
import xyPlot
import displayGroupOdbToolset as dgo
import connectorBehavior
import numpy as np
import xlsxwriter
def GetShearData(fileName):
    result = { }
    liness = open(fileName, 'rt').readlines()
    lines = tuple(liness)
    story = 0
    frame = -1
    for line in lines:
        if line.startswith('Frame ='):
            frame = int(line.split(' ')[-1]) + 1
            if frame == 1:
                story = story + 1
                result[story] = { }
                continue
            continue
        
        if line.startswith('Resultant force ='):
            shear = map(float, line.split(' ')[-3:])
            result[story][frame] = shear
            continue
    
    return result

def GetMomentData(fileName):
    result = { }
    liness = open(fileName, 'rt').readlines()
    lines = tuple(liness)
    story = 0
    frame = -1
    for line in lines:
        if line.startswith('Frame ='):
            frame = int(line.split(' ')[-1]) + 1
            if frame == 1:
                story = story + 1
                result[story] = { }
                continue
            continue
        
        if line.startswith('Resultant moment ='):
            shear = map(float, line.split(' ')[-3:])
            result[story][frame] = shear
            continue
    
    return result

mname = getInput('odb: ')
odb=openOdb(path=mname, readOnly=True)
viewPort = session.Viewport(name='Viewport: 1')  
viewPort.setValues(displayedObject=odb)
stepname=odb.steps.keys()[-1]
mystep=odb.steps['Step-2']
myframeNO=len(mystep.frames)
myframeNO1=range(myframeNO)
xx1=[0 for i in myframeNO1]
nodeset=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT1']
nodeset0=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT2']
nodeset1=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT3']
nodeset2=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT4']
nodeset3=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT5']
nodeset4=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT6']
nodesetf=odb.rootAssembly.instances['SOLID-1'].nodeSets['NSET_LOAD_PT'].nodes[-1]
zuobiaox=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT1'].nodes[-1].coordinates[0]
zuobiaoxL=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT2'].nodes[-1].coordinates[0]
zuobiaoxM=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT3'].nodes[-1].coordinates[0]
zuobiaox1=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT4'].nodes[-1].coordinates[0]
zuobiaox2=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT5'].nodes[-1].coordinates[0]
zuobiaox3=odb.rootAssembly.instances['SOLID-1'].nodeSets['PT6'].nodes[-1].coordinates[0]
zbx=[zuobiaox, zuobiaoxL, zuobiaoxL, zuobiaox1, zuobiaox2, zuobiaox3]
L=abs(zuobiaox-zuobiaoxL)
stepseis = odb.steps.values()[-1]
time=[0 for i in myframeNO1]
ux=0
uxx=0
t = odb.name
tt = list(t)
tt[-1] = 's'
tt[-2] = 'l'
tt[-3] = 'x'
str = ''.join(tt)
wb = xlsxwriter.Workbook(str)
ws = wb.add_worksheet('tongji')
ws.write(0, 0, 'time')
ws.write(0, 1, 'weiyi')
ws.write(0, 2, 'f1z')
ws.write(0, 3, 'f2z')
ws.write(0, 4, 'f3z')
ws.write(0, 5, 'f4z')
ws.write(0, 6, 'f5z')
ws.write(0, 7, 'f6z')
ws.write(0, 8, 'm1y')
ws.write(0, 9, 'm2y')
ws.write(0, 10, 'm3y')
ws.write(0, 11, 'm4y')
ws.write(0, 12, 'm5y')
ws.write(0, 13, 'm6y')
ws.write(0, 14, 'rf')
ws.write(1, 2, 0)
ws.write(1, 3, 0.2)
ws.write(1, 4, 0.4)
ws.write(1, 5, 0.6)
ws.write(1, 6, 0.8)
ws.write(1, 7, 1)
ws.write(1, 8, 0)
ws.write(1, 9, 0.2)
ws.write(1, 10, 0.4)
ws.write(1, 11, 0.6)
ws.write(1, 12, 0.8)
ws.write(1, 13, 1)
leaf = dgo.LeafFromElementSets(elementSets=('LIANLIANGCZ', ))
session.viewports['Viewport: 1'].odbDisplay.displayGroup.replace(leaf=leaf)
session.viewports['Viewport: 1'].odbDisplay.setValues(viewCutNames = ('X-Plane',), viewCut = ON)
session.viewports['Viewport: 1'].odbDisplay.viewCuts['X-Plane'].setValues(showFreeBodyCut = True)
f = open('BodyCut.dat', 'w')
hhigh = tuple(zbx)
for h in hhigh:
    session.viewports['Viewport: 1'].odbDisplay.viewCuts['X-Plane'].setValues(position = float(h))
    path = odb.name
    for stepfram in range(0, len(stepseis.frames)):
        session.viewports['Viewport: 1'].odbDisplay.setFrame( step = 1  ,frame = stepfram)
        odb = session.odbs[path]
        session.writeFreeBodyReport(fileName = 'BodyCut.dat', append = ON, step = 1 , frame = stepfram, odb = odb)
r = GetShearData('BodyCut.dat')
for xx in myframeNO1:
		frames=mystep.frames[xx]
		time[xx]=mystep.frames[xx].frameValue
		displacement=frames.fieldOutputs['U']
		nodedisplacement=displacement.getSubset(region=nodesetf)
		nodevalues=nodedisplacement.values
		uu1=[]
		for v in nodevalues:
		  uu1.append(v.data[0])
		  ux=uu1[0]
		xx1[xx]=ux
xarray=np.array(xx1, dtype=float)
for k, v in r.items():
    for frm, shear in v.items():
        ws.write(frm+1, k+1, shear[2])
mon = GetMomentData('BodyCut.dat')
for k, v in mon.items():
    for frm, moent in v.items():
        ws.write(frm+1, k+7, moent[1])
for i in range(myframeNO):
	ws.write(i+2, 0, time[i])
	ws.write(i+2, 1, xarray[i])
#wb.filename = 'lianliang2.xlsm'
#wb.add_vba_project('C:/Temp/vbaProject.bin')
#chart = wb.add_chart({'type':'scatter'})
leaf = dgo.Leaf(leafType=DEFAULT_MODEL)
session.viewports['Viewport: 1'].odbDisplay.displayGroup.replace(leaf=leaf)
session.viewports['Viewport: 1'].odbDisplay.setValues(viewCutNames = ('Z-Plane',), viewCut = ON)
session.viewports['Viewport: 1'].odbDisplay.viewCuts['Z-Plane'].setValues(showFreeBodyCut = True)
f = open('RF.dat', 'w')
session.viewports['Viewport: 1'].odbDisplay.viewCuts['Z-Plane'].setValues(position = 0.01)
for stepfram in range(0, len(stepseis.frames)):
    session.viewports['Viewport: 1'].odbDisplay.setFrame( step = 1  ,frame = stepfram)
    odb = session.odbs[path]
    session.writeFreeBodyReport(fileName = 'RF.dat', append = ON, step = 1 , frame = stepfram, odb = odb)
rr = GetShearData('RF.dat')
for k, v in rr.items():
    for frm, shear in v.items():
        ws.write(frm+1, 14, shear[0])
chart = wb.add_chart({'type': 'scatter','subtype': 'smooth'})
chart.add_series({
    'name':       '=tongji!$O$1',
    'categories': '=tongji!$B$2:$B$1000',
    'values':     '=tongji!$O$2:$O$1000',
})
ws.insert_chart('U8',chart,{'x_offset':20,'y_offset':5})
chart1 = wb.add_chart({'type': 'scatter','subtype': 'smooth'})
chart1.add_series({
    'name':       '=tongji!$C$1',
    'categories': '=tongji!$C$2:$H$2',
    'values':     '=tongji!$C$15:$H$15',
})
ws.insert_chart('U28',chart1,{'x_offset':20,'y_offset':5})
wb.close()
