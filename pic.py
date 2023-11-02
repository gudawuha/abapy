#***coding:UTF-8***
from abaqus import *
from odbAccess import *
from abaqusConstants import *
from viewerModules import *
import __main__
import displayGroupMdbToolset as dgm
import displayGroupOdbToolset as dgo
import os, glob

mname = getInput('odb:' )
odb=openOdb(path=mname, readOnly=True)
stepname=odb.steps.keys()[-1]
mystep=odb.steps['Step-2']
myframeNO=len(mystep.frames)
myframeNO1=range(myframeNO)
viewPort = session.Viewport(name='Viewport: 1') 
session.viewports['Viewport: 1'].makeCurrent()
session.viewports['Viewport: 1'].maximize()
viewPort.setValues(displayedObject=odb)
session.viewports['Viewport: 1'].view.setProjection(projection=PARALLEL)
session.viewports['Viewport: 1'].view.setValues(session.views['Bottom'])
session.pngOptions.setValues(imageSize=(4096,1792))
session.printOptions.setValues(reduceColors=False)
session.printToFile(fileName='C:\Users\Administrator\Desktop\damge3',format=PNG)
RootPath = 'Post-' + mname[0:-4]
RootPath = os.path.abspath(RootPath)
if not os.path.exists(RootPath):
    os.makedirs(RootPath)
os.chdir(RootPath) 
for i in range(0,9):
	session.viewports['Viewport: 1'].odbDisplay.setPrimaryVariable(
	variableLabel='DAMAGEC', outputPosition=INTEGRATION_POINT, )
	session.viewports['Viewport: 1'].odbDisplay.display.setValues(plotState=(
    CONTOURS_ON_DEF, ))
	session.viewports[session.currentViewportName].odbDisplay.setFrame(
	step='Step-2', frame=8*i)
	session.printOptions.setValues(rendition=COLOR,
	vpDecorations=OFF, vpBackground=OFF)
	session.pngOptions.setValues(imageSize=(4096, 4096))
	session.printToFile(fileName="%s frame=%s, time=%s" % (mname, 8*i, i), format=PNG, canvasObjects=(viewPort,))
