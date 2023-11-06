#***coding:UTF-8***
from abaqus import *
from odbAccess import *
from abaqusConstants import *
from viewerModules import *
import displayGroupOdbToolset as dgo
import os

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
	vpDecorations=OFF, vpBackground=OFF, reduceColors=False)
	session.pngOptions.setValues(imageSize=(4096, 4096))
	session.printToFile(fileName="DAMAGEC% sframe=%s, time=%s" % (mname[0:-4], 8*i, i), format=PNG, canvasObjects=(viewPort,))
session.viewports['Viewport: 1'].view.setValues(session.views['Bottom'])
for i in range(0,9):
	leaf = dgo.LeafFromElementSets(elementSets=("BAR-1.SET_ALL_BAR", ))
	session.viewports['Viewport: 1'].odbDisplay.displayGroup.replace(leaf=leaf)
	session.viewports['Viewport: 1'].odbDisplay.setPrimaryVariable(
	variableLabel='S', outputPosition=INTEGRATION_POINT, refinement=(INVARIANT, 
	'Mises'), )
	session.viewports['Viewport: 1'].odbDisplay.display.setValues(plotState=(
	CONTOURS_ON_DEF, ))
	session.viewports[session.currentViewportName].odbDisplay.setFrame(
	step='Step-2', frame=8*i)
	session.printOptions.setValues(rendition=COLOR,
	vpDecorations=OFF, vpBackground=OFF, reduceColors=False)
	session.pngOptions.setValues(imageSize=(4096, 4096))
	session.printToFile(fileName="S% sframe=%s, time=%s" % (mname[0:-4], 8*i, i), format=PNG, canvasObjects=(viewPort,))
session.viewports['Viewport: 1'].view.setValues(session.views['Bottom'])
for i in range(0,9):
	leaf = dgo.LeafFromElementSets(elementSets=("BAR-1.SET_ALL_BAR", ))
	session.viewports['Viewport: 1'].odbDisplay.displayGroup.replace(leaf=leaf)
	session.viewports['Viewport: 1'].odbDisplay.setPrimaryVariable(
	variableLabel='PE', outputPosition=INTEGRATION_POINT, refinement=(
	INVARIANT, 'Max. In-Plane Principal'), )
	session.viewports['Viewport: 1'].odbDisplay.display.setValues(plotState=(
	CONTOURS_ON_DEF, ))
	session.viewports[session.currentViewportName].odbDisplay.setFrame(
	step='Step-2', frame=8*i)
	session.printOptions.setValues(rendition=COLOR,
	vpDecorations=OFF, vpBackground=OFF, reduceColors=False)
	session.pngOptions.setValues(imageSize=(4096, 4096))
	session.printToFile(fileName="PE% sframe=%s, time=%s" % (mname[0:-4], 8*i, i), format=PNG, canvasObjects=(viewPort,))
