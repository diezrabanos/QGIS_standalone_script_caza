import os
import xlrd


#import para procesar
import qgis.core as qgisCore
from qgis.core import QgsProject, QgsVectorLayer,QgsField,QgsExpression,QgsExpressionContext,QgsExpressionContextScope,QgsVectorFileWriter, QgsMarkerSymbol,QgsRendererCategory,QgsCategorizedSymbolRenderer,QgsPointXY, QgsPoint,QgsFeature,QgsGeometry,QgsLineSymbol,QgsExpressionContextUtils,QgsPalLayerSettings,QgsTextFormat,QgsVectorLayerSimpleLabeling,QgsExpressionContextUtils,QgsCoordinateTransform,QgsCoordinateReferenceSystem,QgsApplication,QgsProcessingFeedback
from qgis.PyQt.QtCore import QVariant,QDate, QTime, QDateTime, Qt
from qgis.utils import iface
from qgis.analysis import QgsNativeAlgorithms



import os
import glob
import re
import sys


import math
import time

import pyproj
import webbrowser

os.environ['PROJ_LIB'] = r'C:\Program Files\QGIS 3.10\apps\Python37\lib\site-packages\pyproj\proj_dir\share\proj'
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = r'C:\Program Files\QGIS 3.10\apps\Qt5\plugins'
os.environ['PATH'] += r'C:\Program Files\QGIS 3.10\apps\qgis-ltr\apps\qgis\bin;C:\Program Files\QGIS 3.10\apps\qgis-ltr\apps\Qt5\bin'

from qgis.core import *
print("1")
# Supply path to qgis install location
QgsApplication.setPrefixPath(r"C:\Program Files\QGIS 3.10\apps\qgis-ltr", True)
print("2")
# Create a reference to the QgsApplication.  Setting the
# second argument to False disables the GUI.
qgs = QgsApplication([], False)
print("3")
# Load providers
qgs.initQgis()
print("4")
# Append the path where processing plugin can be found
sys.path.append(r'C:/Program Files/QGIS 3.10/apps/qgis-ltr/python/plugins')
import processing
from processing.core.Processing import Processing
Processing.initialize()
QgsApplication.processingRegistry().addProvider(QgsNativeAlgorithms())

print ("5")
#print ("empiezo", file=open(r"c:/work/paraimprimir.txt", "a"))
print ("")
print ("")
print ("")


#defino las rutas de las capas con las que va a trabajar                                                                              #OJO ESTO LO TENDREIS QUE CAMBIAR EN OTRAS PROVINCIAS
rutacapademanchas=r"O:/sigmena/usuarios/caza_y_pesca/manchas 2020-2021/Manchas_2020-2021.shp"
columnafecha='Fecha_caz'
columnamatricula='P_Matricul'
rutacapademontes=r"O:/sigmena/carto/PROPIEDA/MONTES/PERTENEN/catalogo/42_mup_catalogo_ex_etrs89.shp"
rutatablalicencias=r"G:/Area_Comun_MN/Aprovechamientos/MUP_LICENCIA_CAZA_BASE_MARIO.xlsx"
hojalicencias='DATOS BASE MARIO'
columnamontes='MUP'

diasaconsiderar = int(input('Introducir cuantos dias consideramos desde hoy para hacer el analisis '))-1



#comienzo el analisis

#leo la tabla excel PARA VER que montes han pagado___________________________________________________________________________________
listado_montes_pagado=[]
#Abrimos el fichero excel
documento_excel = xlrd.open_workbook(rutatablalicencias)
hoja_de_interes = documento_excel.sheet_by_name(hojalicencias)#sheet_by_index(0)
#Leemos el numero de filas y columnas de la hoja de libros
filas = hoja_de_interes.nrows
columnas = hoja_de_interes.ncols
#print("el excel tiene " + str(filas) + " filas y " + str(columnas) + " columnas")
for i in range(1, hoja_de_interes.nrows): #Ignoramos la primera fila, que indica los campos
    #for j in range(libros.ncols):
    linea=float(repr(hoja_de_interes.cell_value(i,2)))#columna que tiene la etiqueta de los montes, con indice cero                    #OJO ESTO HAY QUE CAMBIAR EL NUMERO
    if str(repr(hoja_de_interes.cell_value(i,10))) > "":    #columna que tiene que tener algun dato                                    #OJO la columna 10 es la que contiene una fecha si se ha pagado, en caso contrario esta en blanco
        listado_montes_pagado.append(linea)
#print("listado_montes_pagado")
#print(listado_montes_pagado)

#leo la capa de manchas
vlmanchas=QgsVectorLayer(rutacapademanchas ,"Manchas","ogr")
#y la reparo
proceso1=processing.run("native:fixgeometries",{ 'INPUT' : rutacapademanchas, 'OUTPUT' : 'TEMPORARY_OUTPUT' })
vlmanchas=proceso1['OUTPUT']
#y la de montes
vlmontes=QgsVectorLayer(rutacapademontes ,"Montes","ogr")

#miro las fechas hoy y el dia hasta el que queremos considerar
from datetime import date
import datetime
today = date.today()
hoy = today.strftime("%Y-%m-%d")# YY-mm-dd
print("hoy=", hoy)
fechalimite=today+datetime.timedelta(days=diasaconsiderar)
print("fecha limite=", fechalimite)

#selecciono las manchas que tengo que estudiar por fecha        
vlmanchas.selectByExpression("\"{}\"".format(columnafecha)+" >= '{}' ".format(hoy)+" AND \"{}\"".format(columnafecha)+" <= '{}' ".format(fechalimite),QgsVectorLayer.SetSelection)
selection = vlmanchas.selectedFeatures()
QgsProject.instance().addMapLayer(vlmanchas)
listadecotosaestudiar=[]

#seleciono los montes que tocan con todas las manchas
processing.run('qgis:selectbylocation',{ 'INPUT' : vlmontes, 'INTERSECT' : vlmanchas, 'METHOD' : 0, 'PREDICATE' : [0] })
selection_montes = vlmontes.selectedFeatures()

#creo la web para los resultados
#crea el principio de la web
cabecera="""<!DOCTYPE HTML><html lang="es"><head>  <title>Resultado Analisis de Manchas y Montes de UP</title><meta charset="cp1252">  <meta name="description" content="Resultado del analisis de las manchas de caza"></head><body>"""
#crea la cabecera de la web
noticia0="""<h2>Licencias de aprovechamiento de caza en montes de utilidad publica </h2><p>Analisis de las manchas de caza que se cazaran hasta el %s. </p><p>En primer lugar se cruzan las manchas con la capa de montes de utilidad publica para conocer que porcentaje de la superficie de la mancha se encuentra dentro de monte de UP. </p><p>De esos montes analiza cruzandolo con la tabla de aprovechamientos si el titular del coto ha pagado la licencia para poder caza en el.</p><p>________________________________________________________ </p><p> </p><p>"""%fechalimite
web=cabecera+noticia0
#me quedo con una foto fija para ver si cambia porque haya montes sin licencia
web1=web


print ("")


listademontesaestudiar=[]
for feature in selection:
    #attrs = feature.attributes()
    #me quedo con el indice de la columna que contiene la matricula en la capa de manchas
    idx = feature.fieldNameIndex('P_Matricul')
    idx2=feature.fieldNameIndex(columnafecha)
    #print(feature.attributes()[idx])
    #me quedo con la geometria de las manchas
    geom_1 = feature.geometry()
    for fea in selection_montes:
        geom_montes = fea.geometry()
        #tengo que ver cuales de esas manchas interseccionan con montes de up.
        if geom_1.intersects(geom_montes) is True:
            geom3 = geom_1.intersection(geom_montes)
            #calculo esta superficie y su porcentaje de geom3 para ver si es despreciable
            area_ratio = 100.0*geom3.area() / geom_1.area()
            
            #saco las columnas
            n_mon=fea["Etiqueta"]#etiqueta es la columna de la capa de montes con el numero de los montes
            matricula = feature.attribute(columnamatricula)
            #veo si esta en la lista de los que han pagado
            if float(n_mon) not in listado_montes_pagado:
                #preparo la linea de texto para presentar en la web
                print(" El "+ str(round(area_ratio,2))+" % de la mancha del Coto "+feature.attributes()[idx]+" se encuentra en el MUP "+str(n_mon) +" que no tiene licencia")
                web=web+("""<p> El %s  de la mancha del Coto %s que se caza el %s se encuentra en el MUP %s que no tiene licencia.</p> """%(str(round(area_ratio,2))+" %" ,feature.attributes()[idx],feature.attributes()[idx2].toString(Qt.ISODate),str(n_mon) ))
            
                
            listadecotosaestudiar.append(matricula)
            listademontesaestudiar.append(n_mon)
#la foto fija una vez analizadas todas las manchas
web2=web
if web1==web2:
    web=web+"<p> No existen manchas dentro de montes que no tengan licencia</p>"

#crea el final de la web
final="""</body></html>"""
web=web+final

#creo el archivo para la web

ruta='licencias.html'#'R:/SIGMENA/prueba/2021/01/12/licencias.html'
sys.stdout=open(ruta, 'w')
print(web)
sys.stdout.close()


#abro el archivo
import webbrowser
new = 2 # open in a new tab, if possible
"""# open a public URL, in this case, the webbrowser docs
#url = "http://docs.python.org/library/webbrowser.html"
#webbrowser.open(url,new=new)"""
#// open an HTML file on my own (Windows) computer
url = ruta#"file://"+ruta#"file://d/testdata.html"
webbrowser.open(url,new=new)    

#cierro la aplicacion                             
qgs.exitQgis()
