import win32com.client #le decimos que usamos catia
from pycatia import catia
from sys import exit    #para que exit tenga sentido en el codigo

#get catia

CATIA = win32com.client.Dispatch("catia.Application")


############    ABRE UN DOCUMENTO EXISTENTE #################


## Ruta del archivo que deseas abrir
# ruta_archivo = r"C:\Users\Probook\Desktop\AERO\TFG\python\ejercicios\pruebaPython.CATDrawing"

## Intenta abrir el archivo
# try:
#     documento = CATIA.Documents.Open(ruta_archivo)
#     print("El archivo se ha abierto correctamente")
# except:
#     print("No se pudo abrir el archivo")


# Comprueba si hay un documento activo

## if CATIA.ActiveDocument is not None:
#     documento = CATIA.ActiveDocument
#     print("Hay un documento de CATIA abierto")
# else:
#     print("No hay un documento de CATIA abierto")

##  Cierra el documento activo
# if CATIA.ActiveDocument is not None:   
#     CATIA.ActiveDocument.Close()
#     print("El documento activo ha sido cerrado")
# else:
#     print("No hay un documento de CATIA abierto")



##############################################################

if CATIA.ActiveDocument is not None:
    documento = CATIA.ActiveDocument
    # Obtiene el nombre del documento activo
    activeDocName = documento.Name
    print("El nombre del documento activo es:", activeDocName)
else:
    print("No hay un documento de CATIA abierto")

# print(CATIA.ActiveDocument.name)
# activeDocName = CATIA.activedocument.name

if ".CATDrawing" not in activeDocName:
    print("Not a drawing open")
    exit()
else:
    print("a drawing is open")

#elementos seleccionados

oSel = documento.Selection
print("elementos seleccionados =", oSel.count)

#list for parts 

partlist = []

dataview = {}      #crea un diccionario

if oSel.count > 0:
    for i in range(oSel.count):
        selection = oSel.item(i+1)
        print(f'name: {selection.value.name} type: {selection.type}')
        typeDoc = selection.type
        if typeDoc == "DrawingView":
            view = oSel.item(i+1).value
            # parentLink = view.generativelinks.firstlink().parent.name
            # partlist.append(parentLink)
            # print(f"{view.name}: link: {parentLink}")

            x1 = float()
            y1 = float()
            z1 = float()
            x2 = float()
            y2 = float()
            z2 = float()

            x1,y1,z1,x2,y2,z2 = view.generativebehavior.GetProjectionPlane()

            print(x1,y1,z1,x2,y2,z2)

            dataview[view.name] = (x1,y1,z1,x2,y2,z2)
            print(dataview)

#grab CatPart

print(partlist) 

#thispart = CATIA.documents.item(partlist[0]).Part
# thispart = CATIA.documents.item("DrawingView").part
thispart = CATIA.ActiveDocument.item("DrawingView").part

#create the geometrical set

geoSet = thispart.Hybridbodies.add()
geoSet.name = "export plane from 2D"

hsb = thispart.HybridShapeFactory

#create first point

pointMain = hsb.AddNewpointCoord(0,0,0)
pointMain.name = "Origin"
geoSet.AppendHybridShape(pointMain)

ref = thispart.CreateReferenceFromObject(pointMain)

thispart.update()       #se supone que hasta aqui crea un punto en el part como geometrical set

for item,value in dataview.items():

    #create direction reference
    dirByCoord1 = hsb.AddNewDirectionByCoord(value[0],value[1],value[2])
    dirByCoord2 = hsb.AddNewDirectionByCoord(value[3],value[4],value[5])

    #create lines
    LinepointDir1 = hsb.AddNewLinePtDir(ref, dirByCoord1, 0.0, 35, False)
    LinepointDir1.name = item + "X"
    geoSet.AppendHybridShape(LinepointDir1)

    LinepointDir2 = hsb.AddNewLinePtDir(ref, dirByCoord2, 0.0, 35, False)
    LinepointDir2.name = item + "Y"
    geoSet.AppendHybridShape(LinepointDir2)


    #create plane
    planeLine = hsb.AddNewPlane2Lines(LinepointDir1,LinepointDir2)
    planeLine.name = item
    geoSet.AppendHybridShape(planeLine)


    thispart.update()


#para crear una aplicion (exe) de nuestro codigo vamos a la terminal y ponemos Auto-py-to-exe y seleccionamos el archivo .py
#seleccionamos la carpeta donde queremos que se ubique y le damos a "one folder" para q solo nos guarde un .exe y no 
#un porron de documentos y mierdas innecesarias