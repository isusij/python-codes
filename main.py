import win32com.client.dynamic # Module for COM­Client
import os   # Module for File­Handling
import pandas as pd    #Module for creating data frames
import numpy as np 
import pyautogui
import time

print('-----new execution-----')

CATIA = win32com.client.Dispatch("CATIA.Application")

# *******************************************************************************
#           LOS SIGUIENTES PARAMETROS DEBEN SER DEFINIDOS POR EL USUARIO
#*******************************************************************************

# Ruta de la carpeta que contiene el/los archivo/s CAD con los que se quiere trabajar
directory_path = "C:\\Users\\Probook\\Desktop\\AERO\\TFG\\python\\CADs_laminados"

# Nombre de la carpeta donde se guardaran los resultados
new_folder_name = "Resultados"

# Ancho de la cinta unidireccional de la que se dispone para que se tracen los patrones correctamente
tape_width = 150.0  # en mm

# *******************************************************************************

def main(CAD_file):
    """"
    Función principal, realiza todo el proceso de obtencion de los patrones de corte de una pieza a partir del archivo CAD
    - Inputs
        - CAD_file [file]: archivo CAD de la pieza de la que se van a obtener los patrones de corte
    - Output
    """

    def flattening(ply_number):
        """"
        Hace el flattening de una capa del laminado de la pieza
        - Inputs
            - ply_number [int]: numero de la capa de la que se hace el flattening
        - Output
        """
        
        def click_on_img(img, confidence_number,espera):
            """
            Hace click sobre un icono
            - Inputs
                - img [string]: icono a seleccionar guardado como archivo .png
                - confidence_number [long]: porcentaje de similitud que debe encontrar con la imagen
                - espera [long]: tiempo de espera antes de ejecutar la función
            - Output
            """
            time.sleep(espera)

            img_location = pyautogui.locateCenterOnScreen(img, confidence=confidence_number)
            img_X = img_location[0]
            img_Y = img_location[1]

            pyautogui.click(img_X,img_Y)

        def doubleClick_on_img(img, confidence_number, espera):
            """
            Hace doble click sobre un icono
            - Inputs
                - img [string]: icono a seleccionar guardado como archivo .png
                - confidence_number [long]: porcentaje de similitud que debe encontrar con la imagen
                - espera [long]: tiempo de espera antes de ejecutar la función
            - Output
            """
            time.sleep(espera)

            img_location = pyautogui.locateCenterOnScreen(img, confidence=confidence_number)
            img_X = img_location[0]
            img_Y = img_location[1]

            pyautogui.doubleClick(img_X,img_Y)

        def ply_selection():
            """
            Selecciona una capa para trabajar con ella
            - Inputs
            - Output
            """
            selection = partDocument1.Selection
            selection.Clear()

            # Ruta para llegar a las capas de la pieza
            partHBs = part1.HybridBodies
            stackHB = partHBs.Item("Stacking")
            stackHBs = stackHB.HybridBodies
            pliesHB = stackHBs.Item("Plies Group.1")
            pliesHBs = pliesHB.HybridBodies
            sequenceHB = pliesHBs.Item(f"Sequence.{ply_number}")
            sequenceHBs = sequenceHB.HybridBodies
            plyHB = sequenceHBs.Item(f"Ply.{ply_number}")

            selection = partDocument1.Selection
            selection.Add(plyHB)    # selecciona la capa

        #*********COMIEZO DEL PROCESO PARA UNA CAPA*************

        ply_selection()

        doubleClick_on_img('flattening_icon.png', 0.9, 0.0)

        click_on_img('xy_plane.png', 0.6, 1)

        click_on_img('ok_button.png', 0.8, 1)

        part1.Update()

    def cinta_UD(ply_number):
        """
        Saca los patrones de corte de cinta unidireccional de la pieza a partir del flattening
        - Inputs
            - ply_number [int]: capa de la que se va a sacar el patron de corte
        - Output
        """ 

        def create_sketch(sketch_name):
            """"
            Crea un nuevo sketch en el plano xy
            - Inputs
                - sketch_name [string]: nombre del nuevo sketch que queremos crear
            - Output
                - new_sketch [geometric element]: sketch resultante
            """

            partHBs = part1.HybridBodies
            flattGeoSet = partHBs.Item("Flattening_Geometry")   
            part1.InWorkObject = flattGeoSet

            sketches1 = flattGeoSet.HybridSketches
            XYPlane = part1.OriginElements.PlaneXY
            new_sketch = sketches1.Add(XYPlane)   # Se crea el sketch en el plano XY 
            new_sketch.Name = sketch_name

            part1.Update()

            return new_sketch

        def flatten_projection():
            """
            Saca la proyeccion del flattening de una capa en un sketch a parte
            - Inputs
            - Output
                - projection [geometric element]: proyeccion del flattening en un sketch
            """

            part1.InWorkObject = sketch
            factory2D1 = sketch.OpenEdition()

            # Ruta en el arbol de CATIA para llegar hasta el contorno que queremos proyectar
            pliesHBs = pliesHB.HybridBodies
            sequenceHB = pliesHBs.Item(f"Sequence.{ply_number}")
            sequenceHBs = sequenceHB.HybridBodies
            plyHB = sequenceHBs.Item(f"Ply.{ply_number}")
            plyHBs = plyHB.HybridBodies
            flattBodyHB = plyHBs.Item("Flatten Body")
            flattBodyHBs = flattBodyHB.HybridBodies
            flatteningHB = flattBodyHBs.Item("Flattening")
            flatteningHBs = flatteningHB.HybridSketches
            sketch_contour = flatteningHBs.Item("Sketch.FlattenContour.1")

            reference1 = part1.CreateReferenceFromObject(sketch_contour)
            Projection = factory2D1.CreateProjections(reference1)

            sketch.CloseEdition()
            part1.InWorkObject = flattGeoSet

            part1.Update()

            return Projection

        def limits_vertical(x, y1, y2, line_name):
            """
            Dibuja una linea vertical a una distancia de otra dada
            - Inputs
                - x [float]: coordenada x de la linea a dibujar
                - y1 [float]: coordenada y del punto 1
                - y2 [float]: coordenada y del punto 2
                - line_name [string]: nombre de la linea en CATIA
            - Output  
                - line2D1 [geometric element]: la linea creada
                - startPoint [float]: extremo 1 de la linea creada
                - endPoint [float]: extreo 2 de la linea creada
            """

            part1.InWorkObject = sketch

            factory2D1 = sketch.OpenEdition()

            # Se crea la linea
            startPoint = factory2D1.CreatePoint(x, y1)
            endPoint = factory2D1.CreatePoint(x, y2)
            line2D1 = factory2D1.CreateLine(x, y1, x, y2)
            line2D1.StartPoint = startPoint
            line2D1.EndPoint = endPoint
            line2D1.Name = line_name

            # Se impone la condicion de verticalidad
            constraints1 = sketch.Constraints
            geometricElements = sketch.GeometricElements
            reference1 = part1.CreateReferenceFromObject(line2D1)
            axis2D1 = geometricElements.Item("AbsoluteAxis")
            line2D2 = axis2D1.GetItem("VDirection")
            reference2 = part1.CreateReferenceFromObject(line2D2)
            constraint1 = constraints1.AddBiEltCst(13, reference1, reference2)  #13=catCstTypeVerticality
            constraint1.Mode = 0    #0=catCstModeDrivingDimension, 1=catCstModeDrivenDimension

            # Se impone la condicion de que mantenga una distancia de 10mm con el flattening
            geometry2D1 = projection.Item("Mark.1")
            reference3 = part1.CreateReferenceFromObject(geometry2D1)
            constraint2 = constraints1.AddBiEltCst(1, reference1, reference3)   #1=catCstTypeDistance
            constraint2.Mode = 0    #0=catCstModeDrivingDimension, 1=catCstModeDrivenDimension
            length1 = constraint2.Dimension
            length1.Value = 10.0
            
            sketch.CloseEdition()
            part1.InWorkObject = flattGeoSet

            part1.Update()

            return line2D1, startPoint, endPoint

        def limits_hztal(y, x1, x2, line_name):
            """
            Dibuja una linea horizontal a una distancia de otra dada
            - Inputs
                - y [float]: coordenada y de la linea a dibujar
                - x1 [float]: coordenada x del punto 1
                - x2 [float]: coordenada x del punto 2
                - line_name [string]: nombre de la linea en CATIA
            - Output  
                - line2D1 [geometric element]: la linea creada
                - startPoint [object]: extreno 1 de la linea creada
                - endPoint [object]: extremo 2 de la linea creada
            """

            part1.InWorkObject = sketch
            factory2D1 = sketch.OpenEdition()

            # Se crea la linea
            startPoint = factory2D1.CreatePoint(x1, y)
            endPoint = factory2D1.CreatePoint(x2, y)
            line2D1 = factory2D1.CreateLine(x1, y, x2, y)
            line2D1.StartPoint = startPoint
            line2D1.EndPoint = endPoint
            line2D1.Name = line_name

            # Se impone la condicion de horizontalidad
            constraints1 = sketch.Constraints
            geometricElements = sketch.GeometricElements
            reference1 = part1.CreateReferenceFromObject(line2D1)
            axis2D1 = geometricElements.Item("AbsoluteAxis")
            line2D2 = axis2D1.GetItem("HDirection")
            reference2 = part1.CreateReferenceFromObject(line2D2)
            constraint1 = constraints1.AddBiEltCst(10, reference1, reference2)  #catCstTypeHorizontality
            constraint1.Mode = 0    #catCstModeDrivingDimension

            # Se impone la condicion de que mantenga una distancia de 10mm con el flattening
            geometry2D1 = projection.Item("Mark.1")
            reference3 = part1.CreateReferenceFromObject(geometry2D1)
            constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance
            constraint4.Mode = 0    #catCstModeDrivingDimension
            length1 = constraint4.Dimension
            length1.Value = 10.0

            sketch.CloseEdition()
            part1.InWorkObject = flattGeoSet

            part1.Update()

            return line2D1, startPoint, endPoint

        def point_coincidence(point1, point2):
            """
            Hace coincidir dos puntos
            - Inputs
                - point1 [float]: primer punto
                - point2 [float]: segundo punto
            - Output 
            """
                
            part1.InWorkObject = sketch
            sketch.OpenEdition()

            constraints1 = sketch.Constraints
            reference1 = part1.CreateReferenceFromObject(point1)
            reference2 = part1.CreateReferenceFromObject(point2)
            constraint1 = constraints1.AddBiEltCst(2, reference1, reference2)    #catCstTypeOn=2
            constraint1.Mode = 0    #catCstModeDrivingDimension

            sketch.CloseEdition()
            part1.InWorkObject = flattGeoSet

            part1.Update()

        def measure(element):
            """"
            Devuelve la magnitud del elemento pedido
            - Inputs
                - element [geometry]: elemento que se quiere dimensionar
            - Output
                - length [long]: longitud del elemento medido
            """

            theSPAWorkbench = partDocument1.GetWorkbench("SPAWorkbench")
            measurableObject = theSPAWorkbench.GetMeasurable(element)
            length = measurableObject.Length

            return length

        def first_line_UD(y, x1, x2):
            """
            Dibuja la primera linea de cinta
            - Inputs
                - y [float]: coordenada y de la linea a dibujar
                - x1 [float]: coordenada x del punto 1
                - x2 [float]: coordenada x del punto 2
            - Output  
                - line2D1 [geometric element]: la linea creada
            """

            part1.InWorkObject = sketch
            factory2D1 = sketch.OpenEdition()

            # Se crea la linea
            startPoint = factory2D1.CreatePoint(x1, y)
            endPoint = factory2D1.CreatePoint(x2, y)
            line2D1 = factory2D1.CreateLine(x1, y, x2, y)
            line2D1.StartPoint = startPoint
            line2D1.EndPoint = endPoint
            line2D1.Name = "first"

            # Se impone la condicion de horizontalidad
            constraints1 = sketch.Constraints
            geometricElements = sketch.GeometricElements
            reference1 = part1.CreateReferenceFromObject(line2D1)
            axis2D1 = geometricElements.Item("AbsoluteAxis")
            line2D2 = axis2D1.GetItem("HDirection")
            reference2 = part1.CreateReferenceFromObject(line2D2)
            constraint1 = constraints1.AddBiEltCst(10, reference1, reference2)  #10=catCstTypeHorizontality
            constraint1.Mode = 0    #catCstModeDrivingDimension

            # Se impone la condicion de que coincida con el limite del flattening
            reference3 = part1.CreateReferenceFromObject(linea_10inf)
            constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #1=catCstTypeDistance
            constraint4.Mode = 0    #catCstModeDrivingDimension
            length1 = constraint4.Dimension
            length1.Value = 0.0

            sketch.CloseEdition()
            part1.InWorkObject = flattGeoSet

            part1.Update()

            return line2D1

        def lines_UD(y, x1, x2):
            """
            Dibuja una linea vertical a una distancia de otra dada
            - Inputs
                - y [float]: coordenada y de la linea a dibujar
                - x1 [float]: coordenada x del punto 1
                - x2 [float]: coordenada x del punto 2
            - Output  
                - line2D1 [geometric element]: la linea creada
            """

            part1.InWorkObject = sketch
            factory2D1 = sketch.OpenEdition()

            # Se crea la linea
            startPoint = factory2D1.CreatePoint(x1, y)
            endPoint = factory2D1.CreatePoint(x2, y)
            line2D1 = factory2D1.CreateLine(x1, y, x2, y)
            line2D1.StartPoint = startPoint
            line2D1.EndPoint = endPoint

            # Se impone la condicion de horizontalidad
            constraints1 = sketch.Constraints
            geometricElements = sketch.GeometricElements
            reference1 = part1.CreateReferenceFromObject(line2D1)
            axis2D1 = geometricElements.Item("AbsoluteAxis")
            line2D2 = axis2D1.GetItem("HDirection")
            reference2 = part1.CreateReferenceFromObject(line2D2)
            constraint1 = constraints1.AddBiEltCst(10, reference1, reference2)  #catCstTypeHorizontality
            constraint1.Mode = 0    #catCstModeDrivingDimension

            # Se impone la condicion de que las lineas que representan los cortes esten separadas 150mm (ancho de la cinta)
            reference3 = part1.CreateReferenceFromObject(linea)
            constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance
            constraint4.Mode = 0    #catCstModeDrivingDimension
            length1 = constraint4.Dimension
            length1.Value = tape_width

            sketch.CloseEdition()
            part1.InWorkObject = flattGeoSet

            part1.Update()

            return line2D1

        def hide_sketch():
            """"
            Oculta el sketch seleccionado
            - Inputs
            - Output
            """

            selection1 = partDocument1.Selection
            visPropertySet1 = selection1.VisPropertie
            selection1.Add(sketch)     
            visPropertySet1.SetShow(1)      # 0=show    1=hide
            selection1.Clear()

            part1.Update()

        def create_drawing():
            """"
            Crea un drawing del patron de corte ya diseñado
            - Inputs
            - Output
            """

            # Se crea una nueva hoja dentro del Drawing
            drawingSheets1 = drawingDocument1.Sheets
            drawingSheet1 = drawingSheets1.Add("New Sheet")
            drawingSheet1.Name = f"ply{ply_number}"
            drawingSheet1.Activate()
            drawingSheet1.Scale = 1.0
            drawingViews1 = drawingSheet1.Views
            drawingView1 = drawingViews1.Add("AutomaticNaming")
            drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior

            # Se obtiene la vista frontal de los patrones dibujados
            partDoc = documents2.Item(f"{partDocument1.Name}")
            product1 = partDoc.GetItem(f"{part1.Name}")
            drawingViewGenerativeBehavior1.Document = product1
            drawingViewGenerativeBehavior1.DefineFrontView(1.0, 0.0, 0.0, 0.0, 1.0, 0.0) #se escoge la vista XY
            drawingView1.Scale = 1.0
            drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior

            drawingViewGenerativeBehavior1.Update()

            partDocument1.Activate()


        sketch = create_sketch(f"sketch{ply_number}")

        projection = flatten_projection()

        linea_10izq, startPoint_10izq, endPoint_10izq = limits_vertical(-5000.0, 5000.0, -5000.0, "lim_izq")
        linea_10dcha, startPoint_10dcha, endPoint_10dcha = limits_vertical(5000.0, 5000.0, -5000.0, "lim_dcha")
        linea_10sup, startPoint_10sup, endPoint_10sup = limits_hztal(5000.0, 5000.0, -5000.0, "lim_sup")
        linea_10inf, startPoint_10inf, endPoint_10inf = limits_hztal(-5000.0, 5000.0, -5000.0, "lim_inf")

        point_coincidence(startPoint_10izq, startPoint_10sup)
        point_coincidence(startPoint_10dcha, endPoint_10sup)
        point_coincidence(endPoint_10dcha, endPoint_10inf)
        point_coincidence(endPoint_10izq, startPoint_10inf)

        length = measure(linea_10dcha)
        division = length/150

        # numero de cortes que necesitara la capa
        num_tape = int(np.ceil(division))
        cortes.append(num_tape)
        print(f"numero de cortes de la capa {s} = {num_tape}")

        # longitud de los cortes de la capa
        width = measure(linea_10sup)
        long_trozo.append(width)
        print(f"longitud de la cinta {s} = {width}")

        # longitud total que ha de ser cortada para la capa
        total_tape = num_tape * width
        long_tot.append(total_tape)
        print(f"longitud total necesaria para la capa {s} = {total_tape}")

        linea = first_line_UD(5000.0, width/2 + 50, -(width/2 + 50))

        lines_drawn = 0

        while lines_drawn < num_tape:

            linea = lines_UD(5000.0, width/2 + 50, -(width/2 + 50))

            lines_drawn = lines_drawn + 1

        create_drawing()

        hide_sketch()
            

    # ******CREAMOS LAS LISTAS DE LOS DATOS OBJETIVO DEL CODIGO**********
    #       se almacenara un valor en cada lista por cada capa

    long_tot = []
    cortes = []
    long_trozo = []

    #*************FLATTENING DE TODAS LAS CAPAS QUE TENGA LA PIEZA******************

    # Access to Plies Group items

    partDocument1 = CATIA.ActiveDocument
    part1 = partDocument1.Part
    partHBs = part1.HybridBodies
    stackHB = partHBs.Item("Stacking")
    stackHBs = stackHB.HybridBodies
    pliesHB = stackHBs.Item("Plies Group.1")
    pliesHBs = pliesHB.HybridBodies

    # Contar el numero de capas que tiene la pieza

    for s in range(1, pliesHBs.Count + 1):

        print(f"-- creando flattening de la capa {s}...")
        flattening(s)

    print("***flattenings terminados con exito***")

    #*********OCULTAR TODO LO QUE NO ES NECESARIO **********

    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    selection1.Add(stackHB)     #STACKING
    visPropertySet1.SetShow(1)      # 0=show    1=hide
    selection1.Clear()
    bodies1 = part1.Bodies
    partBodyHB = bodies1.Item("PartBody")    #PART BODY 
    selection1.Add(partBodyHB)     
    visPropertySet1.SetShow(1)
    selection1.Clear()
    geometricalSetHB = partHBs.Item("Geometrical Set.1")  #GEOMETRICAL SET
    selection1.Add(geometricalSetHB)     
    visPropertySet1.SetShow(1)
    selection1.Clear()

    #****************CREAR UN GEOMETRICAL SET ESPECIAL QUE CONTENGA LOS RESULTADOS**************

    flattGeoSet = partHBs.Add()
    flattGeoSet.Name = "Flattening_Geometry"
    part1.InWorkObject = flattGeoSet
    part1.Update()

    # *********CREAR UN DRAWING DONDE SE IRAN AÑADIENDO HOJAS POR CADA CAPA**********

    documents2 = CATIA.Documents
    drawingDocument1 = documents2.Add("Drawing")
    partDocument1.Activate()

    # ************CREAR LOS SKETCH CON LOS PATRONES DE CORTE Y PASARLO A UN DRAWING********************

    for s in range(1, pliesHBs.Count + 1):
       
        print(f"--sacando patrones de corte de la capa {s}...")
        cinta_UD(s)

    print("***Cantidad de material obtenida correctamente***")

    # ********CERRAMOS EL ARCHIVO Y DEJAMOS ACTIVA LA VISTA DEL DRAWING PARA FINALIZAR***********

    partDocument1.Close()
    drawingDocument1.Activate()

    # *******CREAMOS TABLAS CON LOS DATOS OBTENIDOS************

    data = {'longitud por corte': long_trozo,
            'numero de cortes': cortes,
            'Longitud total': long_tot
            }


    total_long = sum(long_tot)
    print("longitud total de la pieza = ", total_long)

    total_cuts = sum(cortes)
    print("numero total de cortes necesarios = ", total_cuts)

    # hacemos una ultima fila con los numeros para la pieza entera
    final_row = ['--', total_cuts, total_long ]

    df = pd.DataFrame(data)

    df.index = np.arange(1, len(df)+1)

    df.loc[len(df)+1] = final_row

    last_row = len(df)
    df.rename(index={last_row : 'Total pieza'}, inplace=True)
    print(df)

    # ***********CREAR DOCUMENTO DE EXCEL CON LOS RESULTADOS****************

    df.to_excel(f"{results_folder_path}\patrones_corte_{CAD_file}.xlsx")

    print("--excel creado correctamente")


# ********* CREAMOS UNA CARPETA DONDE SE GUARDEN LOS RESULTADOS ******************

# Ruta completa de la carpeta nueva donde se guardarán los resultados
results_folder_path = os.path.join(directory_path, new_folder_name)

try:
    os.mkdir(results_folder_path)
    print("Carpeta creada exitosamente.")
except OSError as e:
    # Si hay algún error al crear la carpeta, muestra el mensaje de error
    print(f"Error al crear la carpeta: {e}")

extension = ".CATPart"

file_names = [f for f in os.listdir(directory_path) if f.endswith(extension)]

# Imprime la lista con los archivos con la extension deseada
print("Files with '{}' extension:".format(extension))
for i, file_name in enumerate(file_names):
    print(f"{i + 1}. {file_name}")


# ************ SE PIDE AL USUARIO INTRODUCIR EL/LOS ARCHIVO/S DE LA LISTA PARA OBTENER LOS PATRONES DE CORTE ***********

user_imput = input("Introduce el numero del documento con el que quieres trabajar. Si quieres abrir todos los archivos escribe cualquier letra: ")

# Comprueba el tipo de imput

if user_imput.isalpha():
    print("--Ha elegido abrir todos los archivos de la carpeta")

    for j in range(len(file_names)):        # Se hace un bucle para seleccionar todos los archivos
        selected_file_index = j
        selected_file = file_names[selected_file_index]

        file_path = os.path.join(directory_path, selected_file)
        documents1 = CATIA.Documents
        Document = documents1.Open(file_path)

        print(f"trabajando con el archivo {selected_file}")
        time.sleep(2)

        partName = selected_file.split(".")[0]    # para quedarme con el nombre del archivo sin el ".CATPart"

        main(partName)  # Funcion que saca los patrones de corte

else:
    try:
        selected_file_index = int(user_imput) - 1
        selected_file = file_names[selected_file_index]
        print(f"Archivo seleccionado: {selected_file}")

        file_path = os.path.join(directory_path, selected_file)
        documents1 = CATIA.Documents
        Document = documents1.Open(file_path)
        time.sleep(2)

        partName = selected_file.split(".")[0]  # para quedarme con el nombre del archivo sin el ".CATPart"

        main(partName)      # Funcion que saca los patrones de corte
        
    except ValueError:
        print("Input not valid. Try again.")




