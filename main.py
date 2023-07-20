import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­Handling
import win32gui # Module for MessageBox
import numpy as np 
import pyautogui
import time

print('-----new execution-----')

CATIA = win32com.client.Dispatch("CATIA.Application")

def flattening(numero_capa):
    """"
    Hace el flattening de una capa del laminado de la pieza
    - Inputs
        - numero_capa [int]: numero de la capa de la que se hace el flattening
    - Output
    """
    
    def click_on_img(img, confidence_number,espera):
        """
        Doble click sobre un icono
        - Inputs
            - img [string]: icono a seleccionar .png
            - confidence_number [long]: porcentaje de similitud que debe encontrar con la imagen
            - espera [long]: tiempo de espera antes de ejecutar la función
        - Output
        """
        time.sleep(espera)
        img_location = pyautogui.locateCenterOnScreen(img, confidence=confidence_number)
        print(img_location)
        img_X = img_location[0]
        img_Y = img_location[1]

        pyautogui.click(img_X,img_Y)

    def doubleClick_on_img(img, confidence_number, espera):
        """
        Doble click sobre un icono
        - Inputs
            - img [string]: icono a seleccionar .png
            - confidence_number [long]: porcentaje de similitud que debe encontrar con la imagen
            - espera [long]: tiempo de espera antes de ejecutar la función
        - Output
        """
        time.sleep(espera)

        img_location = pyautogui.locateCenterOnScreen(img, confidence=confidence_number)
        img_X = img_location[0]
        img_Y = img_location[1]

        pyautogui.doubleClick(img_X,img_Y)

    def ply_selection(ply_number):
        """
        Seleccionar piel
        - Inputs
            - ply_number [int]: numero de la piel a seleccionar
        - Output
        """
        selection = partDocument1.Selection
        selection.Clear()
        hybridBodies1 = part1.HybridBodies

        hybridBody1 = hybridBodies1.Item("Stacking")

        hybridBodies2 = hybridBody1.HybridBodies

        hybridBody3 = hybridBodies2.Item("Plies Group.1")

        hybridBodies3 = hybridBody3.HybridBodies

        hybridBody4 = hybridBodies3.Item(f"Sequence.{ply_number}")

        hybridBodies4 = hybridBody4.HybridBodies

        plyHB = hybridBodies4.Item(f"Ply.{ply_number}")

        selection = partDocument1.Selection

        hybridBodies1 = hybridBody1.Parent

        selection.Add(plyHB)

    #*********COMIEZO DEL PROCESO PARA UNA CAPA*************

    ply_selection(numero_capa)

    doubleClick_on_img('flattening_icon.png', 0.9, 0.0)

    click_on_img('xy_plane.png', 0.6, 1)

    click_on_img('ok_button.png', 0.8, 1)

    part1.Update()

def create_sketch(sketch_name):
        """"
        Crea un nuevo sketch en el plano xy
        - Inputs
            - sketch_name [string]: nombre del nuevo sketch
        - Output
            - new_sketch [geometric element]: sketch resultante
        """

        partHBs = part1.HybridBodies

        flattGeoSet = partHBs.Item("Flattening_Geometry")   

        part1.InWorkObject = flattGeoSet

        sketches1 = flattGeoSet.HybridSketches

        XYPlane = part1.OriginElements.PlaneXY

        new_sketch = sketches1.Add(XYPlane)   

        new_sketch.Name = sketch_name

        part1.Update()

        return new_sketch

def cinta_UD(ply_number):
    """
    Saca los patrones de corte de cinta unidireccional de la pieza
    - Inputs
        - ply_number [int]: capa de la que se va a sacar el patron de corte
    - Output
        - projection [geometric element]: proyeccion del flattening en el sketch
    """ 

    def flatten_projection():
        """
        Saca la proyeccion del flattening de una capa en un sketch a parte
        - Inputs
            - active_sketch [object]: sketch en el que se trabaja
        - Output
            - projection [geometric element]: proyeccion del flattening en el sketch
        """

        part1.InWorkObject = sketch

        factory2D1 = sketch.OpenEdition()

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

    def limits_vertical(x, y1, y2, offset, active_sketch, line_name):
        """
        Dibuja una linea vertical a una distancia de otra dada
        - Inputs
            - x [float]: coordenada x de la linea a dibujar
            - y1 [float]: coordenada y del punto 1
            - y2 [float]: coordenada y del punto 2
            - offset [long]: distancia de la linea a la linea de referencia
            - active_sketch [object]: sketch en el que se dibuja la linea
            - line_name [string]: nombre de la linea en CATIA
        - Output  
            - line2D1 [geometric element]: la linea creada
            - startPoint [float]
            - endPoint [float]
        """

        part1.InWorkObject = active_sketch

        factory2D1 = active_sketch.OpenEdition()

        startPoint = factory2D1.CreatePoint(x, y1)

        endPoint = factory2D1.CreatePoint(x, y2)

        line2D1 = factory2D1.CreateLine(x, y1, x, y2)

        line2D1.StartPoint = startPoint

        line2D1.EndPoint = endPoint

        line2D1.Name = line_name

        constraints1 = active_sketch.Constraints

        geometricElements = active_sketch.GeometricElements

        reference1 = part1.CreateReferenceFromObject(line2D1)

        axis2D1 = geometricElements.Item("AbsoluteAxis")

        line2D2 = axis2D1.GetItem("VDirection")

        reference2 = part1.CreateReferenceFromObject(line2D2)

        constraint1 = constraints1.AddBiEltCst(13, reference1, reference2)  #catCstTypeVerticality

        constraint1.Mode = 0    #catCstModeDrivingDimension

        geometry2D1 = projection.Item("Mark.1")

        reference3 = part1.CreateReferenceFromObject(geometry2D1)

        constraint2 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance

        constraint2.Mode = 0    #catCstModeDrivingDimension

        length1 = constraint2.Dimension

        length1.Value = offset
        
        active_sketch.CloseEdition()

        part1.InWorkObject = flattGeoSet

        part1.Update()

        

        return line2D1, startPoint, endPoint

    def limits_hztal(y, x1, x2, offset, active_sketch, line_name):
        """
        Dibuja una linea horizontal a una distancia de otra dada
        - Inputs
            - y [float]: coordenada y de la linea a dibujar
            - x1 [float]: coordenada x del punto 1
            - x2 [float]: coordenada x del punto 2
            - offset [long]: distancia de la linea a la linea de referencia
            - active_sketch [object]: sketch en el que se dibuja la linea
            - line_name [string]: nombre de la linea en CATIA
        - Output  
            - line2D1 [geometric element]: la linea creada
            - startPoint [object]
            - endPoint [object]
        """

        part1.InWorkObject = active_sketch

        factory2D1 = active_sketch.OpenEdition()

        startPoint = factory2D1.CreatePoint(x1, y)

        endPoint = factory2D1.CreatePoint(x2, y)

        line2D1 = factory2D1.CreateLine(x1, y, x2, y)

        line2D1.StartPoint = startPoint

        line2D1.EndPoint = endPoint

        line2D1.Name = line_name

        constraints1 = active_sketch.Constraints

        geometricElements = active_sketch.GeometricElements

        reference1 = part1.CreateReferenceFromObject(line2D1)

        axis2D1 = geometricElements.Item("AbsoluteAxis")

        line2D2 = axis2D1.GetItem("HDirection")

        reference2 = part1.CreateReferenceFromObject(line2D2)

        constraint1 = constraints1.AddBiEltCst(10, reference1, reference2)  #catCstTypeHorizontality

        constraint1.Mode = 0    #catCstModeDrivingDimension

        geometry2D1 = projection.Item("Mark.1")

        reference3 = part1.CreateReferenceFromObject(geometry2D1)

        constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance

        constraint4.Mode = 0    #catCstModeDrivingDimension

        length1 = constraint4.Dimension

        length1.Value = offset

        active_sketch.CloseEdition()

        part1.InWorkObject = flattGeoSet

        part1.Update()

        return line2D1, startPoint, endPoint

    def point_coincidence(point1, point2, active_sketch):
        """
        Hace coincidir dos puntos
        - Inputs
            - point1 [float]: primer punto
            - point2 [float]: segundo punto
            - active_sketch [object]: sketch en el que se dibuja la linea
        - Output  
            - vertex [float]: punto de union de los dos anteiores
        """
            
        part1.InWorkObject = active_sketch

        active_sketch.OpenEdition()

        constraints1 = active_sketch.Constraints

        reference1 = part1.CreateReferenceFromObject(point1)

        reference2 = part1.CreateReferenceFromObject(point2)

        constraint1 = constraints1.AddBiEltCst(2, reference1, reference2)    #catCstTypeOn=2

        constraint1.Mode = 0    #catCstModeDrivingDimension

        vertex = point1

        active_sketch.CloseEdition()

        part1.InWorkObject = flattGeoSet

        part1.Update()

        return vertex

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
        print(f'{element.Name} length: {measurableObject.Length} mm')

        return length

    def first_line_UD(x, y1, y2, active_sketch):
        """
        Dibuja la primera linea de cinta
        - Inputs
            - x [float]: coordenada x de la linea a dibujar
            - y1 [float]: coordenada y del punto 1
            - y2 [float]: coordenada y del punto 2
            - active_sketch [object]: sketch en el que se dibuja la linea
        - Output  
            - line2D1 [geometric element]: la linea creada
        """

        part1.InWorkObject = active_sketch

        factory2D1 = active_sketch.OpenEdition()

        startPoint = factory2D1.CreatePoint(x, y1)

        endPoint = factory2D1.CreatePoint(x, y2)

        line2D1 = factory2D1.CreateLine(x, y1, x, y2)

        line2D1.StartPoint = startPoint

        line2D1.EndPoint = endPoint

        line2D1.Name = "first"

        constraints1 = active_sketch.Constraints

        geometricElements = active_sketch.GeometricElements

        reference1 = part1.CreateReferenceFromObject(line2D1)

        axis2D1 = geometricElements.Item("AbsoluteAxis")

        line2D2 = axis2D1.GetItem("VDirection")

        reference2 = part1.CreateReferenceFromObject(line2D2)

        constraint1 = constraints1.AddBiEltCst(13, reference1, reference2)  #catCstTypeVerticality

        constraint1.Mode = 0    #catCstModeDrivingDimension

        reference3 = part1.CreateReferenceFromObject(linea_10izq)

        constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance

        constraint4.Mode = 0    #catCstModeDrivingDimension

        length1 = constraint4.Dimension

        length1.Value = 0.0

        active_sketch.CloseEdition()

        part1.InWorkObject = flattGeoSet

        part1.Update()

        return line2D1

    def lines_UD(x, y1, y2, active_sketch):
        """
        Dibuja una linea vertical a una distancia de otra dada
        - Inputs
            - x [float]: coordenada x de la linea a dibujar
            - y1 [float]: coordenada y del punto 1
            - y2 [float]: coordenada y del punto 2
            - active_sketch [object]: sketch en el que se dibuja la linea
        - Output  
            - line2D1 [geometric element]: la linea creada
        """

        part1.InWorkObject = active_sketch

        factory2D1 = active_sketch.OpenEdition()

        startPoint = factory2D1.CreatePoint(x, y1)

        endPoint = factory2D1.CreatePoint(x, y2)

        line2D1 = factory2D1.CreateLine(x, y1, x, y2)

        line2D1.StartPoint = startPoint

        line2D1.EndPoint = endPoint

        constraints1 = active_sketch.Constraints

        geometricElements = active_sketch.GeometricElements

        reference1 = part1.CreateReferenceFromObject(line2D1)

        axis2D1 = geometricElements.Item("AbsoluteAxis")

        line2D2 = axis2D1.GetItem("VDirection")

        reference2 = part1.CreateReferenceFromObject(line2D2)

        constraint1 = constraints1.AddBiEltCst(13, reference1, reference2)  #catCstTypeVerticality

        constraint1.Mode = 0    #catCstModeDrivingDimension

        reference3 = part1.CreateReferenceFromObject(linea)

        constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance

        constraint4.Mode = 0    #catCstModeDrivingDimension

        length1 = constraint4.Dimension

        length1.Value = 150.0

        active_sketch.CloseEdition()

        part1.InWorkObject = flattGeoSet

        part1.Update()

        return line2D1

    def hide_sketch():
        """"
        Oculta el sketch seleccionado
        - Inputs
            - sketch_name [object]: sketch a ocultar
        - Output
        """

        selection1 = partDocument1.Selection

        visPropertySet1 = selection1.VisProperties

        selection1.Add(sketch)     

        visPropertySet1.SetShow(1)

        selection1.Clear()

        part1.Update()

    def create_drawing():
        """"
        Crea un drawing del patron de corte ya diseñado
        - Inputs
            - direction_degrees [string]: nombre de la hoja del drawing, para nombrar la direccion representada
        - Output
        """

        drawingSheets1 = drawingDocument1.Sheets

        drawingSheet1 = drawingSheets1.Add("New Sheet")
    
        drawingSheet1.Name = f"ply{ply_number}"

        drawingSheet1.Activate()

        drawingSheet1.Scale = 1.0

        drawingViews1 = drawingSheet1.Views

        drawingView1 = drawingViews1.Add("AutomaticNaming")

        drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks

        drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior

        partDoc = documents2.Item(f"{partDocument1.Name}")

        product1 = partDoc.GetItem(f"{part1.Name}")

        drawingViewGenerativeBehavior1.Document = product1

        drawingViewGenerativeBehavior1.DefineFrontView(1.0, 0.0, 0.0, 0.0, 1.0, 0.0)

        # drawingView1.x = 105.0

        # drawingView1.y = 148.5

        drawingView1.Scale = 1.0

        drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior

        drawingViewGenerativeBehavior1.Update()

        partDocument1.Activate()





    sketch = create_sketch(f"sketch{ply_number}")

    projection = flatten_projection()

    linea_10izq, startPoint_10izq, endPoint_10izq = limits_vertical(-5000.0, 5000.0, -5000.0, 10, sketch, "lim_izq")

    linea_10dcha, startPoint_10dcha, endPoint_10dcha = limits_vertical(5000.0, 5000.0, -5000.0, 10, sketch, "lim_dcha")

    linea_10sup, startPoint_10sup, endPoint_10sup = limits_hztal(5000.0, 5000.0, -5000.0, 10, sketch, "lim_sup")

    linea_10inf, startPoint_10inf, endPoint_10inf = limits_hztal(-5000.0, 5000.0, -5000.0, 10, sketch, "lim_inf")

    vertex110 = point_coincidence(startPoint_10izq, startPoint_10sup, sketch)

    vertex210 = point_coincidence(startPoint_10dcha, endPoint_10sup, sketch)

    vertex310 = point_coincidence(endPoint_10dcha, endPoint_10inf, sketch)

    vertex410 = point_coincidence(endPoint_10izq, startPoint_10inf, sketch)

    length = measure(linea_10sup)

    division = length/150

    num_tape = int(np.ceil(division))

    print(num_tape)

    width = measure(linea_10dcha)

    linea = first_line_UD(5000.0, width/2 + 50, -(width/2 + 50), sketch)

    lines_drawn = 0

    while lines_drawn < num_tape:

        linea = lines_UD(5000.0, width/2 + 50, -(width/2 + 50), sketch)

        lines_drawn = lines_drawn + 1

    create_drawing()

    hide_sketch()
        
    




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

ply_count = 0

for s in range(1, pliesHBs.Count + 1):

    sequence = pliesHBs.Item(s)

    print(sequence.Name)    #devuelve sequence.1, sequence.2 ...
   
    print(sequence)
    
    # while s < ply_count:

    flattening(s)

        # ply_count = ply_count + 1



print("bucle terminado con exito")

#*********OCULTAR TODO LO QUE NO ES NECESARIO **********

selection1 = partDocument1.Selection

visPropertySet1 = selection1.VisProperties

selection1.Add(stackHB)     #STACKING

visPropertySet1.SetShow(1)

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


#****************CREAR UN GEOMETRICAL SET ESPECIAL PARA EL PROCESO**************

flattGeoSet = partHBs.Add()

flattGeoSet.Name = "Flattening_Geometry"

part1.InWorkObject = flattGeoSet

part1.Update()

# *********CREAR UN DRAWING DONDE SE IRAN GUARDANDO LOS PATRONES DE CORTE**********

documents2 = CATIA.Documents

drawingDocument1 = documents2.Add("Drawing")

partDocument1.Activate()

# ************CREAR EL SKETCH DE CANTIDAD MATERIAL********************

for s in range(1, pliesHBs.Count + 1):

    sequence = pliesHBs.Item(s)

    print(sequence.Name)    #devuelve sequence.1, sequence.2 ...
   

    # Funcion que saca el material
    cinta_UD(s)

drawingDocument1.Activate()