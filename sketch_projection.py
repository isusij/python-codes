import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­Handling
import win32gui # Module for MessageBox
import numpy as np 
import pyautogui
import time

print('-----new execution-----')

CATIA = win32com.client.Dispatch("CATIA.Application")


def create_sketch(sketch_name):
    """"
    Crea un nuevo sketch en el plano xy
    - Inputs
        - sketch_name [string]: nombre del nuevo sketch
    - Output
        - new_sketch [geometric element]: sketch resultante
    """

    hybridBodies1 = part1.HybridBodies

    hybridBody1 = hybridBodies1.Item("Geometrical Set.1")   

    part1.InWorkObject = hybridBody1

    sketches1 = hybridBody1.HybridSketches

    XYPlane = part1.OriginElements.PlaneXY

    new_sketch = sketches1.Add(XYPlane)   

    new_sketch.Name = sketch_name

    part1.Update()

    return new_sketch

def flatten_projection(ply_number, active_sketch):
    """
    Dibuja una linea en un sketch
    - Inputs
        - ply_number [int]: ply cuyo flattening se quiere proyectar
        - active_sketch [object]: sketch en el que se dibuja la linea
    - Output
        - projection [geometric element]: proyeccion del flattening en el sketch
    """

    part1.InWorkObject = active_sketch

    factory2D1 = active_sketch.OpenEdition()

    hybridBody2 = hybridBodies1.Item("Stacking")

    hybridBodies2 = hybridBody2.HybridBodies

    hybridBody3 = hybridBodies2.Item("Plies Group.1")

    hybridBodies3 = hybridBody3.HybridBodies

    hybridBody4 = hybridBodies3.Item(f"Sequence.{ply_number}")

    hybridBodies4 = hybridBody4.HybridBodies

    hybridBody5 = hybridBodies4.Item(f"Ply.{ply_number}")

    hybridBodies5 = hybridBody5.HybridBodies

    hybridBody6 = hybridBodies5.Item("Flatten Body")

    hybridBodies6 = hybridBody6.HybridBodies

    hybridBody7 = hybridBodies6.Item("Flattening")

    sketches_contour = hybridBody7.HybridSketches

    sketch_contour = sketches_contour.Item("Sketch.FlattenContour.1")

    reference1 = part1.CreateReferenceFromObject(sketch_contour)

    Projection = factory2D1.CreateProjections(reference1)

    active_sketch.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return Projection

def limits_vertical(x, y1, y2, offset, active_sketch):
    """
     Dibuja una linea vertical a una distancia de otra dada
    - Inputs
        - x [float]: coordenada x de la linea a dibujar
        - y1 [float]: coordenada y del punto 1
        - y2 [float]: coordenada y del punto 2
        - offset [long]: distancia de la linea a la linea de referencia
        - active_sketch [object]: sketch en el que se dibuja la linea
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

    constraints1 = active_sketch.Constraints

    geometricElements = active_sketch.GeometricElements

    reference1 = part1.CreateReferenceFromObject(line2D1)

    axis2D1 = geometricElements.Item("AbsoluteAxis")

    line2D2 = axis2D1.GetItem("VDirection")

    reference2 = part1.CreateReferenceFromObject(line2D2)

    constraint1 = constraints1.AddBiEltCst(13, reference1, reference2)  #catCstTypeVerticality

    constraint1.Mode = 0    #catCstModeDrivingDimension

    geometry2D1 = geometricElements1.Item("Mark.1")

    reference3 = part1.CreateReferenceFromObject(geometry2D1)

    constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance

    constraint4.Mode = 0    #catCstModeDrivingDimension

    length1 = constraint4.Dimension

    length1.Value = offset

    active_sketch.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return line2D1, startPoint, endPoint

def limits_hztal(y, x1, x2, offset, active_sketch):
    """
     Dibuja una linea horizontal a una distancia de otra dada
    - Inputs
        - y [float]: coordenada y de la linea a dibujar
        - x1 [float]: coordenada x del punto 1
        - x2 [float]: coordenada x del punto 2
        - offset [long]: distancia de la linea a la linea de referencia
        - active_sketch [object]: sketch en el que se dibuja la linea
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

    constraints1 = active_sketch.Constraints

    geometricElements1 = active_sketch.GeometricElements

    reference1 = part1.CreateReferenceFromObject(line2D1)

    axis2D1 = geometricElements1.Item("AbsoluteAxis")

    line2D2 = axis2D1.GetItem("HDirection")

    reference2 = part1.CreateReferenceFromObject(line2D2)

    constraint1 = constraints1.AddBiEltCst(10, reference1, reference2)  #catCstTypeHorizontality

    constraint1.Mode = 0    #catCstModeDrivingDimension

    geometry2D1 = geometricElements1.Item("Mark.1")

    reference3 = part1.CreateReferenceFromObject(geometry2D1)

    constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance

    constraint4.Mode = 0    #catCstModeDrivingDimension

    length1 = constraint4.Dimension

    length1.Value = offset

    active_sketch.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return line2D1, startPoint, endPoint

def paint_lines_vertical(x,y1,y2, line_name, active_sketch):
    """
    Dibuja una linea en un sketch
    - Inputs
        - active_sketch [object]: sketch sobre el que se trabaja
        - x [float]: coordenada x de la linea
        - y1 [float]: coordenada y del punto 1
        - y2 [float]: coordenada y del punto 2
        - line_name [string]: nombre de la linea
    - Output
        - startPoint [float]: (x,y1)
        - endPoint [float]: (x,y2)
        - line2D1 [geometric element]: la linea dibujada por la funcion
    """

    sketch_current = active_sketch 

    factory2D1 = sketch_current.OpenEdition()

    startPoint = factory2D1.CreatePoint(x, y1)

    endPoint = factory2D1.CreatePoint(x, y2)

    line2D1 = factory2D1.CreateLine(x, y1, x, y2)

    line2D1.Name = line_name

    line2D1.StartPoint = startPoint

    line2D1.EndPoint = endPoint

    constraints1 = sketch_current.Constraints

    reference1 = part1.CreateReferenceFromObject(line2D1)

    axis = sketch_current.GeometricElements

    axis2D1 = axis.Item("AbsoluteAxis")

    line2D2 = axis2D1.GetItem("VDirection")

    reference2 = part1.CreateReferenceFromObject(line2D2)

    constraint1 = constraints1.AddBiEltCst(13, reference1, reference2)    #catCstTypeVerticality==13

    constraint1.Mode = 0    #catCstModeDrivingDimension==0

    # hasta aqui para que se vertical, ahora tenemos que hacer que sea tangente con la geometria de proyeccion

    geometry2D1 = geometricElements1.Item("Mark.1")

    reference3 = part1.CreateReferenceFromObject(geometry2D1)

    constraint2 = constraints1.AddBiEltCst(4, reference1, reference3)    #catCstTypeTangency

    constraint2.Mode = 0    #catCstModeDrivingDimension

    sketch_current.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return startPoint, endPoint, line2D1

def paint_lines_hztal(y,x1,x2, line_name, active_sketch):
    """
     Dibuja una linea en un sketch
    - Inputs
        - y [float]: coordenada y de la linea
        - x1 [float]: coordenada x del punto 1
        - x2 [float]: coordenada x del punto 2
        - line_name [string]: nombre de la linea
        - active_sketch [object]: sketch en el que se dibuja la linea
    - Output
        - startPoint [float]: (x1,y)
        - endPoint [float]: (x2,y)   
        - line2D1 [geometric element]: la linea dibujada por la funcion
    """

    sketch_current = active_sketch

    factory2D1 = sketch_current.OpenEdition()

    startPoint = factory2D1.CreatePoint(x1, y)

    endPoint = factory2D1.CreatePoint(x2, y)

    line2D1 = factory2D1.CreateLine(x1, y, x2, y)

    line2D1.Name = line_name

    line2D1.StartPoint = startPoint

    line2D1.EndPoint = endPoint

    constraints1 = sketch_current.Constraints

    reference1 = part1.CreateReferenceFromObject(line2D1)

    axis = sketch_current.GeometricElements

    axis2D1 = axis.Item("AbsoluteAxis")

    line2D2 = axis2D1.GetItem("HDirection")

    reference2 = part1.CreateReferenceFromObject(line2D2)

    constraint1 = constraints1.AddBiEltCst(10, reference1, reference2)    #catCstTypeHorizontality==10

    constraint1.Mode = 0    #catCstModeDrivingDimension==0

    # hasta aqui para que se vertical, ahora tenemos que hacer que sea tangente con la geometria de proyeccion

    # geometricElements1 = factory2D1.Item("Projection.1")

    geometry2D1 = geometricElements1.Item("Mark.1")

    reference3 = part1.CreateReferenceFromObject(geometry2D1)

    constraint2 = constraints1.AddBiEltCst(4, reference1, reference3)    #catCstTypeTangency

    constraint2.Mode = 0    #catCstModeDrivingDimension

    sketch_current.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return startPoint, endPoint, line2D1

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

    part1.InWorkObject = hybridBody1

    part1.Update()

    return vertex

def get_point_coordinates(point):
    """
     devuelve las coordenadas del punto dado
    - Inputs
        - point1 [float]: punto
    - Output  
        - x
        - y
        - z
    """
    
    x = point.X.Value
    y = point.Y.Value
    z = point.Z.Value
    # x = point[0]
    # y = point[1]
    # z = point[2]

    print("Coordenadas del punto:", point.Value)
    print("X:", x)
    print("Y:", y)
    print("Z:", z)

def line_offset_vertical(x, y1, y2, offset, reference_line, active_sketch):
    """
     Dibuja una linea vertical a una distancia de otra dada
    - Inputs
        - x [float]: coordenada x de la linea a dibujar
        - y1 [float]: coordenada y del punto 1
        - y2 [float]: coordenada y del punto 2
        - off[long]: distancia de la linea a la linea de referencia
        - reference_line [geometric element]: linea sobre la que se calcula la distancia
        - active_sketch [object]: sketch en el que se dibuja la linea
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

    constraints1 = active_sketch.Constraints

    geometricElements = active_sketch.GeometricElements

    reference1 = part1.CreateReferenceFromObject(line2D1)

    axis2D1 = geometricElements.Item("AbsoluteAxis")

    line2D2 = axis2D1.GetItem("VDirection")

    reference2 = part1.CreateReferenceFromObject(line2D2)

    constraint1 = constraints1.AddBiEltCst(13, reference1, reference2)  #catCstTypeVerticality

    constraint1.Mode = 0    #catCstModeDrivingDimension

    reference3 = part1.CreateReferenceFromObject(reference_line)

    constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance

    constraint4.Mode = 0    #catCstModeDrivingDimension

    length1 = constraint4.Dimension

    length1.Value = offset

    active_sketch.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return line2D1, startPoint, endPoint

def line_offset_hztal(y, x1, x2, offset, reference_line, active_sketch):
    """
     Dibuja una linea horizontal a una distancia de otra dada
    - Inputs
        - y [float]: coordenada y de la linea a dibujar
        - x1 [float]: coordenada x del punto 1
        - x2 [float]: coordenada x del punto 2
        - off[long]: distancia de la linea a la linea de referencia
        - reference_line [geometric element]: linea sobre la que se calcula la distancia
        - active_sketch [object]: sketch en el que se dibuja la linea
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

    constraints1 = active_sketch.Constraints

    geometricElements1 = active_sketch.GeometricElements

    reference1 = part1.CreateReferenceFromObject(line2D1)

    axis2D1 = geometricElements1.Item("AbsoluteAxis")

    line2D2 = axis2D1.GetItem("HDirection")

    reference2 = part1.CreateReferenceFromObject(line2D2)

    constraint1 = constraints1.AddBiEltCst(10, reference1, reference2)  #catCstTypeHorizontality

    constraint1.Mode = 0    #catCstModeDrivingDimension

    reference3 = part1.CreateReferenceFromObject(reference_line)

    constraint4 = constraints1.AddBiEltCst(1, reference1, reference3)   #catCstTypeDistance

    constraint4.Mode = 0    #catCstModeDrivingDimension

    length1 = constraint4.Dimension

    length1.Value = offset

    active_sketch.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return line2D1, startPoint, endPoint

def create_drawing(direction_degrees):
    """"
    Crea un drawing de la front view del part donde ver el sketch creado anteriormente
    - Inputs
        - direction_degrees [string]: nombre de la hoja del drawing, para nombrar la direccion representada
    - Output
    """

    drawingSheets1 = drawingDocument1.Sheets

    drawingSheet1 = drawingSheets1.Add("New Sheet")
 
    drawingSheet1.Name = direction_degrees

    drawingSheet1.Activate()

    drawingSheet1.Scale = 1.0

    drawingViews1 = drawingSheet1.Views

    drawingView1 = drawingViews1.Add("AutomaticNaming")

    drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks

    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior

    partDocument1 = documents2.Item("flattening.bis.CATPart")

    product1 = partDocument1.GetItem("Rib_M_G")

    drawingViewGenerativeBehavior1.Document = product1

    drawingViewGenerativeBehavior1.DefineFrontView(1.0, 0.0, 0.0, 0.0, 1.0, 0.0)

    # drawingView1.x = 105.0

    # drawingView1.y = 148.5

    drawingView1.Scale = 1.0

    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior

    drawingViewGenerativeBehavior1.Update()

def hide_sketch(sketch_name):
    """"
    Oculta el sketch seleccionado
    - Inputs
        - sketch_name [object]: sketch a ocultar
    - Output
    """

    selection1 = partDocument1.Selection

    visPropertySet1 = selection1.VisProperties

    selection1.Add(sketch_name)     

    visPropertySet1.SetShow(1)

    selection1.Clear()

    part1.Update()

def hide_limits():
    """"
    Oculta los limites de la pieza para tomar como referencia los limites con el offset de 10mm
    - Inputs
    - Output
    """

    selection1 = partDocument1.Selection

    part1.InWorkObject = sketch1

    geometricElements = sketch1.GeometricElements

    sketch1.OpenEdition()

    visPropertySet1 = selection1.VisProperties

    lim_superior = geometricElements.Item("lim_superior")

    selection1.Add(lim_superior)    

    visPropertySet1.SetShow(1)

    selection1.Clear()

    lim_inferior = geometricElements.Item("lim_inferior")

    selection1.Add(lim_inferior)    

    visPropertySet1.SetShow(1)

    selection1.Clear()

    limite_dcha = geometricElements.Item("limite_dcha")

    selection1.Add(limite_dcha)    

    visPropertySet1.SetShow(1)

    selection1.Clear()

    limite_izq = geometricElements.Item("limite_izq")

    selection1.Add(limite_izq)    

    visPropertySet1.SetShow(1)

    selection1.Clear()

    sketch1.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

def delete_line(line_to_hide, active_sketch):
    """"
    Oculta un elemento en concreto
    - Input
        - line_to_hide [string]: elemento (linea) que queremos ocultar
        - active_sketch [object]: sketch en el que se oculta la linea
    - Output
    """

    selection1 = partDocument1.Selection

    part1.InWorkObject = active_sketch
    
    active_sketch.OpenEdition()

    geometry = active_sketch.GeometricElements

    line = geometry.Item(line_to_hide)

    selection1.Add(line)  

    selection1.Delete()  

    selection1.Clear()

    active_sketch.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

def get_distance(element1, element2, active_sketch):
    """"
    Mide la distancia entre dos elementos 
    - Inputs
        - element1 [object]: elemento 1
        - element2 [object]: elemento 2
        - active_sketch [object]: sketch en el que se encuentran los elementos
    - Output
        - length1 [long]: distancia entre los dos elementos
    """

    parameters1 = part1.Parameters
 
    length1 = parameters1.CreateDimension("", "LENGTH", 0.000000)
    
    # 'if you want to rename the parameter 
    length1.Rename("MeasureDistance")
    
    relations1 = part1.Relations
 
    formula1 = relations1.CreateFormula("Formula.2", "", length1, f"distance(`Geometrical Set.1\{active_sketch}\{element1}` ,`Geometrical Set.1\{active_sketch}\{element2}` ) ")
    
    # 'rename the formula 
    formula1.Rename("Distance")

    return length1

def paint_lines_45(x1, y1, x2, y2, line_name, active_sketch):
    """
     Dibuja una linea en un sketch
    - Inputs
        - x1 [float]: coordenada x del punto 1
        - y1 [float]: coordenada y del punto 1
        - x2 [float]: coordenada x del punto 2
        - y2 [float]: coordenada y del punto 2
        - line_name [string]: nombre de la linea
        - active_sketch [object]: sketch en el que se dibuja la linea
    - Output
        - startPoint [float]: (x1,y1)
        - endPoint [float]: (x2,y2)   
        - line2D1 [geometric element]: la linea dibujada por la funcion
    """


    part1.InWorkObject = active_sketch

    factory2D1 = active_sketch.OpenEdition()

    startPoint = factory2D1.CreatePoint(x1, y1)

    endPoint = factory2D1.CreatePoint(x2, y2)

    line2D1 = factory2D1.CreateLine(x1, y1, x2, y2)

    line2D1.StartPoint = startPoint

    line2D1.EndPoint = endPoint

    line2D1.Name = line_name

    geometry2D1 = geometricElements1.Item("Mark.1")

    geometry2D1.Construction = True

    constraints1 = active_sketch.Constraints

    reference1 = part1.CreateReferenceFromObject(line2D1)

    # reference2 = part1.CreateReferenceFromObject(geometry2D1)
    reference2 = part1.CreateReferenceFromObject(vertex310)

    constraint1 = constraints1.AddBiEltCst(2, reference1, reference2)   #catCstTypeOn

    constraint1.Mode = 0    #catCstModeDrivingDimension

    geometry2D1.Construction = True

    reference3 = part1.CreateReferenceFromObject(geometry2D1)

    constraint2 = constraints1.AddBiEltCst(6, reference1, reference3)   #catCstTypeAngle

    constraint2.Mode = 0    #catCstModeDrivingDimension

    # constraint2.AngleSector = catCstAngleSector1

    angle1 = constraint2.Dimension

    angle1.Value = 45.000000

    active_sketch.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return startPoint, endPoint, line2D1



#********CREAR UN NUEVO SKETCH EN EL PLANO XY********

partDocument1 = CATIA.ActiveDocument

part1 = partDocument1.Part

hybridBodies1 = part1.HybridBodies

hybridBody1 = hybridBodies1.Item("Geometrical Set.1")   

part1.InWorkObject = hybridBody1

sketch1 = create_sketch("cantidad_material")

#*******PROYECCION DEL FLATTEN CONTOUR EN EL SKETCH RECIEN CREADO************

geometricElements1 = flatten_projection(1, sketch1)

#*****OCULTAR STACKING**********

selection1 = partDocument1.Selection

visPropertySet1 = selection1.VisProperties

hybridBody2 = hybridBodies1.Item("Stacking")

selection1.Add(hybridBody2)     #PARA SELECCIONAR ALGO    hybridBody2=stacking

visPropertySet1.SetShow(1)

selection1.Clear()

#*************CREAR LAS LINEAS DENTRO DEL SKETCH RESPECTO A LA PROYECCION******************

# startPoint_izq, endPoint_izq, limite_izq = paint_lines_vertical(-100.0, 100.0, -100.0, "limite_izq", sketch1)

# startPoint_dcha, endPoint_dcha, limite_dcha = paint_lines_vertical(100.0, 100.0, -100.0, "limite_dcha", sketch1)

# startPoint_sup, endPoint_sup, lim_superior = paint_lines_hztal(300.0, 100.0, -100.0, "lim_superior", sketch1)

# startPoint_inf, endPoint_inf, lim_inferior = paint_lines_hztal(-300.0, 100.0, -100.0, "lim_inferior", sketch1)

# #*********uNE LOS EXTREMOS DE LAS RECTAS PARA OBTENER EL PARALELOGRAMO LIMITE DE LA GEOMETRIA********

# vertex1 = point_coincidence(startPoint_izq, startPoint_sup, sketch1)

# vertex2 = point_coincidence(endPoint_sup, startPoint_dcha, sketch1)

# vertex3 = point_coincidence(endPoint_dcha, endPoint_inf, sketch1)

# vertex4 = point_coincidence(endPoint_izq, startPoint_inf, sketch1)

# #*********SUPONEMOS UNA DESVIACION DE 10mm EN CADA DIRECCION********************

# linea_10izq, startPoint_10izq, endPoint_10izq = line_offset_vertical(-500.0, 200.0, -200.0, 10, limite_izq, sketch1)

# linea_10dcha, startPoint_10dcha, endPoint_10dcha = line_offset_vertical(500.0, 200.0, -200.0, 10, limite_dcha, sketch1)

# linea_10sup, startPoint_10sup, endPoint_10sup = line_offset_hztal(500.0, 100.0, -100.0, 10, lim_superior, sketch1)

# linea_10inf, startPoint_10inf, endPoint_10inf = line_offset_hztal(-500.0, 100.0, -100.0, 10, lim_inferior, sketch1)

linea_10izq, startPoint_10izq, endPoint_10izq = limits_vertical(-500.0, 200.0, -200.0, 10, sketch1)

linea_10dcha, startPoint_10dcha, endPoint_10dcha = limits_vertical(500.0, 200.0, -200.0, 10, sketch1)

linea_10sup, startPoint_10sup, endPoint_10sup = limits_hztal(500.0, 100.0, -100.0, 10, sketch1)

linea_10inf, startPoint_10inf, endPoint_10inf = limits_hztal(-500.0, 100.0, -100.0, 10, sketch1)

vertex110 = point_coincidence(startPoint_10izq, startPoint_10sup, sketch1)

vertex210 = point_coincidence(startPoint_10dcha, endPoint_10sup, sketch1)

vertex310 = point_coincidence(endPoint_10dcha, endPoint_10inf, sketch1)

vertex410 = point_coincidence(endPoint_10izq, startPoint_10inf, sketch1)

#*************METER EL SKETCH EN UN DRAWING************************

documents2 = CATIA.Documents

drawingDocument1 = documents2.Add("Drawing")

create_drawing("cantidad_material")

partDocument1.Activate()

#*********ENTRE CADA DRAWING OCULTAMOS LO QUE YA HEMOS SACADO*********


# hide_sketch(sketch1)  #no lo oculto porque quiero tener el limite
# hide_limits()



#***********AHORA SACAMOS LAS TIRAS DE CINTA UNIDIRECCIONAL DIRRECION 0***********

sketch2 = create_sketch("flatten_geometry_0")

geometricElements1 = flatten_projection(1, sketch2)

startPoint_izq, endPoint_izq, limite_izq = paint_lines_vertical(-100.0, 100.0, -100.0, "limite_izq", sketch2)

linea1_v, _, _ = line_offset_vertical(-100.0, 200.0 , -200.0, 10, limite_izq, sketch2)

linea2_v, _, _ = line_offset_vertical(0.0, 200.0, -200.0, 150.0, linea1_v, sketch2)

delete_line("limite_izq", sketch2)

create_drawing("direccion_0")

hide_sketch(sketch2)

partDocument1.Activate()

#**************TIRAS DIRECCION 90********************

sketch3 = create_sketch("flatten_geometry_90")

geometricElements1 = flatten_projection(1, sketch3)

startPoint_sup, endPoint_sup, lim_superior = paint_lines_hztal(300.0, 100.0, -100.0, "lim_superior", sketch3)

linea1_h, _, _ = line_offset_hztal(500.0, 100.0, -100.0, 10, lim_superior, sketch3)

linea2_h, _, _ = line_offset_hztal(0.0, 100.0, -100.0, 150.0, linea1_h, sketch3)

linea3_h, _, _ = line_offset_hztal(-100.0, 100.0, -100.0, 150.0, linea2_h, sketch3)

delete_line("lim_superior", sketch3)

create_drawing("direccion_90")

hide_sketch(sketch3)

partDocument1.Activate()

#************TIRAS DIRECCION +45********************************

sketch4 = create_sketch("flatten_geometry_+45")

geometricElements1 = flatten_projection(1, sketch4)

startPoint_45, endPoint_45, line45 = paint_lines_45(-100.0, -300.0, 200.0, 100.0, "lim_45", sketch4)