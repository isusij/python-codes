import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­Handling
import win32gui # Module for MessageBox
import numpy as np 
import pyautogui
import time

print('-----new execution-----')

CATIA = win32com.client.Dispatch("CATIA.Application")


def flatten_projection(ply_number):
    """
    Dibuja una linea en un sketch
    - Inputs
        - ply_number [int]: ply cuyo flattening se quiere proyectar
    - Output
        - projection [geometric element]: proyeccion del flattening en el sketch
    """
    #CREAR UN NUEVO SKETCH EN EL PLANO XY

    # partDocument1 = CATIA.ActiveDocument

    # part1 = partDocument1.Part

    # hybridBodies1 = part1.HybridBodies

    # hybridBody1 = hybridBodies1.Item("Geometrical Set.1")   

    # part1.InWorkObject = hybridBody1

    # sketches1 = hybridBody1.HybridSketches

    # XYPlane = part1.OriginElements.PlaneXY

    # sketch1 = sketches1.Add(XYPlane)   

    # sketch1.Name = f"proyection_flat_ply{ply_number}"

    part1.InWorkObject = sketch1

    factory2D1 = sketch1.OpenEdition()

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

    sketches2 = hybridBody7.HybridSketches

    sketch2 = sketches2.Item("Sketch.FlattenContour.1")

    reference1 = part1.CreateReferenceFromObject(sketch2)

    Projection = factory2D1.CreateProjections(reference1)

    sketch1.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return Projection

def paint_lines_vertical(x,y1,y2, line_name):
    """
    Dibuja una linea en un sketch
    - Inputs
        - x [float]: coordenada x de la linea
        - y1 [float]: coordenada y del punto 1
        - y2 [float]: coordenada y del punto 2
        - line_name [string]: nombre de la linea
    - Output
        - startPoint [float]: (x,y1)
        - endPoint [float]: (x,y2)
        - line2D1 [geometric element]: la linea dibujada por la funcion
    """

    factory2D1 = sketch1.OpenEdition()

    startPoint = factory2D1.CreatePoint(x, y1)

    endPoint = factory2D1.CreatePoint(x, y2)

    line2D1 = factory2D1.CreateLine(x, y1, x, y2)

    line2D1.Name = line_name

    line2D1.StartPoint = startPoint

    line2D1.EndPoint = endPoint

    constraints1 = sketch1.Constraints

    reference1 = part1.CreateReferenceFromObject(line2D1)

    axis = sketch1.GeometricElements

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

    sketch1.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return startPoint, endPoint, line2D1

def paint_lines_hztal(y,x1,x2, line_name):
    """
     Dibuja una linea en un sketch
    - Inputs
        - y [float]: coordenada y de la linea
        - x1 [float]: coordenada x del punto 1
        - x2 [float]: coordenada x del punto 2
        - line_name [string]: nombre de la linea
    - Output
        - startPoint [float]: (x1,y)
        - endPoint [float]: (x2,y)   
        - line2D1 [geometric element]: la linea dibujada por la funcion
    """

    factory2D1 = sketch1.OpenEdition()

    startPoint = factory2D1.CreatePoint(x1, y)

    endPoint = factory2D1.CreatePoint(x2, y)

    line2D1 = factory2D1.CreateLine(x1, y, x2, y)

    line2D1.Name = line_name

    line2D1.StartPoint = startPoint

    line2D1.EndPoint = endPoint

    constraints1 = sketch1.Constraints

    reference1 = part1.CreateReferenceFromObject(line2D1)

    axis = sketch1.GeometricElements

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

    sketch1.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return startPoint, endPoint, line2D1

def point_coincidence(point1, point2):
    """
     Hace coincidir dos puntos
    - Inputs
        - point1 [float]: primer punto
        - point2 [float]: segundo punto
    - Output  
        - vertex [float]: punto de union de los dos anteiores
    """
        
    part1.InWorkObject = sketch1

    sketch1.OpenEdition()

    constraints1 = sketch1.Constraints

    reference1 = part1.CreateReferenceFromObject(point1)

    reference2 = part1.CreateReferenceFromObject(point2)

    constraint1 = constraints1.AddBiEltCst(2, reference1, reference2)    #catCstTypeOn=2

    constraint1.Mode = 0    #catCstModeDrivingDimension

    vertex = point1

    sketch1.CloseEdition()

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

def line_offset_vertical(x, y1, y2, offset, reference_line):
    """
     Dibuja una linea vertical a una distancia de otra dada
    - Inputs
        - x [float]: coordenada x de la linea a dibujar
        - y1 [float]: coordenada y del punto 1
        - y2 [float]: coordenada y del punto 2
        - offset [long]: distancia de la linea a la linea de referencia
        - reference_line [geometric element]: linea sobre la que se calcula la distancia
    - Output  
        - line2D1 [geometric element]: la linea creada
    """

    part1.InWorkObject = sketch1

    factory2D1 = sketch1.OpenEdition()

    point2D1 = factory2D1.CreatePoint(x, y1)

    point2D2 = factory2D1.CreatePoint(x, y2)

    line2D1 = factory2D1.CreateLine(x, y1, x, y2)

    line2D1.StartPoint = point2D1

    line2D1.EndPoint = point2D2

    constraints1 = sketch1.Constraints

    geometricElements = sketch1.GeometricElements

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

    sketch1.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return line2D1

def line_offset_hztal(y, x1, x2, offset, reference_line):
    """
     Dibuja una linea horizontal a una distancia de otra dada
    - Inputs
        - y [float]: coordenada y de la linea a dibujar
        - x1 [float]: coordenada x del punto 1
        - x2 [float]: coordenada x del punto 2
        - offset [long]: distancia de la linea a la linea de referencia
        - reference_line [geometric element]: linea sobre la que se calcula la distancia
    - Output  
        - line2D1 [geometric element]: la linea creada
    """

    part1.InWorkObject = sketch1

    factory2D1 = sketch1.OpenEdition()

    point2D1 = factory2D1.CreatePoint(x1, y)

    point2D2 = factory2D1.CreatePoint(x2, y)

    line2D1 = factory2D1.CreateLine(x1, y, x2, y)

    line2D1.StartPoint = point2D1

    line2D1.EndPoint = point2D2

    constraints1 = sketch1.Constraints

    geometricElements1 = sketch1.GeometricElements

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

    sketch1.CloseEdition()

    part1.InWorkObject = hybridBody1

    part1.Update()

    return line2D1




#********CREAR UN NUEVO SKETCH EN EL PLANO XY********

partDocument1 = CATIA.ActiveDocument

part1 = partDocument1.Part

hybridBodies1 = part1.HybridBodies

hybridBody1 = hybridBodies1.Item("Geometrical Set.1")   

part1.InWorkObject = hybridBody1

sketches1 = hybridBody1.HybridSketches

XYPlane = part1.OriginElements.PlaneXY

sketch1 = sketches1.Add(XYPlane)   

sketch1.Name = "proyection_flat"

#*******PROYECCION DEL FLATTEN CONTOUR EN EL SKETCH RECIEN CREADO************

geometricElements1 = flatten_projection(1)

# part1.InWorkObject = sketch1

# factory2D1 = sketch1.OpenEdition()

# hybridBody2 = hybridBodies1.Item("Stacking")

# hybridBodies2 = hybridBody2.HybridBodies

# hybridBody3 = hybridBodies2.Item("Plies Group.1")

# hybridBodies3 = hybridBody3.HybridBodies

# hybridBody4 = hybridBodies3.Item("Sequence.1")

# hybridBodies4 = hybridBody4.HybridBodies

# hybridBody5 = hybridBodies4.Item("Ply.1")

# hybridBodies5 = hybridBody5.HybridBodies

# hybridBody6 = hybridBodies5.Item("Flatten Body")

# hybridBodies6 = hybridBody6.HybridBodies

# hybridBody7 = hybridBodies6.Item("Flattening")

# sketches2 = hybridBody7.HybridSketches

# sketch2 = sketches2.Item("Sketch.FlattenContour.1")

# reference1 = part1.CreateReferenceFromObject(sketch2)

# geometricElements1 = factory2D1.CreateProjections(reference1)

# sketch1.CloseEdition()

# part1.InWorkObject = hybridBody1

# part1.Update()

#*****OCULTAR STACKING**********

selection1 = partDocument1.Selection

visPropertySet1 = selection1.VisProperties

hybridBody2 = hybridBodies1.Item("Stacking")

selection1.Add(hybridBody2)     #PARA SELECCIONAR ALGO    hybridBody2=stacking

visPropertySet1.SetShow(1)

selection1.Clear()

#*************CREAR LAS LINEAS DENTRO DEL SKETCH RESPECTO A LA PROYECCION******************

startPoint_izq, endPoint_izq, limite_izq = paint_lines_vertical(-100.0, 100.0, -100.0, "limite_izq")

startPoint_dcha, endPoint_dcha, limite_dcha = paint_lines_vertical(100.0, 100.0, -100.0, "limite_dcha")

startPoint_sup, endPoint_sup, lim_superior = paint_lines_hztal(300.0, 100.0, -100.0, "lim_superior")

startPoint_inf, endPoint_inf, lim_inferior = paint_lines_hztal(-300.0, 100.0, -100.0, "lim_inferior")

#*********uNE LOS EXTREMOS DE LAS RECTAS PARA OBTENER EL PARALELOGRAMO LIMITE DE LA GEOMETRIA********

vertex1 = point_coincidence(startPoint_izq, startPoint_sup)

vertex2 = point_coincidence(endPoint_sup, startPoint_dcha)

vertex3 = point_coincidence(endPoint_dcha, endPoint_inf)

vertex4 = point_coincidence(endPoint_izq, startPoint_inf)


linea1_v = line_offset_vertical(-100.0, 200.0 , -200.0, 10, limite_izq)

linea2_v = line_offset_vertical(0.0, 200.0, -200.0, 150.0, linea1_v)

linea1_h = line_offset_hztal(500.0, 100.0, -100.0, 10, lim_superior)

linea2_h = line_offset_hztal(0.0, 100.0, -100.0, 150.0, linea1_h)

linea3_h = line_offset_hztal(-100.0, 100.0, -100.0, 150.0, linea2_h)
