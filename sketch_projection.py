import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­Handling
import win32gui # Module for MessageBox
import numpy as np 
import pyautogui
import time

print('-----new execution-----')

CATIA = win32com.client.Dispatch("CATIA.Application")


def paint_lines_vertical(x,y1,y2, line_name):
    """
    Dibuja una linea en un sketch
    - Inputs
        - x [float]: coordenada x de la linea
        - y1 [float]: coordenada y del punto 1
        - y2 [float]: coordenada y del punto 2
        - line_name [string]: nombre de la linea
    - Output
        - line2D
    """

    factory2D1 = sketch1.OpenEdition()

    point2D1 = factory2D1.CreatePoint(x, y1)

    point2D2 = factory2D1.CreatePoint(x, y2)

    line2D1 = factory2D1.CreateLine(x, y1, x, y2)

    line2D1.Name = line_name

    line2D1.StartPoint = point2D1

    line2D1.EndPoint = point2D2

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

def paint_lines_hztal(y,x1,x2, line_name):
    """
    Dibuja una linea en un sketch
    - Inputs
        - y [float]: coordenada y de la linea
        - x1 [float]: coordenada x del punto 1
        - x2 [float]: coordenada x del punto 2
        - line_name [string]: nombre de la linea
    - Output
        - line2D
    """

    factory2D1 = sketch1.OpenEdition()



    point2D1 = factory2D1.CreatePoint(x1, y)

    point2D2 = factory2D1.CreatePoint(x2, y)

    line2D1 = factory2D1.CreateLine(x1, y, x2, y)

    line2D1.Name = line_name

    line2D1.StartPoint = point2D1

    line2D1.EndPoint = point2D2

    constraints1 = sketch1.Constraints

    reference1 = part1.CreateReferenceFromObject(line2D1)

    axis = sketch1.GeometricElements

    axis2D1 = axis.Item("AbsoluteAxis")

    line2D2 = axis2D1.GetItem("HDirection")

    reference2 = part1.CreateReferenceFromObject(line2D2)

    constraint1 = constraints1.AddBiEltCst(10, reference1, reference2)    #catCstTypeHorizontality==10

    constraint1.Mode = 0    #catCstModeDrivingDimension==0

    # hasta aqui para que se vertical, ahora tenemos que hacer que sea tangente con la geometria de proyeccion

    geometry2D1 = geometricElements1.Item("Mark.1")

    reference3 = part1.CreateReferenceFromObject(geometry2D1)

    constraint2 = constraints1.AddBiEltCst(4, reference1, reference3)    #catCstTypeTangency

    constraint2.Mode = 0    #catCstModeDrivingDimension

    sketch1.CloseEdition()

    part1.InWorkObject = hybridBody1


#CREAR UN NUEVO SKETCH EN EL PLANO XY

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

part1.InWorkObject = sketch1

factory2D1 = sketch1.OpenEdition()

hybridBody2 = hybridBodies1.Item("Stacking")

hybridBodies2 = hybridBody2.HybridBodies

hybridBody3 = hybridBodies2.Item("Plies Group.1")

hybridBodies3 = hybridBody3.HybridBodies

hybridBody4 = hybridBodies3.Item("Sequence.1")

hybridBodies4 = hybridBody4.HybridBodies

hybridBody5 = hybridBodies4.Item("Ply.1")

hybridBodies5 = hybridBody5.HybridBodies

hybridBody6 = hybridBodies5.Item("Flatten Body")

hybridBodies6 = hybridBody6.HybridBodies

hybridBody7 = hybridBodies6.Item("Flattening")

sketches2 = hybridBody7.HybridSketches

sketch2 = sketches2.Item("Sketch.FlattenContour.1")

reference1 = part1.CreateReferenceFromObject(sketch2)

geometricElements1 = factory2D1.CreateProjections(reference1)

sketch1.CloseEdition()

part1.InWorkObject = hybridBody1

part1.Update()

#*****OCULTAR STACKING**********

selection1 = partDocument1.Selection

visPropertySet1 = selection1.VisProperties

selection1.Add(hybridBody2)     #PARA SELECCIONAR ALGO    hybridBody2=stacking

visPropertySet1.SetShow(1)

selection1.Clear()

#*************CREAR LAS LINEAS DENTRO DEL SKETCH RESPECTO A LA PROYECCION******************
""""
factory2D1 = sketch1.OpenEdition()

line2D1 = factory2D1.CreateLine(-100.0, 100.0, -100.0, -100.0)

line2D1.Name = "limite_izq"

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


"""

line2D1 = paint_lines_vertical(-100.0, 100.0, -100.0, "limite_izq")

line2D2 = paint_lines_vertical(100.0, 100.0, -100.0, "limite_dcha")

line2D3 = paint_lines_hztal(300.0, 100.0, -100.0, "lim_superior")

line2D4 = paint_lines_hztal(-300.0, 100.0, -100.0, "lim_inferior")

part1.Update()