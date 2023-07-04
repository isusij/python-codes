import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­Handling
import win32gui # Module for MessageBox
import numpy as np 
import pyautogui
import time

print('-----new execution-----')

CATIA = win32com.client.Dispatch("CATIA.Application")

def mover_sketch(ply_number):
    """"
    corta el sketch del flattening de dentro del stacking y lo pega en el geometrical set de la geometria de los flattenings
    - Imput
        - ply_number [int]: numero de la capa cuyo flattening mavos a mover
    - Output
    """
    partDocument1 = CATIA.ActiveDocument

    selection1 = partDocument1.Selection

    hybridBodies2 = hybridBody1.HybridBodies

    hybridBody2 = hybridBodies2.Item("Plies Group.1")

    hybridBodies3 = hybridBody2.HybridBodies

    hybridBody3 = hybridBodies3.Item(f"Sequence.{ply_number}")

    hybridBodies4 = hybridBody3.HybridBodies

    hybridBody4 = hybridBodies4.Item(f"Ply.{ply_number}")

    hybridBodies5 = hybridBody4.HybridBodies

    hybridBody5 = hybridBodies5.Item("Flatten Body")

    hybridBodies6 = hybridBody5.HybridBodies

    hybridBody6 = hybridBodies6.Item("Flattening")

    sketches1 = hybridBody6.HybridSketches

    sketch1 = sketches1.Item(f"Sketch.FlattenContour.{ply_number}")

    selection1.Add(sketch1)

    selection1.Cut()

    partDocument1 = CATIA.ActiveDocument

    selection2 = partDocument1.Selection

    hybridBody7 = hybridBodies1.Item("Flettening_Geometry")

    selection2.Add(hybridBody7)

    selection2.Paste()



partDocument1 = CATIA.ActiveDocument

selection1 = partDocument1.Selection

selection1.Clear()

part1 = partDocument1.Part

hybridBodies1 = part1.HybridBodies

hybridBody1 = hybridBodies1.Item("Stacking")

#PARA OCULTAR EL CUERPO Y SOLO SE QUEDE VISIBLE ES SKETCH

selection1 = partDocument1.Selection

visPropertySet1 = selection1.VisProperties

hybridBodies1 = hybridBody1.Parent

bSTR1 = hybridBody1.Name

selection1.Add(hybridBody1)     #PARA SELECCIONAR ALGO

visPropertySet1 = visPropertySet1.Parent

bSTR2 = visPropertySet1.Name

bSTR3 = visPropertySet1.Name

visPropertySet1.SetShow(1)

selection1.Clear()

#***********CREAR UN NUEVO GEOMETRICAL SET PARA METER AHI EL SKETCH DEL FLATTENING*************

geometricalSet1 = part1.HybridBodies

flattenGeometrical = geometricalSet1.Add()

flattenGeometrical.Name = "Flettening_Geometry"

part1.Update()

#******************MOVER EL KETCH AL GEOMETRICAL SET Y DEFINIR IN WORK OBJECT******************************

mover_sketch(1)

hybridBody7 = hybridBodies1.Item("Flettening_Geometry")

part1.InWorkObject = hybridBody7


part1.Update()

#*******************************************

