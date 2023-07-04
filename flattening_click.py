import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­Handling
import win32gui # Module for MessageBox
import numpy as np 
import pyautogui
import time

print('-----new execution-----')


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

    hybridBody5 = hybridBodies4.Item(f"Ply.{ply_number}")

    selection = partDocument1.Selection

    hybridBodies1 = hybridBody1.Parent

    selection.Add(hybridBody5)
   



CATIA = win32com.client.Dispatch("CATIA.Application")

partDocument1 = CATIA.ActiveDocument

part1 = partDocument1.Part

ply_selection(2)
 
# click_on_img('producibility.png', 0.9, 0.1)

# click_on_img('ply.png', 0.8, 0.1)

doubleClick_on_img('flattening_icon.png', 0.9, 0.0)

# flattening_location = pyautogui.locateCenterOnScreen('flattening_icon.png', confidence=0.9)


# # print(flattening_location)

# flat_X = flattening_location[0]
# flat_Y = flattening_location[1]

# pyautogui.doubleClick(flat_X,flat_Y)

#**********SEECION DEL PLANO DEL FLATTENING***************************


if False:
    time.sleep(2)

    selection1 = partDocument1.Selection
    selection1.Clear()
    originElements1 = part1.OriginElements
    plane_XY = originElements1.PlaneXY
    selection1 = partDocument1.Selection
    selection1.Add(plane_XY)

    selection1.Select() #NO SE QUE FUNCION SIRVE


#***********************************************

# click_on_img('xy_plane.png', 0.6, 1)

# xy_plane_location = pyautogui.locateCenterOnScreen('xy_plane.png', confidence=0.8)


# xy_plane_X = xy_plane_location[0]
# xy_plane_Y = xy_plane_location[1]

# pyautogui.click(xy_plane_X,xy_plane_Y)

# time.sleep(1)

click_on_img('ok_button.png', 0.8, 1)

"""


# ok_location = pyautogui.locateCenterOnScreen('ok_button.png', confidence=0.9)

# ok_X = ok_location[0]
# ok_Y = ok_location[1]

# pyautogui.click(ok_X,ok_Y)


#mejor no hacer la optimizacion
# click_on_img('flatten_optimization.png', 0.9, 1.2)  

# click_on_img('ply.png', 0.8, 0.5)

# click_on_img('optimization_ok.png',0.9, 0.5)

click_on_img('save_asDXF.png',0.9, 0.6)

# click_on_img('ply.png', 0.8, 0.5)
ply_selection(1)

click_on_img('optimization_ok.png',0.9, 0.5)

click_on_img('aceptar_button_DXF.png', 0.9, 1)

#************** ya esta guardado el flttening como archivo DXF, ahora hay que abrirlo en un drawing************

documents2 = CATIA.Documents

ruta_DXF = r"C:\\Users\\Probook\\Desktop\\AERO\\TFG\\python\\patron_laminado\\Plies Group.1_Ply.1_prueba_0.dxf"

OpenDocument2 = documents2.Open(ruta_DXF)

time.sleep(2)

drawingDocument2 = CATIA.ActiveDocument

drawingDocument2.Update()


#######################################################

# # Obtener la selección activa
# selection = CATIA.ActiveDocument.Selection

# # Abrir el árbol de especificación (Specification Tree)
# spec_tree = CATIA.ActiveDocument.Part
# spec_tree.OpenEdition()

# # Seleccionar el objeto deseado del árbol por su nombre
# objeto_nombre = "Sequence.1"
# selection.Clear()
# selection.Search(objeto_nombre, "CATPrtSearch")
# # selection.Search("CATPrtSearch", objeto_nombre, False)

# # Realizar la función deseada sobre el objeto seleccionado
# if selection.Count > 0:
#     objeto_seleccionado = selection.Item(1)



#     flattening_location = pyautogui.locateCenterOnScreen('flattening_icon.png')

#     # print(flattening_location)

#     flat_X = flattening_location[0]
#     flat_Y = flattening_location[1]

#     pyautogui.doubleClick(flat_X,flat_Y)



# else:
#     print("El objeto no fue encontrado en el árbol.")

# # Cerrar el árbol de especificación (Specification Tree)
# spec_tree.CloseEdition()

###############################################################

# hybridBodies1 = part1.HybridBodies

# hybridBody1 = hybridBodies1.Item("Stacking")

# hybridBodies2 = hybridBody1.HybridBodies

# hybridBody2 = hybridBodies2.Item("Plies Group.1")

# hybridBodies3 = hybridBody2.HybridBodies

# hybridBody3 = hybridBodies3.Item("Sequence.1")

# hybridBodies4 = hybridBody3.HybridBodies

# hybridBody4 = hybridBodies4.Item("Ply.1")

#######################

# hybridBody3.DoClick()
# tree = CATIA.ActiveDocument.Product

# tree.DoClick("Sequence.1")

# camino_objeto = ["Stacking", "Plies Group.1", "Sequence.1"]

# # Obtiene el árbol de especificación de la parte activa
# tree = CATIA.ActiveDocument.Product

# # Encuentra el elemento final del camino en el árbol
# elemento_objeto = tree.GetItem(camino_objeto)

# elemento_objeto.DoClick()
# tree.SelectNode(elemento_objeto)

"""