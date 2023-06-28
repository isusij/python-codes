import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­Handling
import win32gui # Module for MessageBox
import numpy as np 
import pyautogui
import time

print('-----new execution-----')


def click_on_img(img, confidence_number,espera):

    img_location = pyautogui.locateCenterOnScreen(img, confidence=confidence_number)
    img_X = img_location[0]
    img_Y = img_location[1]

    pyautogui.click(img_X,img_Y)

    time.sleep(espera)

def doubleClick_on_img(img, confidence_number, espera):
    
    img_location = pyautogui.locateCenterOnScreen(img, confidence=confidence_number)
    img_X = img_location[0]
    img_Y = img_location[1]

    pyautogui.doubleClick(img_X,img_Y)

    time.sleep(espera)

CATIA = win32com.client.Dispatch("CATIA.Application")

partDocument1 = CATIA.ActiveDocument

part1 = partDocument1.Part

# click_on_img('producibility.png', 0.9, 0.1)

# click_on_img('ply.png', 0.8, 0.1)

doubleClick_on_img('flattening_icon.png', 0.9, 0.1)

# flattening_location = pyautogui.locateCenterOnScreen('flattening_icon.png', confidence=0.9)


# # print(flattening_location)

# flat_X = flattening_location[0]
# flat_Y = flattening_location[1]

# pyautogui.doubleClick(flat_X,flat_Y)

click_on_img('xy_plane.png', 0.8, 0.5)

# xy_plane_location = pyautogui.locateCenterOnScreen('xy_plane.png', confidence=0.8)


# xy_plane_X = xy_plane_location[0]
# xy_plane_Y = xy_plane_location[1]

# pyautogui.click(xy_plane_X,xy_plane_Y)

# time.sleep(1)

click_on_img('ok_button.png', 0.9, 1)

# ok_location = pyautogui.locateCenterOnScreen('ok_button.png', confidence=0.9)

# ok_X = ok_location[0]
# ok_Y = ok_location[1]

# pyautogui.click(ok_X,ok_Y)

# time.sleep(1)

click_on_img('flatten_optimization.png', 0.9, 0.1)

click_on_img('ply.png', 0.8, 0.1)

click_on_img('optimization_ok.png',0.9, 0.1)

click_on_img('save_asDXF.png',0.9, 0.1)

click_on_img('ply.png', 0.8, 0.1)

click_on_img('optimization_ok.png',0.9, 0.5)

click_on_img('aceptar_button_DXF.png', 0.9, 0.1)



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

