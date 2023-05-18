import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­Handling
import win32gui # Module for MessageBox
import numpy as np 

# Some basic geometrical data
halfSpan = 1000.0
rootLength = 100.0
tipLength = 50.0
rootTwist = 0.0
tipTwist = 5.0

# Binding python session into CATIA
CATIA = win32com.client.Dispatch("CATIA.Application")
documents1 = CATIA.Documents # CATIA object for managing documents
partDocument1 = documents1.Add("Part") # Starting new part
part1 = partDocument1.Part

#Shape factory provides generating of shapes
ShFactory = part1.HybridShapeFactory

# Starting new body (geometrical set) in part1
bodies1 = part1.HybridBodies

# Adding new body to part1
body1 = bodies1.Add()

# Naming new body as "wireframe"
body1.Name = "Wireframe"

bodies2 = body1.hybridBodies # Starting new geometrical set in Wireframe

body2 = bodies2.Add() # Adding new body to Wireframe

body2.Name = "RootSection" # Naming new body as "RootSection"

body3 = bodies2.Add()

body3.Name = "TipSection"

body4 = bodies1.Add() # Adding new body in part1

body4.Name = "Surfaces" # Naming new body as "Surfaces"

# Loading point coordinated from text file
RootAirfoil = np.array([[10.00000, 10.00000, 0.000000], 
                        [20.00000, 20.00000, 0.000000]])
#RootAirfoil = np.loadtxt('data/clarky.dat',skiprows=1)

TipAirfoil = np.array([[10.000000, 25.000000, 10.000000], 
                       [20.000000, 25.000000, 10.000000]])
#TipAirfoil = np.loadtxt('data/clarky.dat',skiprows=1)

# Creating new point [0,0,0] in Wireframe
point0 = ShFactory.AddNewPointCoord(0.000000, 0.000000, 0.000000)

body1.AppendHybridShape(point0)

# part1 should be updated after every new object
part1.Update() 

#Creatinging Z­direction for translating wing sections
wingAxis1 = ShFactory.AddNewDirectionByCoord(0.000000, 0.000000, 1.000000)

#Creating twist point, sections will be twisted around this point
twistPoint1 = ShFactory.AddNewPointCoord(25.0,0.0,0.0)

twistRef1 = part1.CreateReferenceFromObject(twistPoint1)

#Creating Z­direction for translating wing sections
twistDir1 = ShFactory.AddNewDirectionByCoord(0.000000, 0.000000, 1.000000)

#Creating [POINT­DIRECTION] axis for twisting wing sections 
twistAxis1 = ShFactory.AddNewLinePtDir(twistRef1, twistDir1, 0.000000, 20.000000, False)

# Starting new spline for root section
spline1 = ShFactory.AddNewSpline()
spline1.SetSplineType(0)
spline1.SetClosing(0)

# print("##########")
# print(f'RootAirfoil.shape = {RootAirfoil.shape[:]}')
# print(RootAirfoil)

# Filling the spline with points
def Fill_Spline_With_Points(pointsAirfoil, visibility=False):
    
    spline = ShFactory.AddNewSpline()
    spline.SetSplineType(0)
    spline.SetClosing(0)
    
    for i in range(0,len(pointsAirfoil)):
        
        PT = pointsAirfoil[i] # coordinates are 0..1 which is too small for CATIA
        print(PT)
        point = ShFactory.AddNewPointCoord(PT[0],PT[1],0.0)# coordinates are 2D, Z=0.0
        spline.AddPoint(point) # new point to spline is added
    if visibility:
        ShFactory.GSMVisibility(spline,1) # show the spline
    else:
        ShFactory.GSMVisibility(spline,0) # hide the spline

    return spline


spline1 = Fill_Spline_With_Points(RootAirfoil,0)

spline2 = Fill_Spline_With_Points(TipAirfoil,0)



# for i in range(0,len(RootAirfoil)):

#     PT = RootAirfoil[i]
#     point = ShFactory.AddNewPointCoord(PT[0],PT[1],0.0)# coordinates are 2D, Z=0.0
#     spline1.AddPoint(point) # new point to spline is added
# ShFactory.GSMVisibility(spline1,0) # hide the spline

# # Starting new spline for tip section
# spline2 = ShFactory.AddNewSpline()
# spline2.SetSplineType(0)
# spline2.SetClosing(0)


# for i in range(0,len(TipAirfoil)):

#     PT = TipAirfoil[i]
#     point = ShFactory.AddNewPointCoord(PT[0],PT[1],0.0)
#     spline2.AddPoint(point)
# ShFactory.GSMVisibility(spline2,1)

#Scale [REFERENCE POINT ­ RATIO] the root section
ref1 = part1.CreateReferenceFromObject(spline1)

ref2 = part1.CreateReferenceFromObject(twistPoint1)

scaling1 = ShFactory.AddNewHybridScaling(ref1,ref2, rootLength/100.0)

scaling1.VolumeResult = False

body2.AppendHybridShape(scaling1)

ShFactory.GSMVisibility(scaling1,0)

#Rotate [AXIS] the root section
rotate1= ShFactory.AddNewEmptyRotate()

ref1= part1.CreateReferenceFromObject(scaling1)

ref2 = part1.CreateReferenceFromObject(twistAxis1)

rotate1.ElemToRotate = ref1

rotate1.VolumeResult = False

rotate1.RotationType = 0

rotate1.Axis = twistAxis1

rotate1.AngleValue = rootTwist

body2.AppendHybridShape(rotate1)

ShFactory.GSMVisibility(rotate1,0)

#Translate [DIRECTION ­ DISTANCE] the root section
# is actually not necessary here
translate1 = ShFactory.AddNewEmptyTranslate()

ref1 = part1.CreateReferenceFromObject(rotate1)

translate1.ElemToTranslate = rotate1

translate1.VectorType = 0

translate1.Direction = wingAxis1

translate1.DistanceValue = 0.00

translate1.VolumeResult = False

translate1.Name = "rootShape" # Naming result "rootShape" IMPORTANT!!!

body2.AppendHybridShape(translate1)


###############TIP SECTION#############


#Scale [REFERENCE POINT ­ RATIO] the root section
ref3 = part1.CreateReferenceFromObject(spline1)

ref4 = part1.CreateReferenceFromObject(twistPoint1)

scaling2 = ShFactory.AddNewHybridScaling(ref3,ref4, tipLength/100.0)

scaling2.VolumeResult = False

body3.AppendHybridShape(scaling2)

ShFactory.GSMVisibility(scaling2,0)





#Rotate [AXIS] the root section
rotate2= ShFactory.AddNewEmptyRotate()

ref3= part1.CreateReferenceFromObject(scaling2)

ref4 = part1.CreateReferenceFromObject(twistAxis1)

rotate2.ElemToRotate = ref3

rotate2.VolumeResult = False

rotate2.RotationType = 0

rotate2.Axis = twistAxis1

rotate2.AngleValue = rootTwist

body2.AppendHybridShape(rotate2)

ShFactory.GSMVisibility(rotate2,0)









translate2 = ShFactory.AddNewEmptyTranslate()

translate2.ElemToTranslate = rotate2

translate2.VectorType = 0

translate2.Direction = wingAxis1

translate2.DistanceValue = 0.00

translate2.VolumeResult = False

translate2.Name = "tipShape" # Naming result "tipShape" IMPORTANT!!!

body3.AppendHybridShape(translate2)

######################################

#Create new loft ­ MULTISECTION SURFACE
loft1 = ShFactory.AddNewLoft()

loft1.SectionCoupling = 1

loft1.Relimitation = 1

loft1.CanonicalDetection = 2

#Adding root section to the loft
shapes1 = body2.HybridShapes

# getting item from pool!!
result1 = shapes1.Item("rootShape")

ref1root = part1.CreateReferenceFromObject(result1)

ref2root = None

loft1.AddSectionToLoft(ref1root, 1, ref2root)

#Adding tip section to the loft
shapes2 = body3.HybridShapes

# getting item from pool!!
result2 = shapes2.Item("tipShape")

ref1tip = part1.CreateReferenceFromObject(result2)

ref2tip = None

loft1.AddSectionToLoft(ref1tip, 1, ref2tip)

loft1.Name = "masterSurface"

#Adding loft to Surfaces geometrical set
body4.AppendHybridShape(loft1)

part1.Update()