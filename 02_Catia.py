print('\t- New execution.')

import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­Handling
import win32gui # Module for MessageBox
import numpy as np # Module for numerical computing

CATIA = win32com.client.Dispatch("CATIA.Application")
documents1 = CATIA.Documents

partDocument1 = documents1.Add("Part")

part1 = partDocument1.Part

hybridBodies1 = part1.HybridBodies

hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

sketches1 = hybridBody1.HybridSketches

originElements1 = part1.OriginElements

reference1 = originElements1.PlaneXY

sketch1 = sketches1.Add(reference1)

arrayOfVariantOfDouble1 = [0.000000,0.000000,0.000000,1.000000,0.000000,0.000000,0.000000,1.000000,0.000000]
sketch1.SetAbsoluteAxisData(arrayOfVariantOfDouble1)

part1.InWorkObject = sketch1


factory2D1 = sketch1.OpenEdition()


geometricElements1 = sketch1.GeometricElements


axis2D1 = geometricElements1.Item("AbsoluteAxis")



line2D1 = axis2D1.GetItem("HDirection")

line2D1.ReportName = 1

line2D2 = axis2D1.GetItem("VDirection")

line2D2.ReportName = 2

point2D1 = factory2D1.CreatePoint(5.000000, 5.000000)

point2D1.ReportName = 3

point2D2 = factory2D1.CreatePoint(5.000000, -5.000000)

point2D2.ReportName = 4

line2D3 = factory2D1.CreateLine(5.000000, 5.000000, 5.000000, -5.000000)

line2D3.ReportName = 5

line2D3.StartPoint = point2D1

line2D3.EndPoint = point2D2

point2D3 = factory2D1.CreatePoint(-5.000000, -5.000000)

point2D3.ReportName = 6

line2D4 = factory2D1.CreateLine(5.000000, -5.000000, -5.000000, -5.000000)

line2D4.ReportName = 7

line2D4.StartPoint = point2D2

line2D4.EndPoint = point2D3

point2D4 = factory2D1.CreatePoint(-5.000000, 5.000000)

point2D4.ReportName = 8

line2D5 = factory2D1.CreateLine(-5.000000, -5.000000, -5.000000, 5.000000)

line2D5.ReportName = 9

line2D5.StartPoint = point2D3

line2D5.EndPoint = point2D4

line2D6 = factory2D1.CreateLine(-5.000000, 5.000000, 5.000000, 5.000000)

line2D6.ReportName = 10

line2D6.StartPoint = point2D4

line2D6.EndPoint = point2D1

constraints1 = sketch1.Constraints
reference2 = part1.CreateReferenceFromObject(line2D3)
reference3 = part1.CreateReferenceFromObject(line2D2)
constraint1 = constraints1.AddBiEltCst(13, reference2, reference3)
constraint1.Mode = 0

reference4 = part1.CreateReferenceFromObject(line2D4)

reference5 = part1.CreateReferenceFromObject(line2D1)

constraint2 = constraints1.AddBiEltCst(10, reference4, reference5)

constraint2.Mode = 0

reference6 = part1.CreateReferenceFromObject(line2D5)

reference7 = part1.CreateReferenceFromObject(line2D2)

constraint3 = constraints1.AddBiEltCst(13, reference6, reference7)

constraint3.Mode = 0

reference8 = part1.CreateReferenceFromObject(line2D6)

reference9 = part1.CreateReferenceFromObject(line2D1)

constraint4 = constraints1.AddBiEltCst(10, reference8, reference9)

constraint4.Mode = 0

reference10 = part1.CreateReferenceFromObject(line2D3)

reference11 = part1.CreateReferenceFromObject(line2D5)

point2D5 = axis2D1.GetItem("Origin")

reference12 = part1.CreateReferenceFromObject(point2D5)

constraint5 = constraints1.AddTriEltCst(17, reference10, reference11, reference12)

constraint5.Mode = 0

reference13 = part1.CreateReferenceFromObject(line2D4)

reference14 = part1.CreateReferenceFromObject(line2D6)

reference15 = part1.CreateReferenceFromObject(point2D5)

constraint6 = constraints1.AddTriEltCst(17, reference13, reference14, reference15)

constraint6.Mode = 0

reference16 = part1.CreateReferenceFromObject(line2D3)

reference17 = part1.CreateReferenceFromObject(line2D5)

constraint7 = constraints1.AddBiEltCst(1, reference16, reference17)

constraint7.Mode = 0

length1 = constraint7.Dimension

length1.Value = 10.000000

reference18 = part1.CreateReferenceFromObject(line2D6)

reference19 = part1.CreateReferenceFromObject(line2D4)

constraint8 = constraints1.AddBiEltCst(1, reference18, reference19)

constraint8.Mode = 0

length2 = constraint8.Dimension

length2.Value = 10.000000

sketch1.CloseEdition 

part1.InWorkObject = hybridBody1

part1.Update 

bodies1 = part1.Bodies

body1 = bodies1.Item("PartBody")

part1.InWorkObject = body1

part1.InWorkObject = body1

shapeFactory1 = part1.ShapeFactory

reference20 = part1.CreateReferenceFromName("")

pad1 = shapeFactory1.AddNewPadFromRef(reference20, 20.000000)

reference21 = part1.CreateReferenceFromObject(sketch1)

pad1.SetProfileElement(reference21)

limit1 = pad1.FirstLimit

length3 = limit1.Dimension

length3.Value = 10.000000

part1.Update()
