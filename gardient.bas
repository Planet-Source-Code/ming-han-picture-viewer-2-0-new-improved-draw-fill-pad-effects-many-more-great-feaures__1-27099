Attribute VB_Name = "gardient"
Public Function Horizontal(Newform As Form, Colour1 As ColorConstants, Colour2 As ColorConstants)
'angle = (Vertical / Horizontal)
'calculation variables for r,g,b gradiency
Dim VR, VG, VB As Single
'colors of the picture boxes
Dim Color1, Color2 As Long
'r,g,b variables for each picture box
Dim R, G, B, R2, G2, B2 As Integer
'calculation variable for extracting the rgb values
Dim temp As Long

Color1 = Colour1
Color2 = Colour2

'extract the r,g,b values from the first picture box
temp = (Color1 And 255)
R = temp And 255
temp = Int(Color1 / 256)
G = temp And 255
temp = Int(Color1 / 65536)
B = temp And 255
temp = (Color2 And 255)
R2 = temp And 255
temp = Int(Color2 / 256)
G2 = temp And 255
temp = Int(Color2 / 65536)
B2 = temp And 255

'create a calculation variable for determining the step between
'each level of the gradient; this also allows the user to create
'a perfect gradient regardless of the form size
VR = Abs(R - R2) / Newform.ScaleWidth
VG = Abs(G - G2) / Newform.ScaleWidth
VB = Abs(B - B2) / Newform.ScaleWidth
'if the second value is lower then the first value, make the step
'negative
If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < B Then VB = -VB
'run a loop through the form height, incrementing the gradient color
'according to the height of the line being drawn
For x = 0 To Newform.ScaleWidth
R2 = R + VR * x
G2 = G + VG * x
B2 = B + VB * x
'draw the line and continue through the loop
Newform.Line (x, 0)-(x, Newform.ScaleHeight), RGB(R2, G2, B2)
Next x

End Function
