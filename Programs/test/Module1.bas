Attribute VB_Name = "Module1"
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const SIZE_SE = &HF008& 'COOOL!!!!!
Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

Sub ShapeTheForm(TheForm As Form)                                                                                                                                                                                                         'This code and all parts of it Copyright (C) 2001 Jeff Katz
With TheForm
'this gets called alot so make it quick
thematrix = CreateRectRgn(0, 0, .ScaleWidth, .ScaleHeight)   'The Whole Form
notthematrix = CreateRectRgn(0, 0, .ScaleWidth, .ScaleHeight) 'The Whole Form

a = CreateRectRgn(10, 0, .ScaleWidth - 10, .ScaleHeight) '[] the form
b = CreateRectRgn(0, 10, .ScaleWidth, .ScaleHeight - 10) ' = the form

 c = CreateEllipticRgn(0, 0, 20, 20) 'upper left corner
 d = CreateEllipticRgn(0, .ScaleHeight, 20, .ScaleHeight - 20)
 e = CreateEllipticRgn(.ScaleWidth, 0, .ScaleWidth - 20, 20)
 f = CreateEllipticRgn(.ScaleWidth, .ScaleHeight, .ScaleWidth - 20, .ScaleHeight - 20)


g = CombineRgn(thematrix, thematrix, a, 4) 'cut out pieces
g = CombineRgn(thematrix, thematrix, b, 4) 'cut out pieces
g = CombineRgn(thematrix, thematrix, c, 4) 'cut out pieces
g = CombineRgn(thematrix, thematrix, d, 4) 'cut out pieces
g = CombineRgn(thematrix, thematrix, e, 4) 'cut out pieces
g = CombineRgn(thematrix, thematrix, f, 4) 'cut out pieces
g = CombineRgn(thematrix, notthematrix, thematrix, 4) 'invert

m = SetWindowRgn(.hwnd, thematrix, True)
DeleteObject thematrix
DeleteObject notthematrix
DeleteObject a
DeleteObject b
DeleteObject c
DeleteObject d
DeleteObject e
DeleteObject f
DeleteObject g
DeleteObject m
'deleteing objects let this get called more than say, once.
End With
End Sub

