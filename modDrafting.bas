Attribute VB_Name = "modDrafting"
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.
'---------------------------------------------------------------------------

'***************************************************************************
Option Explicit
Public Const Pi = 3.14159265358979
Public Const vbGrey = &H7F7F7F
Public Const vbPink = &H7F7FFF
Type CADLayer
    ID As Integer
    Name As String
    Color As Long
    Width As Integer
    style As Integer
    Mode As Integer
    FontName As String
    Frozen As Boolean
    Locked As Boolean
    Hidden As Boolean
End Type
Type CadPoint
    x As Single
    y As Single
    Layer As CADLayer
End Type
Type CadFace
    Vertex(2) As CadPoint
    Layer As CADLayer
End Type
Type CadLine
    P1 As CadPoint
    P2 As CadPoint
    Layer As CADLayer
End Type
Type CadArc
    Center As CadPoint
    Radius As Single
    Angle1 As Single
    Angle2 As Single
    Layer As CADLayer
End Type
Type CadSpline
    Vertex() As CadPoint
    Layer As CADLayer
End Type
Type CadPolyLine
    Vertex() As CadPoint
    Layer As CADLayer
End Type
Type CadEllipse
    F1 As CadPoint
    F2 As CadPoint
    P1 As CadPoint
    Angle1 As Single
    Angle2 As Single
    NumPoints As Integer
    Layer As CADLayer
End Type
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE)
  lfFacename As String * 33
End Type
Type CadText
    Start As CadPoint
    Size As Single
    Angle As Single
    Text As String
    Length As Single
    Layer As CADLayer
End Type
Type CadInsert
    Name As String
    ScaleX As Single
    ScaleY As Single
    Angle As Single
    Base As CadPoint
    Layer As CADLayer
End Type
Type Geometry
    Name As String
    Points() As CadPoint
    Lines() As CadLine
    Arcs() As CadArc
    Ellipses() As CadEllipse
    Splines() As CadSpline
    PolyLines() As CadPolyLine
    Text() As CadText
    Inserts() As CadInsert
    Faces() As CadFace
End Type
Type SelSet
    Type As String
    Index As Integer
End Type
Type KeySet
    Key As String
    Value As Variant
End Type
Type DataSet
    Type As String
    Data() As KeySet
End Type
'-----Spline Stuff-----
Private P() As Single
Private u() As Single
'-----Font stuff-------
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Sub DeleteSelection(MySel() As SelSet, MyGeo As Geometry)
If Not isSelected(MySel()) Then Exit Sub
Dim i As Integer
Dim J As Integer
For i = 0 To UBound(MySel)
    RemoveGeo MyGeo, MySel(i).Type, MySel(i).Index
    For J = i + 1 To UBound(MySel)
        If MySel(J).Type = MySel(i).Type And MySel(J).Index > MySel(i).Index Then
            MySel(J).Index = MySel(J).Index - 1
        End If
    Next J
Next i
End Sub


Function GeoLength(MyGeo As Geometry) As Single
On Local Error Resume Next
Dim i As Integer
With MyGeo
    For i = 0 To UBound(.Lines)
        GeoLength = GeoLength + LineLen(.Lines(i))
    Next i
    For i = 0 To UBound(.Arcs)
        GeoLength = GeoLength + ArcLen(.Arcs(i))
    Next i
    For i = 0 To UBound(.Ellipses)
        GeoLength = GeoLength + EllipseLen(.Ellipses(i))
    Next i
    For i = 0 To UBound(.Splines)
        GeoLength = GeoLength + SplineLen(.Splines(i))
    Next i
    For i = 0 To UBound(.PolyLines)
        GeoLength = GeoLength + PolyLineLen(.PolyLines(i))
    Next i
End With
End Function

Function GetSelLayer(MySel() As SelSet, MyGeo As Geometry) As CADLayer
Select Case MySel(0).Type
    Case "Point": GetSelLayer = MyGeo.Points(MySel(0).Index).Layer
    Case "Line": GetSelLayer = MyGeo.Lines(MySel(0).Index).Layer
    Case "Arc": GetSelLayer = MyGeo.Arcs(MySel(0).Index).Layer
    Case "Ellipse": GetSelLayer = MyGeo.Ellipses(MySel(0).Index).Layer
    Case "Spline": GetSelLayer = MyGeo.Splines(MySel(0).Index).Layer
    Case "PolyLine": GetSelLayer = MyGeo.PolyLines(MySel(0).Index).Layer
    Case "Text": GetSelLayer = MyGeo.Text(MySel(0).Index).Layer
    Case "Face": GetSelLayer = MyGeo.Faces(MySel(0).Index).Layer
    Case "Insert": GetSelLayer = MyGeo.Inserts(MySel(0).Index).Layer
End Select
End Function

Sub SetSelLayer(MySel() As SelSet, MyGeo As Geometry, MyLayer As CADLayer)
Dim i As Integer
For i = 0 To UBound(MySel)
    Select Case MySel(0).Type
        Case "Point": MyGeo.Points(MySel(0).Index).Layer = MyLayer
        Case "Line": MyGeo.Lines(MySel(0).Index).Layer = MyLayer
        Case "Arc": MyGeo.Arcs(MySel(0).Index).Layer = MyLayer
        Case "Ellipse": MyGeo.Ellipses(MySel(0).Index).Layer = MyLayer
        Case "Spline": MyGeo.Splines(MySel(0).Index).Layer = MyLayer
        Case "PolyLine": MyGeo.PolyLines(MySel(0).Index).Layer = MyLayer
        Case "Text": MyGeo.Text(MySel(0).Index).Layer = MyLayer
        Case "Face": MyGeo.Faces(MySel(0).Index).Layer = MyLayer
        Case "Insert": MyGeo.Inserts(MySel(0).Index).Layer = MyLayer
    End Select
Next i
End Sub
Sub SwapSelection(ByRef SelA As SelSet, ByRef SelB As SelSet)
Dim Buf As SelSet
Buf = SelA
SelA = SelB
SelB = Buf
End Sub

Sub ZZ_OLD_CombineSelection(MyGeo As Geometry, SelA As SelSet, SelB As SelSet)
On Error GoTo eTrap
Dim pLineA As CadPolyLine
Dim pLineB As CadPolyLine
Dim i As Integer
Dim J As Integer
Select Case SelA.Type
    Case "Point"
        pLineA.Layer = MyGeo.Points(SelA.Index).Layer
        ReDim pLineA.Vertex(0) As CadPoint
        pLineA.Vertex(0) = MyGeo.Points(SelA.Index)
    Case "Line"
        pLineA.Layer = MyGeo.Lines(SelA.Index).Layer
        ReDim pLineA.Vertex(1) As CadPoint
        pLineA.Vertex(0) = MyGeo.Lines(SelA.Index).P1
        pLineA.Vertex(1) = MyGeo.Lines(SelA.Index).P2
    Case "Arc": pLineA = ArcToPolyLine(MyGeo.Arcs(SelA.Index))
    Case "Ellipse": pLineA = EllipseToPolyLine(MyGeo.Ellipses(SelA.Index))
    Case "Spline": pLineA = SplineToPolyLine(MyGeo.Splines(SelA.Index))
    Case "PolyLine": pLineA = MyGeo.PolyLines(SelA.Index)
End Select
Select Case SelB.Type
    Case "Point"
        pLineB.Layer = MyGeo.Points(SelB.Index).Layer
        ReDim pLineB.Vertex(0) As CadPoint
        pLineB.Vertex(0) = MyGeo.Points(SelB.Index)
    Case "Line"
        pLineB.Layer = MyGeo.Lines(SelB.Index).Layer
        ReDim pLineB.Vertex(1) As CadPoint
        pLineB.Vertex(0) = MyGeo.Lines(SelB.Index).P1
        pLineB.Vertex(1) = MyGeo.Lines(SelB.Index).P2
    Case "Arc": pLineB = ArcToPolyLine(MyGeo.Arcs(SelB.Index))
    Case "Ellipse": pLineB = EllipseToPolyLine(MyGeo.Ellipses(SelB.Index))
    Case "Spline": pLineB = SplineToPolyLine(MyGeo.Splines(SelB.Index))
    Case "PolyLine": pLineB = MyGeo.PolyLines(SelB.Index)
End Select
i = UBound(pLineA.Vertex) + 1
For J = 0 To UBound(pLineB.Vertex)
    If pLineA.Vertex(i - 1).x <> pLineB.Vertex(J).x And pLineA.Vertex(i - 1).y <> pLineB.Vertex(J).y Then
        ReDim Preserve pLineA.Vertex(i) As CadPoint
        pLineA.Vertex(i) = pLineB.Vertex(J)
        i = i + 1
    End If
Next J
RemoveGeo MyGeo, SelA.Type, SelA.Index
RemoveGeo MyGeo, SelB.Type, SelB.Index
i = UBound(MyGeo.PolyLines) + 1
ReDim Preserve MyGeo.PolyLines(i) As CadPolyLine
MyGeo.PolyLines(i) = pLineA
Exit Sub
eTrap:
    i = 0
    Resume Next
End Sub


Function ConvColor(Color As Integer) As Long
'THANKS to Fabio Guerrazzi for the Autocad color conversion
' VB QbColor             AUTOCAD
' =============         ===========
'0   Nero                  7
'1   Blu                   50
'2   Verde                 41
'3   Azzurro               45
'4   Rosso                 34
'5   Fucsia                55
'6   Giallo                38
'7   Bianco                9
'8   Grigio                8
'9   Blu chiaro            5
'10  Verde limone          3
'11  Azzurro chiaro        4
'12  Rosso chiaro          1
'13  Fucsia chiaro         6
'14  Giallo chiaro         2
'15  Bianco brillante      0


 ' Converts a Color from  AutoCAD 10-14 to VB
 
 Dim c As Integer
 
 Select Case Color
    Case 7: c = 0
    Case 50: c = 1
    Case 42: c = 2
    Case 45: c = 3
    Case 34: c = 4
    Case 55: c = 5
    Case 38: c = 6
    Case 9: c = 7
    Case 8: c = 8
    Case 5: c = 9
    Case 3: c = 10
    Case 4: c = 11
    Case 1: c = 12
    Case 6: c = 13
    Case 2: c = 14
    Case 7: c = 15
 End Select
 
 ConvColor = QBColor(c)

End Function
Function ACOS(x As Single)
ACOS = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.
'---------------------------------------------------------------------------

'***************************************************************************
End Function

Sub AddGeo(ByRef Source() As Geometry, ByRef Dest() As Geometry)
Dim i As Integer
Dim K As Integer
AddPoints Source(0).Points(), Dest(0).Points()
AddLines Source(0).Lines(), Dest(0).Lines()
AddEllipses Source(0).Ellipses(), Dest(0).Ellipses()
AddArcs Source(0).Arcs(), Dest(0).Arcs()
AddSplines Source(0).Splines(), Dest(0).Splines()
AddPolyLines Source(0).PolyLines(), Dest(0).PolyLines()
AddText Source(0).Text(), Dest(0).Text()
AddInserts Source(0).Inserts(), Dest(0).Inserts()
For i = 1 To UBound(Source)
    K = UBound(Dest) + 1
    AddPoints Source(i).Points(), Dest(K).Points()
    AddLines Source(i).Lines(), Dest(K).Lines()
    AddEllipses Source(i).Ellipses(), Dest(K).Ellipses()
    AddArcs Source(i).Arcs(), Dest(K).Arcs()
    AddSplines Source(i).Splines(), Dest(K).Splines()
    AddPolyLines Source(i).PolyLines(), Dest(K).PolyLines()
    AddText Source(i).Text(), Dest(K).Text()
    AddInserts Source(i).Inserts(), Dest(K).Inserts()
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************


'***************************************************************************
End Sub

Sub AddSelectionToGeo(SourceGeo As Geometry, MySel As SelSet, DestGeo As Geometry)
On Error GoTo eTrap
Dim K As Integer
Select Case MySel.Type
    Case "Point"
        K = UBound(DestGeo.Points) + 1
        ReDim Preserve DestGeo.Points(K)
        DestGeo.Points(K) = SourceGeo.Points(MySel.Index)
    Case "Line"
        K = UBound(DestGeo.Lines) + 1
        ReDim Preserve DestGeo.Lines(K)
        DestGeo.Lines(K) = SourceGeo.Lines(MySel.Index)
    Case "Arc"
        K = UBound(DestGeo.Arcs) + 1
        ReDim Preserve DestGeo.Arcs(K)
        DestGeo.Arcs(K) = SourceGeo.Arcs(MySel.Index)
    Case "Ellipse"
        K = UBound(DestGeo.Ellipses) + 1
        ReDim Preserve DestGeo.Ellipses(K)
        DestGeo.Ellipses(K) = SourceGeo.Ellipses(MySel.Index)
    Case "Spline"
        K = UBound(DestGeo.Splines) + 1
        ReDim Preserve DestGeo.Splines(K)
        DestGeo.Splines(K) = SourceGeo.Splines(MySel.Index)
    Case "PolyLine"
        K = UBound(DestGeo.PolyLines) + 1
        ReDim Preserve DestGeo.PolyLines(K)
        DestGeo.PolyLines(K) = SourceGeo.PolyLines(MySel.Index)
    Case "Text"
        K = UBound(DestGeo.Text) + 1
        ReDim Preserve DestGeo.Text(K)
        DestGeo.Text(K) = SourceGeo.Text(MySel.Index)
    Case "Insert"
        K = UBound(DestGeo.Inserts) + 1
        ReDim Preserve DestGeo.Inserts(K)
        DestGeo.Inserts(K) = SourceGeo.Inserts(MySel.Index)
End Select
Exit Sub
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function ArcToPolyLine(MyArc As CadArc) As CadPolyLine
On Local Error GoTo eTrap
Dim i As Integer
Dim P As Single
Dim cx As Single
Dim cy As Single
Dim rad As Single
Dim ang1 As Single
Dim ang2 As Single
Dim aLen As Single
Dim div As Integer
rad = MyArc.Radius
ang1 = MyArc.Angle1
ang2 = MyArc.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
div = (ang2 - ang1) / 8
If div = 0 Then div = 1
cx = MyArc.Center.x
cy = MyArc.Center.y
ReDim ArcToPolyLine.Vertex(div + 1) As CadPoint
ArcToPolyLine.Layer = MyArc.Layer
aLen = (ang2 - ang1) / (div)
For i = 0 To div
    P = ang1 + (i * aLen)
    ArcToPolyLine.Vertex(i).x = (rad * Cos(P * Pi / 180)) + cx
    ArcToPolyLine.Vertex(i).y = (rad * Sin(P * Pi / 180)) + cy
Next i
ArcToPolyLine.Vertex(i).x = (rad * Cos(ang2 * Pi / 180)) + cx
ArcToPolyLine.Vertex(i).y = (rad * Sin(ang2 * Pi / 180)) + cy
'For p = MyArc.Angle1 To MyArc.Angle2 Step aLen
'    ArcToPolyLine.Vertex(i).X = (Rad * Cos(p * Pi / 180)) + cx
'    ArcToPolyLine.Vertex(i).Y = (Rad * Sin(p * Pi / 180)) + cy
'    i = i + 1
'Next p
Exit Function
eTrap:
    MsgBox Err.Description
    Resume Next
End Function

Function cNormal(ChkLine As CadLine, Length As Single) As CadLine
Dim ang1 As Single
ang1 = cAngle(ChkLine) + 90
cNormal = cAngLine(ang1, MidPoint(ChkLine), Length)
End Function
Function GetExtents(MyGeo() As Geometry, GNum As Integer) As CadLine
On Local Error Resume Next
Dim eSel() As SelSet
SelectAllGeo eSel(), MyGeo(GNum)
GetExtents = GetSelectionExtents(eSel(), MyGeo(), GNum)
End Function



Function LineInBox(SelLine As CadLine, tLine As CadLine) As Boolean
If BoxLineIntersect(SelLine, tLine) Or PtInBox(SelLine, tLine.P1) Or PtInBox(SelLine, tLine.P2) Then LineInBox = True
End Function

Sub SwapPt(ByRef a As CadPoint, ByRef B As CadPoint)
Dim x As CadPoint
x = a
a = B
B = x
End Sub
Sub ParseDXFLayers(sArray() As String, ByRef DXFLayers() As CADLayer)
Dim i As Long
Dim K As Integer
Dim c As Integer
Do
    i = SearchSection(sArray(), i, "LAYER")
    If i = -1 Then Exit Do
    ReDim Preserve DXFLayers(K) As CADLayer
    i = SearchSection(sArray(), i, "2")
    DXFLayers(K).Name = Trim(sArray(i + 1))
    i = SearchSection(sArray(), i, "70")
    Select Case sArray(i + 1)
        Case 0: DXFLayers(K).Frozen = False: DXFLayers(K).Locked = False
        Case 1: DXFLayers(K).Frozen = True: DXFLayers(K).Locked = False
        Case 2: DXFLayers(K).Frozen = True: DXFLayers(K).Locked = False
        Case 3: DXFLayers(K).Frozen = True: DXFLayers(K).Locked = False
        Case 4: DXFLayers(K).Frozen = False: DXFLayers(K).Locked = True
        Case 5: DXFLayers(K).Frozen = True: DXFLayers(K).Locked = True
        Case 6: DXFLayers(K).Frozen = True: DXFLayers(K).Locked = True
    End Select
    i = SearchSection(sArray(), i, "62")
    c = CInt(sArray(i + 1))
    If c < 0 Then DXFLayers(K).Hidden = True
    c = Abs(c)
    DXFLayers(K).Color = ConvColor(c)
    i = SearchSection(sArray(), i, "6")
    If IsNumeric(sArray(i + 1)) Then
        DXFLayers(K).style = sArray(i + 1)
    End If
    If DXFLayers(K).style > 4 Then DXFLayers(K).style = 0
    DXFLayers(K).Width = 1
    DXFLayers(K).FontName = "Arial Black"
    K = K + 1
Loop
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub PrepareCanvas(Canvas As PictureBox, Layer As CADLayer)
Canvas.DrawMode = Layer.Mode
Canvas.ForeColor = Layer.Color
Canvas.DrawWidth = Layer.Width
Canvas.DrawStyle = Layer.style
End Sub

Sub SplinePoints(MySPline As CadSpline, sPoints() As CadPoint)
'EXTRA-Special thanks to: Franco Languasco for his spline module!!!!!!
Dim n As Integer
n = UBound(MySPline.Vertex)
If n < 2 Then ReDim sPoints(0) As CadPoint: sPoints(0) = MySPline.Vertex(0): Exit Sub
ReDim sPoints(15 + n ^ 2) As CadPoint
'T_Spline MySPline.Vertex(), 15, sPoints()
C_Spline MySPline.Vertex(), sPoints()
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub CreateInsertGeo(MyGeo() As Geometry, MyInsert As CadInsert, ByRef NewGeo() As Geometry)
On Error GoTo eTrap
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim g As Integer
Dim g2 As Integer
Dim CPt As CadPoint
Dim tPoint As CadPoint
Dim tLine As CadLine
Dim tArc As CadArc
Dim tEllipse As CadEllipse
Dim tSpline As CadSpline
Dim tPolyLine As CadPolyLine
Dim tText As CadText
Dim tInsert As CadInsert
If MyInsert.ScaleX = 0 Then MyInsert.ScaleX = 1
If MyInsert.ScaleY = 0 Then MyInsert.ScaleY = 1
K = FindGeo(MyGeo(), MyInsert.Name)
ReDim NewGeo(UBound(MyGeo)) As Geometry
NewGeo(0).Name = MyInsert.Name
'-------------------Create the geometry
J = UBound(MyGeo(K).Points)
g = 0
For i = 0 To J
    tPoint = MyGeo(K).Points(i)
    tPoint = ScalePoint(tPoint, MyInsert.ScaleX, MyInsert.ScaleY)
    tPoint = RotatePoint(tPoint, CPt, MyInsert.Angle)
    tPoint = MovePoint(tPoint, MyInsert.Base.x, MyInsert.Base.y)
    ReDim Preserve NewGeo(0).Points(g) As CadPoint
    NewGeo(0).Points(g) = tPoint
    NewGeo(0).Points(g).Layer = MyInsert.Layer
    NewGeo(0).Points(g).Layer.Color = tPoint.Layer.Color
    NewGeo(0).Points(g).Layer.Width = tPoint.Layer.Width
    NewGeo(0).Points(g).Layer.style = tPoint.Layer.style
    g = g + 1
Next i
J = UBound(MyGeo(K).Lines)
g = 0
For i = 0 To J
    tLine = MyGeo(K).Lines(i)
    tLine = ScaleLine(tLine, MyInsert.ScaleX, MyInsert.ScaleY)
    tLine = RotateLine(tLine, CPt, MyInsert.Angle)
    tLine = MoveLine(tLine, MyInsert.Base.x, MyInsert.Base.y)
    ReDim Preserve NewGeo(0).Lines(g) As CadLine
    NewGeo(0).Lines(g) = tLine
    NewGeo(0).Lines(g).Layer = MyInsert.Layer
    NewGeo(0).Lines(g).Layer.Color = tLine.Layer.Color
    NewGeo(0).Lines(g).Layer.Width = tLine.Layer.Width
    NewGeo(0).Lines(g).Layer.style = tLine.Layer.style
    g = g + 1
Next i
J = UBound(MyGeo(K).Arcs)
g = 0
g2 = 0
For i = 0 To J
    tArc = MyGeo(K).Arcs(i)
    If MyInsert.ScaleX <> 1 Or MyInsert.ScaleY <> 1 Then
        tPolyLine = ArcToPolyLine(tArc)
        tPolyLine = ScalePolyLine(tPolyLine, MyInsert.ScaleX, MyInsert.ScaleY)
        tPolyLine = RotatePolyLine(tPolyLine, CPt, MyInsert.Angle)
        tPolyLine = MovePolyLine(tPolyLine, MyInsert.Base.x, MyInsert.Base.y)
        ReDim Preserve NewGeo(0).PolyLines(g2) As CadPolyLine
        NewGeo(0).PolyLines(g2) = tPolyLine
        NewGeo(0).PolyLines(g2).Layer = MyInsert.Layer
        NewGeo(0).PolyLines(g2).Layer.Color = tPolyLine.Layer.Color
        NewGeo(0).PolyLines(g2).Layer.Width = tPolyLine.Layer.Width
        NewGeo(0).PolyLines(g2).Layer.style = tPolyLine.Layer.style
        g2 = g2 + 1
    Else
        tArc = RotateArc(tArc, CPt, MyInsert.Angle)
        tArc = MoveArc(tArc, MyInsert.Base.x, MyInsert.Base.y)
        ReDim Preserve NewGeo(0).Arcs(g) As CadArc
        NewGeo(0).Arcs(g) = tArc
        NewGeo(0).Arcs(g).Layer = MyInsert.Layer
        NewGeo(0).Arcs(g).Layer.Color = tArc.Layer.Color
        NewGeo(0).Arcs(g).Layer.Width = tArc.Layer.Width
        NewGeo(0).Arcs(g).Layer.style = tArc.Layer.style
        g = g + 1
    End If
Next i
J = UBound(MyGeo(K).Ellipses)
g = 0
For i = 0 To J
    tEllipse = MyGeo(K).Ellipses(i)
    If MyInsert.ScaleX <> 1 Or MyInsert.ScaleY <> 1 Then
        tPolyLine = EllipseToPolyLine(tEllipse)
        tPolyLine = ScalePolyLine(tPolyLine, MyInsert.ScaleX, MyInsert.ScaleY)
        tPolyLine = RotatePolyLine(tPolyLine, CPt, MyInsert.Angle)
        tPolyLine = MovePolyLine(tPolyLine, MyInsert.Base.x, MyInsert.Base.y)
        ReDim Preserve NewGeo(0).PolyLines(g2) As CadPolyLine
        NewGeo(0).PolyLines(g2) = tPolyLine
        NewGeo(0).PolyLines(g2).Layer = MyInsert.Layer
        NewGeo(0).PolyLines(g2).Layer.Color = tPolyLine.Layer.Color
        NewGeo(0).PolyLines(g2).Layer.Width = tPolyLine.Layer.Width
        NewGeo(0).PolyLines(g2).Layer.style = tPolyLine.Layer.style
        g2 = g2 + 1
    Else
        tEllipse = RotateEllipse(tEllipse, CPt, MyInsert.Angle)
        tEllipse = MoveEllipse(tEllipse, MyInsert.Base.x, MyInsert.Base.y)
        ReDim Preserve NewGeo(0).Ellipses(g) As CadEllipse
        NewGeo(0).Ellipses(g) = tEllipse
        NewGeo(0).Ellipses(g).Layer = MyInsert.Layer
        NewGeo(0).Ellipses(g).Layer.Color = tEllipse.Layer.Color
        NewGeo(0).Ellipses(g).Layer.Width = tEllipse.Layer.Width
        NewGeo(0).Ellipses(g).Layer.style = tEllipse.Layer.style
        g = g + 1
    End If
Next i
J = UBound(MyGeo(K).Splines)
g = 0
For i = 0 To J
    tSpline = MyGeo(K).Splines(i)
    tSpline = ScaleSpline(tSpline, MyInsert.ScaleX, MyInsert.ScaleY)
    tSpline = RotateSpline(tSpline, CPt, MyInsert.Angle)
    tSpline = MoveSpline(tSpline, MyInsert.Base.x, MyInsert.Base.y)
    ReDim Preserve NewGeo(0).Splines(g) As CadSpline
    NewGeo(0).Splines(g) = tSpline
    NewGeo(0).Splines(g).Layer = MyInsert.Layer
    NewGeo(0).Splines(g).Layer.Color = tSpline.Layer.Color
    NewGeo(0).Splines(g).Layer.Width = tSpline.Layer.Width
    NewGeo(0).Splines(g).Layer.style = tSpline.Layer.style
    g = g + 1
Next i
J = UBound(MyGeo(K).PolyLines)
For i = 0 To J
    tPolyLine = MyGeo(K).PolyLines(i)
    tPolyLine = ScalePolyLine(tPolyLine, MyInsert.ScaleX, MyInsert.ScaleY)
    tPolyLine = RotatePolyLine(tPolyLine, CPt, MyInsert.Angle)
    tPolyLine = MovePolyLine(tPolyLine, MyInsert.Base.x, MyInsert.Base.y)
    ReDim Preserve NewGeo(0).PolyLines(g2) As CadPolyLine
    NewGeo(0).PolyLines(g2) = tPolyLine
    NewGeo(0).PolyLines(g2).Layer = MyInsert.Layer
    NewGeo(0).PolyLines(g2).Layer.Color = tPolyLine.Layer.Color
    NewGeo(0).PolyLines(g2).Layer.Width = tPolyLine.Layer.Width
    NewGeo(0).PolyLines(g2).Layer.style = tPolyLine.Layer.style
    g2 = g2 + 1
Next i
J = UBound(MyGeo(K).Text)
g = 0
For i = 0 To J
    tText = MyGeo(K).Text(i)
    tText = ScaleText(tText, MyInsert.ScaleX, MyInsert.ScaleY)
    tText = RotateText(tText, CPt, MyInsert.Angle)
    tText = MoveText(tText, MyInsert.Base.x, MyInsert.Base.y)
    ReDim Preserve NewGeo(0).Text(g) As CadText
    NewGeo(0).Text(g) = tText
    NewGeo(0).Text(g).Layer = MyInsert.Layer
    NewGeo(0).Text(g).Layer.Color = tText.Layer.Color
    NewGeo(0).Text(g).Layer.Width = tText.Layer.Width
    NewGeo(0).Text(g).Layer.style = tText.Layer.style
    g = g + 1
Next i
J = UBound(MyGeo(K).Inserts)
g = 0
For i = 0 To J
    tInsert = MyGeo(K).Inserts(i)
    tInsert = ScaleInsert(tInsert, MyInsert.ScaleX, MyInsert.ScaleY)
    tInsert = RotateInsert(tInsert, CPt, MyInsert.Angle)
    tInsert = MoveInsert(tInsert, MyInsert.Base.x, MyInsert.Base.y)
    ReDim Preserve NewGeo(0).Inserts(g2) As CadInsert
    NewGeo(0).Inserts(g) = tInsert
    NewGeo(0).Inserts(g).Layer = MyInsert.Layer
    g = g + 1
Next i
For i = 1 To UBound(MyGeo)
    NewGeo(i) = MyGeo(i)
Next i
Exit Sub
eTrap:
    J = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub DrawCadInsert(Canvas As PictureBox, MyInsert As CadInsert, MyGeo() As Geometry, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
Dim iGeo() As Geometry
CreateInsertGeo MyGeo(), MyInsert, iGeo()
If Mode = 6 Then Erase iGeo(0).Text()
DrawCadGeo Canvas, iGeo(), 0, Mode, Color, Width
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub ExplodeInsert(ByRef MyGeo() As Geometry, GNum As Integer, SelNum As Integer)
On Error GoTo eTrap
Dim iGeo() As Geometry
Dim i As Integer
Dim K As Integer
Dim NewNum As Integer
'--------First we get the geometry for the current insert ----
CreateInsertGeo MyGeo(), MyGeo(GNum).Inserts(SelNum), iGeo()
AddGeo iGeo(), MyGeo()
'-----Next we itterate through any inserts that the insert uses-----
NewNum = FindGeo(MyGeo(), MyGeo(GNum).Inserts(SelNum).Name)
K = UBound(MyGeo(NewNum).Inserts)
For i = 0 To K
    ExplodeInsert MyGeo(), NewNum, i
Next i
'------Next remove the insert from the group
RemoveGeo MyGeo(GNum), "Insert", SelNum
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function FindGeo(MyGeo() As Geometry, Name As String)
Dim i As Integer
For i = 0 To UBound(MyGeo)
    If MyGeo(i).Name = Name Then
        FindGeo = i
        Exit Function
    End If
Next i
FindGeo = -1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function FindLayer(MyLayers() As CADLayer, Name As String)
Dim i As Integer
For i = 0 To UBound(MyLayers)
    If MyLayers(i).Name = Name Then
        FindLayer = i
        Exit Function
    End If
Next i
FindLayer = -1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Sub GetDrawGeo(SourceGeo() As Geometry, MyInsert As CadInsert, DestGeo As Geometry)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim tGeo() As Geometry
Dim xGeo As Geometry
'Create the geometry in a temp geo-set
CreateInsertGeo SourceGeo(), MyInsert, tGeo()
'Add all the drawable geometry to the destination
AddPoints tGeo(0).Points(), DestGeo.Points()
AddLines tGeo(0).Lines(), DestGeo.Lines()
AddEllipses tGeo(0).Ellipses(), DestGeo.Ellipses()
AddArcs tGeo(0).Arcs(), DestGeo.Arcs()
AddSplines tGeo(0).Splines(), DestGeo.Splines()
AddPolyLines tGeo(0).PolyLines(), DestGeo.PolyLines()
AddText tGeo(0).Text(), DestGeo.Text()
K = UBound(tGeo(0).Inserts)
For i = 0 To K
    GetDrawGeo tGeo(), tGeo(0).Inserts(i), DestGeo
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function VLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As CadLine
VLine.P1.x = X1
VLine.P1.y = Y1
VLine.P2.x = X2
VLine.P2.y = Y2
End Function

Function ZZ_OLD_GetExtents(MyGeo() As Geometry, GNum As Integer) As CadLine
'THis function returns the zoomable extents of a geometry set
'This is done by converting all geometry (and inserts) into lines and then
'finding the max and min of all the points of those lines - -
'we then buffer the edges so that it doesn't touch the borders of the canvas
On Error GoTo eTrap:
Dim eGeo As Geometry
Dim eSel() As SelSet
Dim i As Integer
Dim K As Integer
Dim tInsert As CadInsert
'Setup our Extent Line so that it is a "negatively defined box"
ZZ_OLD_GetExtents.P1.x = 32000
ZZ_OLD_GetExtents.P1.y = 32000
ZZ_OLD_GetExtents.P2.x = -32000
ZZ_OLD_GetExtents.P2.y = -32000
'Get all drawable geometry into the primary view of our temporary geometry set
'This inclusdes making sure all the inserts are converted to the primary view of the set
tInsert.Name = MyGeo(GNum).Name
GetDrawGeo MyGeo(), tInsert, eGeo
'Explode all of the geometry into lines
SelectAllGeo eSel(), eGeo
For i = 0 To UBound(eSel)
    Select Case eSel(i).Type
            Case "Arc": ExplodeArc eGeo, eSel(i).Index
            Case "Ellipse": ExplodeEllipse eGeo, eSel(i).Index
            Case "Spline": ExplodeSpline eGeo, eSel(i).Index
            Case "PolyLine": ExplodePolyline eGeo, eSel(i).Index
    End Select
Next i
'--------------------Get the extents of the lines (all geometry is now lines)-----
For i = 0 To UBound(eGeo.Lines)
    If eGeo.Lines(i).P1.x < ZZ_OLD_GetExtents.P1.x Then ZZ_OLD_GetExtents.P1.x = eGeo.Lines(i).P1.x
    If eGeo.Lines(i).P1.y < ZZ_OLD_GetExtents.P1.y Then ZZ_OLD_GetExtents.P1.y = eGeo.Lines(i).P1.y
    If eGeo.Lines(i).P1.x > ZZ_OLD_GetExtents.P2.x Then ZZ_OLD_GetExtents.P2.x = eGeo.Lines(i).P1.x
    If eGeo.Lines(i).P1.y > ZZ_OLD_GetExtents.P2.y Then ZZ_OLD_GetExtents.P2.y = eGeo.Lines(i).P1.y
    If eGeo.Lines(i).P2.x < ZZ_OLD_GetExtents.P1.x Then ZZ_OLD_GetExtents.P1.x = eGeo.Lines(i).P2.x
    If eGeo.Lines(i).P2.y < ZZ_OLD_GetExtents.P1.y Then ZZ_OLD_GetExtents.P1.y = eGeo.Lines(i).P2.y
    If eGeo.Lines(i).P2.x > ZZ_OLD_GetExtents.P2.x Then ZZ_OLD_GetExtents.P2.x = eGeo.Lines(i).P2.x
    If eGeo.Lines(i).P2.y > ZZ_OLD_GetExtents.P2.y Then ZZ_OLD_GetExtents.P2.y = eGeo.Lines(i).P2.y
Next i
'--------------Buffer edges so the zoom shows the outer elements
'GetExtents.P1.X = GetExtents.P1.X - ((GetExtents.P2.X - GetExtents.P1.X) / 6)
'GetExtents.P1.Y = GetExtents.P1.Y - ((GetExtents.P2.Y - GetExtents.P1.Y) / 6)
'GetExtents.P2.X = GetExtents.P2.X + ((GetExtents.P2.X - GetExtents.P1.X) / 6)
'GetExtents.P2.Y = GetExtents.P2.Y + ((GetExtents.P2.Y - GetExtents.P1.Y) / 6)

Exit Function
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function GetSelectionExtents(eSel() As SelSet, MyGeo() As Geometry, GNum As Integer) As CadLine
'THis function returns the extents of a selection set
'This is done by converting all geometry (and inserts) into lines and then
'finding the max and min of all the points of those lines - -
If Not isSelected(eSel()) Then Exit Function
On Error GoTo eTrap:
Dim i As Integer
Dim K As Integer
Dim eGeo() As Geometry
Dim NewSel() As SelSet
ReDim eGeo(UBound(MyGeo))
'----------Copy Our Geometry Set
For i = 0 To UBound(MyGeo)
    eGeo(i) = MyGeo(i)
Next i
'---------Mark the selected geometry with an ID = -99------
For i = 0 To UBound(eSel)
    Select Case eSel(i).Type
        Case "Point": eGeo(GNum).Points(eSel(i).Index).Layer.ID = -99
        Case "Line": eGeo(GNum).Lines(eSel(i).Index).Layer.ID = -99
        Case "Arc": eGeo(GNum).Arcs(eSel(i).Index).Layer.ID = -99
        Case "Ellipse": eGeo(GNum).Ellipses(eSel(i).Index).Layer.ID = -99
        Case "Spline": eGeo(GNum).Splines(eSel(i).Index).Layer.ID = -99
        Case "PolyLine": eGeo(GNum).PolyLines(eSel(i).Index).Layer.ID = -99
        Case "Insert": eGeo(GNum).Inserts(eSel(i).Index).Layer.ID = -99
    End Select
Next i
'-Explode all the inserts
K = UBound(eGeo(GNum).Inserts)
For i = K To 0 Step -1
    ExplodeInsert eGeo(), GNum, i
Next i
'---------Next we select all geometry, and explode it-------
SelectAllGeo NewSel(), eGeo(GNum)
For i = UBound(NewSel) To 0 Step -1
    ExplodeSelection eGeo(GNum), NewSel(i)
Next i
'---------Explode all of the lines into points--------
For i = UBound(eGeo(GNum).Lines) To 0 Step -1
    ExplodeLine eGeo(GNum), i
Next i
'Setup our Extent Line so that it is a "negatively defined box"
GetSelectionExtents.P1.x = 32000
GetSelectionExtents.P1.y = 32000
GetSelectionExtents.P2.x = -32000
GetSelectionExtents.P2.y = -32000
'--------------------Get the extents of the lines (all geometry is now lines)-----
For i = 0 To UBound(eGeo(GNum).Points)
    If eGeo(GNum).Points(i).Layer.ID = -99 Then
        If eGeo(GNum).Points(i).x < GetSelectionExtents.P1.x Then GetSelectionExtents.P1.x = eGeo(GNum).Points(i).x
        If eGeo(GNum).Points(i).y < GetSelectionExtents.P1.y Then GetSelectionExtents.P1.y = eGeo(GNum).Points(i).y
        If eGeo(GNum).Points(i).x > GetSelectionExtents.P2.x Then GetSelectionExtents.P2.x = eGeo(GNum).Points(i).x
        If eGeo(GNum).Points(i).y > GetSelectionExtents.P2.y Then GetSelectionExtents.P2.y = eGeo(GNum).Points(i).y
    End If
Next i
'--------------Buffer edges so the zoom shows the outer elements
'GetExtents.P1.X = GetExtents.P1.X - ((GetExtents.P2.X - GetExtents.P1.X) / 6)
'GetExtents.P1.Y = GetExtents.P1.Y - ((GetExtents.P2.Y - GetExtents.P1.Y) / 6)
'GetExtents.P2.X = GetExtents.P2.X + ((GetExtents.P2.X - GetExtents.P1.X) / 6)
'GetExtents.P2.Y = GetExtents.P2.Y + ((GetExtents.P2.Y - GetExtents.P1.Y) / 6)

Exit Function
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Sub ImportDXF(Status As Object, ByRef DXFGeo() As Geometry, ByRef DXFLayers() As CADLayer, FileDXF As String)
Dim FF As Integer
Dim DXFLine As String
Dim Version As String
Dim ENDSEC As Boolean
Dim Section() As String
Dim GetNext As Boolean
Dim bCount As Integer
ReDim DXFGeo(0) As Geometry  'Zero will be the PV, the others will be blocks if they exists.
FF = FreeFile
Open FileDXF For Input As #FF
'First we need to find the version number . . .
FindCommand FF, "$ACADVER"
FindCommand FF, "1"
Line Input #FF, Version
'Skip to the TABLES section of the DXF file
FindCommand FF, "TABLES"
GetSection FF, "LAYER", "ENDSEC", "ENDSEC", Section()
ParseDXFLayers Section(), DXFLayers()
'Next we skip all the header stuff and get to the section called 'BLOCKS'
FindCommand FF, "BLOCKS"
'---------------------------
'BLOCKS are groups of geometry that
'are re-useable within the drawing
'they may appear several times within one drawing
'and if the block is modified it automatically
'modifies each time wherever it's used within the drawing
GetNext = True
bCount = 1
Do While Not ENDSEC
    'First we load in a SECTION into an array (BLOCK) to (ENDBLK)
    'we do this until we come across the "ENDSEC" command
    If GetSection(FF, "BLOCK", "ENDBLK", "ENDSEC", Section()) Then
        'We have a "BLOCK" in the array
        'So we have to advance our array of BLOCKS (Geometry)
        ReDim Preserve DXFGeo(bCount) As Geometry
        If ParseDXF(Section(), DXFGeo(UBound(DXFGeo)), DXFLayers(), Version, True) Then bCount = bCount + 1
        Status.Caption = "Importing Block" & DXFGeo(UBound(DXFGeo)).Name
    Else
        ENDSEC = True
    End If
Loop
'Now we go after the 'Primary View Entities
ENDSEC = False
GetSection FF, "ENTITIES", "ENDSEC", "ENDSEC", Section()
'This grabs ALL PV ENTITIES . . . kind of like one huge block
Close #FF 'We can close the file because we're finished with it
'Next we fill the array with geometry data
Status.Caption = "Importing PV"
ParseDXF Section(), DXFGeo(0), DXFLayers(), Version, False
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub


Function MoveInsert(MyInsert As CadInsert, dX As Single, dY As Single) As CadInsert
MoveInsert = MyInsert
MoveInsert.Base.x = MyInsert.Base.x + dX
MoveInsert.Base.y = MyInsert.Base.y + dY
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ParseDXF(sArray() As String, ByRef bGeo As Geometry, MyLayers() As CADLayer, Version As String, isBlock As Boolean) As Boolean
'On Local Error GoTo exitMe:
Dim i As Long
Dim J As Long
Dim K As Long
Dim P As Long
Dim MyKeys() As DataSet
Dim Endword As String
If isBlock Then i = SearchSection(sArray(), i, "2") + 1
bGeo.Name = sArray(i)
If Not isBlock Then
    DoEvents
End If
For J = i To UBound(sArray)
    If IsDXFCommand(sArray(J)) Then 'We Found an ENTITY COMMAND
        ReDim Preserve MyKeys(K) As DataSet
        MyKeys(K).Type = sArray(J)
        'I am not sure if a BLOCK can use a block.
        'Either way, this is designed to work even if you can
        Select Case MyKeys(K).Type
            Case "INSERT", "DIMENSION"
                'KEY "2" on an INSERT provides the BLOCK name to be inserted
                J = SearchSection(sArray(), J, "2")
            Case Else
                J = FindStart(sArray(), J, Version)
        End Select
        If MyKeys(K).Type = "POLYLINE" Then
        'Endword = "ENDSEC" Else Endword = "0"
            Do While UCase(sArray(J)) <> "ENDSEC" And UCase(sArray(J)) <> "SEQEND" And UCase(sArray(J + 1)) <> "ENDSEC" And UCase(sArray(J + 1)) <> "SEQEND" And J + 1 <= UBound(sArray)
                ReDim Preserve MyKeys(K).Data(P)
                MyKeys(K).Data(P).Key = sArray(J)
                MyKeys(K).Data(P).Value = sArray(J + 1)
                P = P + 1
                J = J + 2
            Loop
        Else
            Do While UCase(sArray(J)) <> "0" And J + 1 <= UBound(sArray)
                ReDim Preserve MyKeys(K).Data(P)
                MyKeys(K).Data(P).Key = sArray(J)
                MyKeys(K).Data(P).Value = sArray(J + 1)
                P = P + 1
                J = J + 2
            Loop
        End If
        PrepareDXFEntity MyKeys(K), bGeo, MyLayers()
        K = K + 1
        P = 0
    End If
Next J
ParseDXF = True
Exit Function
ExitMe:
MsgBox "ERROR  " & Err.Description
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Sub PrepareDXFEntity(ByRef MyKey As DataSet, ByRef MyGeo As Geometry, MyLayers() As CADLayer)
'This may take a little more time during the "load"
'and it may take a little more memory, but in the end
'it will draw much faster
On Error GoTo eTrap
Dim i As Long
Dim J As Long
Dim K As Long
Dim P As Long
Dim Layer As CADLayer
i = FindLayer(MyLayers(), kVal(MyKey.Data(), 8))
Layer = MyLayers(i)
Select Case MyKey.Type
    Case "POINT"
        K = UBound(MyGeo.Points) + 1
        ReDim Preserve MyGeo.Points(K) As CadPoint
        MyGeo.Points(K).x = kVal(MyKey.Data(), 10)
        MyGeo.Points(K).y = kVal(MyKey.Data(), 20)
        MyGeo.Points(K).Layer = Layer
    Case "LINE"
        K = UBound(MyGeo.Lines) + 1
        ReDim Preserve MyGeo.Lines(K) As CadLine
        MyGeo.Lines(K).P1.x = kVal(MyKey.Data(), 10)
        MyGeo.Lines(K).P1.y = kVal(MyKey.Data(), 20)
        MyGeo.Lines(K).P2.x = kVal(MyKey.Data(), 11)
        MyGeo.Lines(K).P2.y = kVal(MyKey.Data(), 21)
        MyGeo.Lines(K).Layer = Layer
    Case "ARC", "CIRCLE"
        K = UBound(MyGeo.Arcs) + 1
        ReDim Preserve MyGeo.Arcs(K) As CadArc
        MyGeo.Arcs(K).Center.x = kVal(MyKey.Data(), 10)
        MyGeo.Arcs(K).Center.y = kVal(MyKey.Data(), 20)
        MyGeo.Arcs(K).Radius = kVal(MyKey.Data(), 40)
        MyGeo.Arcs(K).Angle1 = kVal(MyKey.Data(), 50)
        MyGeo.Arcs(K).Angle2 = kVal(MyKey.Data(), 51)
        MyGeo.Arcs(K).Layer = Layer
    Case "CIRCLE"
        K = UBound(MyGeo.Arcs) + 1
        ReDim Preserve MyGeo.Arcs(K) As CadArc
        MyGeo.Arcs(K).Center.x = kVal(MyKey.Data(), 10)
        MyGeo.Arcs(K).Center.y = kVal(MyKey.Data(), 20)
        MyGeo.Arcs(K).Radius = kVal(MyKey.Data(), 40)
        MyGeo.Arcs(K).Angle1 = 0
        MyGeo.Arcs(K).Angle2 = 360
        MyGeo.Arcs(K).Layer = Layer
    Case "ELLIPSE"
        K = UBound(MyGeo.Ellipses) + 1
        ReDim Preserve MyGeo.Ellipses(K) As CadEllipse
        Dim Ratio As Single
        Dim Angle As Single
        Dim Center As CadPoint
        Dim Edge As CadPoint
        Dim a As Single
        Dim B As Single
        Ratio = kVal(MyKey.Data(), 40)
        Center.x = kVal(MyKey.Data(), 10)
        Center.y = kVal(MyKey.Data(), 20)
        Edge.x = Center.x + kVal(MyKey.Data(), 11)
        Edge.y = Center.y + kVal(MyKey.Data(), 21)
        Angle = PtPtAngle(Center, Edge)
        a = PtLen(Center, Edge)
        B = Sqr((a ^ 2 - (a * Ratio) ^ 2))
        MyGeo.Ellipses(K).F1 = cAngPt(Angle, Center, -B)
        MyGeo.Ellipses(K).F2 = cAngPt(Angle, Center, B)
        MyGeo.Ellipses(K).P1 = cAngPt(90 + Angle, Center, a * Ratio)
        MyGeo.Ellipses(K).Angle1 = kVal(MyKey.Data(), 41) * 180 / Pi
        MyGeo.Ellipses(K).Angle2 = kVal(MyKey.Data(), 42) * 180 / Pi
        MyGeo.Ellipses(K).NumPoints = 32
        MyGeo.Ellipses(K).Layer = Layer
        'DrawCadLine frmDraw.picDraw, PtLine(Center, MyGeo.Ellipses(k).P1), , vbGreen, 3
        'DrawCadLine frmDraw.picDraw, cAngLine(90 + Angle, Center, Ratio * (PtLen(Center, MyGeo.Ellipses(k).P1))), , vbRed, 3
        'DrawCadLine frmDraw.picDraw, tLine, , vbYellow, 1
        'DrawCadEllipse frmDraw.picDraw, MyGeo.Ellipses(k), , vbGreen
        'frmDraw.picDraw.Picture = frmDraw.picDraw.Image
    Case "POLYLINE"
        K = UBound(MyGeo.PolyLines) + 1
        ReDim Preserve MyGeo.PolyLines(K) As CadPolyLine
        Do While P < UBound(MyKey.Data)
            Do While MyKey.Data(P).Value <> "VERTEX" And P < UBound(MyKey.Data)
                P = P + 1
            Loop
            If P < UBound(MyKey.Data) Then
                ReDim Preserve MyGeo.PolyLines(K).Vertex(J) As CadPoint
                Do While MyKey.Data(P).Key <> 10 And P < UBound(MyKey.Data)
                    P = P + 1
                Loop
                MyGeo.PolyLines(K).Vertex(J).x = MyKey.Data(P).Value
                Do While MyKey.Data(P).Key <> 20 And P < UBound(MyKey.Data)
                    P = P + 1
                Loop
                MyGeo.PolyLines(K).Vertex(J).y = MyKey.Data(P).Value
                Do While MyKey.Data(P).Key <> 0 And P < UBound(MyKey.Data)
                    P = P + 1
                Loop
                J = J + 1
            End If
        Loop
        MyGeo.PolyLines(K).Layer = Layer
    Case "SPLINE"
        Dim Knots() As Single
        Dim cPts() As CadPoint
        Dim fPts() As CadPoint
        Do While P < UBound(MyKey.Data)
            If MyKey.Data(P).Key = 40 Then 'KNots
                ReDim Preserve Knots(i) As Single
                Knots(i) = MyKey.Data(P).Value
                i = i + 1
            End If
            If MyKey.Data(P).Key = 10 Then 'Control Points (the spline SHOULD touch these points) - - but I have.t got it figured out perfectly yet
                ReDim Preserve cPts(J) As CadPoint
                cPts(J).x = MyKey.Data(P).Value
                cPts(J).y = MyKey.Data(P + 1).Value
                J = J + 1
            End If
            If MyKey.Data(P).Key = 11 Then ' Fit Points - These constrain the spline to touch the control points
                ReDim Preserve fPts(K) As CadPoint
                fPts(K).x = MyKey.Data(P).Value
                fPts(K).y = MyKey.Data(P + 1).Value
                K = K + 1
            End If
            P = P + 1
        Loop
        '-------------Make Spline-----------------------
        P = 0
        K = UBound(MyGeo.Splines) + 1
        ReDim Preserve MyGeo.Splines(K) As CadSpline
        For J = 0 To UBound(fPts)
            ReDim Preserve MyGeo.Splines(K).Vertex(P) As CadPoint
            MyGeo.Splines(K).Vertex(P) = fPts(J)
            P = P + 1
        Next J
        MyGeo.Splines(K).Layer = Layer
    Case "TEXT", "MTEXT"
        K = UBound(MyGeo.Text) + 1
        ReDim Preserve MyGeo.Text(K) As CadText
        MyGeo.Text(K).Start.x = kVal(MyKey.Data(), 10)
        MyGeo.Text(K).Start.y = kVal(MyKey.Data(), 20)
        MyGeo.Text(K).Size = kVal(MyKey.Data(), 40)
        MyGeo.Text(K).Angle = kVal(MyKey.Data(), 50)
        MyGeo.Text(K).Text = kVal(MyKey.Data(), 1)
        If MyKey.Type = "MTEXT" Then MyGeo.Text(K).Text = Right(MyGeo.Text(K).Text, Len(MyGeo.Text(K).Text) - 4)
        MyGeo.Text(K).Length = 1
        MyGeo.Text(K).Layer = Layer
    Case "INSERT"
        K = UBound(MyGeo.Inserts) + 1
        ReDim Preserve MyGeo.Inserts(K) As CadInsert
        MyGeo.Inserts(K).Name = kVal(MyKey.Data(), 2)
        MyGeo.Inserts(K).Base.x = kVal(MyKey.Data(), 10)
        MyGeo.Inserts(K).Base.y = kVal(MyKey.Data(), 20)
        MyGeo.Inserts(K).ScaleX = kVal(MyKey.Data(), 41)
        MyGeo.Inserts(K).ScaleY = kVal(MyKey.Data(), 42)
        MyGeo.Inserts(K).Angle = kVal(MyKey.Data(), 50)
        MyGeo.Inserts(K).Layer = Layer
    Case "DIMENSION"
        K = UBound(MyGeo.Inserts) + 1
        ReDim Preserve MyGeo.Inserts(K) As CadInsert
        MyGeo.Inserts(K).Name = kVal(MyKey.Data(), 2)
        MyGeo.Inserts(K).Layer = Layer
End Select
ReDim MyKey.Data(0) As KeySet
Exit Sub
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function CircleIntersectRect(SelLine As CadLine, MyArc As CadArc, ret As Integer) As Boolean

'CONVERTED FROM:
'Fast Circle-Rectangle Intersection Checking
'by Clifford A. Shaffer
'from "Graphics Gems", Academic Press, 1990


' Return TRUE iff rectangle R intersects
' circle with centerpoint C and
' radius Rad.



Dim r As CadLine

Dim Rad2 As Double

Rad2 = MyArc.Radius ^ 2

' Translate coordinates, placing C at the origin.
r = SelLine
r.P2.x = r.P2.x - MyArc.Center.x: r.P2.y = r.P2.y - MyArc.Center.y
r.P1.x = r.P1.x - MyArc.Center.x: r.P1.y = r.P1.y - MyArc.Center.y


If (r.P2.x < 0) Then ' /* R to left of circle center */
' Exit Function
If (r.P2.y < 0) Then '/* R in lower left corner */
CircleIntersectRect = ((r.P2.x * r.P2.x + r.P2.y * r.P2.y) < Rad2)
ret = 1
ElseIf (r.P1.y > 0) Then '/* R in upper left corner */
CircleIntersectRect = ((r.P2.x * r.P2.x + r.P1.y * r.P1.y) < Rad2)
ret = 2
Else ' /* R due West of circle */
CircleIntersectRect = (Abs(r.P2.x) < MyArc.Radius)
ret = 3
End If

ElseIf (r.P1.x > 0) Then ' /* R to right of circle center */

If (r.P2.y < 0) Then ' /* R in lower right corner */
CircleIntersectRect = ((r.P1.x * r.P1.x + r.P2.y * r.P2.y) < Rad2)
ret = 4
ElseIf (r.P1.y > 0) Then ' /* R in upper right corner */
CircleIntersectRect = ((r.P1.x * r.P1.x + r.P1.y + r.P1.y) < Rad2)
ret = 5
Else ' /* R due East of circle */
CircleIntersectRect = (r.P1.x < MyArc.Radius)
ret = 6
End If
Else ' /* R on circle vertical centerline */
If (r.P2.y < 0) Then ' /* R due South of circle */
CircleIntersectRect = (Abs(r.P2.y) < MyArc.Radius)
ret = 7
ElseIf (r.P1.y > 0) Then ' /* R due North of circle */
CircleIntersectRect = (r.P1.y < MyArc.Radius)
ret = 8
Else ' /* R contains circle centerpoint */
CircleIntersectRect = True
ret = 9
End If
End If


'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function






Function VPt(x As Single, y As Single) As CadPoint
VPt.x = x
VPt.y = y
End Function

Sub ZZ_OLD_SplinePoints(MySPline As CadSpline, sPoints() As CadPoint)
'Special thanks to: Warzi for sections of this routine.
Erase sPoints()
Dim a As CadPoint, B As CadPoint, c As CadPoint, D As CadPoint
Dim K As Double, J As Integer, n As Integer, P As Integer
Dim x0 As Single, y0 As Single
Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim x3 As Single, y3 As Single
Dim i As Integer
n = UBound(MySPline.Vertex)
If n = 0 Then ReDim sPoints(0) As CadPoint: sPoints(0) = MySPline.Vertex(0): Exit Sub
a = MySPline.Vertex(0) ': B = A: C = B: D = C
For i = -3 To n
    If i >= 0 Then a = MySPline.Vertex(i)
    If i + 1 <= n And i + 1 >= 0 Then B = MySPline.Vertex(i + 1) Else B = a
    If i + 2 <= n And i + 1 >= 0 Then c = MySPline.Vertex(i + 2) Else c = B
    If i + 3 <= n And i + 1 >= 0 Then D = MySPline.Vertex(i + 3) Else D = c
    x0 = (a.x + 4 * B.x + c.x) / 6!
    y0 = (a.y + 4 * B.y + c.y) / 6!
    X1 = (c.x - a.x) / 2!
    Y1 = (c.y - a.y) / 2!
    X2 = (a.x - 2 * B.x + c.x) / 2!
    Y2 = (a.y - 2 * B.y + c.y) / 2!
    x3 = (-a.x + 3 * (B.x - c.x) + D.x) / 6!
    y3 = (-a.y + 3 * (B.y - c.y) + D.y) / 6!
    
    For J = 0 To n
        ReDim Preserve sPoints(P) As CadPoint
        K = J / n
        sPoints(P).x = ((x3 * K + X2) * K + X1) * K + x0
        sPoints(P).y = ((y3 * K + Y2) * K + Y1) * K + y0
        If P = 0 Then P = P + 1 Else If sPoints(P).x <> sPoints(P - 1).x And sPoints(P).y <> sPoints(P - 1).y Then P = P + 1
    Next J
Next i
ReDim Preserve sPoints(P - 1) As CadPoint
sPoints(P - 1) = MySPline.Vertex(n)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub




Function kVal(Data() As KeySet, Key As String) As Variant
Dim i As Integer
For i = 0 To UBound(Data)
    If Data(i).Key = Key Then
        kVal = Data(i).Value
        Exit Function
    End If
Next i
kVal = 0
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function PtLine(PtA As CadPoint, PtB As CadPoint) As CadLine
PtLine.P1 = PtA
PtLine.P2 = PtB
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function RotateInsert(MyInsert As CadInsert, Pivot As CadPoint, Angle As Single) As CadInsert
RotateInsert = MyInsert
RotateInsert.Angle = MyInsert.Angle + Angle
RotateInsert.Base = RotatePoint(MyInsert.Base, Pivot, Angle)
'RotateInsert.Base = MovePoint(MyInsert.Base, Pivot.X - MyInsert.Base.X, Pivot.y - MyInsert.Base.y)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ScaleInsert(MyInsert As CadInsert, ScaleX As Single, ScaleY As Single) As CadInsert
ScaleInsert = MyInsert
ScaleInsert.ScaleX = MyInsert.ScaleX * ScaleX
ScaleInsert.ScaleY = MyInsert.ScaleY * ScaleY
ScaleInsert.Base = ScalePoint(MyInsert.Base, ScaleX, ScaleY)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function SearchSection(sArray() As String, Start As Long, Value As String) As Long
Dim i As Long
For i = Start To UBound(sArray)
    If sArray(i) = Value Then
        SearchSection = i
        Exit Function
    End If
Next i
SearchSection = -1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function FindStart(sArray() As String, Start As Long, Version As String)
Dim i As Long
Select Case Version
    Case "AC1012", "AC1013", "AC1014"
        i = SearchSection(sArray(), Start, "100") + 1
        i = SearchSection(sArray(), i, "100") + 2
        FindStart = i
        Exit Function
Case Else
    For i = Start To UBound(sArray)
        If sArray(i) = "10" Then
            FindStart = i
            Exit Function
        End If
    Next i
End Select
FindStart = -1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function IsDXFCommand(InText As String)
Select Case UCase(InText)
    Case "POINT", "LINE", "VERTEX", "POLYLINE", "CIRCLE", "ARC", "ELLIPSE", "TEXT", "INSERT", "DIMENSION", "SPLINE", "MTEXT"
        'These are the basic ENTITY COMMANDS available in the DXF language
        IsDXFCommand = True
    Case Else
        IsDXFCommand = False
End Select
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function GetSection(FileNum As Integer, Start As String, Finish As String, EndString As String, sArray() As String) As Boolean
ReDim sArray(0) As String
Dim Temp As String
Dim i As Long
Do While Temp <> Start
    Line Input #FileNum, Temp
    Temp = UCase(Trim(Temp))
    If Temp = EndString Then
        GetSection = False
        Exit Function
    End If
Loop
Do While Temp <> Finish
    Line Input #FileNum, Temp
    Temp = UCase(Trim(Temp))
    If Temp <> Finish Then
        ReDim Preserve sArray(i) As String
        sArray(i) = Temp
        i = i + 1
    End If
Loop
GetSection = True
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub FindCommand(FileNum As Integer, Command As String)
Dim x As String
Do While UCase(Trim(x)) <> UCase(Command)
    Line Input #FileNum, x
Loop
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub SelectAllGeo(ByRef MySel() As SelSet, MyGeo As Geometry)
On Error GoTo eTrap
Erase MySel()
Dim i As Integer
Dim J As Integer
Dim K As Integer
'-------------------------
K = UBound(MyGeo.Points)
For i = 0 To K
    ReDim Preserve MySel(J) As SelSet
    MySel(J).Type = "Point"
    MySel(J).Index = i
    J = J + 1
Next i
'-------------------------
K = UBound(MyGeo.Lines)
For i = 0 To K
    ReDim Preserve MySel(J) As SelSet
    MySel(J).Type = "Line"
    MySel(J).Index = i
    J = J + 1
Next i
'-------------------------
K = UBound(MyGeo.Arcs)
For i = 0 To K
    ReDim Preserve MySel(J) As SelSet
    MySel(J).Type = "Arc"
    MySel(J).Index = i
    J = J + 1
Next i
'-------------------------
K = UBound(MyGeo.Ellipses)
For i = 0 To K
    ReDim Preserve MySel(J) As SelSet
    MySel(J).Type = "Ellipse"
    MySel(J).Index = i
    J = J + 1
Next i
'-------------------------
K = UBound(MyGeo.Splines)
For i = 0 To K
    ReDim Preserve MySel(J) As SelSet
    MySel(J).Type = "Spline"
    MySel(J).Index = i
    J = J + 1
Next i
'-------------------------
K = UBound(MyGeo.PolyLines)
For i = 0 To K
    ReDim Preserve MySel(J) As SelSet
    MySel(J).Type = "PolyLine"
    MySel(J).Index = i
    J = J + 1
Next i
'-------------------------
K = UBound(MyGeo.Text)
For i = 0 To K
    ReDim Preserve MySel(J) As SelSet
    MySel(J).Type = "Text"
    MySel(J).Index = i
    J = J + 1
Next i
K = UBound(MyGeo.Inserts)
For i = 0 To K
    ReDim Preserve MySel(J) As SelSet
    MySel(J).Type = "Insert"
    MySel(J).Index = i
    J = J + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub AddPoints(ByRef Source() As CadPoint, ByRef Dest() As CadPoint)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim P As Integer
K = UBound(Source) ' if source is empty, the error will stop the stansfer at the for-loop
P = K
K = UBound(Dest) ' if dest is empty, it will start at array zero
K = K + 1
For i = 0 To P
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub


Sub AddInserts(ByRef Source() As CadInsert, ByRef Dest() As CadInsert)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim P As Integer
K = UBound(Source) ' if source is empty, the error will stop the stansfer at the for-loop
P = K
K = UBound(Dest) ' if dest is empty, it will start at array zero
K = K + 1
For i = 0 To P
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub AddFaces(ByRef Source() As CadFace, ByRef Dest() As CadFace)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim P As Integer
K = UBound(Source) ' if source is empty, the error will stop the stansfer at the for-loop
P = K
K = UBound(Dest) ' if dest is empty, it will start at array zero
K = K + 1
For i = 0 To P
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub AddArcs(ByRef Source() As CadArc, ByRef Dest() As CadArc)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim P As Integer
K = UBound(Source) ' if source is empty, the error will stop the stansfer at the for-loop
P = K
K = UBound(Dest) ' if dest is empty, it will start at array zero
K = K + 1
For i = 0 To P
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub AddSelection(ByRef Source() As SelSet, ByRef Dest() As SelSet)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim J As Integer
Dim Match As Boolean
K = UBound(Source) ' if source is empty, the error will stop the stansfer at the for-loop
J = K
K = UBound(Dest) ' if dest is empty, it will start at array zero
K = K + 1
For i = 0 To J
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
'-----------remove duplicates
K = UBound(Dest)
For i = 0 To K - 1
    For J = i + 1 To K
        If Dest(i).Type = Dest(J).Type And Dest(i).Index = Dest(J).Index Then Dest(J).Type = "BLANK"
    Next J
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub AddLines(ByRef Source() As CadLine, ByRef Dest() As CadLine)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim P As Integer
K = UBound(Source)
P = K
K = UBound(Dest)
K = K + 1
For i = 0 To P
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub AddEllipses(ByRef Source() As CadEllipse, ByRef Dest() As CadEllipse)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim P As Integer
K = UBound(Source)
P = K
K = UBound(Dest)
K = K + 1
For i = 0 To P
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub AddSplines(ByRef Source() As CadSpline, ByRef Dest() As CadSpline)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim P As Integer
K = UBound(Source)
P = K
K = UBound(Dest)
K = K + 1
For i = 0 To P
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub AddPolyLines(ByRef Source() As CadPolyLine, ByRef Dest() As CadPolyLine)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim P As Integer
K = UBound(Source)
P = K
K = UBound(Dest)
K = K + 1
For i = 0 To P
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub AddText(ByRef Source() As CadText, ByRef Dest() As CadText)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim P As Integer
K = UBound(Source)
P = K
K = UBound(Dest)
K = K + 1
For i = 0 To P
    ReDim Preserve Dest(K)
    Dest(K) = Source(i)
    K = K + 1
Next i
Exit Sub
eTrap:
    K = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function ArcArcIntersect(ArcA As CadArc, ArcB As CadArc, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
'returns -1 if the arcs do not intersect
'Returns 0 if the arcs intersect at 1 point
'Returns 1 if the arcs intersect at 2 points
Dim a As Single ' This is the distance from the center of ArcA to the midpoint vetween the intersection
Dim B As Single ' This is the distance from the center of ArcB to the midpoint vetween the intersection
Dim D As Single ' This is the distance between 2 centers of the arcs
Dim h As Single ' This is the distance from the base line to the intersecting point
Dim X2 As Single 'The Xpoint between the
Dim Y2 As Single 'The Ypoint between the intersection
D = Sqr((ArcB.Center.y - ArcA.Center.y) ^ 2 + (ArcB.Center.x - ArcA.Center.x) ^ 2)
If D = 0 Then ArcArcIntersect = -1: Exit Function
'-----------If the arcs are too far apart--------
If (ArcA.Radius + ArcB.Radius) < D Then ArcArcIntersect = -1: Exit Function
'------------If the Arcs are too close together------------------------
If D + ArcB.Radius < ArcA.Radius Then ArcArcIntersect = -1: Exit Function
If D + ArcA.Radius < ArcB.Radius Then ArcArcIntersect = -1: Exit Function
'-----------------------------------------------------------
a = ((ArcA.Radius ^ 2) - (ArcB.Radius ^ 2) + (D ^ 2)) / (2 * D)
h = Sqr(Abs((ArcA.Radius ^ 2) - (a ^ 2)))
X2 = ArcA.Center.x + a * ((ArcB.Center.x - ArcA.Center.x) / D)
Y2 = ArcA.Center.y + a * ((ArcB.Center.y - ArcA.Center.y) / D)
If ArcA.Radius = a Then  'circles touch at 1 point (tangent)
    ReDim iPoints(0) As CadPoint
    iPoints(0).x = X2 + h * (ArcB.Center.y - ArcA.Center.y) / D
    iPoints(0).y = Y2 - h * (ArcB.Center.x - ArcA.Center.x) / D
    iPoints(0).Layer.Width = 3
    iPoints(0).Layer.Color = vbBlue
Else ' 2 intersection points
    ReDim iPoints(1) As CadPoint
    iPoints(0).x = X2 + h * (ArcB.Center.y - ArcA.Center.y) / D
    iPoints(0).y = Y2 - h * (ArcB.Center.x - ArcA.Center.x) / D
    iPoints(1).x = X2 - h * (ArcB.Center.y - ArcA.Center.y) / D
    iPoints(1).y = Y2 + h * (ArcB.Center.x - ArcA.Center.x) / D
    iPoints(0).Layer.Width = 3
    iPoints(1).Layer.Width = 3
    iPoints(0).Layer.Color = vbBlue
    iPoints(1).Layer.Color = vbBlue
End If
ArcArcIntersect = UBound(iPoints)
'--------------Check for imaginary intersection points------
Dim i As Integer
For i = 0 To UBound(iPoints)
    If Not InsideAngles(ArcPtAngle(ArcA, iPoints(i)), ArcA.Angle1, ArcA.Angle2) Then iPoints(i).Layer.Color = vbGreen
Next i
For i = 0 To UBound(iPoints)
    If Not InsideAngles(ArcPtAngle(ArcB, iPoints(i)), ArcB.Angle1, ArcB.Angle2) Then
        If iPoints(i).Layer.Color <> vbBlue Then iPoints(i).Layer.Color = vbRed Else iPoints(i).Layer.Color = vbGreen
    End If
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Function ArcLen(MyArc As CadArc) As Single
If MyArc.Angle1 > MyArc.Angle2 Then
    ArcLen = ((MyArc.Angle2 + 360 - MyArc.Angle1) / 360) * 2 * Pi * MyArc.Radius
Else
    ArcLen = ((MyArc.Angle2 - MyArc.Angle1) / 360) * 2 * Pi * MyArc.Radius
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function ArcPt(MyArc As CadArc, MyAngle As Single) As CadPoint
Dim a1 As Single, a2 As Single
ArcPt.x = MyArc.Center.x + (MyArc.Radius * Cos(MyAngle * Pi / 180))
ArcPt.y = MyArc.Center.y + (MyArc.Radius * Sin(MyAngle * Pi / 180))
ArcPt.Layer.Color = vbBlue
ArcPt.Layer.Width = 3
a1 = MyArc.Angle1
a2 = MyArc.Angle2
If a2 < a1 Then a2 = a2 + (360)
If MyAngle < a1 Then
    If a2 > a1 Then
        ArcPt.Layer.Color = vbRed
    ElseIf MyAngle > a2 Then
        ArcPt.Layer.Color = vbRed
    End If
ElseIf MyAngle > a1 And MyAngle > a2 And a2 > a1 Then
    ArcPt.Layer.Color = vbRed
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ArcPtAngle(MyArc As CadArc, MyPoint As CadPoint)
Dim tLine As CadLine
tLine.P1 = MyArc.Center
tLine.P2 = MyPoint
ArcPtAngle = cAngle(tLine)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Function ArcToEllipse(MyArc As CadArc) As CadEllipse
ArcToEllipse.F1 = MyArc.Center
ArcToEllipse.F2 = MyArc.Center
ArcToEllipse.P1.x = MyArc.Center.x + MyArc.Radius
ArcToEllipse.P1.y = MyArc.Center.y
If MyArc.Angle1 = 0 And MyArc.Angle2 = 360 Then
    ArcToEllipse.Angle1 = 0
    ArcToEllipse.Angle2 = 360
Else
    ArcToEllipse.Angle1 = EllipseAngle(MyArc.Angle1, ArcToEllipse)
    ArcToEllipse.Angle2 = EllipseAngle(MyArc.Angle2, ArcToEllipse)
End If
ArcToEllipse.NumPoints = 32
ArcToEllipse.Layer = MyArc.Layer
'DrawCadEllipse frmDraw.picDraw, ArcToEllipse, , vbGreen, 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ASIN(x As Single)
ASIN = Atn(x / Sqr(-x * x + 1))
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ATAN(x As Single)
ATAN = Atn(x)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function BoxLineIntersect(SelLine As CadLine, MyLine As CadLine) As Boolean
Dim iPt() As CadPoint
Dim tLine As CadLine
tLine.P1 = SelLine.P1
tLine.P2.x = SelLine.P1.x
tLine.P2.y = SelLine.P2.y
If LineLineIntersect(tLine, MyLine, iPt()) <> -1 Then If iPt(0).Layer.Color = vbBlue Then BoxLineIntersect = True: Exit Function
tLine.P1 = SelLine.P1
tLine.P2.y = SelLine.P1.y
tLine.P2.x = SelLine.P2.x
If LineLineIntersect(tLine, MyLine, iPt()) <> -1 Then If iPt(0).Layer.Color = vbBlue Then BoxLineIntersect = True: Exit Function
tLine.P1 = SelLine.P2
tLine.P2.x = SelLine.P1.x
tLine.P2.y = SelLine.P2.y
If LineLineIntersect(tLine, MyLine, iPt()) <> -1 Then If iPt(0).Layer.Color = vbBlue Then BoxLineIntersect = True: Exit Function
tLine.P1 = SelLine.P2
tLine.P2.y = SelLine.P1.y
tLine.P2.x = SelLine.P2.x
If LineLineIntersect(tLine, MyLine, iPt()) <> -1 Then If iPt(0).Layer.Color = vbBlue Then BoxLineIntersect = True: Exit Function
BoxLineIntersect = False
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function BoxArcIntersect(SelLine As CadLine, MyArc As CadArc) As Boolean
Dim iPts() As CadPoint
Dim tLine As CadLine
Dim i As Integer
Dim iRes As Integer
tLine.P1 = SelLine.P1
tLine.P2.x = SelLine.P1.x
tLine.P2.y = SelLine.P2.y
'DrawCadLine frmDraw.picDraw, tLine, , vbGreen, 2
iRes = LineArcIntersect(tLine, MyArc, iPts())
For i = 0 To iRes
    If iPts(i).Layer.Color = vbBlue Then BoxArcIntersect = True: Exit Function
Next i
tLine.P1 = SelLine.P1
tLine.P2.y = SelLine.P1.y
tLine.P2.x = SelLine.P2.x
'DrawCadLine frmDraw.picDraw, tLine, , vbGreen, 2
iRes = LineArcIntersect(tLine, MyArc, iPts())
For i = 0 To iRes
    If iPts(i).Layer.Color = vbBlue Then BoxArcIntersect = True: Exit Function
Next i
tLine.P1 = SelLine.P2
tLine.P2.x = SelLine.P1.x
tLine.P2.y = SelLine.P2.y
'DrawCadLine frmDraw.picDraw, tLine, , vbGreen, 2
iRes = LineArcIntersect(tLine, MyArc, iPts())
For i = 0 To iRes
    If iPts(i).Layer.Color = vbBlue Then BoxArcIntersect = True: Exit Function
Next i
tLine.P1 = SelLine.P2
tLine.P2.y = SelLine.P1.y
tLine.P2.x = SelLine.P2.x
'DrawCadLine frmDraw.picDraw, tLine, , vbGreen, 2
iRes = LineArcIntersect(tLine, MyArc, iPts())
For i = 0 To iRes
    If iPts(i).Layer.Color = vbBlue Then BoxArcIntersect = True: Exit Function
Next i
BoxArcIntersect = False
Exit Function
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function BoxEllipseIntersect(SelLine As CadLine, MyEllipse As CadEllipse) As Boolean
Dim iPts() As CadPoint
Dim tLine As CadLine
Dim i As Integer
Dim iRes As Integer
tLine.P1 = SelLine.P1
tLine.P2.x = SelLine.P1.x
tLine.P2.y = SelLine.P2.y
iRes = LineEllipseIntersect(tLine, MyEllipse, iPts())
For i = 0 To iRes
    If iPts(i).Layer.Color = vbBlue Then BoxEllipseIntersect = True: Exit Function
Next i
tLine.P1 = SelLine.P1
tLine.P2.y = SelLine.P1.y
tLine.P2.x = SelLine.P2.x
iRes = LineEllipseIntersect(tLine, MyEllipse, iPts())
For i = 0 To iRes
    If iPts(i).Layer.Color = vbBlue Then BoxEllipseIntersect = True: Exit Function
Next i
tLine.P1 = SelLine.P2
tLine.P2.x = SelLine.P1.x
tLine.P2.y = SelLine.P2.y
iRes = LineEllipseIntersect(tLine, MyEllipse, iPts())
For i = 0 To iRes
    If iPts(i).Layer.Color = vbBlue Then BoxEllipseIntersect = True: Exit Function
Next i
tLine.P1 = SelLine.P2
tLine.P2.y = SelLine.P1.y
tLine.P2.x = SelLine.P2.x
iRes = LineEllipseIntersect(tLine, MyEllipse, iPts())
For i = 0 To iRes
    If iPts(i).Layer.Color = vbBlue Then BoxEllipseIntersect = True: Exit Function
Next i
BoxEllipseIntersect = False
Exit Function
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function BoxPolyLineIntersect(SelLine As CadLine, MyPolyLine As CadPolyLine) As Boolean
Dim i As Integer
Dim tLine As CadLine
For i = 1 To UBound(MyPolyLine.Vertex)
    tLine.P1 = MyPolyLine.Vertex(i - 1)
    tLine.P2 = MyPolyLine.Vertex(i)
    If BoxLineIntersect(SelLine, tLine) Then
        BoxPolyLineIntersect = True
        Exit Function
    End If
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub BreakLine(ByRef MyGeo As Geometry, LineNum As Integer, bPt As CadPoint)
Dim i As Integer
i = UBound(MyGeo.Lines) + 1
ReDim Preserve MyGeo.Lines(i) As CadLine
MyGeo.Lines(i) = MyGeo.Lines(LineNum)
MyGeo.Lines(LineNum).P2 = bPt
MyGeo.Lines(i).P1 = bPt
Debug.Print bPt.x
Debug.Print bPt.y
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub BreakArc(ByRef MyGeo As Geometry, ArcNum As Integer, bPt As CadPoint)
Dim i As Integer
Dim bAng As Single
bAng = PtPtAngle(MyGeo.Arcs(ArcNum).Center, bPt)
If MyGeo.Arcs(ArcNum).Angle2 - MyGeo.Arcs(ArcNum).Angle1 = 360 Then
    MyGeo.Arcs(ArcNum).Angle1 = dAngle(bAng + 180)
    MyGeo.Arcs(ArcNum).Angle2 = dAngle(bAng + 180 + 360)
End If
i = UBound(MyGeo.Arcs) + 1
ReDim Preserve MyGeo.Arcs(i) As CadArc
MyGeo.Arcs(i) = MyGeo.Arcs(ArcNum)
MyGeo.Arcs(ArcNum).Angle2 = bAng
MyGeo.Arcs(i).Angle1 = bAng
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub BreakEllipse(ByRef MyGeo As Geometry, EllipseNum As Integer, bPt As CadPoint)
Dim i As Integer
Dim bAng As Single
bAng = EllipseAngle(PtPtAngle(EllipseCenter(MyGeo.Ellipses(EllipseNum)), bPt), MyGeo.Ellipses(EllipseNum))
If MyGeo.Ellipses(EllipseNum).Angle2 - MyGeo.Ellipses(EllipseNum).Angle1 = 360 Then
    MyGeo.Ellipses(EllipseNum).Angle1 = dAngle(bAng + 180)
    MyGeo.Ellipses(EllipseNum).Angle2 = dAngle(bAng + 180 + 360)
End If
i = UBound(MyGeo.Ellipses) + 1
ReDim Preserve MyGeo.Ellipses(i) As CadEllipse
MyGeo.Ellipses(i) = MyGeo.Ellipses(EllipseNum)
MyGeo.Ellipses(EllipseNum).Angle2 = bAng
MyGeo.Ellipses(i).Angle1 = bAng
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub BreakSpline(ByRef MyGeo As Geometry, SplineNum As Integer, bPt As CadPoint)
On Error GoTo eTrap
'The only way to break as spline is to convert it into a polyline first
Dim tPolyLine As CadPolyLine
Dim K As Integer
tPolyLine = SplineToPolyLine(MyGeo.Splines(SplineNum))
RemoveGeo MyGeo, "Spline", SplineNum
K = UBound(MyGeo.PolyLines) + 1
ReDim Preserve MyGeo.PolyLines(K) As CadPolyLine
MyGeo.PolyLines(K) = tPolyLine
BreakPolyLine MyGeo, K, bPt
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub BreakPolyLine(ByRef MyGeo As Geometry, PolyLineNum As Integer, bPt As CadPoint)
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim tLine As CadLine
Dim tPolyLine As CadPolyLine
Dim Found As Boolean
i = UBound(MyGeo.PolyLines) + 1
ReDim Preserve MyGeo.PolyLines(i) As CadPolyLine
MyGeo.PolyLines(i) = MyGeo.PolyLines(PolyLineNum)
tPolyLine = MyGeo.PolyLines(PolyLineNum)
tLine.P1 = tPolyLine.Vertex(0)
ReDim Preserve MyGeo.PolyLines(PolyLineNum).Vertex(0) As CadPoint
ReDim Preserve MyGeo.PolyLines(i).Vertex(1) As CadPoint
MyGeo.PolyLines(PolyLineNum).Vertex(0) = tPolyLine.Vertex(0)
For J = 1 To UBound(tPolyLine.Vertex)
    tLine.P2 = tPolyLine.Vertex(J)
    If Not Found Then
        ReDim Preserve MyGeo.PolyLines(PolyLineNum).Vertex(J) As CadPoint
        MyGeo.PolyLines(PolyLineNum).Vertex(J) = tPolyLine.Vertex(J)
    Else
        ReDim Preserve MyGeo.PolyLines(i).Vertex(J - K) As CadPoint
        MyGeo.PolyLines(i).Vertex(J - K) = tPolyLine.Vertex(J)
    End If
    If PtInLine(tLine, bPt) And Not Found Then 'And (cAngle(tLine) = PtPtAngle(tLine.P1, bPt)) Then
        MyGeo.PolyLines(PolyLineNum).Vertex(J) = bPt
        MyGeo.PolyLines(i).Vertex(0) = bPt
        MyGeo.PolyLines(i).Vertex(1) = tPolyLine.Vertex(J)
        K = J - 1
        Found = True
    End If
    tLine.P1 = tLine.P2
Next J
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function CheckArcStart(ByRef MyArc As CadArc, MPt As CadPoint) As Boolean
If MyArc.Angle2 - MyArc.Angle1 = 360 Or MyArc.Angle2 - MyArc.Angle1 = 0 Then
    Dim ang1 As Single
    ang1 = PtPtAngle(MyArc.Center, MPt)
    MyArc.Angle1 = ang1
    MyArc.Angle2 = ang1 + 360
    CheckArcStart = True
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function CheckEllipseStart(ByRef MyEllipse As CadEllipse, MPt As CadPoint) As Boolean
If MyEllipse.Angle2 - MyEllipse.Angle1 = 360 Or MyEllipse.Angle2 - MyEllipse.Angle1 = 0 Then
    Dim ang1 As Single
    ang1 = PtPtAngle(EllipseCenter(MyEllipse), MPt)
    MyEllipse.Angle1 = ang1
    MyEllipse.Angle2 = ang1 + 360
    CheckEllipseStart = True
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Sub ClearGeo(MyGeo() As Geometry)
Erase MyGeo()
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub CornerGroup(ByRef MyGeo As Geometry, MySel() As SelSet)
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim M As Integer
Dim E1Pts() As CadPoint
Dim E2Pts() As CadPoint
Dim iPts() As CadPoint
Dim E1 As Integer
Dim E2 As Integer
Dim cSet(1) As SelSet
For i = 0 To UBound(MySel) - 1
    For J = i + 1 To UBound(MySel)
        GeoSelEndPoints MyGeo, MySel(i), E1Pts()
        GeoSelEndPoints MyGeo, MySel(J), E2Pts()
        If GeoIntersect(MyGeo, MySel(i), MySel(J), iPts()) >= 0 Then
            For K = 0 To UBound(iPts)
                E1 = ClosestPoint(iPts(K), E1Pts())
                E2 = ClosestPoint(iPts(K), E2Pts())
                If iPts(K).Layer.Color = vbBlue Then
                    If MySel(i).Type = "Spline" Then
                        For M = 0 To UBound(MySel)
                            If MySel(M).Type = "Spline" And MySel(M).Index > MySel(i).Index Then MySel(M).Index = MySel(M).Index - 1
                        Next M
                    End If
                    RelimitSelection MyGeo, MySel(i), E1, iPts(K)
                    If E1 = 0 And MySel(i).Type = "PolyLine" Then
                        For M = 0 To UBound(MySel)
                            If MySel(M).Type = "PolyLine" And MySel(M).Index > MySel(i).Index Then MySel(M).Index = MySel(M).Index - 1
                        Next M
                        MySel(i).Index = UBound(MyGeo.PolyLines)
                    End If
                    If MySel(J).Type = "Spline" Then
                        For M = 0 To UBound(MySel)
                            If MySel(M).Type = "Spline" And MySel(M).Index > MySel(J).Index Then MySel(M).Index = MySel(M).Index - 1
                        Next M
                    End If
                    RelimitSelection MyGeo, MySel(J), E2, iPts(K)
                    If E2 = 0 And MySel(J).Type = "PolyLine" Then
                        For M = 0 To UBound(MySel)
                            If MySel(M).Type = "PolyLine" And MySel(M).Index > MySel(J).Index Then MySel(M).Index = MySel(M).Index - 1
                        Next M
                        MySel(J).Index = UBound(MyGeo.PolyLines)
                    End If
                End If
            Next K
        End If
    Next J
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function cPtLine(PtA As CadPoint, PtB As CadPoint) As CadLine
cPtLine.P1 = PtA
cPtLine.P2 = PtB
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function Direction(P1 As CadPoint, P2 As CadPoint, P3 As CadPoint) As Integer
'-1 = CCW
'0 = Straight
'1 = CW
Dim CPt As CadPoint
If DelaunayCenter(P1, P2, P3, CPt) = -1 Then Exit Function
If PtPtAngle(CPt, P3) - PtPtAngle(CPt, P1) > 0 Then Direction = 1 Else Direction = -1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub DivideSelection(ByRef MyGeo As Geometry, MySel As SelSet, Count As Integer)
Dim i As Integer
Dim dPoints() As CadPoint
Dim sAngles() As Single
i = MySel.Index
Select Case MySel.Type
    Case "Line": SpaceLine MyGeo.Lines(i), MyGeo.Lines(i).P1, LineLen(MyGeo.Lines(i)) / Count, Count + 1, MyGeo.Lines(i).P2, dPoints(), sAngles()
    Case "Arc": SpaceArc MyGeo.Arcs(i), ArcPt(MyGeo.Arcs(i), MyGeo.Arcs(i).Angle1), ArcLen(MyGeo.Arcs(i)) / Count, Count + 1, ArcPt(MyGeo.Arcs(i), MyGeo.Arcs(i).Angle1 + 1), dPoints(), sAngles()
    Case "Ellipse": SpaceEllipse MyGeo.Ellipses(i), EllipsePt(MyGeo.Ellipses(i), MyGeo.Ellipses(i).Angle1), EllipseLen(MyGeo.Ellipses(i)) / Count, Count + 1, EllipsePt(MyGeo.Ellipses(i), MyGeo.Ellipses(i).Angle1 + 1), dPoints(), sAngles()
    Case "Spline": SpaceSpline MyGeo.Splines(i), MyGeo.Splines(i).Vertex(0), SplineLen(MyGeo.Splines(i)) / Count, Count + 1, SplineToPolyLine(MyGeo.Splines(i)).Vertex(1), dPoints(), sAngles()
    Case "PolyLine": SpacePolyLine MyGeo.PolyLines(i), MyGeo.PolyLines(i).Vertex(0), PolyLineLen(MyGeo.PolyLines(i)) / Count, Count + 1, MyGeo.PolyLines(i).Vertex(1), dPoints(), sAngles()
End Select
AddPoints dPoints(), MyGeo.Points()
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub DrawCadArc(Canvas As PictureBox, MyArc As CadArc, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
On Error Resume Next 'THis is only for times when the radius is zero
MyArc.Angle1 = dAngle(MyArc.Angle1)
MyArc.Angle2 = dAngle(MyArc.Angle2)
Dim i As Single
Dim interval As Single
If Width = 0 Then Width = MyArc.Layer.Width
Canvas.DrawWidth = Width
Canvas.DrawMode = Mode
Canvas.DrawStyle = MyArc.Layer.style
If Color = -1 Then Color = MyArc.Layer.Color
If MyArc.Angle1 > MyArc.Angle2 Then
    If MyArc.Angle1 <> 360 Then Canvas.Circle (MyArc.Center.x, -MyArc.Center.y), MyArc.Radius, Color, MyArc.Angle1 * Pi / 180, 2 * Pi
    If MyArc.Angle2 <> 0 Then Canvas.Circle (MyArc.Center.x, -MyArc.Center.y), MyArc.Radius, Color, 0, MyArc.Angle2 * Pi / 180
Else
    'It's a good practice to ALWAYS split your arcs into sections
    interval = IIf((MyArc.Angle2 - MyArc.Angle1) / Pi > 0.1, (MyArc.Angle2 - MyArc.Angle1) / Pi, Pi)
    For i = MyArc.Angle1 To MyArc.Angle2 - interval Step interval
        Canvas.Circle (MyArc.Center.x, -MyArc.Center.y), MyArc.Radius, Color, i * Pi / 180, (i + interval) * Pi / 180
    Next i
    Canvas.Circle (MyArc.Center.x, -MyArc.Center.y), MyArc.Radius, Color, i * Pi / 180, (MyArc.Angle2) * Pi / 180
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function EllipseLen(MyEllipse As CadEllipse) As Single
Dim a As Single, B As Single ' ellipse parameters
Dim u As Single, v As Single
Dim i As Single, J As Single
Dim ang1 As Single, ang2 As Single 'adjusted ellipse angles
Dim X1 As Single, Y1 As Single 'plotting coordinates
Dim cx As Single, cy As Single ' center of the ellipse
Dim First As Boolean
Dim TotLen As Single, Flen As Single
Dim tLine As CadLine
'-----------------Angle adjustment-------
ang1 = MyEllipse.Angle1
ang2 = MyEllipse.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
i = (ang2 - ang1) / MyEllipse.NumPoints
If i = 0 Then Exit Function
'-------------------------------------------
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2) * Pi / 180
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
If a = 0 Then Exit Function
For u = ang1 To ang2 Step i
    J = u * Pi / 180
    X1 = Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    Y1 = Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    tLine.P2.x = X1
    tLine.P2.y = Y1
    If First Then EllipseLen = EllipseLen + LineLen(tLine)
    First = True
    tLine.P1 = tLine.P2
Next u
If u - i < ang2 Then
    J = ang2 * Pi / 180
    X1 = Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    Y1 = Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    tLine.P2.x = X1
    tLine.P2.y = Y1
    EllipseLen = EllipseLen + LineLen(tLine)
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function EllipseLenPt(MyEllipse As CadEllipse, sAngle As Single, Dist As Single, Side As Integer) As CadPoint
Dim a As Single, B As Single ' ellipse parameters
Dim u As Single, v As Single
Dim i As Single, J As Single
Dim X1 As Single, Y1 As Single 'plotting coordinates
Dim cx As Single, cy As Single ' center of the ellipse
Dim First As Boolean
Dim TotLen As Single, Flen As Single
'------------------------------------
Dim tDist As Single
Dim tLine As CadLine
Dim M As Single
Dim eLen As Single
'-------------------------------------------
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2) * Pi / 180
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
If a = 0 Then Exit Function
'Get first point of line
J = sAngle * Pi / 180
tLine.P1.x = (Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))) + cx
tLine.P1.y = (Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))) + cy
Do While tDist < Dist
    tLine.P2.x = (Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))) + cx
    tLine.P2.y = (Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))) + cy
    eLen = LineLen(tLine)
    ' mathematically, you can only APPROXIMATE an ellipse-arc length using integration (such as used here)
    If tDist + eLen < Dist Then
        tDist = tDist + eLen
        tLine.P1 = tLine.P2
        J = J + (0.001 * Side)
    Else
        EllipseLenPt = tLine.P2
        Exit Do
    End If
Loop
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub ExplodePolyline(ByRef MyGeo As Geometry, SelNum As Integer)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
K = UBound(MyGeo.Lines) + 1
ReDim Preserve MyGeo.Lines(K + UBound(MyGeo.PolyLines(SelNum).Vertex) - 1)
For i = 0 To UBound(MyGeo.PolyLines(SelNum).Vertex) - 1
    MyGeo.Lines(i + K).P1 = MyGeo.PolyLines(SelNum).Vertex(i)
    MyGeo.Lines(i + K).P2 = MyGeo.PolyLines(SelNum).Vertex(i + 1)
    MyGeo.Lines(i + K).Layer = MyGeo.PolyLines(SelNum).Layer
Next i
RemoveGeo MyGeo, "PolyLine", SelNum
Exit Sub
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub ExplodeSpline(ByRef MyGeo As Geometry, SelNum As Integer)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim tPoly As CadPolyLine
tPoly = SplineToPolyLine(MyGeo.Splines(SelNum))
K = UBound(MyGeo.Lines) + 1
ReDim Preserve MyGeo.Lines(K + UBound(tPoly.Vertex) - 1)
For i = 0 To UBound(tPoly.Vertex) - 1
    MyGeo.Lines(i + K).P1 = tPoly.Vertex(i)
    MyGeo.Lines(i + K).P2 = tPoly.Vertex(i + 1)
    MyGeo.Lines(i + K).Layer = MyGeo.Splines(SelNum).Layer
Next i
RemoveGeo MyGeo, "Spline", SelNum
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub ExplodeEllipse(ByRef MyGeo As Geometry, SelNum As Integer)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim tPoly As CadPolyLine
tPoly = EllipseToPolyLine(MyGeo.Ellipses(SelNum))
K = UBound(MyGeo.Lines) + 1
ReDim Preserve MyGeo.Lines(K + UBound(tPoly.Vertex) - 1)
For i = 0 To UBound(tPoly.Vertex) - 1
    MyGeo.Lines(i + K).P1 = tPoly.Vertex(i)
    MyGeo.Lines(i + K).P2 = tPoly.Vertex(i + 1)
    MyGeo.Lines(i + K).Layer = MyGeo.Ellipses(SelNum).Layer
Next i
RemoveGeo MyGeo, "Ellipse", SelNum
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub ExplodeArc(ByRef MyGeo As Geometry, SelNum As Integer)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
Dim tPoly As CadPolyLine
tPoly = ArcToPolyLine(MyGeo.Arcs(SelNum))
K = UBound(MyGeo.Lines) + 1
ReDim Preserve MyGeo.Lines(K + UBound(tPoly.Vertex) - 1)
For i = 0 To UBound(tPoly.Vertex) - 1
    MyGeo.Lines(i + K).P1 = tPoly.Vertex(i)
    MyGeo.Lines(i + K).P2 = tPoly.Vertex(i + 1)
    MyGeo.Lines(i + K).Layer = MyGeo.Arcs(SelNum).Layer
Next i
RemoveGeo MyGeo, "Arc", SelNum
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub


Sub ExplodeLine(ByRef MyGeo As Geometry, SelNum As Integer)
On Error GoTo eTrap
Dim i As Integer
Dim K As Integer
K = UBound(MyGeo.Points) + 1
ReDim Preserve MyGeo.Points(K + 1)
MyGeo.Points(K) = MyGeo.Lines(SelNum).P1
MyGeo.Points(K + 1) = MyGeo.Lines(SelNum).P2
MyGeo.Points(K).Layer = MyGeo.Lines(SelNum).Layer
MyGeo.Points(K + 1).Layer = MyGeo.Lines(SelNum).Layer
RemoveGeo MyGeo, "Line", SelNum
Exit Sub
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub ExplodeSelection(MyGeo As Geometry, ExSel As SelSet)
Select Case ExSel.Type
    Case "Arc": ExplodeArc MyGeo, ExSel.Index
    Case "Ellipse": ExplodeEllipse MyGeo, ExSel.Index
    Case "Spline": ExplodeSpline MyGeo, ExSel.Index
    Case "PolyLine": ExplodePolyline MyGeo, ExSel.Index
End Select
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub


Sub GeoSelEndMidPoints(MyGeo As Geometry, MySel As SelSet, ByRef dPoints() As CadPoint)
Dim i As Integer
Dim sAngles() As Single
Dim Count As Integer
Count = 2
i = MySel.Index
Select Case MySel.Type
    Case "Line": SpaceLine MyGeo.Lines(i), MyGeo.Lines(i).P1, LineLen(MyGeo.Lines(i)) / Count, Count + 1, MyGeo.Lines(i).P2, dPoints(), sAngles()
    Case "Arc": SpaceArc MyGeo.Arcs(i), ArcPt(MyGeo.Arcs(i), MyGeo.Arcs(i).Angle1), ArcLen(MyGeo.Arcs(i)) / Count, Count + 1, ArcPt(MyGeo.Arcs(i), MyGeo.Arcs(i).Angle1 + 1), dPoints(), sAngles()
    Case "Ellipse": SpaceEllipse MyGeo.Ellipses(i), EllipsePt(MyGeo.Ellipses(i), MyGeo.Ellipses(i).Angle1), EllipseLen(MyGeo.Ellipses(i)) / Count, Count + 1, EllipsePt(MyGeo.Ellipses(i), MyGeo.Ellipses(i).Angle1 + 1), dPoints(), sAngles()
    Case "Spline": SpaceSpline MyGeo.Splines(i), MyGeo.Splines(i).Vertex(0), SplineLen(MyGeo.Splines(i)) / Count, Count + 1, SplineToPolyLine(MyGeo.Splines(i)).Vertex(1), dPoints(), sAngles()
    Case "PolyLine": SpacePolyLine MyGeo.PolyLines(i), MyGeo.PolyLines(i).Vertex(0), PolyLineLen(MyGeo.PolyLines(i)) / Count, Count + 1, MyGeo.PolyLines(i).Vertex(1), dPoints(), sAngles()
    Case "Insert"
        ReDim dPoints(0) As CadPoint
        dPoints(0) = MyGeo.Inserts(MySel.Index).Base
        dPoints(0).Layer.Color = vbBlue
        dPoints(0).Layer.Width = 3
End Select
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub



Function SplineToPolyLine(MySPline As CadSpline) As CadPolyLine
Dim tPts() As CadPoint
Dim i As Integer
SplinePoints MySPline, tPts()
SplineToPolyLine.Layer = MySPline.Layer
ReDim Preserve SplineToPolyLine.Vertex(UBound(tPts)) As CadPoint
For i = 0 To UBound(tPts)
    SplineToPolyLine.Vertex(i) = tPts(i)
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Function ZZZ_OLD_GetExtents(MyGeo() As Geometry, GNum As Integer) 'As CadLine
On Error Resume Next
Dim extSel() As SelSet
Dim lineSel() As SelSet
Dim ePts() As CadPoint
'Dim tPts() As CadPoint
Dim tSel As SelSet
Dim i As Integer
Dim OldCount As Integer
Dim tLine As CadLine
Dim eVar As Single ' THis is the "resolution" of the check for extents
eVar = 1
'--------Select all the geometry
SelectAllGeo extSel(), MyGeo(GNum)
If Not isSelected(extSel()) Then Exit Function
'-------Get a startPoint within the geometry
tSel = extSel(0)
GeoSelEndPoints MyGeo(GNum), tSel, ePts()
'-------Find Top-----------
tLine.P1.x = -10000000000#
tLine.P2.x = 10000000000#
tLine.P1.y = ePts(0).y
tLine.P2.y = ePts(0).y
SelectFromBox tLine, MyGeo(), GNum, lineSel()
Do While isSelected(lineSel)
    tLine.P1.y = tLine.P1.y - eVar
    tLine.P2.y = tLine.P1.y
    SelectFromBox tLine, MyGeo, GNum, lineSel()
Loop
'GetExtents.P1.y = tLine.P1.y
'-------Find Bottom-----------
tLine.P1.x = -10000000000#
tLine.P2.x = 10000000000#
tLine.P1.y = ePts(0).y
tLine.P2.y = ePts(0).y
SelectFromBox tLine, MyGeo, GNum, lineSel()
Do While isSelected(lineSel)
    tLine.P1.y = tLine.P1.y + eVar
    tLine.P2.y = tLine.P1.y
    SelectFromBox tLine, MyGeo, GNum, lineSel()
Loop
'GetExtents.P2.y = tLine.P2.y
'-------Find Right-----------
tLine.P1.y = -10000000000#
tLine.P2.y = 10000000000#
tLine.P1.x = ePts(0).x
tLine.P2.x = ePts(0).x
SelectFromBox tLine, MyGeo, GNum, lineSel()
Do While isSelected(lineSel)
    tLine.P1.x = tLine.P1.x + eVar
    tLine.P2.x = tLine.P1.x
    SelectFromBox tLine, MyGeo, GNum, lineSel()
Loop
'GetExtents.P2.X = tLine.P2.X
'-------Find Left-----------
tLine.P1.y = -10000000000#
tLine.P2.y = 10000000000#
tLine.P1.x = ePts(0).x
tLine.P2.x = ePts(0).x
SelectFromBox tLine, MyGeo, GNum, lineSel()
Do While isSelected(lineSel)
    tLine.P1.x = tLine.P1.x - eVar
    tLine.P2.x = tLine.P1.x
    SelectFromBox tLine, MyGeo, GNum, lineSel()
Loop
'GetExtents.P1.X = tLine.P1.X
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub Zoom(zLine As CadLine, Canvas As PictureBox)
On Error Resume Next
If zLine.P2.x < zLine.P1.x Then Swap zLine.P1.x, zLine.P2.x
If zLine.P2.y < zLine.P1.y Then Swap zLine.P1.y, zLine.P2.y
'zline.P2.y = zline.P1.y + Abs(zline.P2.X - zline.P1.X)
If zLine.P2.x = zLine.P1.x Then Exit Sub
If zLine.P2.y = zLine.P1.y Then Exit Sub
Canvas.ScaleLeft = zLine.P1.x
Canvas.ScaleTop = -zLine.P2.y
If Abs(zLine.P2.x - zLine.P1.x) > Abs(zLine.P2.y - zLine.P1.y) Then
    Canvas.ScaleWidth = Abs(zLine.P2.x - zLine.P1.x)
    Canvas.ScaleHeight = Canvas.ScaleWidth * (Canvas.Height / Canvas.Width)
Else
    Canvas.ScaleHeight = Abs(zLine.P2.y - zLine.P1.y)
    Canvas.ScaleWidth = Canvas.ScaleHeight * (Canvas.Width / Canvas.Height)
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub Center(MyGeo() As Geometry, GNum As Integer, zLine As CadLine, Canvas As PictureBox)
Dim tLine As CadLine
Dim CPt As CadPoint
tLine = GetExtents(MyGeo(), GNum)
CPt.x = tLine.P1.x + ((tLine.P2.x - tLine.P1.x) / 2)
CPt.y = tLine.P1.y + ((tLine.P2.y - tLine.P1.y) / 2)
zLine.P1.x = CPt.x - (Canvas.ScaleWidth / 2)
zLine.P2.x = CPt.x + (Canvas.ScaleWidth / 2)
zLine.P1.y = CPt.y - (Canvas.ScaleHeight / 2)
zLine.P2.y = CPt.y + (Canvas.ScaleHeight / 2)
Zoom zLine, Canvas
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function isSelected(MySel() As SelSet) As Boolean
On Error GoTo eTrap
Dim i As Integer
i = UBound(MySel)
isSelected = True
Exit Function
eTrap:
    isSelected = False
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function PolyLineLen(MyPolyLine As CadPolyLine)
Dim i As Integer
For i = 1 To UBound(MyPolyLine.Vertex)
    PolyLineLen = PolyLineLen + PtLen(MyPolyLine.Vertex(i - 1), MyPolyLine.Vertex(i))
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function SelectFromPoint(Canvas As PictureBox, MyGeo() As Geometry, GNum As Integer, SelPoint As CadPoint, ByRef MySel As SelSet, ByRef ActualPoint As CadPoint) As Boolean
On Error GoTo eTrap
Dim pLine(3) As CadLine
Dim x As Single
Dim y As Single
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim M As Integer
Dim iPts() As CadPoint
Dim tLine As CadLine
Dim tSel() As SelSet
Dim iGeo() As Geometry
Dim SelLine As CadLine
x = Canvas.ScaleX(10, vbPixels, vbUser) / 2
y = Canvas.ScaleY(10, vbPixels, vbUser) / 2
pLine(0).P1.x = SelPoint.x - x
pLine(0).P1.y = SelPoint.y - y
pLine(0).P2.x = SelPoint.x + x
pLine(0).P2.y = SelPoint.y - y
pLine(1).P1.x = SelPoint.x - x
pLine(1).P1.y = SelPoint.y + y
pLine(1).P2.x = SelPoint.x + x
pLine(1).P2.y = SelPoint.y + y
pLine(2).P1 = pLine(0).P1
pLine(2).P2 = pLine(1).P1
pLine(3).P1 = pLine(0).P2
pLine(3).P2 = pLine(1).P2
SelLine.P1 = pLine(0).P1
SelLine.P2 = pLine(1).P2
'For k = 0 To 3
'    DrawCadLine Canvas, pLine(k), , vbGreen, 1
'Next k
J = UBound(MyGeo(GNum).Points)
For i = 0 To J
    If PtInBox(SelLine, MyGeo(GNum).Points(i)) Then
        MySel.Type = "Point"
        MySel.Index = i
        ActualPoint = MyGeo(GNum).Points(i)
        SelectFromPoint = True: Exit Function
    End If
Next i
J = UBound(MyGeo(GNum).Lines)
For i = 0 To J
    For K = 0 To 3
        If LineLineIntersect(pLine(K), MyGeo(GNum).Lines(i), iPts()) >= 0 Then
            For M = 0 To UBound(iPts)
                If iPts(M).Layer.Color = vbBlue Then
                    MySel.Type = "Line"
                    MySel.Index = i
                    ActualPoint = iPts(M)
                    SelectFromPoint = True: Exit Function
                End If
            Next M
        End If
    Next K
Next i
J = UBound(MyGeo(GNum).Arcs)
For i = 0 To J
    For K = 0 To 3
        If LineArcIntersect(pLine(K), MyGeo(GNum).Arcs(i), iPts()) >= 0 Then
            For M = 0 To UBound(iPts)
                If iPts(M).Layer.Color = vbBlue Then
                    MySel.Type = "Arc"
                    MySel.Index = i
                    ActualPoint = iPts(M)
                    SelectFromPoint = True: Exit Function
                End If
            Next M
        End If
    Next K
Next i
J = UBound(MyGeo(GNum).Ellipses)
For i = 0 To J
    For K = 0 To 3
        If LineEllipseIntersect(pLine(K), MyGeo(GNum).Ellipses(i), iPts()) >= 0 Then
            For M = 0 To UBound(iPts)
                If iPts(M).Layer.Color = vbBlue Then
                    MySel.Type = "Ellipse"
                    MySel.Index = i
                    ActualPoint = iPts(M)
                    SelectFromPoint = True: Exit Function
                End If
            Next M
        End If
    Next K
Next i
J = UBound(MyGeo(GNum).Splines)
For i = 0 To J
    For K = 0 To 3
        If SplineLineIntersect(MyGeo(GNum).Splines(i), pLine(K), iPts()) >= 0 Then
            For M = 0 To UBound(iPts)
                If iPts(M).Layer.Color = vbBlue Then
                    MySel.Type = "Spline"
                    MySel.Index = i
                    ActualPoint = iPts(M)
                    SelectFromPoint = True: Exit Function
                End If
            Next M
        End If
    Next K
Next i
J = UBound(MyGeo(GNum).PolyLines)
For i = 0 To J
    For K = 0 To 3
        If PolyLineLineIntersect(MyGeo(GNum).PolyLines(i), pLine(K), iPts()) >= 0 Then
            For M = 0 To UBound(iPts)
                If iPts(M).Layer.Color = vbBlue Then
                    MySel.Type = "PolyLine"
                    MySel.Index = i
                    ActualPoint = iPts(M)
                    SelectFromPoint = True: Exit Function
                End If
            Next M
        End If
    Next K
Next i
J = UBound(MyGeo(GNum).Text)
For i = 0 To J
    tLine = cAngLine(MyGeo(GNum).Text(i).Angle, MyGeo(GNum).Text(i).Start, MyGeo(GNum).Text(i).Length)
    If LineInBox(SelLine, tLine) Then
        MySel.Type = "Text"
        MySel.Index = i
        ActualPoint = MyGeo(GNum).Text(i).Start
        SelectFromPoint = True: Exit Function
    End If
Next i
J = UBound(MyGeo(GNum).Faces)
tLine.P1 = pLine(0).P1
tLine.P2 = pLine(1).P2
For i = 0 To J
    If LineInBox(tLine, VLine(MyGeo(GNum).Faces(i).Vertex(0).x, MyGeo(GNum).Faces(i).Vertex(0).y, MyGeo(GNum).Faces(i).Vertex(1).x, MyGeo(GNum).Faces(i).Vertex(1).y)) Or _
        LineInBox(tLine, VLine(MyGeo(GNum).Faces(i).Vertex(1).x, MyGeo(GNum).Faces(i).Vertex(1).y, MyGeo(GNum).Faces(i).Vertex(2).x, MyGeo(GNum).Faces(i).Vertex(2).y)) Or _
        LineInBox(tLine, VLine(MyGeo(GNum).Faces(i).Vertex(2).x, MyGeo(GNum).Faces(i).Vertex(2).y, MyGeo(GNum).Faces(i).Vertex(0).x, MyGeo(GNum).Faces(i).Vertex(0).y)) Then
        MySel.Type = "Face"
        MySel.Index = i
        ActualPoint = MyGeo(GNum).Faces(i).Vertex(0)
        SelectFromPoint = True: Exit Function
    End If
Next i
J = UBound(MyGeo(GNum).Inserts)
tLine.P1 = pLine(0).P1
tLine.P2 = pLine(1).P2
For i = 0 To J
    CreateInsertGeo MyGeo(), MyGeo(GNum).Inserts(i), iGeo()
    SelectFromBox tLine, iGeo(), 0, tSel()
    If isSelected(tSel()) Then
        MySel.Type = "Insert"
        MySel.Index = i
        SelectFromPoint = True: Exit Function
    End If
Next i
SelectFromPoint = False
Exit Function
eTrap:
    J = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub RadiusGroup(MyGeo As Geometry, MySel() As SelSet, rad As Single)
On Error GoTo eTrap
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim lTop As Integer
Dim lCount As Integer
Dim lineSel() As SelSet
Dim iPts() As CadPoint
Dim E1 As Integer
Dim E2 As Integer
Dim E1Pts() As CadPoint
Dim E2Pts() As CadPoint
'-------First convert everything to  lines----------
lTop = UBound(MyGeo.Lines) + 1
For i = 0 To UBound(MySel)
    If MySel(i).Type = "Line" Then
        ReDim Preserve lineSel(lCount) As SelSet
        lineSel(lCount) = MySel(i)
        lCount = lCount + 1
    Else
        ExplodeSelection MyGeo, MySel(i)
        For K = lTop To UBound(MyGeo.Lines)
            ReDim Preserve lineSel(lCount) As SelSet
            lineSel(lCount).Type = "Line"
            lineSel(lCount).Index = K
            lCount = lCount + 1
        Next K
        lTop = UBound(MyGeo.Lines) + 1
    End If
Next i
'-------------Next - we will only radius REAL intersections of the SelSet
For i = 0 To UBound(lineSel)
    For J = 0 To UBound(lineSel)
        If LineLineIntersect(MyGeo.Lines(lineSel(i).Index), MyGeo.Lines(lineSel(J).Index), iPts()) >= 0 Then
            If iPts(0).Layer.Color = vbBlue Then
                GeoSelEndPoints MyGeo, lineSel(i), E1Pts()
                GeoSelEndPoints MyGeo, lineSel(J), E2Pts()
                E1 = ClosestPoint(iPts(0), E1Pts())
                E2 = ClosestPoint(iPts(0), E2Pts())
                RadiusSelection MyGeo, lineSel(i), E1, lineSel(J), E2, iPts(0), rad
            End If
        End If
    Next J
Next i
Exit Sub
eTrap:
  lTop = 0
  Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub RelimitGroup(MyGeo As Geometry, MySel() As SelSet, tSel As SelSet, MPt As CadPoint)
On Error GoTo eTrap
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim M As Integer
Dim P As Integer
Dim Min As Single
Dim iPts() As CadPoint
Dim ePts() As CadPoint
Dim lSel As SelSet
Dim lPts() As CadPoint
Dim RNum() As Integer
Dim r As Integer
J = UBound(MyGeo.Lines) + 1
ReDim Preserve MyGeo.Lines(J) As CadLine
MyGeo.Lines(J).P1 = MPt
lSel.Type = "Line"
lSel.Index = J
For i = 0 To UBound(MySel)
    Erase iPts()
    If GeoIntersect(MyGeo, MySel(i), tSel, iPts()) >= 0 Then
        'The group-element intersects with the target
        Erase ePts()
        GeoSelEndPoints MyGeo, MySel(i), ePts()
        r = 0
        Erase RNum()
        For K = 0 To UBound(ePts)
            'we check to see if a line from the endpoint to the
            'trim point crosses the target element
            MyGeo.Lines(J).P2 = ePts(K)
            Erase lPts()
            If GeoIntersect(MyGeo, lSel, tSel, lPts()) >= 0 Then
                For M = 0 To UBound(lPts)
                    If lPts(M).Layer.Color = vbBlue Then
                        ReDim Preserve RNum(r) As Integer
                        RNum(r) = K
                        r = r + 1
                        Exit For
                    End If
                Next M
            End If
        Next K
        P = 0
        If r = 1 Then ' we need to TRIM the closest endpoint (to the intersection)
            For M = 0 To UBound(iPts)
                If iPts(M).Layer.Color = vbBlue Then P = M
            Next M
            'We need to fix the indexes of the splines involved in the Relimit
            'if the current element is a spline - since it will be removed and converted
            'to a polyline
            'We also need to adjust polylines involved because if we Relimit the upper half
            'of a polyline, it will be removed and upset the index sets
            If MySel(i).Type = "Spline" Then '
                If tSel.Type = "Spline" And MySel(i).Index < tSel.Index Then tSel.Index = tSel.Index - 1
                For M = i + 1 To UBound(MySel)
                    If MySel(M).Type = "Spline" And MySel(i).Index < MySel(M).Index Then MySel(M).Index = MySel(M).Index - 1
                Next M
            End If
            If RNum(0) = 0 And MySel(i).Type = "PolyLine" Then
                If tSel.Type = "PolyLine" And MySel(i).Index < tSel.Index Then tSel.Index = tSel.Index - 1
                For M = i + 1 To UBound(MySel)
                    If RNum(0) = 0 And MySel(M).Type = "PolyLine" And MySel(i).Index < MySel(M).Index Then MySel(M).Index = MySel(M).Index - 1
                Next M
            End If
            RelimitSelection MyGeo, MySel(i), RNum(0), iPts(P)
        ElseIf r = 2 Then ' we need to EXTEND the closest endpoint to the intersection
            If PtLen(MPt, ePts(RNum(0))) < PtLen(MPt, ePts(RNum(1))) Then K = RNum(0) Else K = RNum(1)
            For M = 0 To UBound(iPts)
                Select Case tSel.Type
                    Case "Line", "Spline", "PolyLine"
                        If iPts(M).Layer.Color = vbRed Then P = M
                    Case "Arc", "Ellipse"
                        If iPts(M).Layer.Color = vbYellow Then P = M
                        If (iPts(M).Layer.Color = vbYellow Or iPts(M).Layer.Color = vbGreen) And (MySel(i).Type = "Arc" Or MySel(i).Type = "Ellipse") Then P = M
                End Select
            Next M
            RelimitSelection MyGeo, MySel(i), K, iPts(P)
        End If
    End If
Next i
RemoveGeo MyGeo, "Line", J
For i = 0 To UBound(MySel)
    If MySel(i).Type = "Spline" And MySel(i).Index <> tSel.Index Then
        MySel(i).Type = "PolyLine"
        MySel(i).Index = UBound(MyGeo.PolyLines)
    End If
Next i
Exit Sub
eTrap:
    J = 0
    'MsgBox Err.Description
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub ShowSelection(Canvas As PictureBox, MyGeo() As Geometry, GNum As Integer, MySel As SelSet, MyColor As Long)
On Error Resume Next
Dim J As Integer
Dim ePoints() As CadPoint
Select Case MySel.Type
    Case "Point":    DrawCadPoint Canvas, MyGeo(GNum).Points(MySel.Index), 13, MyColor
    Case "Line":     DrawCadLine Canvas, MyGeo(GNum).Lines(MySel.Index), 13, MyColor
    Case "Arc":      DrawCadArc Canvas, MyGeo(GNum).Arcs(MySel.Index), 13, MyColor
    Case "Ellipse":  DrawCadEllipse Canvas, MyGeo(GNum).Ellipses(MySel.Index), 13, MyColor
    Case "Spline":   DrawCadSpline Canvas, MyGeo(GNum).Splines(MySel.Index), 13, MyColor
    Case "PolyLine": DrawCadPolyLine Canvas, MyGeo(GNum).PolyLines(MySel.Index), 13, MyColor
    Case "Text":     DrawCadText Canvas, MyGeo(GNum).Text(MySel.Index), 13, MyColor
    Case "Face":     DrawCadFace Canvas, MyGeo(GNum).Faces(MySel.Index), 13, MyColor
    Case "Insert":   DrawCadInsert Canvas, MyGeo(GNum).Inserts(MySel.Index), MyGeo(), 13, vbPink
End Select
GeoSelEditPoints MyGeo(GNum), MySel, ePoints()
For J = 0 To UBound(ePoints)
    DrawCadPoint Canvas, ePoints(J)
Next J
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub


Sub SelectionIntersect(MyGeo As Geometry, MySel() As SelSet, ByRef tPoints() As CadPoint, Optional Real As Boolean)
On Error GoTo eTrap
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim M As Integer
Dim iPoints() As CadPoint
For i = 0 To UBound(MySel)
    For J = i + 1 To UBound(MySel)
        For K = 0 To GeoIntersect(MyGeo, MySel(i), MySel(J), iPoints())
            If Real Then If iPoints(K).Layer.Color <> vbBlue Then GoTo SkipIt
            ReDim Preserve tPoints(M) As CadPoint
            tPoints(M) = iPoints(K)
            M = M + 1
SkipIt:
        Next K
    Next J
Next i
eTrap:
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub MoveSelection(MyGeo As Geometry, MySel() As SelSet, dX As Single, dY As Single, Copy As Boolean)
On Error GoTo eTrap
Dim i As Integer
Dim J As Integer
Dim K As Integer
For i = 0 To UBound(MySel)
    J = MySel(i).Index
    K = J
    Select Case MySel(i).Type
        Case "Point"
            If Copy Then
                K = UBound(MyGeo.Points) + 1
                ReDim Preserve MyGeo.Points(K) As CadPoint
            End If
            MyGeo.Points(K) = MovePoint(MyGeo.Points(J), dX, dY)
        Case "Line"
            If Copy Then
                K = UBound(MyGeo.Lines) + 1
                ReDim Preserve MyGeo.Lines(K) As CadLine
            End If
            MyGeo.Lines(K) = MoveLine(MyGeo.Lines(J), dX, dY)
        Case "Arc"
            If Copy Then
                K = UBound(MyGeo.Arcs) + 1
                ReDim Preserve MyGeo.Arcs(K) As CadArc
            End If
            MyGeo.Arcs(K) = MoveArc(MyGeo.Arcs(J), dX, dY)
        Case "Ellipse"
            If Copy Then
                K = UBound(MyGeo.Ellipses) + 1
                ReDim Preserve MyGeo.Ellipses(K) As CadEllipse
            End If
            MyGeo.Ellipses(K) = MoveEllipse(MyGeo.Ellipses(J), dX, dY)
        Case "Spline"
            If Copy Then
                K = UBound(MyGeo.Splines) + 1
                ReDim Preserve MyGeo.Splines(K) As CadSpline
            End If
            MyGeo.Splines(K) = MoveSpline(MyGeo.Splines(J), dX, dY)
        Case "PolyLine"
            If Copy Then
                K = UBound(MyGeo.PolyLines) + 1
                ReDim Preserve MyGeo.PolyLines(K) As CadPolyLine
            End If
            MyGeo.PolyLines(K) = MovePolyLine(MyGeo.PolyLines(J), dX, dY)
        Case "Text"
            If Copy Then
                K = UBound(MyGeo.Text) + 1
                ReDim Preserve MyGeo.Text(K) As CadText
            End If
            MyGeo.Text(K) = MoveText(MyGeo.Text(J), dX, dY)
        Case "Insert"
            If Copy Then
                K = UBound(MyGeo.Inserts) + 1
                ReDim Preserve MyGeo.Inserts(K) As CadInsert
            End If
            MyGeo.Inserts(K) = MoveInsert(MyGeo.Inserts(J), dX, dY)
        Case "Face"
            If Copy Then
                K = UBound(MyGeo.Faces) + 1
                ReDim Preserve MyGeo.Faces(K) As CadFace
            End If
            MyGeo.Faces(K) = MoveFace(MyGeo.Faces(J), dX, dY)
    End Select
    MySel(i).Index = K
Next i
eTrap:
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub RelimitSelection(ByRef MyGeo As Geometry, ByRef rSel As SelSet, EndNum As Integer, MPt As CadPoint)
Dim i As Integer
Select Case rSel.Type
    Case "Line": RelimitLine MyGeo, rSel.Index, EndNum, MPt
    Case "Arc": RelimitArc MyGeo, rSel.Index, EndNum, MPt
    Case "Ellipse": RelimitEllipse MyGeo, rSel.Index, EndNum, MPt
    Case "Spline":
        RelimitSpline MyGeo, rSel.Index, EndNum, MPt
        rSel.Type = "PolyLine"
        rSel.Index = UBound(MyGeo.PolyLines)
    Case "PolyLine": RelimitPolyLine MyGeo, rSel.Index, EndNum, MPt
End Select
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub RotateSelection(MyGeo As Geometry, MySel() As SelSet, RotAngle As Single, Pivot As CadPoint, Copy As Boolean)
On Error GoTo eTrap
Dim i As Integer
Dim J As Integer
Dim K As Integer
For i = 0 To UBound(MySel)
    J = MySel(i).Index
    K = J
    Select Case MySel(i).Type
        Case "Point"
            If Copy Then
                K = UBound(MyGeo.Points) + 1
                ReDim Preserve MyGeo.Points(K) As CadPoint
            End If
            MyGeo.Points(K) = RotatePoint(MyGeo.Points(J), Pivot, RotAngle)
        Case "Line"
            If Copy Then
                K = UBound(MyGeo.Lines) + 1
                ReDim Preserve MyGeo.Lines(K) As CadLine
            End If
            MyGeo.Lines(K) = RotateLine(MyGeo.Lines(J), Pivot, RotAngle)
        Case "Arc"
            If Copy Then
                K = UBound(MyGeo.Arcs) + 1
                ReDim Preserve MyGeo.Arcs(K) As CadArc
            End If
            MyGeo.Arcs(K) = RotateArc(MyGeo.Arcs(J), Pivot, RotAngle)
        Case "Ellipse"
            If Copy Then
                K = UBound(MyGeo.Ellipses) + 1
                ReDim Preserve MyGeo.Ellipses(K) As CadEllipse
            End If
            MyGeo.Ellipses(K) = RotateEllipse(MyGeo.Ellipses(J), Pivot, RotAngle)
        Case "Spline"
            If Copy Then
                K = UBound(MyGeo.Splines) + 1
                ReDim Preserve MyGeo.Splines(K) As CadSpline
            End If
            MyGeo.Splines(K) = RotateSpline(MyGeo.Splines(J), Pivot, RotAngle)
        Case "PolyLine"
            If Copy Then
                K = UBound(MyGeo.PolyLines) + 1
                ReDim Preserve MyGeo.PolyLines(K) As CadPolyLine
            End If
            MyGeo.PolyLines(K) = RotatePolyLine(MyGeo.PolyLines(J), Pivot, RotAngle)
        Case "Text"
            If Copy Then
                K = UBound(MyGeo.Text) + 1
                ReDim Preserve MyGeo.Text(K) As CadText
            End If
            MyGeo.Text(K) = RotateText(MyGeo.Text(J), Pivot, RotAngle)
        Case "Insert"
            If Copy Then
                K = UBound(MyGeo.Inserts) + 1
                ReDim Preserve MyGeo.Inserts(K) As CadInsert
            End If
            MyGeo.Inserts(K) = RotateInsert(MyGeo.Inserts(J), Pivot, RotAngle)
        Case "Face"
            If Copy Then
                K = UBound(MyGeo.Faces) + 1
                ReDim Preserve MyGeo.Faces(K) As CadFace
            End If
            MyGeo.Faces(K) = RotateFace(MyGeo.Faces(J), Pivot, RotAngle)
    End Select
    MySel(i).Index = K
Next i
eTrap:
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub ScaleSelection(MyGeo As Geometry, MySel() As SelSet, dX As Single, dY As Single, Pivot As CadPoint, Copy As Boolean)
On Local Error GoTo eTrap
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim Count As Integer
Dim NewPoint As CadPoint
NewPoint = ScalePoint(Pivot, dX, dY)
For i = 0 To UBound(MySel)
    J = MySel(i).Index
    K = J
    Select Case MySel(i).Type
        Case "Point"
            If Copy Then
                K = UBound(MyGeo.Points) + 1
                ReDim Preserve MyGeo.Points(K) As CadPoint
            End If
            MyGeo.Points(K) = ScalePoint(MyGeo.Points(J), dX, dY)
        Case "Line"
            If Copy Then
                K = UBound(MyGeo.Lines) + 1
                ReDim Preserve MyGeo.Lines(K) As CadLine
            End If
            MyGeo.Lines(K) = ScaleLine(MyGeo.Lines(J), dX, dY)
        Case "Arc"
            'An Arc becomes a polyline when it is scaled because there is not always a
            'perfect mathematical expression of an Arc if it is scaled
            Count = UBound(MyGeo.PolyLines) + 1
            ReDim Preserve MyGeo.PolyLines(Count) As CadPolyLine
            MyGeo.PolyLines(Count) = ArcToPolyLine(MyGeo.Arcs(J))
            RemoveGeo MyGeo, "Arc", J
            J = Count
            K = J
            If Copy Then
                K = UBound(MyGeo.PolyLines) + 1
                ReDim Preserve MyGeo.PolyLines(K) As CadPolyLine
            End If
            MyGeo.PolyLines(K) = ScalePolyLine(MyGeo.PolyLines(J), dX, dY)
            MySel(i).Type = "PolyLine"
        Case "Ellipse"
            'An ellipse becomes a Polyline when it is scaled because there is not always a
            'perfect mathematical expression of an ellipse if it is scaled
            Count = UBound(MyGeo.PolyLines) + 1
            ReDim Preserve MyGeo.PolyLines(Count) As CadPolyLine
            MyGeo.PolyLines(Count) = EllipseToPolyLine(MyGeo.Ellipses(J))
            RemoveGeo MyGeo, "Ellipse", J
            J = Count
            K = J
            If Copy Then
                K = UBound(MyGeo.PolyLines) + 1
                ReDim Preserve MyGeo.PolyLines(K) As CadPolyLine
            End If
            MyGeo.PolyLines(K) = ScalePolyLine(MyGeo.PolyLines(J), dX, dY)
            MySel(i).Type = "PolyLine"
        Case "Spline"
            If Copy Then
                K = UBound(MyGeo.Splines) + 1
                ReDim Preserve MyGeo.Splines(K) As CadSpline
            End If
            MyGeo.Splines(K) = ScaleSpline(MyGeo.Splines(J), dX, dY)
        Case "PolyLine"
            If Copy Then
                K = UBound(MyGeo.PolyLines) + 1
                ReDim Preserve MyGeo.PolyLines(K) As CadPolyLine
            End If
            MyGeo.PolyLines(K) = ScalePolyLine(MyGeo.PolyLines(J), dX, dY)
        Case "Text"
            If Copy Then
                K = UBound(MyGeo.Text) + 1
                ReDim Preserve MyGeo.Text(K) As CadText
            End If
            MyGeo.Text(K) = ScaleText(MyGeo.Text(J), dX, dY)
        Case "Insert"
            If Copy Then
                K = UBound(MyGeo.Inserts) + 1
                ReDim Preserve MyGeo.Inserts(K) As CadInsert
            End If
            MyGeo.Inserts(K) = ScaleInsert(MyGeo.Inserts(J), dX, dY)
        Case "Face"
            If Copy Then
                K = UBound(MyGeo.Faces) + 1
                ReDim Preserve MyGeo.Faces(K) As CadFace
            End If
            MyGeo.Faces(K) = ScaleFace(MyGeo.Faces(J), dX, dY)
    End Select
    MySel(i).Index = K
Next i
MoveSelection MyGeo, MySel(), Pivot.x - NewPoint.x, Pivot.y - NewPoint.y, False
eTrap:
    Count = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub BreakSelection(MyGeo As Geometry, ByRef MySel() As SelSet, ByRef bSel As SelSet, bPt As CadPoint)
Select Case bSel.Type
    Case "Line"
        BreakLine MyGeo, bSel.Index, bPt
        bSel.Index = UBound(MyGeo.Lines)
    Case "Arc"
        BreakArc MyGeo, bSel.Index, bPt
        bSel.Index = UBound(MyGeo.Arcs)
    Case "Ellipse"
        BreakEllipse MyGeo, bSel.Index, bPt
        bSel.Index = UBound(MyGeo.Ellipses)
    Case "Spline"
        BreakSpline MyGeo, bSel.Index, bPt
        bSel.Type = "PolyLine"
        bSel.Index = UBound(MyGeo.PolyLines) - 1
        ReDim MySel(0) As SelSet
        MySel(0) = bSel
        bSel.Index = UBound(MyGeo.PolyLines)
    Case "PolyLine"
        BreakPolyLine MyGeo, bSel.Index, bPt
        bSel.Index = UBound(MyGeo.PolyLines)
End Select

'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub GeoSelEndPoints(MyGeo As Geometry, MySel As SelSet, ByRef ePoints() As CadPoint)
Erase ePoints()
Dim i As Integer
Dim J As Integer
i = MySel.Index
Select Case MySel.Type
    Case "Point"
        ReDim ePoints(0) As CadPoint
        ePoints(0) = MyGeo.Points(i)
    Case "Line"
        ReDim ePoints(1) As CadPoint
        ePoints(0) = MyGeo.Lines(i).P1
        ePoints(1) = MyGeo.Lines(i).P2
    Case "Arc"
        ReDim ePoints(1) As CadPoint
        ePoints(0) = ArcPt(MyGeo.Arcs(i), MyGeo.Arcs(i).Angle1)
        ePoints(1) = ArcPt(MyGeo.Arcs(i), MyGeo.Arcs(i).Angle2)
    Case "Ellipse"
        ReDim ePoints(1) As CadPoint
        ePoints(0) = EllipsePt(MyGeo.Ellipses(i), MyGeo.Ellipses(i).Angle1)
        ePoints(1) = EllipsePt(MyGeo.Ellipses(i), MyGeo.Ellipses(i).Angle2)
    Case "Spline"
        ReDim ePoints(1) As CadPoint
        ePoints(0) = MyGeo.Splines(i).Vertex(0)
        ePoints(1) = MyGeo.Splines(i).Vertex(UBound(MyGeo.Splines(i).Vertex))
    Case "PolyLine"
        ReDim ePoints(1) As CadPoint
        ePoints(0) = MyGeo.PolyLines(i).Vertex(0)
        ePoints(1) = MyGeo.PolyLines(i).Vertex(UBound(MyGeo.PolyLines(i).Vertex))
    Case "Text"
        ReDim ePoints(0) As CadPoint
        ePoints(0) = MyGeo.Text(i).Start
    Case "Insert"
        ReDim ePoints(0) As CadPoint
        ePoints(0) = MyGeo.Inserts(i).Base
End Select
For J = 0 To UBound(ePoints)
    ePoints(J).Layer.Color = vbCyan
    ePoints(J).Layer.style = 0
    ePoints(J).Layer.Width = 3
Next J
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub GeoSelEditPoints(MyGeo As Geometry, MySel As SelSet, ByRef ePoints() As CadPoint)
Erase ePoints()
Dim i As Integer
Dim J As Integer
i = MySel.Index
Select Case MySel.Type
    Case "Point"
        ReDim ePoints(0) As CadPoint
        ePoints(0) = MyGeo.Points(i)
    Case "Line"
        ReDim ePoints(1) As CadPoint
        ePoints(0) = MyGeo.Lines(i).P1
        ePoints(1) = MyGeo.Lines(i).P2
    Case "Arc"
        ReDim ePoints(2) As CadPoint
        ePoints(0) = MyGeo.Arcs(i).Center
        ePoints(1) = ArcPt(MyGeo.Arcs(i), MyGeo.Arcs(i).Angle1)
        ePoints(2) = ArcPt(MyGeo.Arcs(i), MyGeo.Arcs(i).Angle2)
    Case "Ellipse"
        ReDim ePoints(4) As CadPoint
        ePoints(0) = MyGeo.Ellipses(i).F1
        ePoints(1) = MyGeo.Ellipses(i).F2
        ePoints(2) = MyGeo.Ellipses(i).P1
        ePoints(3) = EllipsePt(MyGeo.Ellipses(i), MyGeo.Ellipses(i).Angle1)
        ePoints(4) = EllipsePt(MyGeo.Ellipses(i), MyGeo.Ellipses(i).Angle2)
    Case "Spline"
        ReDim ePoints(UBound(MyGeo.Splines(i).Vertex)) As CadPoint
        For J = 0 To UBound(MyGeo.Splines(i).Vertex)
            ePoints(J) = MyGeo.Splines(i).Vertex(J)
        Next J
    Case "PolyLine"
        ReDim ePoints(UBound(MyGeo.PolyLines(i).Vertex)) As CadPoint
        For J = 0 To UBound(MyGeo.PolyLines(i).Vertex)
            ePoints(J) = MyGeo.PolyLines(i).Vertex(J)
        Next J
    Case "Text"
        ReDim ePoints(0) As CadPoint
        ePoints(0) = MyGeo.Text(i).Start
    Case "Insert"
        ReDim ePoints(0) As CadPoint
        ePoints(0) = MyGeo.Inserts(i).Base
End Select
For J = 0 To UBound(ePoints)
    ePoints(J).Layer.Color = vbCyan
    ePoints(J).Layer.style = 0
    ePoints(J).Layer.Width = 3
Next J
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function ClosestPoint(SelPoint As CadPoint, ePoints() As CadPoint) As Integer
Dim Dist As Single
Dim i As Integer
Dim Min As Single
Min = PtLen(SelPoint, ePoints(0))
For i = 0 To UBound(ePoints)
     If PtLen(SelPoint, ePoints(i)) < Min Then
        Min = PtLen(SelPoint, ePoints(i))
        ClosestPoint = i
    End If
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub DrawCadGeo(Canvas As PictureBox, MyGeo() As Geometry, GNum As Integer, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
On Error Resume Next
Dim i As Integer
For i = 0 To UBound(MyGeo(GNum).Points)
    DrawCadPoint Canvas, MyGeo(GNum).Points(i), Mode, Color, Width
Next i
For i = 0 To UBound(MyGeo(GNum).Lines)
    DrawCadLine Canvas, MyGeo(GNum).Lines(i), Mode, Color, Width
Next i
For i = 0 To UBound(MyGeo(GNum).Arcs)
    DrawCadArc Canvas, MyGeo(GNum).Arcs(i), Mode, Color, Width
Next i
For i = 0 To UBound(MyGeo(GNum).Ellipses)
    DrawCadEllipse Canvas, MyGeo(GNum).Ellipses(i), Mode, Color, Width
Next i
For i = 0 To UBound(MyGeo(GNum).Splines)
    DrawCadSpline Canvas, MyGeo(GNum).Splines(i), Mode, Color, Width
Next i
For i = 0 To UBound(MyGeo(GNum).PolyLines)
    DrawCadPolyLine Canvas, MyGeo(GNum).PolyLines(i), Mode, Color, Width
Next i
For i = 0 To UBound(MyGeo(GNum).Text)
    DrawCadText Canvas, MyGeo(GNum).Text(i), Mode, Color
Next i
For i = 0 To UBound(MyGeo(GNum).Faces)
    DrawCadFace Canvas, MyGeo(GNum).Faces(i), Mode, Color, Width
Next i
For i = 0 To UBound(MyGeo(GNum).Inserts)
    DrawCadInsert Canvas, MyGeo(GNum).Inserts(i), MyGeo(), Mode, Color, Width
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function EllipseEllipseIntersect(MyEllipse As CadEllipse, EllipseB As CadEllipse, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim a As Single, B As Single, u As Single, v As Single  ' ellipse parameters
Dim X1 As Single, Y1 As Single 'plotting coordinates
Dim cx As Single, cy As Single ' center of the ellipse
Dim i As Integer, J As Single
Dim First As Boolean
Dim tLine As CadLine
Dim TotLen As Single, Flen As Single
'-----------Intersect Stuff---------
Dim iRes As Integer
Dim iPts() As CadPoint
Dim iLine As CadLine
Dim iCount As Integer
Dim iEllipse As CadEllipse
'------------------Ellipse Parameters-------------------------
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2) * Pi / 180
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
'------------------Intersect Parameters----------------
iLine.P1.x = (Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))) + cx
iLine.P1.y = (Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))) + cy
iEllipse = EllipseB
iEllipse.Angle1 = 0
iEllipse.Angle2 = 360
'-----------------Cycle through the ellipse------------
For J = 0 To 2 * Pi Step Pi / 50
    iLine.P2.x = (Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))) + cx
    iLine.P2.y = (Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))) + cy
    iRes = LineEllipseIntersect(iLine, iEllipse, iPts()) ' check the intersect
    If iRes > 0 Then
        For i = 0 To UBound(iPts)
            If iPts(i).Layer.Color = vbBlue Then
                ReDim Preserve iPoints(iCount) As CadPoint
                iPoints(iCount) = iPts(i)
                iCount = iCount + 1
            End If
        Next i
    End If
    iLine.P1 = iLine.P2
Next J
EllipseEllipseIntersect = iCount - 1
If iCount = 0 Then Exit Function
'-------Check for virtual intersections
For i = 0 To UBound(iPoints)
    If Not InsideAngles(EllipsePtAngle(MyEllipse, iPoints(i)), MyEllipse.Angle1, MyEllipse.Angle2) Then iPoints(i).Layer.Color = vbGreen
Next i
For i = 0 To UBound(iPoints)
    If Not InsideAngles(EllipsePtAngle(EllipseB, iPoints(i)), EllipseB.Angle1, EllipseB.Angle2) Then
        If iPoints(i).Layer.Color <> vbBlue Then iPoints(i).Layer.Color = vbRed Else iPoints(i).Layer.Color = vbGreen
    End If
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function EllipseToPolyLine(MyEllipse As CadEllipse) As CadPolyLine
EllipseToPolyLine.Layer = MyEllipse.Layer
Dim a As Single, B As Single ' ellipse parameters
Dim u As Single, v As Single
Dim i As Single, J As Single
Dim ang1 As Single, ang2 As Single 'adjusted ellipse angles
Dim X1 As Single, Y1 As Single 'plotting coordinates
Dim cx As Single, cy As Single ' center of the ellipse
Dim First As Boolean
Dim TotLen As Single, Flen As Single
Dim vCount As Integer
'-----------------Angle adjustment-------
ang1 = MyEllipse.Angle1
ang2 = MyEllipse.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
i = (ang2 - ang1) / MyEllipse.NumPoints
'-------------------------------------------
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2) * Pi / 180
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
For u = ang1 To ang2 Step i
    J = u * Pi / 180
    X1 = Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    Y1 = Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    ReDim Preserve EllipseToPolyLine.Vertex(vCount) As CadPoint
    EllipseToPolyLine.Vertex(vCount).x = cx + X1
    EllipseToPolyLine.Vertex(vCount).y = cy + Y1
    vCount = vCount + 1
Next u
If u - i < ang2 Then
    J = ang2 * Pi / 180
    X1 = Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    Y1 = Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    ReDim Preserve EllipseToPolyLine.Vertex(vCount) As CadPoint
    EllipseToPolyLine.Vertex(vCount).x = cx + X1
    EllipseToPolyLine.Vertex(vCount).y = cy + Y1
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function GeoIntersect(MyGeo As Geometry, SelA As SelSet, SelB As SelSet, ByRef iPoints() As CadPoint) As Integer
Select Case SelA.Type
    Case "Line"
        Select Case SelB.Type
            Case "Line": GeoIntersect = LineLineIntersect(MyGeo.Lines(SelA.Index), MyGeo.Lines(SelB.Index), iPoints())
            Case "Arc": GeoIntersect = LineArcIntersect(MyGeo.Lines(SelA.Index), MyGeo.Arcs(SelB.Index), iPoints())
            Case "Ellipse": GeoIntersect = LineEllipseIntersect(MyGeo.Lines(SelA.Index), MyGeo.Ellipses(SelB.Index), iPoints())
            Case "Spline": GeoIntersect = SplineLineIntersect(MyGeo.Splines(SelB.Index), MyGeo.Lines(SelA.Index), iPoints())
            Case "PolyLine": GeoIntersect = PolyLineLineIntersect(MyGeo.PolyLines(SelB.Index), MyGeo.Lines(SelA.Index), iPoints())
        End Select
    Case "Arc"
        Select Case SelB.Type
            Case "Line": GeoIntersect = LineArcIntersect(MyGeo.Lines(SelB.Index), MyGeo.Arcs(SelA.Index), iPoints())
            Case "Arc": GeoIntersect = ArcArcIntersect(MyGeo.Arcs(SelA.Index), MyGeo.Arcs(SelB.Index), iPoints())
            Case "Ellipse": GeoIntersect = EllipseArcIntersect(MyGeo.Ellipses(SelB.Index), MyGeo.Arcs(SelA.Index), iPoints())
            Case "Spline": GeoIntersect = SplineArcIntersect(MyGeo.Splines(SelB.Index), MyGeo.Arcs(SelA.Index), iPoints())
            Case "PolyLine": GeoIntersect = PolyLineArcIntersect(MyGeo.PolyLines(SelB.Index), MyGeo.Arcs(SelA.Index), iPoints())
        End Select
    Case "Ellipse"
        Select Case SelB.Type
            Case "Line": GeoIntersect = LineEllipseIntersect(MyGeo.Lines(SelB.Index), MyGeo.Ellipses(SelA.Index), iPoints())
            Case "Arc": GeoIntersect = EllipseArcIntersect(MyGeo.Ellipses(SelA.Index), MyGeo.Arcs(SelB.Index), iPoints())
            Case "Ellipse": GeoIntersect = EllipseEllipseIntersect(MyGeo.Ellipses(SelB.Index), MyGeo.Ellipses(SelA.Index), iPoints())
            Case "Spline": GeoIntersect = SplineEllipseIntersect(MyGeo.Splines(SelB.Index), MyGeo.Ellipses(SelA.Index), iPoints())
            Case "PolyLine": GeoIntersect = PolyLineEllipseIntersect(MyGeo.PolyLines(SelB.Index), MyGeo.Ellipses(SelA.Index), iPoints())
        End Select
    Case "Spline"
        Select Case SelB.Type
            Case "Line": GeoIntersect = SplineLineIntersect(MyGeo.Splines(SelA.Index), MyGeo.Lines(SelB.Index), iPoints())
            Case "Arc": GeoIntersect = SplineArcIntersect(MyGeo.Splines(SelA.Index), MyGeo.Arcs(SelB.Index), iPoints())
            Case "Ellipse": GeoIntersect = SplineEllipseIntersect(MyGeo.Splines(SelA.Index), MyGeo.Ellipses(SelB.Index), iPoints())
            Case "Spline": GeoIntersect = SplineSplineIntersect(MyGeo.Splines(SelA.Index), MyGeo.Splines(SelB.Index), iPoints())
            Case "PolyLine": GeoIntersect = PolyLineSplineIntersect(MyGeo.PolyLines(SelB.Index), MyGeo.Splines(SelA.Index), iPoints())
        End Select
    Case "PolyLine"
        Select Case SelB.Type
            Case "Line": GeoIntersect = PolyLineLineIntersect(MyGeo.PolyLines(SelA.Index), MyGeo.Lines(SelB.Index), iPoints())
            Case "Arc": GeoIntersect = PolyLineArcIntersect(MyGeo.PolyLines(SelA.Index), MyGeo.Arcs(SelB.Index), iPoints())
            Case "Ellipse": GeoIntersect = PolyLineEllipseIntersect(MyGeo.PolyLines(SelA.Index), MyGeo.Ellipses(SelB.Index), iPoints())
            Case "Spline": GeoIntersect = PolyLineSplineIntersect(MyGeo.PolyLines(SelA.Index), MyGeo.Splines(SelB.Index), iPoints())
            Case "PolyLine": GeoIntersect = PolyLinePolyLineIntersect(MyGeo.PolyLines(SelA.Index), MyGeo.PolyLines(SelB.Index), iPoints())
        End Select
    Case Else
        GeoIntersect = -1
End Select
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Function PolyLineArcIntersect(MyPolyLine As CadPolyLine, MyArc As CadArc, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim iRes As Integer
Dim iPts() As CadPoint
Dim iCount As Integer
Dim tLine As CadLine
Dim i As Integer
Dim J As Integer
For i = 1 To UBound(MyPolyLine.Vertex)
    tLine.P1 = MyPolyLine.Vertex(i - 1)
    tLine.P2 = MyPolyLine.Vertex(i)
    iRes = LineArcIntersect(tLine, MyArc, iPts())
    For J = 0 To iRes
        If iPts(J).Layer.Color = vbBlue Or iPts(J).Layer.Color = vbRed Or i = 1 Or i = UBound(MyPolyLine.Vertex) Then
            If i = 1 Then If PtLen(MyPolyLine.Vertex(0), iPts(J)) > PtLen(MyPolyLine.Vertex(1), iPts(J)) Then GoTo SkipIt
            If i = UBound(MyPolyLine.Vertex) Then If PtLen(MyPolyLine.Vertex(i), iPts(J)) > PtLen(MyPolyLine.Vertex(i - 1), iPts(J)) Then GoTo SkipIt
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPts(J)
            iCount = iCount + 1
        End If
SkipIt:
    Next J
Next i
PolyLineArcIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function PolyLineEllipseIntersect(MyPolyLine As CadPolyLine, MyEllipse As CadEllipse, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim iRes As Integer
Dim iPts() As CadPoint
Dim iCount As Integer
Dim tLine As CadLine
Dim i As Integer
Dim J As Integer
For i = 1 To UBound(MyPolyLine.Vertex)
    tLine.P1 = MyPolyLine.Vertex(i - 1)
    tLine.P2 = MyPolyLine.Vertex(i)
    iRes = LineEllipseIntersect(tLine, MyEllipse, iPts())
    For J = 0 To iRes
        If iPts(J).Layer.Color = vbBlue Or iPts(J).Layer.Color = vbRed Or i = 1 Or i = UBound(MyPolyLine.Vertex) Then
            If i = 1 Then If PtLen(MyPolyLine.Vertex(0), iPts(J)) > PtLen(MyPolyLine.Vertex(1), iPts(J)) Then GoTo SkipIt
            If i = UBound(MyPolyLine.Vertex) Then If PtLen(MyPolyLine.Vertex(i), iPts(J)) > PtLen(MyPolyLine.Vertex(i - 1), iPts(J)) Then GoTo SkipIt
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPts(J)
            iCount = iCount + 1
        End If
SkipIt:
    Next J
Next i
PolyLineEllipseIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function PolyLineSplineIntersect(MyPolyLine As CadPolyLine, MySPline As CadSpline, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim iRes As Integer
Dim iPts() As CadPoint
Dim iCount As Integer
Dim tLine As CadLine
Dim i As Integer
Dim J As Integer
For i = 1 To UBound(MyPolyLine.Vertex)
    tLine.P1 = MyPolyLine.Vertex(i - 1)
    tLine.P2 = MyPolyLine.Vertex(i)
    iRes = SplineLineIntersect(MySPline, tLine, iPts())
    For J = 0 To iRes
        'If i = 1 Then If PtLen(MyPolyline.Vertex(0), iPts(j)) > PtLen(MyPolyline.Vertex(1), iPts(j)) Then GoTo SkipIt
        'If i = UBound(MyPolyline.Vertex) Then If PtLen(MyPolyline.Vertex(i), iPts(j)) > PtLen(MyPolyline.Vertex(i - 1), iPts(j)) Then GoTo SkipIt
        ReDim Preserve iPoints(iCount) As CadPoint
        iPoints(iCount) = iPts(J)
        iCount = iCount + 1
    Next J
Next i
PolyLineSplineIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function PolyLineLineIntersect(MyPolyLine As CadPolyLine, MyLine As CadLine, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim iRes As Integer
Dim iPt() As CadPoint
Dim iCount As Integer
Dim tLine As CadLine
Dim i As Integer
Dim J As Integer
For i = 1 To UBound(MyPolyLine.Vertex)
    tLine.P1 = MyPolyLine.Vertex(i - 1)
    tLine.P2 = MyPolyLine.Vertex(i)
    If LineLineIntersect(tLine, MyLine, iPt()) = 0 Then
        If iPt(0).Layer.Color = vbRed Or iPt(0).Layer.Color = vbBlue Or i = 1 Or i = UBound(MyPolyLine.Vertex) Then
            If i = 1 Then If PtLen(MyPolyLine.Vertex(0), iPt(0)) > PtLen(MyPolyLine.Vertex(1), iPt(0)) Then GoTo SkipIt
            If i = UBound(MyPolyLine.Vertex) Then If PtLen(MyPolyLine.Vertex(i), iPt(0)) > PtLen(MyPolyLine.Vertex(i - 1), iPt(0)) Then GoTo SkipIt
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPt(0)
            iCount = iCount + 1
        End If
SkipIt:
    End If
Next i
PolyLineLineIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function PolyLinePolyLineIntersect(PolyLineA As CadPolyLine, PolyLineB As CadPolyLine, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim iRes As Integer
Dim iPts() As CadPoint
Dim iCount As Integer
Dim ALine As CadLine
Dim BLine As CadLine
Dim a As Integer
Dim B As Integer
Dim J As Integer
For a = 1 To UBound(PolyLineA.Vertex)
    ALine.P1 = PolyLineA.Vertex(a - 1)
    ALine.P2 = PolyLineA.Vertex(a)
    If PolyLineLineIntersect(PolyLineB, ALine, iPts()) >= 0 Then
        For J = 0 To UBound(iPts)
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPts(J)
            iCount = iCount + 1
        Next J
    End If
Next a
For B = 1 To UBound(PolyLineB.Vertex)
    BLine.P1 = PolyLineB.Vertex(B - 1)
    BLine.P2 = PolyLineB.Vertex(B)
    If PolyLineLineIntersect(PolyLineA, BLine, iPts()) >= 0 Then
        For J = 0 To UBound(iPts)
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPts(J)
            iCount = iCount + 1
        Next J
    End If
Next B
'For A = 1 To UBound(PolyLineA.Vertex)
'    ALine.P1 = PolyLineA.Vertex(A - 1)
'    ALine.P2 = PolyLineA.Vertex(A)
'    For B = 1 To UBound(PolyLineB.Vertex)
'        BLine.P1 = PolyLineB.Vertex(B - 1)
'        BLine.P2 = PolyLineB.Vertex(B)
'        If LineLineIntersect(ALine, BLine, iPt()) = 0 Then
'            If iPt(0).layer.color = vbBlue Or (A = 1 And B = 1) Or (A = 1 And B = UBound(PolyLineB.Vertex)) Or (B = 1 And A = UBound(PolyLineA.Vertex)) Or (A = UBound(PolyLineA.Vertex) And B = UBound(PolyLineB.Vertex)) Then
'                If A = 1 Then If PtLen(PolyLineA.Vertex(0), iPt(0)) > PtLen(PolyLineA.Vertex(1), iPt(0)) Then GoTo SkipIt
'                If A = UBound(PolyLineA.Vertex) Then If PtLen(PolyLineA.Vertex(A), iPt(0)) > PtLen(PolyLineA.Vertex(A - 1), iPt(0)) Then GoTo SkipIt
'                If B = 1 Then If PtLen(PolyLineB.Vertex(0), iPt(0)) > PtLen(PolyLineB.Vertex(1), iPt(0)) Then GoTo SkipIt
'                If B = UBound(PolyLineB.Vertex) Then If PtLen(PolyLineB.Vertex(B), iPt(0)) > PtLen(PolyLineB.Vertex(B - 1), iPt(0)) Then GoTo SkipIt
'                ReDim Preserve iPoints(iCount) As CadPoint
'                iPoints(iCount) = iPt(0)
'                iCount = iCount + 1
'            End If
'SkipIt:
'        End If
'    Next B
'Next A
PolyLinePolyLineIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function RevLine(MyLine As CadLine) As CadLine
RevLine.P1 = MyLine.P2
RevLine.P2 = MyLine.P1
RevLine.Layer = MyLine.Layer
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Function SIGN(x)
SIGN = Sgn(x)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub SpaceEllipse(MyEllipse As CadEllipse, StartPt As CadPoint, Dist As Single, Count As Integer, SidePt As CadPoint, ByRef sPoints() As CadPoint, ByRef sAngles() As Single)
Erase sPoints()
Erase sAngles()
Dim sAngle As Single
Dim i As Integer
Dim M As Single
Dim P As Single
Dim q As Integer
Dim Side As Integer
M = cAngle(cAngLine(90 + EllipsePtAngle(MyEllipse, StartPt), StartPt, 1, True))
P = M - PtPtAngle(StartPt, SidePt)
q = PtPtAngle(MyEllipse.F1, MyEllipse.F2)
If P >= 0 Then
    Side = IIf(P > 90, -1, 1)
Else
    Side = IIf(P < -90, -1, 1)
End If
If q >= 90 And q <= 270 Then Side = Side * -1
sAngle = EllipsePtAngle(MyEllipse, StartPt)
ReDim Preserve sPoints(0) As CadPoint
sPoints(0) = EllipsePt(MyEllipse, sAngle)
sPoints(0).Layer.Color = vbBlue
sPoints(0).Layer.Width = 3
ReDim Preserve sAngles(0) As Single
sAngles(0) = EllipseAngle(sAngle, MyEllipse)
For i = 1 To Count - 1
    ReDim Preserve sPoints(i) As CadPoint
    sPoints(i) = EllipseLenPt(MyEllipse, sAngle, Dist, Side)
    sPoints(i).Layer.Color = vbBlue
    sPoints(i).Layer.Width = 3
    ReDim Preserve sAngles(i) As Single
    sAngles(i) = EllipseAngle(sAngle, MyEllipse)
    sAngle = EllipsePtAngle(MyEllipse, sPoints(i))
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub SpaceArc(MyArc As CadArc, StartPt As CadPoint, Dist As Single, Count As Integer, SidePt As CadPoint, ByRef sPoints() As CadPoint, ByRef sAngles() As Single)
Erase sPoints()
Erase sAngles()
Dim sAngle As Single
Dim i As Integer
Dim M As Single
Dim P As Single
Dim Side As Integer
Dim interval As Single
M = cAngle(cAngLine(90 + ArcPtAngle(MyArc, StartPt), StartPt, 1, False))
P = M - PtPtAngle(StartPt, SidePt)
If P >= 0 Then
    Side = IIf(P > 90, -1, 1)
Else
    Side = IIf(P < -90, -1, 1)
End If

sAngle = ArcPtAngle(MyArc, StartPt)
interval = (Dist / (2 * Pi * MyArc.Radius) * 360)
For i = 0 To Count - 1
    ReDim Preserve sPoints(i) As CadPoint
    sPoints(i) = ArcPt(MyArc, sAngle)
    sPoints(i).Layer.Color = vbBlue
    sPoints(i).Layer.Width = 3
    ReDim Preserve sAngles(i) As Single
    sAngles(i) = sAngle
    sAngle = sAngle + (interval * Side)
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub SpaceLine(MyLine As CadLine, StartPt As CadPoint, Dist As Single, Count As Integer, SidePt As CadPoint, ByRef sPoints() As CadPoint, ByRef sAngles() As Single)
Erase sPoints()
Erase sAngles()
Dim i As Integer
Dim sCount As Integer
Dim M As Single
Dim P As Single
Dim Side As Integer
M = cAngle(MyLine)
P = M - PtPtAngle(StartPt, SidePt)
If P >= 0 Then
    Side = IIf(P > 90, -1, 1)
Else
    Side = IIf(P < -90, -1, 1)
End If
For i = 0 To Count - 1
    ReDim Preserve sPoints(i) As CadPoint
    sPoints(i).Layer.Color = vbBlue
    sPoints(i).Layer.Width = 3
    sPoints(i).x = StartPt.x + Cos(M * Pi / 180) * Dist * i * Side
    sPoints(i).y = StartPt.y + Sin(M * Pi / 180) * Dist * i * Side
    ReDim Preserve sAngles(i) As Single
    sAngles(i) = M
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub SpacePolyLine(MyPoly As CadPolyLine, StartPt As CadPoint, Dist As Single, Count As Integer, SidePt As CadPoint, ByRef sPoints() As CadPoint, ByRef sAngles() As Single)
Erase sPoints()
Erase sAngles()
Dim i As Integer
Dim sCount As Integer
Dim sV As Integer
Dim M As Single
Dim P As Single
Dim Side As Integer
Dim sLine As CadLine
Dim totDist As Single
Dim rDist As Single
Dim lPt As CadPoint
'First we have to find out which segment that the start point is on
For sV = 1 To UBound(MyPoly.Vertex)
    sLine.P1 = MyPoly.Vertex(sV - 1)
    sLine.P2 = MyPoly.Vertex(sV)
    If PtInLine(sLine, StartPt) Then Exit For
Next sV
'Now we have to figure out the direction to follow the polyline
M = cAngle(sLine)
P = M - PtPtAngle(StartPt, SidePt)
If P >= 0 Then
    Side = IIf(P >= 90, -1, 1)
Else
    Side = IIf(P <= -90, -1, 1)
End If
lPt = StartPt
ReDim Preserve sPoints(sCount) As CadPoint
sPoints(sCount) = StartPt
sPoints(sCount).Layer.Color = vbBlue
sPoints(sCount).Layer.Width = 3
ReDim Preserve sAngles(sCount) As Single
sAngles(sCount) = M
sCount = sCount + 1
If Side > 0 Then
    'space points from the start to the next vertex on the polyline (towards the last vertex)
    Do While totDist + Dist < PtLen(StartPt, MyPoly.Vertex(sV)) And sCount <= Count - 1
        ReDim Preserve sPoints(sCount) As CadPoint
        sPoints(sCount).Layer.Color = vbBlue
        sPoints(sCount).Layer.Width = 3
        sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * Dist
        sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * Dist
        totDist = totDist + Dist
        lPt = sPoints(sCount)
        ReDim Preserve sAngles(sCount) As Single
        sAngles(sCount) = M
        sCount = sCount + 1
    Loop
    rDist = Dist - PtLen(lPt, MyPoly.Vertex(sV))  'remaining distance left
    For i = sV To UBound(MyPoly.Vertex) - 1
        M = PtPtAngle(MyPoly.Vertex(i), MyPoly.Vertex(i + 1))
        'We need to space a point along the current vertex equal to the remaining distance
        If rDist < PtLen(MyPoly.Vertex(i), MyPoly.Vertex(i + 1)) And sCount <= Count - 1 Then
            lPt = MyPoly.Vertex(i)
            ReDim Preserve sPoints(sCount) As CadPoint
            sPoints(sCount).Layer.Color = vbBlue
            sPoints(sCount).Layer.Width = 3
            sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * rDist
            sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * rDist
            totDist = totDist + Dist
            lPt = sPoints(sCount)
            ReDim Preserve sAngles(sCount) As Single
            sAngles(sCount) = M
            sCount = sCount + 1
            Do While PtLen(lPt, MyPoly.Vertex(i + 1)) > Dist And sCount <= Count - 1
                ReDim Preserve sPoints(sCount) As CadPoint
                sPoints(sCount).Layer.Color = vbBlue
                sPoints(sCount).Layer.Width = 3
                sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * Dist
                sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * Dist
                totDist = totDist + Dist
                lPt = sPoints(sCount)
                ReDim Preserve sAngles(sCount) As Single
                sAngles(sCount) = M
                sCount = sCount + 1
            Loop
            rDist = Dist - PtLen(lPt, MyPoly.Vertex(i + 1)) 'remaining distance left
        Else
            rDist = rDist - PtLen(MyPoly.Vertex(i), MyPoly.Vertex(i + 1))
            lPt = MyPoly.Vertex(i + 1)
        End If
    Next i
    lPt = MyPoly.Vertex(i)
    'space points for the remaining distance along the vector of the last segment
    'We need to space a point along the current vertex equal to the remaining distance
    If rDist <> 0 And sCount <= Count - 1 Then
        ReDim Preserve sPoints(sCount) As CadPoint
        sPoints(sCount).Layer.Color = vbBlue
        sPoints(sCount).Layer.Width = 3
        sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * rDist
        sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * rDist
        totDist = totDist + Dist
        lPt = sPoints(sCount)
        ReDim Preserve sAngles(sCount) As Single
        sAngles(sCount) = M
        sCount = sCount + 1
    End If
    For i = sCount To Count - 1
        ReDim Preserve sPoints(i) As CadPoint
        sPoints(sCount).Layer.Color = vbBlue
        sPoints(sCount).Layer.Width = 3
        sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * Dist
        sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * Dist
        totDist = totDist + Dist
        lPt = sPoints(sCount)
        ReDim Preserve sAngles(sCount) As Single
        sAngles(sCount) = M
        sCount = sCount + 1
    Next i
Else
    'space points from the start to the next vertex on the polyline (towards the last vertex)
    Do While totDist + Dist < PtLen(StartPt, MyPoly.Vertex(sV)) And sCount <= Count - 1
        ReDim Preserve sPoints(sCount) As CadPoint
        sPoints(sCount).Layer.Color = vbBlue
        sPoints(sCount).Layer.Width = 3
        sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * Dist
        sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * Dist
        totDist = totDist + Dist
        lPt = sPoints(sCount)
        ReDim Preserve sAngles(sCount) As Single
        sAngles(sCount) = M
        sCount = sCount + 1
    Loop
    rDist = Dist - PtLen(lPt, MyPoly.Vertex(sV)) 'remaining distance left
    For i = sV To 1 Step -1
        M = PtPtAngle(MyPoly.Vertex(i), MyPoly.Vertex(i - 1))
        'We need to space a point along the current vertex equal to the remaining distance
        If rDist < PtLen(MyPoly.Vertex(i), MyPoly.Vertex(i - 1)) And sCount <= Count - 1 Then
            lPt = MyPoly.Vertex(i)
            ReDim Preserve sPoints(sCount) As CadPoint
            sPoints(sCount).Layer.Color = vbBlue
            sPoints(sCount).Layer.Width = 3
            sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * rDist
            sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * rDist
            totDist = totDist + Dist
            lPt = sPoints(sCount)
            ReDim Preserve sAngles(sCount) As Single
            sAngles(sCount) = M
            sCount = sCount + 1
            Do While PtLen(lPt, MyPoly.Vertex(i - 1)) > Dist And sCount <= Count - 1
                ReDim Preserve sPoints(sCount) As CadPoint
                sPoints(sCount).Layer.Color = vbBlue
                sPoints(sCount).Layer.Width = 3
                sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * Dist
                sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * Dist
                totDist = totDist + Dist
                lPt = sPoints(sCount)
                ReDim Preserve sAngles(sCount) As Single
                sAngles(sCount) = M
                sCount = sCount + 1
            Loop
            rDist = Dist - PtLen(lPt, MyPoly.Vertex(i - 1)) 'remaining distance left
        Else
            rDist = rDist - PtLen(MyPoly.Vertex(i), MyPoly.Vertex(i - 1))
            lPt = MyPoly.Vertex(i - 1)
        End If
    Next i
    lPt = MyPoly.Vertex(0)
    'space points for the remaining distance along the vector of the last segment
    'We need to space a point along the current vertex equal to the remaining distance
    If rDist > 0 And sCount <= Count - 1 Then
        ReDim Preserve sPoints(sCount) As CadPoint
        sPoints(sCount).Layer.Color = vbBlue
        sPoints(sCount).Layer.Width = 3
        sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * rDist
        sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * rDist
        totDist = totDist + Dist
        lPt = sPoints(sCount)
        ReDim Preserve sAngles(sCount) As Single
        sAngles(sCount) = M
        sCount = sCount + 1
    End If
    For i = sCount To Count - 1
        ReDim Preserve sPoints(i) As CadPoint
        sPoints(sCount).Layer.Color = vbBlue
        sPoints(sCount).Layer.Width = 3
        sPoints(sCount).x = lPt.x + Cos(M * Pi / 180) * Dist
        sPoints(sCount).y = lPt.y + Sin(M * Pi / 180) * Dist
        totDist = totDist + Dist
        lPt = sPoints(sCount)
        ReDim Preserve sAngles(sCount) As Single
        sAngles(sCount) = M
        sCount = sCount + 1
    Next i
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub SpaceSelection(MyGeo As Geometry, SpaceSel As SelSet, StartPt As CadPoint, Dist As Single, Count As Integer, SidePt As CadPoint, Group As Boolean, PivotPt As CadPoint, MySel() As SelSet, Contour As Boolean)
On Error GoTo eTrap
Dim sPoints() As CadPoint
Dim sAngles() As Single
Dim i As Integer
Dim K As Integer
Dim StartArray As Integer
Dim LastAngle As Single
Select Case SpaceSel.Type
    Case "Line": SpaceLine MyGeo.Lines(SpaceSel.Index), StartPt, Dist, Count, SidePt, sPoints(), sAngles()
    Case "Arc": SpaceArc MyGeo.Arcs(SpaceSel.Index), StartPt, Dist, Count, SidePt, sPoints(), sAngles()
    Case "Ellipse": SpaceEllipse MyGeo.Ellipses(SpaceSel.Index), StartPt, Dist, Count, SidePt, sPoints(), sAngles()
    Case "Spline": SpaceSpline MyGeo.Splines(SpaceSel.Index), StartPt, Dist, Count, SidePt, sPoints(), sAngles()
    Case "PolyLine": SpacePolyLine MyGeo.PolyLines(SpaceSel.Index), StartPt, Dist, Count, SidePt, sPoints(), sAngles()
End Select
If Not Group Then
    K = UBound(MyGeo.Points) + 1
    For i = 0 To UBound(sPoints)
        ReDim Preserve MyGeo.Points(K) As CadPoint
        MyGeo.Points(K) = sPoints(i)
        K = K + 1
    Next i
Else
    If sPoints(0).x = PivotPt.x And sPoints(0).y = PivotPt.y Then StartArray = 1 Else StartArray = 0
    For i = StartArray To UBound(sPoints)
        MoveSelection MyGeo, MySel(), sPoints(i).x - PivotPt.x, sPoints(i).y - PivotPt.y, True
        PivotPt = sPoints(i)
        LastAngle = sAngles(i - 1)
        If Contour Then RotateSelection MyGeo, MySel(), sAngles(i) - LastAngle, sPoints(i), False
    Next i
End If
Exit Sub
eTrap:
    K = 0
    LastAngle = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub SpaceSpline(MySPline As CadSpline, StartPt As CadPoint, Dist As Single, Count As Integer, SidePt As CadPoint, ByRef sPoints() As CadPoint, ByRef sAngles() As Single)
SpacePolyLine SplineToPolyLine(MySPline), StartPt, Dist, Count, SidePt, sPoints(), sAngles()
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function SplineLen(MySPline As CadSpline) As Single
SplineLen = PolyLineLen(SplineToPolyLine(MySPline))
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ZZ_OLD_SplineToPolyLine(MySPline As CadSpline) As CadPolyLine
Dim du As Single
Dim vX As Single
Dim vY As Single
Dim bv As Single
Dim K As Single
Dim u As Single
Dim vCount As Integer
Dim i As Integer
vCount = UBound(MySPline.Vertex)
'If (MySpline.Vertex(VCount).x = MySpline.Vertex(VCount - 1).x) And (MySpline.Vertex(VCount).y = MySpline.Vertex(VCount - 1).y) Then Exit Sub
du = 0.025 'SplineSmooth
ReDim ZZ_OLD_SplineToPolyLine.Vertex(i) As CadPoint
ZZ_OLD_SplineToPolyLine.Vertex(i).x = MySPline.Vertex(0).x
ZZ_OLD_SplineToPolyLine.Vertex(i).y = MySPline.Vertex(0).y
For u = 0 To 1 Step du
    vX = 0: vY = 0
    For K = 0 To vCount ' For Each control point
        bv = sBlend(K, vCount, u) ' Calculate blending Function
        vX = vX + MySPline.Vertex(K).x * bv
        vY = vY + MySPline.Vertex(K).y * bv
    Next K
    ReDim Preserve ZZ_OLD_SplineToPolyLine.Vertex(i) As CadPoint
    ZZ_OLD_SplineToPolyLine.Vertex(i).x = vX
    ZZ_OLD_SplineToPolyLine.Vertex(i).y = vY
    i = i + 1
Next u
ReDim Preserve ZZ_OLD_SplineToPolyLine.Vertex(i) As CadPoint
ZZ_OLD_SplineToPolyLine.Vertex(i).x = MySPline.Vertex(vCount).x
ZZ_OLD_SplineToPolyLine.Vertex(i).y = MySPline.Vertex(vCount).y
ZZ_OLD_SplineToPolyLine.Layer = MySPline.Layer
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub ZZ_NONANG_DrawCadEllipse(Canvas As PictureBox, MyEllipse As CadEllipse, Optional Mode As Integer = 13, Optional Color As Long = -1)
Dim TotLen As Single, Flen As Single ' point measurements
Dim a As Single, B As Single ' ellipse parameters
Dim u As Single, v As Single, r As Single, f As Integer, t As Single ' iiteration parameters
Dim Start As Single, Finish As Single ' determines ellipse angle boundaries
Dim ang1 As Single, ang2 As Single 'adjusted ellipse angles
Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single 'plotting coordinates
Dim cx As Single, cy As Single ' center of the ellipse
Dim i As Single
Dim First As Boolean
'-----------Setup Canvas-----------
Canvas.DrawWidth = MyEllipse.Layer.Width
Canvas.DrawMode = Mode
Canvas.DrawStyle = MyEllipse.Layer.style
If Color = -1 Then Color = MyEllipse.Layer.Color
'-----------------Dimensional Calculations-------
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2)
If a = 0 Then Exit Sub
'-------------Get Boundaries (for angles)
ang1 = MyEllipse.Angle1
ang2 = MyEllipse.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
Start = Format((ang1 / 360), "0.000")
Finish = Format((ang2 / 360), "0.000")
'--------------------------
For r = Start To Finish Step 0.01
    If (r >= 0 And r < 0.5) Or (r >= 1 And r < 1.5) Then f = 1 Else f = -1
    If r > 1 Then t = r - 1 Else t = r
    If f > 0 Then u = (a - (4 * a * t)) Else u = (-a + (4 * a * (t - 0.5)))
    X1 = u
    Y1 = (Sqr((B ^ 2 * (1 - (u ^ 2 / a ^ 2))))) * f
    X2 = RotX(X1, Y1, v) + cx
    Y2 = RotY(X1, Y1, v) + cy
    If Not First Then Canvas.PSet (X2, -Y2), Color: First = True
    Canvas.Line -(X2, -Y2), Color
Next r
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub DrawCadEllipse(Canvas As PictureBox, MyEllipse As CadEllipse, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
Dim a As Single, B As Single ' ellipse parameters
Dim u As Single, v As Single
Dim i As Single, J As Single
Dim ang1 As Single, ang2 As Single 'adjusted ellipse angles
Dim X1 As Single, Y1 As Single 'plotting coordinates
Dim cx As Single, cy As Single ' center of the ellipse
Dim First As Boolean
Dim TotLen As Single, Flen As Single
'-----------Setup Canvas-----------
If Width = 0 Then Width = MyEllipse.Layer.Width
Canvas.DrawWidth = Width
Canvas.DrawMode = Mode
Canvas.DrawStyle = MyEllipse.Layer.style
If Color = -1 Then Color = MyEllipse.Layer.Color
'-----------------Angle adjustment-------
ang1 = MyEllipse.Angle1
ang2 = MyEllipse.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
i = (ang2 - ang1) / MyEllipse.NumPoints
If i = 0 Then Exit Sub
'-------------------------------------------
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2) * Pi / 180
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
If a = 0 Then Exit Sub
For u = ang1 To ang2 Step i
    J = u * Pi / 180
    X1 = Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    Y1 = Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    If Not First Then Canvas.PSet (cx + X1, -cy - Y1), Color: First = True
    If First Then Canvas.Line -(cx + X1, -cy - Y1), Color
Next u
If u - i < ang2 Then
    J = ang2 * Pi / 180
    X1 = Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    Y1 = Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
    Canvas.Line -(cx + X1, -cy - Y1), Color
End If
Exit Sub
'For U = Ang1 To Ang2 Step i
'    X1 = a * Cos(U * pi / 180)
'    Y1 = B * Sin(U * pi / 180)
'    Hyp = IIf(X1 < 0, -Sqr((X1 ^ 2) + (Y1 ^ 2)), Sqr((X1 ^ 2) + (Y1 ^ 2)))
'    j = IIf(X1 = 0, pi / 2, Atn(Y1 / X1))
'    If j + v > 2 * pi Then j = j + (2 * pi)
'    X2 = (Hyp * Cos(v + j))
'    Y2 = (Hyp * Sin(v + j))
'    If First Then Canvas.Line (cx + X3, -cy - Y3)-(cx + X2, -cy - Y2), Color
'    X3 = X2
'    Y3 = Y2
'    First = True
'Next U
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub DrawCadText(Canvas As PictureBox, MyText As CadText, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1)
Dim f As LOGFONT
Dim hPrevFont As Long
Dim hFont As Long
Dim XSIZE As Integer
Dim YSIZE As Integer
Canvas.DrawMode = Mode
If Color = -1 Then Color = MyText.Layer.Color
If MyText.Layer.FontName = "" Then MyText.Layer.FontName = "Arial Black"
f.lfEscapement = 10 * Val(MyText.Angle)  'rotation angle, in tenths
f.lfFacename = MyText.Layer.FontName + Chr$(0)
XSIZE = Canvas.ScaleX(MyText.Size, 0, 2)
YSIZE = Canvas.ScaleY(MyText.Size, 0, 2)
If XSIZE = 0 Then XSIZE = 1
If YSIZE = 0 Then YSIZE = 1
f.lfWidth = (XSIZE * -15) / Screen.TwipsPerPixelY
f.lfHeight = (YSIZE * -20) / Screen.TwipsPerPixelY
hFont = CreateFontIndirect(f)
hPrevFont = SelectObject(Canvas.hDC, hFont)
Canvas.ForeColor = Color
Canvas.CurrentX = MyText.Start.x
Canvas.CurrentY = -MyText.Start.y - MyText.Size
Canvas.Print MyText.Text
MyText.Length = Canvas.TextWidth(MyText.Text)
'  Clean up, restore original font
hFont = SelectObject(Canvas.hDC, hPrevFont)
DeleteObject hFont
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub DrawCadPoint(Canvas As PictureBox, MyPoint As CadPoint, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
If Color = -1 Then Color = MyPoint.Layer.Color
If Width = 0 Then Width = MyPoint.Layer.Width
Canvas.DrawWidth = Width
Canvas.DrawMode = Mode
Canvas.DrawStyle = MyPoint.Layer.style
Canvas.PSet (MyPoint.x, -MyPoint.y), Color
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub DrawPointArray(Canvas As PictureBox, pts() As CadPoint, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
On Error GoTo eTrap
Dim i As Integer
For i = 0 To UBound(pts)
    DrawCadPoint Canvas, pts(i), Mode, Color, Width
Next i
eTrap:
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub ZZ_OLD_DrawCadEllipse(Canvas As PictureBox, MyEllipse As CadEllipse, Optional Mode As Integer = 13, Optional Color As Long = -1)
'Dim a As Single, b As Single
'Dim RotAngle As Single
'Dim A1 As Single, A2 As Single
''----------------
'A1 = MyEllipse.Angle1
'A2 = MyEllipse.Angle2
'If A2 < A1 Then A2 = A2 + (360)
''----------------
'Dim X1 As Single, Y1 As Single
'Dim X2 As Single, Y2 As Single
'Dim X3 As Single, Y3 As Single
'Dim Ratio As Single
'Dim Hyp As Single
'Dim j As Single
'Dim u As Single
'Dim Count As Integer
'Dim interval As Single
'Dim L1 As CadLine, L2 As CadLine
'Dim hPt As CadPoint
'L1.P1 = MyEllipse.Center
'L1.P2 = MyEllipse.P1
'L2.P1 = MyEllipse.Center
'L2.P2 = MyEllipse.P2
'RotAngle = cAngle(L1) * pi / 180
'Ratio = LineLen(L2) / LineLen(L1)
'hPt.X = LineLen(L1) * Cos(RotAngle)
'hPt.y = LineLen(L1) * Sin(RotAngle)
''--------------FOCI STUFF------------
'L1 = cAngLine(cAngle(L1), MyEllipse.Center, (LineLen(L1) + LineLen(L2)) / 2, True)
'L1.Width = 1
'L1.Color = vbRed
''DrawCadLine Canvas, L1, Mode
''--------------Get A and B------------
'a = Sqr((hPt.X ^ 2) + (hPt.y ^ 2))
'b = Ratio * a
'interval = (A2 - A1) / MyEllipse.NumPoints
'If interval = 0 Then Exit Sub
''-----------Setup Canvas-----------
'Canvas.DrawWidth = MyEllipse.Width
'Canvas.DrawMode = Mode
'Canvas.DrawStyle = MyEllipse.Style
''------------------------
'If Color = -1 Then Color = MyEllipse.Color
'For u = A1 To A2 Step interval
'    X1 = a * Cos(u * pi / 180)
'    Y1 = b * Sin(u * pi / 180)
'    Hyp = Sqr((X1 ^ 2) + (Y1 ^ 2))
'    If X1 = 0 Then j = pi / 2 Else j = Atn(Y1 / X1)
'    If X1 < 0 Then Hyp = -Hyp
'    X2 = (Hyp * Cos(RotAngle + j))
'    Y2 = (Hyp * Sin(RotAngle + j))
'
'    If Count > 0 Then Canvas.Line (MyEllipse.Center.X + X3, -MyEllipse.Center.y - Y3)-(MyEllipse.Center.X + X2, -MyEllipse.Center.y - Y2), Color
'    X3 = X2
'    Y3 = Y2
'    Count = Count + 1
'Next u
'If u - interval <> A2 Then
'    X1 = a * Cos(A2 * pi / 180)
'    Y1 = b * Sin(A2 * pi / 180)
'    Hyp = Sqr((X1 ^ 2) + (Y1 ^ 2))
'    If X1 = 0 Then j = pi / 2 Else j = Atn(Y1 / X1)
'    If X1 < 0 Then Hyp = -Hyp
'    X2 = (Hyp * Cos(RotAngle + j))
'    Y2 = (Hyp * Sin(RotAngle + j))
'    Canvas.Line (MyEllipse.Center.X + X3, -MyEllipse.Center.y - Y3)-(MyEllipse.Center.X + X2, -MyEllipse.Center.y - Y2), Color
'End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function DelaunayCenter(P1 As CadPoint, P2 As CadPoint, P3 As CadPoint, ByRef CenterPoint As CadPoint) As Integer
On Local Error Resume Next
Dim L1 As CadLine
Dim L2 As CadLine
Dim L3 As CadLine
Dim Mid1 As CadLine
Dim Mid2 As CadLine
Dim Mid3 As CadLine
Dim tPt() As CadPoint
L1.P1 = P1
L1.P2 = P2
L2.P1 = P2
L2.P2 = P3
L3.P1 = P3
L3.P2 = P1
Mid1 = cAngLine(cAngle(L1) + 90, MidPoint(L1), 50, True)
Mid2 = cAngLine(cAngle(L2) + 90, MidPoint(L2), 50, True)
Mid3 = cAngLine(cAngle(L3) + 90, MidPoint(L3), 50, True)
DelaunayCenter = LineLineIntersect(Mid1, Mid2, tPt())
CenterPoint = tPt(0)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function cAngLine(Angle As Single, pt As CadPoint, h As Single, Optional Extend As Boolean) As CadLine
cAngLine.P1 = pt
If Extend Then
    cAngLine.P1.x = -h * Cos(Angle * Pi / 180) + pt.x
    cAngLine.P1.y = -h * Sin(Angle * Pi / 180) + pt.y
End If
cAngLine.P2.x = h * Cos(Angle * Pi / 180) + pt.x
cAngLine.P2.y = h * Sin(Angle * Pi / 180) + pt.y
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function cAngPt(Angle As Single, pt As CadPoint, h As Single) As CadPoint
cAngPt.x = h * Cos(Angle * Pi / 180) + pt.x
cAngPt.y = h * Sin(Angle * Pi / 180) + pt.y
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function AnglePt(Angle As Single, pt As CadPoint, h As Single) As CadPoint
AnglePt.x = h * Cos(Angle * Pi / 180) + pt.x
AnglePt.y = h * Sin(Angle * Pi / 180) + pt.y
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub RadiusSelection(MyGeo As Geometry, SelA As SelSet, EndA As Integer, SelB As SelSet, EndB As Integer, MovePt As CadPoint, rad As Single)
On Error GoTo eTrap
Dim Dist1 As Single
Dim Angle1 As Single
Dim Angle2 As Single
Dim Angle3 As Single
Dim K As Integer
Dim iPts() As CadPoint
'-----------------Get the inside angle of the lines
Angle1 = cAngle(MyGeo.Lines(SelA.Index))
Angle3 = cAngle(MyGeo.Lines(SelB.Index))
If EndA = 1 Then Angle1 = dAngle(Angle1 + 180)
If EndB = 1 Then Angle3 = dAngle(Angle3 + 180)
If Angle3 < Angle1 Xor Abs(Angle3 - Angle1) > 180 Then
    Swap Angle1, Angle3
    Swap SelA.Index, SelB.Index
    Swap EndA, EndB
End If
Angle2 = (Angle3 - Angle1) / 2
'---------------------Get the bisecting line---------
Dist1 = rad / Sin(Angle2 * Pi / 180)
'--------Add the arc to the geometry------------
K = UBound(MyGeo.Arcs) + 1
ReDim Preserve MyGeo.Arcs(K) As CadArc
'mygeo.Arcs(k).Center = CurLine.P2
MyGeo.Arcs(K).Center = AnglePt(Angle1 + Angle2, MovePt, Dist1)
MyGeo.Arcs(K).Radius = rad
MyGeo.Arcs(K).Angle1 = Angle1 + Angle2 + 90
MyGeo.Arcs(K).Angle2 = Angle1 + Angle2 - 90
'mygeo.Arcs(k).angle1 = -(angle1 + angle2) + (270 + angle2)
'mygeo.Arcs(k).angle2 = (angle1 + angle2) + (270 - angle2)
MyGeo.Arcs(K).Layer = MyGeo.Lines(SelA.Index).Layer
'---------Corner / Relimit the lines to the Tangent point
LineArcIntersect MyGeo.Lines(SelA.Index), MyGeo.Arcs(K), iPts
RelimitLine MyGeo, SelA.Index, EndA, iPts(0)
RelimitArc MyGeo, K, 1, iPts(0)
LineArcIntersect MyGeo.Lines(SelB.Index), MyGeo.Arcs(K), iPts
RelimitLine MyGeo, SelB.Index, EndB, iPts(0)
RelimitArc MyGeo, K, 0, iPts(0)
'-------------Add Arc to the selection---------
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function ZZ_nonAng_EllipseAngle(MyAngle As Single, MyEllipse As CadEllipse) As Single
'----------------------------
Dim TotLen As Single, Flen As Single ' point measurements
Dim a As Single, B As Single, v As Single, u As Single ' ellipse parameters
Dim X1 As Single, Y1 As Single
Dim cx As Single, cy As Single ' center of the ellipse
Dim MyLine As CadLine
Dim tLine As CadLine
Dim M As Single, c As Single
Dim tPt As CadPoint
'-----------------Dimensional Calculations-------
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2)
'-----------Setup out intersection line---------
MyLine = cAngLine(MyAngle, EllipseCenter(MyEllipse), TotLen)
'-----------First we rotate our intersection line negative the rotation of the ellipse
tLine = RotateLine(MyLine, EllipseCenter(MyEllipse), -v)
'-----------Next we move the line so it is based upon the center 0,0
tLine = MoveLine(tLine, -cx, -cy)
'----------Get our variables for: Y = mx+c
M = CSlope(tLine)
c = (tLine.P1.y) - (M * (tLine.P1.x))
'----------substitue (mx+c) into the equation for ellipse (x^2/a^2 +y^2/b^2 = 1) as variable "y" and solve for "x"
If MyAngle <= 180 Then
    X1 = (a * (B * Sqr(Abs(a ^ 2 * M ^ 2 + B ^ 2 - c ^ 2)) - a * c * M)) / (a ^ 2 * M ^ 2 + B ^ 2)
    Y1 = M * X1 + c
    If MyAngle <= 90 Then
        ZZ_nonAng_EllipseAngle = ((a - X1) / (4 * a)) * 360
    Else
        ZZ_nonAng_EllipseAngle = ((a + X1) / (4 * a)) * 360
    End If
Else
    X1 = -(a * (B * Sqr(Abs(a ^ 2 * M ^ 2 + B ^ 2 - c ^ 2)) + a * c * M)) / (a ^ 2 * M ^ 2 + B ^ 2)
    Y1 = M * X1 + c
    If MyAngle <= 270 Then
        ZZ_nonAng_EllipseAngle = ((3 * a + X1) / (4 * a)) * 360
    Else
        ZZ_nonAng_EllipseAngle = ((3 * a - X1) / (4 * a)) * 360
    End If
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function EllipseAngle(MyAngle As Single, MyEllipse As CadEllipse) As Single
'----------------------------
Dim TotLen As Single, Flen As Single ' point measurements
Dim a As Single, B As Single, v As Single, u As Single, z As Single ' ellipse parameters
Dim X1 As Single, Y1 As Single
Dim cx As Single, cy As Single ' center of the ellipse
'-----------------Dimensional Calculations-------
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2)
'tEllipse = RotateEllipse(MyEllipse, EllipseCenter(MyEllipse), -v)
'-------Derive Criteria------------- = (-pi / 2 < z - v < pi / 2)
z = (MyAngle - v) * Pi / 180
If z < 0 Then z = z + 2 * Pi
If z < Pi / 2 Then
    u = ATAN((a * Tan(z)) / B)
ElseIf z >= Pi / 2 And z < 1.5 * Pi Then
    u = ATAN((a * Tan(z)) / B) + Pi
ElseIf z >= 1.5 * Pi Then
    u = ATAN((a * Tan(z)) / B) + 2 * Pi
End If
If u < 0 Then u = u + 2 * Pi
EllipseAngle = (u * 180 / Pi)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function EllipseArcIntersect(MyEllipse As CadEllipse, MyArc As CadArc, ByRef iPoints() As CadPoint) As Integer
Dim tEllipse As CadEllipse
tEllipse = ArcToEllipse(MyArc)
EllipseArcIntersect = EllipseEllipseIntersect(MyEllipse, tEllipse, iPoints())
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function EllipseCenter(MyEllipse As CadEllipse) As CadPoint
EllipseCenter.x = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
EllipseCenter.y = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ZZ_OLD_EllipseEllipseIntersect(EllipseA As CadEllipse, EllipseB As CadEllipse, ByRef iPoints() As CadPoint) As Integer
On Local Error GoTo eTrap
Erase iPoints()
'---INtersept stuff-------
Dim iRes As Integer
Dim iPt() As CadPoint
Dim iCount As Integer
'----------EllipseA------------
Dim TotLenA As Single, fLenA As Single ' point measurements
Dim aA As Single, bA As Single ' ellipse parameters
Dim uA As Single, vA As Single, rA As Single, fA As Integer ' iiteration parameters
Dim X1A As Single, Y1A As Single
Dim cxA As Single, cyA As Single ' center of the ellipse
Dim iLineA As CadLine
'-----------------Dimensional Calculations-------
fLenA = PtLen(EllipseA.F1, EllipseA.F2)
TotLenA = PtLen(EllipseA.F1, EllipseA.P1) + PtLen(EllipseA.F2, EllipseA.P1)
aA = TotLenA / 2
bA = Sqr((TotLenA / 2) ^ 2 - (fLenA / 2) ^ 2)
cxA = (EllipseA.F1.x + EllipseA.F2.x) / 2
cyA = (EllipseA.F1.y + EllipseA.F2.y) / 2
vA = PtPtAngle(EllipseA.F1, EllipseA.F2)
'----------EllipseB------------
Dim TotLenB As Single, fLenB As Single ' point measurements
Dim aB As Single, bB As Single ' ellipse parameters
Dim uB As Single, vB As Single, rB As Single, fB As Integer ' iiteration parameters
Dim X1B As Single, Y1B As Single
Dim cxB As Single, cyB As Single ' center of the ellipse
Dim iLineB As CadLine
'-----------------Dimensional Calculations-------
fLenB = PtLen(EllipseB.F1, EllipseB.F2)
TotLenB = PtLen(EllipseB.F1, EllipseB.P1) + PtLen(EllipseB.F2, EllipseB.P1)
aB = TotLenB / 2
bB = Sqr((TotLenB / 2) ^ 2 - (fLenB / 2) ^ 2)
cxB = (EllipseB.F1.x + EllipseB.F2.x) / 2
cyB = (EllipseB.F1.y + EllipseB.F2.y) / 2
vB = PtPtAngle(EllipseB.F1, EllipseB.F2)
'-------------------------------------------------------------------------
'---------------
iLineA.P1.x = RotX(aA, 0, vA) + cxA
iLineA.P1.y = RotY(aA, 0, vA) + cyA
For rA = 0 To 1 Step 0.01
    If (rA >= 0 And rA < 0.5) Or (rA >= 1 And rA < 1.5) Then fA = 1 Else fA = -1
    If fA > 0 Then uA = (aA - (4 * aA * rA)) Else uA = (-aA + (4 * aA * (rA - 0.5)))
    X1A = uA
    Y1A = (Sqr((bA ^ 2 * (1 - (uA ^ 2 / aA ^ 2))))) * fA
    iLineA.P2.x = RotX(X1A, Y1A, vA) + cxA
    iLineA.P2.y = RotY(X1A, Y1A, vA) + cyA
    '---------------------EllipseB---------------
    iLineB.P1.x = RotX(aB, 0, vB) + cxB
    iLineB.P1.y = RotY(aB, 0, vB) + cyB
    For rB = 0 To 1 Step 0.01
        If (rB >= 0 And rB < 0.5) Or (rB >= 1 And rB < 1.5) Then fB = 1 Else fB = -1
        If fB > 0 Then uB = (aB - (4 * aB * rB)) Else uB = (-aB + (4 * aB * (rB - 0.5)))
        X1B = uB
        Y1B = (Sqr((bB ^ 2 * (1 - (uB ^ 2 / aB ^ 2))))) * fB
        iLineB.P2.x = RotX(X1B, Y1B, vB) + cxB
        iLineB.P2.y = RotY(X1B, Y1B, vB) + cyB
    '---------------------Intersect Check-----------
        If LineLineIntersect(iLineA, iLineB, iPt()) = 0 Then
            If iPt(0).Layer.Color = vbBlue Then
                ReDim Preserve iPoints(iCount) As CadPoint
                iPoints(iCount) = iPt(0)
                '--------------------Check For Virtual intsersects
                'If Not InsideAngles(EllipsePtAngle(EllipseA, iPt), EllipseA.Angle1, EllipseA.Angle2) Then iPoints(iCount).layer.color = vbRed
                'If iPoints(iCount).layer.color = vbBlue And Not PtInLine(MyLine, iPoints(iCount)) Then iPoints(iCount).layer.color = vbGreen
                iCount = iCount + 1
            End If
        End If
    '--------------------------------------------
        iLineB.P1 = iLineB.P2
    Next rB
    iLineA.P1 = iLineA.P2
Next rA
'------------------------------
ZZ_OLD_EllipseEllipseIntersect = iCount - 1
Exit Function
eTrap:
    MsgBox Err.Number & " - " & Err.Description
    DoEvents
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function EllipseInBox(SelLine As CadLine, MyEllipse As CadEllipse) As Boolean
Dim sPt As CadPoint
sPt = EllipsePt(MyEllipse, 45)
If sPt.Layer.Color = vbBlue And (sPt.x > SelLine.P2.x Or sPt.y < SelLine.P1.y) Then EllipseInBox = False: Exit Function
sPt = EllipsePt(MyEllipse, 135)
If sPt.Layer.Color = vbBlue And (sPt.x < SelLine.P1.x Or sPt.y < SelLine.P1.y) Then EllipseInBox = False: Exit Function
sPt = EllipsePt(MyEllipse, 225)
If sPt.Layer.Color = vbBlue And (sPt.x < SelLine.P1.x Or sPt.y > SelLine.P2.y) Then EllipseInBox = False: Exit Function
sPt = EllipsePt(MyEllipse, 315)
If sPt.Layer.Color = vbBlue And (sPt.x > SelLine.P2.x Or sPt.y > SelLine.P2.y) Then EllipseInBox = False: Exit Function
EllipseInBox = True
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ZZ_nonang_EllipsePt(MyEllipse As CadEllipse, MyAngle As Single) As CadPoint
If MyAngle >= 360 Then MyAngle = MyAngle - 360
Dim TotLen As Single, Flen As Single ' point measurements
Dim a As Single, B As Single ' ellipse parameters
Dim u As Single, v As Single, r As Single, f As Integer ' iiteration parameters
Dim ang1 As Single, ang2 As Single 'adjusted ellipse angles
Dim X1 As Single, Y1 As Single
Dim cx As Single, cy As Single ' center of the ellipse
'-----------------Dimensional Calculations-------
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2)
'--------------------------
ang1 = MyEllipse.Angle1
ang2 = MyEllipse.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
r = MyAngle / 360
'--------------------------
If (r >= 0 And r < 0.5) Or (r >= 1 And r < 1.5) Then f = 1 Else f = -1
If f > 0 Then u = (a - (4 * a * r)) Else u = (-a + (4 * a * (r - 0.5)))
X1 = u
Y1 = (Sqr((B ^ 2 * (1 - (u ^ 2 / a ^ 2))))) * f

'EllipsePt.X = RotX(X1, Y1, v) + cx
'EllipsePt.y = RotY(X1, Y1, v) + cy
'EllipsePt.Color = vbBlue
'EllipsePt.Width = 3
'If Not InsideAngles(MyAngle, Ang1, Ang2) Then EllipsePt.Color = vbRed
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function EllipsePt(MyEllipse As CadEllipse, ByVal MyAngle As Single) As CadPoint
If MyAngle >= 360 Then MyAngle = MyAngle - 360
Dim TotLen As Single, Flen As Single ' point measurements
Dim a As Single, B As Single ' ellipse parameters
Dim J As Single, v As Single ' iiteration parameters
Dim ang1 As Single, ang2 As Single 'adjusted ellipse angles
Dim X1 As Single, Y1 As Single
Dim cx As Single, cy As Single ' center of the ellipse
'-----------------Dimensional Calculations-------
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2) * Pi / 180
'--------------------------
ang1 = MyEllipse.Angle1
ang2 = MyEllipse.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
'--------------------------
J = MyAngle * Pi / 180
X1 = Cos(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
Y1 = Sin(Atn(B * Tan(J) / a) + v) * Sqr((a ^ 2 - B ^ 2) * Cos(J) ^ 2 + B ^ 2) * Sgn(a * Cos(J))
'--------------------------
EllipsePt.x = X1 + cx
EllipsePt.y = Y1 + cy
EllipsePt.Layer.Color = vbBlue
EllipsePt.Layer.Width = 3
If Not InsideAngles(MyAngle, ang1, ang2) Then EllipsePt.Layer.Color = vbRed
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function EllipsePtAngle(MyEllipse As CadEllipse, MyPoint As CadPoint) As Single
Dim tLine As CadLine
Dim tAng As Single
tLine.P1.x = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
tLine.P1.y = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
tLine.P2 = MyPoint
tAng = cAngle(tLine)
EllipsePtAngle = EllipseAngle(tAng, MyEllipse)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub DelaunayArc(MyPoints() As CadPoint, ByRef MyArc As CadArc)
Dim MidAng As Single
DelaunayCenter MyPoints(0), MyPoints(1), MyPoints(2), MyArc.Center
MyArc.Radius = PtLen(MyArc.Center, MyPoints(0))
MyArc.Angle1 = ArcPtAngle(MyArc, MyPoints(0))
MyArc.Angle2 = ArcPtAngle(MyArc, MyPoints(2))
If MyArc.Angle2 < MyArc.Angle1 Then MyArc.Angle2 = MyArc.Angle2 + 360
MidAng = ArcPtAngle(MyArc, MyPoints(1))
If MidAng < MyArc.Angle1 Then MidAng = MidAng + 360
If Not IsBetween(MidAng, MyArc.Angle1, MyArc.Angle2) Then Swap MyArc.Angle1, MyArc.Angle2
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function InsideAngles(MidAng As Single, a1 As Single, a2 As Single) As Boolean
If a2 < a1 Then a2 = a2 + 360
If MidAng < a1 Then MidAng = MidAng + 360
InsideAngles = IsBetween(MidAng, a1, a2)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function LineArcIntersect(MyLine As CadLine, MyArc As CadArc, ByRef iPoints() As CadPoint) As Integer
'LineArcIntersect = LineEllipseIntersect(MyLine, ArcToEllipse(MyArc), iPoints)
'Exit Function
Erase iPoints()
Dim M As Single
Dim B As Single
Dim r As Single
Dim i As Integer
Dim ang1 As Single, ang2 As Single
Dim tLine As CadLine
Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
tLine = MoveLine(MyLine, -MyArc.Center.x, -MyArc.Center.y)
M = CSlope(tLine)
B = (tLine.P1.y) - (M * (tLine.P1.x))
r = MyArc.Radius
X1 = -(Sqr(Abs(r ^ 2 * (M ^ 2 + 1) - B ^ 2)) + B * M) / (M ^ 2 + 1)
Y1 = M * X1 + B
X2 = (Sqr(Abs(r ^ 2 * (M ^ 2 + 1) - B ^ 2)) - B * M) / (M ^ 2 + 1)
Y2 = M * X2 + B
If X1 = X2 And Y1 = Y2 Then
    ReDim iPoints(0) As CadPoint
    iPoints(0).x = X1 + MyArc.Center.x
    iPoints(0).y = Y1 + MyArc.Center.y
    iPoints(0).Layer.Color = vbBlue
    iPoints(0).Layer.Width = 3
    LineArcIntersect = 0
Else
    ReDim iPoints(1) As CadPoint
    iPoints(0).x = X1 + MyArc.Center.x
    iPoints(0).y = Y1 + MyArc.Center.y
    iPoints(0).Layer.Color = vbBlue
    iPoints(0).Layer.Width = 3
    iPoints(1).x = X2 + MyArc.Center.x
    iPoints(1).y = Y2 + MyArc.Center.y
    iPoints(1).Layer.Color = vbBlue
    iPoints(1).Layer.Width = 3
    LineArcIntersect = 1
End If
For i = 0 To UBound(iPoints)
    If PtLen(iPoints(i), MyArc.Center) > r Then
        Erase iPoints()
        LineArcIntersect = -1
        Exit Function
    End If
Next i
'---------------Get angles to test if the intersection points are 'real' or 'virtual'
ang1 = MyArc.Angle1
ang2 = MyArc.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
'------Check for virtual intersections
For i = 0 To UBound(iPoints)
    If Not InsideAngles(PtPtAngle(MyArc.Center, iPoints(i)), ang1, ang2) Then iPoints(i).Layer.Color = vbRed
    If Not PtInLine(MyLine, iPoints(i)) Then
        If iPoints(i).Layer.Color = vbBlue Then
            iPoints(i).Layer.Color = vbYellow
        Else
            iPoints(i).Layer.Color = vbGreen
        End If
    End If
Next i
'LineArcIntersect = PolyLineLineIntersect(ArcToPolyLine(MyArc), MyLine, iPoints)
'X = -(Sqr(r ^ 2 * (m ^ 2 + 1) - b ^ 2) + b * m) / (m ^ 2 + 1)
'X = (Sqr(r ^ 2 * (m ^ 2 + 1) - b ^ 2) - b * m) / (m ^ 2 + 1)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function LineEllipseIntersect(MyLine As CadLine, MyEllipse As CadEllipse, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
'-----------------------------------------------------------------------------------
Dim TotLen As Single, Flen As Single ' point measurements
Dim a As Single, B As Single, v As Single ' ellipse parameters
Dim ang1 As Single, ang2 As Single 'adjusted ellipse angles
Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
Dim cx As Single, cy As Single ' center of the ellipse
Dim M As Single
Dim c As Single
Dim tLine As CadLine
Dim iCount As Integer
Dim tPt As CadPoint
Dim Is90 As Boolean
'---------------Get angles to test if the intersection points are 'real' or 'virtual'
ang1 = MyEllipse.Angle1
ang2 = MyEllipse.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
'-----------------Get our variables for x^2/a^2 +y^2/b^2 = 1-------------------------
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2)
If MyEllipse.F1.x = MyEllipse.F2.x And MyEllipse.F1.y = MyEllipse.F2.y Then v = 0
'-----------First we rotate our intersection line negative the rotation of the ellipse
tLine = MyLine
If v <> 0 Then tLine = RotateLine(MyLine, EllipseCenter(MyEllipse), -v)
'-----------Next we move the line so it is based upon the center 0,0
tLine = MoveLine(tLine, -cx, -cy)
'----------Get our variables for: Y = mx+c
M = CSlope(tLine)
c = (tLine.P1.y) - (M * (tLine.P1.x))
If cAngle(tLine) = 90 Or cAngle(tLine) = 270 Then Is90 = True
If Abs(v - cAngle(MyLine)) = 90 Or Abs(v - cAngle(MyLine)) = 270 Then Is90 = True
'----------substitue (mx+c) into the equation for ellipse (x^2/a^2 +y^2/b^2 = 1) as variable "y" and solve for "x"
If Not Is90 Then
    X1 = -(a * (B * Sqr(Abs(a ^ 2 * M ^ 2 + B ^ 2 - c ^ 2)) + a * c * M)) / (a ^ 2 * M ^ 2 + B ^ 2)
    Y1 = M * X1 + c
Else
    X1 = tLine.P1.x
    Y1 = B * (B * c - a * M * Sqr(Abs(a ^ 2 * M ^ 2 + B ^ 2 - c ^ 2))) / (a ^ 2 * M ^ 2 + B ^ 2)
End If
'-------Create Point - moving it back to the centers and rotating it back ---------
tPt.Layer.Width = 3
tPt.Layer.Color = vbBlue
tPt.x = RotX(X1, Y1, v) + cx
tPt.y = RotY(X1, Y1, v) + cy
'-------------Add Point to set of intersections-------
'DrawCadEllipse frmDraw.picDraw, MyEllipse, , vbGreen, 1
'DrawCadPoint frmDraw.picDraw, tPt, , vbWhite, 3
'if the sum of the distances from the focci to the point are equal to totlen, then the point lies on the ellipse
If Format((PtLen(tPt, MyEllipse.F1) + PtLen(tPt, MyEllipse.F2)), "0.000") = Format(TotLen, "0.000") Then
    ReDim iPoints(iCount) As CadPoint
    iPoints(iCount) = tPt
    '------Check for virtual intersections
    If Not InsideAngles(EllipseAngle(PtPtAngle(EllipseCenter(MyEllipse), tPt), MyEllipse), ang1, ang2) Then iPoints(iCount).Layer.Color = vbRed
    If Not PtInLine(MyLine, tPt) Then
        If iPoints(iCount).Layer.Color = vbBlue Then
            iPoints(iCount).Layer.Color = vbYellow
        Else
            iPoints(iCount).Layer.Color = vbGreen
        End If
    End If
    iCount = iCount + 1
Else
    LineEllipseIntersect = -1
    Exit Function
End If
'----------substitue (mx+c) into the equation for ellipse (x^2/a^2 +y^2/b^2 = 1) as variable "y" and solve for "x"
If Not Is90 Then
    X2 = (a * (B * Sqr(Abs(a ^ 2 * M ^ 2 + B ^ 2 - c ^ 2)) - a * c * M)) / (a ^ 2 * M ^ 2 + B ^ 2)
    Y2 = M * X2 + c
Else
    X2 = tLine.P1.x
    Y2 = -B * (B * c - a * M * Sqr(Abs(a ^ 2 * M ^ 2 + B ^ 2 - c ^ 2))) / (a ^ 2 * M ^ 2 + B ^ 2)
End If
'-------Create Point - moving it back to the centers and rotating it back ---------
tPt.x = RotX(X2, Y2, v) + cx
tPt.y = RotY(X2, Y2, v) + cy
'----------Test For Tangency--------
If X1 <> X2 Or Y1 <> Y2 Then ' point is not tangent - otherwise . . we need return only the first point
    '---------Add Point to intersection collection--------
    ReDim Preserve iPoints(iCount) As CadPoint
    iPoints(iCount) = tPt
    '------Check for virtual intersections
    If Not InsideAngles(EllipseAngle(PtPtAngle(EllipseCenter(MyEllipse), tPt), MyEllipse), ang1, ang2) Then iPoints(iCount).Layer.Color = vbRed
    If Not PtInLine(MyLine, tPt) Then
        If iPoints(iCount).Layer.Color = vbBlue Then
            iPoints(iCount).Layer.Color = vbYellow
        Else
            iPoints(iCount).Layer.Color = vbGreen
        End If
    End If
    iCount = iCount + 1
End If
LineEllipseIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function ZZ_MATH_ArcEllipseIntersect(MyArc As CadArc, MyEllipse As CadEllipse, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
'-----------------------------------------------------------------------------------
Dim TotLen As Single, Flen As Single ' point measurements
Dim a As Single, B As Single, v As Single ' ellipse parameters
Dim ang1 As Single, ang2 As Single 'adjusted ellipse angles
Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
Dim cx As Single, cy As Single ' center of the ellipse
Dim r As Single
Dim tArc As CadArc
Dim iCount As Integer
Dim tPt As CadPoint
Dim Is90 As Boolean
'---------------Get angles to test if the intersection points are 'real' or 'virtual'
ang1 = MyEllipse.Angle1
ang2 = MyEllipse.Angle2
If ang2 < ang1 Then ang2 = ang2 + 360
'-----------------Get our variables for x^2/a^2 +y^2/b^2 = 1-------------------------
Flen = PtLen(MyEllipse.F1, MyEllipse.F2)
TotLen = PtLen(MyEllipse.F1, MyEllipse.P1) + PtLen(MyEllipse.F2, MyEllipse.P1)
a = TotLen / 2
B = Sqr((TotLen / 2) ^ 2 - (Flen / 2) ^ 2)
cx = (MyEllipse.F1.x + MyEllipse.F2.x) / 2
cy = (MyEllipse.F1.y + MyEllipse.F2.y) / 2
v = PtPtAngle(MyEllipse.F1, MyEllipse.F2)
If MyEllipse.F1.x = MyEllipse.F2.x And MyEllipse.F1.y = MyEllipse.F2.y Then v = 0
'-----------First we rotate our intersection line negative the rotation of the ellipse
tArc = RotateArc(MyArc, EllipseCenter(MyEllipse), -v)
'-----------Next we move the line so it is based upon the center 0,0
tArc = MoveArc(tArc, -cx, -cy)
'----------Get our variables for: x^2 + y^2 = r^2
r = tArc.Radius
'----------substitue (sqr(r^2-x^2)) into the equation for ellipse (x^2/a^2 +y^2/b^2 = 1) as variable "y" and solve for "x"
X1 = a * Sqr(Abs(r ^ 2 - B ^ 2)) / Sqr(Abs(a ^ 2 - B ^ 2))
Y1 = Sqr(Abs(r ^ 2 - X1 ^ 2))
'-------Create Point - moving it back to the centers and rotating it back ---------
tPt.Layer.Width = 3
tPt.Layer.Color = vbBlue
tPt.x = RotX(X1, Y1, v) + cx
tPt.y = RotY(X1, Y1, v) + cy
'-------------Add Point to set of intersections-------
'DrawCadEllipse frmDraw.picDraw, MyEllipse, , vbGreen, 1
'DrawCadPoint frmDraw.picDraw, tPt, , vbWhite, 3
'if the sum of the distances from the focci to the point are equal to totlen, then the point lies on the ellipse
If Format((PtLen(tPt, MyEllipse.F1) + PtLen(tPt, MyEllipse.F2)), "0.000") = Format(TotLen, "0.000") Then
    ReDim iPoints(iCount) As CadPoint
    iPoints(iCount) = tPt
    '------Check for virtual intersections
    If Not InsideAngles(EllipseAngle(PtPtAngle(EllipseCenter(MyEllipse), tPt), MyEllipse), ang1, ang2) Then iPoints(iCount).Layer.Color = vbRed
    
    iCount = iCount + 1
Else
    ZZ_MATH_ArcEllipseIntersect = -1
    Exit Function
End If
'----------substitue (sqr(r^2-x^2)) into the equation for ellipse (x^2/a^2 +y^2/b^2 = 1) as variable "y" and solve for "x"
X2 = -a * Sqr(Abs(r ^ 2 - B ^ 2)) / Sqr(Abs(a ^ 2 - B ^ 2))
Y2 = Sqr(Abs(r ^ 2 - X1 ^ 2))
'-------Create Point - moving it back to the centers and rotating it back ---------
tPt.x = RotX(X2, Y2, v) + cx
tPt.y = RotY(X2, Y2, v) + cy
'----------Test For Tangency--------
If X1 <> X2 Or Y1 <> Y2 Then ' point is not tangent - otherwise . . we need return only the first point
    '---------Add Point to intersection collection--------
    ReDim Preserve iPoints(iCount) As CadPoint
    iPoints(iCount) = tPt
    '------Check for virtual intersections
    If Not InsideAngles(EllipseAngle(PtPtAngle(EllipseCenter(MyEllipse), tPt), MyEllipse), ang1, ang2) Then iPoints(iCount).Layer.Color = vbRed
    
    iCount = iCount + 1
End If
ZZ_MATH_ArcEllipseIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function LineLineIntersect(Line1 As CadLine, Line2 As CadLine, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
'Calculate the intersection point of any two given non-parallel lines.
'
'Returns:  -1 = lines are parallel (no intersection).
'           0 = Neither line contains the intersect point between its points.**
'           1 = Line1 contains the intersect point between its points.**
'           2 = Line2 contains the intersect point between its points.**
'           3 = Both Lines contain the intersect point between their points.**
'           ** Lines Do intersect; Also fills in the ptIntersect point.
'
'BTW:       There are 18 lines of pure code, 25 lines of pure comments and 6
'           mixed lines in this function, just in case you were wondering. (:oþ}

Dim bIntersect  As Boolean
Dim iReturn     As Integer
Dim dDenom      As Double
Dim dPctDelta1  As Double
Dim dPctDelta2  As Double
Dim Delta(2)    As CadPoint

        'Calculate the Deltas (distance of X2 - X1 or Y2 - Y1 of any 2 points)
        Delta(0).x = Line1.P1.x - Line2.P1.x   'Line1-Line2.p1 X-Cross-Delta
        Delta(0).y = Line1.P1.y - Line2.P1.y   'Line1-Line2.p1 Y-Cross-Delta
        Delta(1).x = Line1.P2.x - Line1.P1.x   'Line1 X-Delta
        Delta(1).y = Line1.P2.y - Line1.P1.y   'Line1 Y-Delta
        Delta(2).x = Line2.P2.x - Line2.P1.x   'Line2 X-Delta
        Delta(2).y = Line2.P2.y - Line2.P1.y   'Line2 Y-Delta
        
        'Calculate the denominator (zero = parallel (no intersection))
        'Formula: (L2Dy * L1Dx) - (L2Dx * L1Dy)
        iReturn = -1
        dDenom = (Delta(2).y * Delta(1).x) - (Delta(2).x * Delta(1).y)
        bIntersect = (dDenom <> 0)
        
        If bIntersect Then
            ReDim iPoints(0) As CadPoint
            'The lines will intersect somewhere.
            'Solve for both lines using the Cross-Deltas (Delta(0))
            
            'This yields percentage (0.1 = 10%; 1 = 100%) of the distance
            'between ptStart and ptEnd, of the opposite line, where the line used
            'in the calculation will cross it.
            '0 = ptStart direct hit; 1 = ptEnd direct hit; 0.5 = Centered between Pts; etc.
            'If < 0 or > 1 then the lines still intersect, just not between the points.
            
            'Solve for Line1 where Line2 will cross it.
            dPctDelta1 = ((Delta(2).x * Delta(0).y) - (Delta(2).y * Delta(0).x)) / dDenom
            
            'Solve for Line2 where Line1 will cross it.
            dPctDelta2 = ((Delta(1).x * Delta(0).y) - (Delta(1).y * Delta(0).x)) / dDenom
        
            'Check for absolute intersection. If the percentage is not between
            '0 and 1 then the lines will not intersect between their points.
            'Returns 0, 1, 2 or 3.
            iReturn = IIf(IsBetween(dPctDelta1, 0#, 1#), 1, 0) _
                Or IIf(IsBetween(dPctDelta2, 0#, 1#), 2, 0)
            
            'Calculate point of intersection on Line1 and fill ptIntersect.
            iPoints(0).x = Line1.P1.x + (dPctDelta1 * Delta(1).x)
            iPoints(0).y = Line1.P1.y + (dPctDelta1 * Delta(1).y)
        
        End If
        
        'Return the results.
        Select Case iReturn
            Case -1
                LineLineIntersect = -1
            Case 0
                iPoints(0).Layer.Color = vbYellow
                iPoints(0).Layer.Width = 3
                LineLineIntersect = 0
            Case 1
                iPoints(0).Layer.Color = vbRed
                iPoints(0).Layer.Width = 3
                LineLineIntersect = 0
            Case 2
                iPoints(0).Layer.Color = vbGreen
                iPoints(0).Layer.Width = 3
                LineLineIntersect = 0
            Case 3
                iPoints(0).Layer.Color = vbBlue
                iPoints(0).Layer.Width = 3
                LineLineIntersect = 0
        End Select
        
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function TriPtAngle(P1 As CadPoint, P2 As CadPoint, P3 As CadPoint) As Single
Dim L1 As CadLine
Dim L2 As CadLine
L1.P1 = P2
L1.P2 = P1
L2.P1 = P2
L2.P2 = P3
If P1.x = P2.x And P1.y = P2.y Then
    TriPtAngle = cAngle(L2)
ElseIf P2.x = P3.x And P2.y = P3.y Then
    TriPtAngle = cAngle(L1)
ElseIf P3.x = P1.x And P3.y = P1.y Then
    TriPtAngle = cAngle(L2)
Else
    TriPtAngle = cAngle(L2) - cAngle(L1)
End If
TriPtAngle = dAngle(TriPtAngle)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Public Function IsBetween(ByVal vTestData As Variant, ByVal vLowerBound As Variant, ByVal vUpperBound As Variant, Optional ByVal bInclusive As Boolean = True) As Boolean

'Returns True if vTestData is between vLowerBound and vUpperBound.
'bInclusive = Are the bounds included in the test?

Dim vTemp   As Variant

    If vLowerBound = vUpperBound Then
        Exit Function   'Returns false if upper and lower bounds are equal.
    Else
        If vLowerBound > vUpperBound Then
            'If bounds are reversed, swap them.
            vTemp = vLowerBound
            vLowerBound = vUpperBound
            vUpperBound = vTemp
        End If
        If bInclusive Then
            'If bounds are included in test (use >= and <=).
            IsBetween = (vTestData >= vLowerBound) And (vTestData <= vUpperBound)
        Else
            'If bounds are not included in test (use > and <).
            IsBetween = (vTestData > vLowerBound) And (vTestData < vUpperBound)
        End If
    End If
    
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub SelectFromBox(SelLine As CadLine, MyGeo() As Geometry, GNum As Integer, ByRef SelArray() As SelSet)
On Local Error GoTo eTrap
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim GotPoint As Boolean
Dim Count As Integer
Dim gCount As Integer
Dim tLine As CadLine
Dim iGeo() As Geometry
Dim iSel() As SelSet ' insert selection
If SelLine.P1.x > SelLine.P2.x Then Swap SelLine.P1.x, SelLine.P2.x
If SelLine.P1.y > SelLine.P2.y Then Swap SelLine.P1.y, SelLine.P2.y
Erase SelArray()
gCount = UBound(MyGeo(GNum).Points)
For i = 0 To gCount
    If PtInBox(SelLine, MyGeo(GNum).Points(i)) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Point"
        SelArray(Count).Index = i
        Count = Count + 1
    End If
Next i
gCount = UBound(MyGeo(GNum).Lines)
For i = 0 To gCount
    If LineInBox(SelLine, MyGeo(GNum).Lines(i)) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Line"
        SelArray(Count).Index = i
        Count = Count + 1
    End If
Next i
gCount = UBound(MyGeo(GNum).Arcs)
For i = 0 To gCount
    If PtInBox(SelLine, MyGeo(GNum).Arcs(i).Center) And (MyGeo(GNum).Arcs(i).Radius * 2 < Abs(SelLine.P2.x - SelLine.P1.x) And MyGeo(GNum).Arcs(i).Radius * 2 < Abs(SelLine.P2.y - SelLine.P1.y)) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Arc"
        SelArray(Count).Index = i
        Count = Count + 1
    ElseIf PtInBox(SelLine, ArcPt(MyGeo(GNum).Arcs(i), MyGeo(GNum).Arcs(i).Angle1)) Or PtInBox(SelLine, ArcPt(MyGeo(GNum).Arcs(i), MyGeo(GNum).Arcs(i).Angle2)) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Arc"
        SelArray(Count).Index = i
        Count = Count + 1
    ElseIf BoxArcIntersect(SelLine, MyGeo(GNum).Arcs(i)) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Arc"
        SelArray(Count).Index = i
        Count = Count + 1
    End If
Next i
gCount = UBound(MyGeo(GNum).Ellipses)
For i = 0 To gCount
    If PtInBox(SelLine, EllipseCenter(MyGeo(GNum).Ellipses(i))) And EllipseInBox(SelLine, MyGeo(GNum).Ellipses(i)) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Ellipse"
        SelArray(Count).Index = i
        Count = Count + 1
    ElseIf PtInBox(SelLine, EllipsePt(MyGeo(GNum).Ellipses(i), MyGeo(GNum).Ellipses(i).Angle1)) Or PtInBox(SelLine, EllipsePt(MyGeo(GNum).Ellipses(i), MyGeo(GNum).Ellipses(i).Angle2)) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Ellipse"
        SelArray(Count).Index = i
        Count = Count + 1
    ElseIf BoxEllipseIntersect(SelLine, MyGeo(GNum).Ellipses(i)) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Ellipse"
        SelArray(Count).Index = i
        Count = Count + 1
    End If
Next i
gCount = UBound(MyGeo(GNum).Splines)
For i = 0 To gCount
    If SplineInBox(MyGeo(GNum).Splines(i), SelLine) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Spline"
        SelArray(Count).Index = i
        Count = Count + 1
    End If
Next i
gCount = UBound(MyGeo(GNum).PolyLines)
For i = 0 To gCount
    'GotPoint = False
    For J = 0 To UBound(MyGeo(GNum).PolyLines(i).Vertex) - 1
        'If PtInBox(SelLine, MyGeo(GNum).PolyLines(i).Vertex(j)) Then
        If LineInBox(SelLine, VLine(MyGeo(GNum).PolyLines(i).Vertex(J).x, MyGeo(GNum).PolyLines(i).Vertex(J).y, MyGeo(GNum).PolyLines(i).Vertex(J + 1).x, MyGeo(GNum).PolyLines(i).Vertex(J + 1).y)) Then
            ReDim Preserve SelArray(Count) As SelSet
            SelArray(Count).Type = "PolyLine"
            SelArray(Count).Index = i
            Count = Count + 1
            GotPoint = True
            Exit For
        End If
    Next J
    'If Not GotPoint And BoxPolyLineIntersect(SelLine, MyGeo(GNum).PolyLines(i)) Then
    '
    '    ReDim Preserve SelArray(Count) As SelSet
    '    SelArray(Count).Type = "PolyLine"
    '    SelArray(Count).Index = i
    '    Count = Count + 1
    'End If
Next i
gCount = UBound(MyGeo(GNum).Text)
For i = 0 To gCount
    tLine = cAngLine(MyGeo(GNum).Text(i).Angle, MyGeo(GNum).Text(i).Start, MyGeo(GNum).Text(i).Length)
    If LineInBox(SelLine, tLine) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Text"
        SelArray(Count).Index = i
        Count = Count + 1
    End If
Next i
gCount = UBound(MyGeo(GNum).Faces)
For i = 0 To gCount
    If LineInBox(SelLine, VLine(MyGeo(GNum).Faces(i).Vertex(0).x, MyGeo(GNum).Faces(i).Vertex(0).y, MyGeo(GNum).Faces(i).Vertex(1).x, MyGeo(GNum).Faces(i).Vertex(1).y)) Or _
        LineInBox(SelLine, VLine(MyGeo(GNum).Faces(i).Vertex(1).x, MyGeo(GNum).Faces(i).Vertex(1).y, MyGeo(GNum).Faces(i).Vertex(2).x, MyGeo(GNum).Faces(i).Vertex(2).y)) Or _
        LineInBox(SelLine, VLine(MyGeo(GNum).Faces(i).Vertex(2).x, MyGeo(GNum).Faces(i).Vertex(2).y, MyGeo(GNum).Faces(i).Vertex(0).x, MyGeo(GNum).Faces(i).Vertex(0).y)) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Face"
        SelArray(Count).Index = i
        Count = Count + 1
    End If
Next i
gCount = UBound(MyGeo(GNum).Inserts)
For i = 0 To gCount
    CreateInsertGeo MyGeo(), MyGeo(GNum).Inserts(i), iGeo()
    SelectFromBox SelLine, iGeo(), 0, iSel()
    If isSelected(iSel()) Then
        ReDim Preserve SelArray(Count) As SelSet
        SelArray(Count).Type = "Insert"
        SelArray(Count).Index = i
        Count = Count + 1
    End If
Next i
'------------ERROR HANDLER------------\
Exit Sub
eTrap:
    If Err.Number = 9 Then
        gCount = -1
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub


Sub LoadGeo(zLine As CadLine, MyGeo() As Geometry, FName As String)
On Error GoTo eTrap
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim FF As Integer
Dim gType As String
Dim g As Integer
Dim gCount As Integer
Dim Name As String
FF = FreeFile
ClearGeo MyGeo
Open FName For Input As #FF
Do While Not EOF(FF)
    Input #FF, gType
    Select Case gType
        Case "Zoom"
            Input #FF, zLine.P1.x, zLine.P1.y, zLine.P2.x, zLine.P2.y
        Case "GEO"
            Input #FF, Name
            Input #FF, gCount
            ReDim Preserve MyGeo(gCount) As Geometry
            MyGeo(gCount).Name = Name
        Case "Point"
            g = UBound(MyGeo(gCount).Points) + 1
            ReDim Preserve MyGeo(gCount).Points(g)
            With MyGeo(gCount).Points(g)
                Input #FF, .x, .y, .Layer.Color, .Layer.style, .Layer.Width
            End With
        Case "Line"
            g = UBound(MyGeo(gCount).Lines) + 1
            ReDim Preserve MyGeo(gCount).Lines(g)
            With MyGeo(gCount).Lines(g)
                Input #FF, .P1.x, .P1.y, .P2.x, .P2.y, .Layer.Color, .Layer.style, .Layer.Width
            End With
        Case "Arc"
            g = UBound(MyGeo(gCount).Arcs) + 1
            ReDim Preserve MyGeo(gCount).Arcs(g)
            With MyGeo(gCount).Arcs(g)
                Input #FF, .Center.x, .Center.y, .Radius, .Angle1, .Angle2, .Layer.Color, .Layer.style, .Layer.Width
            End With
        Case "Ellipse"
            g = UBound(MyGeo(gCount).Ellipses) + 1
            ReDim Preserve MyGeo(gCount).Ellipses(g)
            With MyGeo(gCount).Ellipses(g)
                Input #FF, .F1.x, .F1.y, .F2.x, .F2.y, .P1.x, .P1.y, .Angle1, .Angle2, .NumPoints, .Layer.Color, .Layer.style, .Layer.Width
            End With
        Case "Spline"
            g = UBound(MyGeo(gCount).Splines) + 1
            ReDim Preserve MyGeo(gCount).Splines(g)
            With MyGeo(gCount).Splines(g)
                Input #FF, J, .Layer.Color, .Layer.style, .Layer.Width
                ReDim .Vertex(J)
                For K = 0 To J
                    With .Vertex(K)
                        Input #FF, .x, .y
                    End With
                Next K
            End With
        Case "PolyLine"
            g = UBound(MyGeo(gCount).PolyLines) + 1
            ReDim Preserve MyGeo(gCount).PolyLines(g)
            With MyGeo(gCount).PolyLines(g)
                Input #FF, J, .Layer.Color, .Layer.style, .Layer.Width
                ReDim .Vertex(J)
                For K = 0 To J
                    With .Vertex(K)
                        Input #FF, .x, .y
                    End With
                Next K
            End With
        Case "Face"
            g = UBound(MyGeo(gCount).Faces) + 1
            ReDim Preserve MyGeo(gCount).Faces(g)
            With MyGeo(gCount).Faces(g)
                Input #FF, .Layer.Color, .Layer.style, .Layer.Width
                For K = 0 To 2
                    With .Vertex(K)
                        Input #FF, .x, .y
                    End With
                Next K
            End With
        Case "Text"
            g = UBound(MyGeo(gCount).Text) + 1
            ReDim Preserve MyGeo(gCount).Text(g)
            With MyGeo(gCount).Text(g)
                Input #FF, .Start.x, .Start.y, .Angle, .Length, .Size, .Layer.FontName, .Text, .Layer.Color
            End With
        Case "Insert"
            g = UBound(MyGeo(gCount).Inserts) + 1
            ReDim Preserve MyGeo(gCount).Inserts(g)
            With MyGeo(gCount).Inserts(g)
                Input #FF, .Name, .Base.x, .Base.y, .ScaleX, .ScaleY, .Angle
            End With
    End Select
Loop
Close #FF
Exit Sub
eTrap:
    If Err.Number = 53 Then Exit Sub
    g = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function MidPoint(L1 As CadLine) As CadPoint
MidPoint.x = (L1.P1.x + L1.P2.x) / 2
MidPoint.y = (L1.P1.y + L1.P2.y) / 2
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function dAngle(Angle As Single) As Single
If Angle > 360 Then
    dAngle = Angle - 360
ElseIf Angle < 0 Then
    dAngle = Angle + 360
Else
    dAngle = Angle
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function cHyp(X1 As Single, Y1 As Single) As Single
cHyp = Sqr((X1 * X1) + (Y1 * Y1))
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function MoveArc(MyArc As CadArc, dX As Single, dY As Single) As CadArc
MoveArc = MyArc
MoveArc.Center = MovePoint(MyArc.Center, dX, dY)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function MoveText(MyText As CadText, dX As Single, dY As Single) As CadText
MoveText = MyText
MoveText.Start = MovePoint(MyText.Start, dX, dY)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function MoveEllipse(MyEllipse As CadEllipse, dX As Single, dY As Single) As CadEllipse
MoveEllipse = MyEllipse
MoveEllipse.F1 = MovePoint(MyEllipse.F1, dX, dY)
MoveEllipse.F2 = MovePoint(MyEllipse.F2, dX, dY)
MoveEllipse.P1 = MovePoint(MyEllipse.P1, dX, dY)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function MoveSpline(MySPline As CadSpline, dX As Single, dY As Single) As CadSpline
MoveSpline = MySPline
Dim i As Integer
For i = 0 To UBound(MoveSpline.Vertex)
    MoveSpline.Vertex(i) = MovePoint(MySPline.Vertex(i), dX, dY)
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Function MovePolyLine(MyPolyLine As CadPolyLine, dX As Single, dY As Single) As CadPolyLine
MovePolyLine = MyPolyLine
Dim i As Integer
For i = 0 To UBound(MovePolyLine.Vertex)
    MovePolyLine.Vertex(i) = MovePoint(MyPolyLine.Vertex(i), dX, dY)
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function MoveLine(MyLine As CadLine, dX As Single, dY As Single) As CadLine
MoveLine = MyLine
MoveLine.P1 = MovePoint(MyLine.P1, dX, dY)
MoveLine.P2 = MovePoint(MyLine.P2, dX, dY)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function MovePoint(MyPoint As CadPoint, dX As Single, dY As Single) As CadPoint
MovePoint = MyPoint
MovePoint.x = MyPoint.x + dX
MovePoint.y = MyPoint.y + dY
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function MoveFace(MyFace As CadFace, dX As Single, dY As Single) As CadFace
MoveFace = MyFace
MoveFace.Vertex(0) = MovePoint(MyFace.Vertex(0), dX, dY)
MoveFace.Vertex(1) = MovePoint(MyFace.Vertex(1), dX, dY)
MoveFace.Vertex(2) = MovePoint(MyFace.Vertex(2), dX, dY)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function PerpLineCenter(MyLine As CadLine) As CadLine
Dim DeltaX As Single
Dim DeltaY As Single
PerpLineCenter.P1.x = (MyLine.P1.x + MyLine.P2.x) / 2
PerpLineCenter.P1.y = (MyLine.P1.y + MyLine.P2.y) / 2
DeltaX = PerpLineCenter.P1.x - MyLine.P1.x
DeltaY = PerpLineCenter.P1.y - MyLine.P1.y
PerpLineCenter.P2.x = PerpLineCenter.P1.x + -DeltaY
PerpLineCenter.P2.y = PerpLineCenter.P1.y + DeltaX
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function PtAng(X1 As Single, Y1 As Single) As Single
If X1 = 0 Then
    If Y1 >= 0 Then
        PtAng = 90
    Else
        PtAng = 270
    End If
    Exit Function
ElseIf Y1 = 0 Then
    If X1 >= 0 Then
        PtAng = 0
    Else
        PtAng = 180
    End If
    Exit Function
Else
    PtAng = Atn(Y1 / X1)
    PtAng = PtAng * 180 / Pi
    If PtAng < 0 Then PtAng = PtAng + 360
    If PtAng > 360 Then PtAng = PtAng - 360
    '----------Test for direction-(quadrant check)-------
    If X1 < 0 Then PtAng = PtAng + 180
    If Y1 < 0 And PtAng < 90 Then PtAng = PtAng + 180
    'If X1 < 0 And PtAng <> 180 Then PtAng = PtAng + 180
    'If Y1 < 0 And PtAng = 90 Then PtAng = PtAng + 180
    
    'One final check
    If PtAng < 0 Then PtAng = PtAng + 360
    If PtAng > 360 Then PtAng = PtAng - 360
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function cAngle(MyLine As CadLine) As Single
If MyLine.P1.x = MyLine.P2.x Then
    If MyLine.P1.y < MyLine.P2.y Then
        cAngle = 90
    Else
        cAngle = 270
    End If
    Exit Function
ElseIf MyLine.P1.y = MyLine.P2.y Then
    If MyLine.P1.x < MyLine.P2.x Then
        cAngle = 0
    Else
        cAngle = 180
    End If
    Exit Function
Else
    cAngle = Atn(CSlope(MyLine))
    cAngle = cAngle * 180 / Pi
    If cAngle < 0 Then cAngle = cAngle + 360
    '----------Test for direction--------
    If MyLine.P1.x > MyLine.P2.x And cAngle <> 180 Then cAngle = cAngle + 180
    If MyLine.P1.y > MyLine.P2.y And cAngle = 90 Then cAngle = cAngle + 180
    If cAngle > 360 Then cAngle = cAngle - 360
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function TrueAngle(MyLine As CadLine) As Single
If MyLine.P1.x = MyLine.P2.x Then
    If MyLine.P1.y < MyLine.P2.y Then
        TrueAngle = 90
    Else
        TrueAngle = 270
    End If
    Exit Function
ElseIf MyLine.P1.y = MyLine.P2.y Then
    If MyLine.P1.x < MyLine.P2.x Then
        TrueAngle = 0
    Else
        TrueAngle = 180
    End If
    Exit Function
Else
    TrueAngle = Atn(CSlope(MyLine))
    TrueAngle = TrueAngle * 180 / Pi
    If TrueAngle < 0 Then TrueAngle = TrueAngle + 360
    '----------Test for direction--------
    If MyLine.P1.x > MyLine.P2.x And TrueAngle <> 180 Then TrueAngle = TrueAngle + 180
    If MyLine.P1.y > MyLine.P2.y And TrueAngle = 90 Then TrueAngle = TrueAngle + 180
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function PtPtAngle(P1 As CadPoint, P2 As CadPoint) As Single
Dim MyLine As CadLine
MyLine.P1 = P1
MyLine.P2 = P2
PtPtAngle = cAngle(MyLine)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function CSlope(MyLine As CadLine) As Single
'if the line is VERTICAL we need to tweak the line so that the the slope is not UNDEFINED
If MyLine.P1.x = MyLine.P2.x Then
    If MyLine.P1.y < MyLine.P2.y Then
        CSlope = 32000000000000#
    Else
        CSlope = -32000000000000#
    End If
Else
    CSlope = (MyLine.P2.y - MyLine.P1.y) / (MyLine.P2.x - MyLine.P1.x)
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function CPtSlope(PtA As CadPoint, PtB As CadPoint) As Single
'if the line is VERTICAL we need to tweak the line so that the the slope is not UNDEFINED
If PtA.x = PtB.x Then
    If PtA.y < PtB.y Then
        CPtSlope = 32000000000000#
    Else
        CPtSlope = -32000000000000#
    End If
Else
    CPtSlope = (PtB.y - PtA.y) / (PtB.x - PtA.x)
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Sub DefineLine(sX As Single, sY As Single, ByRef x As Single, ByRef y As Single, Angle As Single, Length As Single, DoAngle As Boolean, DoLength As Boolean, RestrictX As Boolean, RestrictY As Boolean, CurX As Single, CurY As Single)
Dim h As Single
Dim a As Single
If Not DoAngle And Not DoLength Then
    If RestrictX Then y = CurY: Exit Sub
    If RestrictY Then x = CurX: Exit Sub
    Exit Sub
End If
If RestrictX Then y = CurY
If RestrictY Then x = CurX
If DoAngle Then
    a = Angle
    If x < sX And y > sY Then a = 180 - a
    If x < sX And y < sY Then a = 180 + a
    If x > sX And y < sY Then a = 360 - a
Else
    Dim tLine As CadLine
    tLine.P1.x = sX
    tLine.P1.y = sY
    tLine.P2.x = x
    tLine.P2.y = y
    a = cAngle(tLine)
End If
If DoLength Then
    h = Length
Else
    h = Sqr((y - sY) ^ 2 + (x - sX) ^ 2)
End If

If Not RestrictY Then x = h * Cos(a * Pi / 180) + sX
If Not RestrictX Then y = h * Sin(a * Pi / 180) + sY
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function LineLen(MyLine As CadLine) As Single
LineLen = Sqr((MyLine.P2.y - MyLine.P1.y) ^ 2 + (MyLine.P2.x - MyLine.P1.x) ^ 2)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function PtInArc(MyArc As CadArc, MyPoint As CadPoint) As Boolean
Dim a1 As Single, a2 As Single, MyAngle As Single
a1 = MyArc.Angle1
a2 = MyArc.Angle2
If a2 < a1 Then a2 = a2 + (360)
MyAngle = ArcPtAngle(MyArc, MyPoint)
PtInArc = True
If MyAngle < a1 Then
    If a2 > a1 Then
        PtInArc = False
        Exit Function
    ElseIf MyAngle > a2 Then
        PtInArc = False
        Exit Function
    End If
ElseIf MyAngle > a1 And MyAngle > a2 And a2 > a1 Then
    PtInArc = False
    Exit Function
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function PtLen(PtA As CadPoint, PtB As CadPoint) As Single
PtLen = Sqr(Abs(PtB.y - PtA.y) ^ 2 + Abs(PtB.x - PtA.x) ^ 2)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function PtInBox(SelLine As CadLine, MyPoint As CadPoint) As Boolean
If MyPoint.x >= SelLine.P1.x And MyPoint.x <= SelLine.P2.x And MyPoint.y >= SelLine.P1.y And MyPoint.y <= SelLine.P2.y Then PtInBox = True
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Function PtInLine(MyLine As CadLine, MyPoint As CadPoint) As Boolean
Dim a As Single, B As Single, c As Single, D As Single
Dim iPts() As CadPoint
Dim iRes As Integer
If MyLine.P1.x = MyPoint.x And MyLine.P1.y = MyPoint.y Then PtInLine = True: Exit Function
If MyLine.P2.x = MyPoint.x And MyLine.P2.y = MyPoint.y Then PtInLine = True: Exit Function
iRes = LineLineIntersect(MyLine, cAngLine(cAngle(MyLine) + 90, MyPoint, 0.0005, True), iPts())
If iRes <> -1 Then If iPts(0).Layer.Color <> vbBlue Then Exit Function
a = Sqr((MyLine.P2.y - MyLine.P1.y) ^ 2 + (MyLine.P2.x - MyLine.P1.x) ^ 2)
B = Sqr((MyLine.P2.y - MyPoint.y) ^ 2 + (MyLine.P2.x - MyPoint.x) ^ 2)
c = Sqr((MyPoint.y - MyLine.P1.y) ^ 2 + (MyPoint.x - MyLine.P1.x) ^ 2)
D = IIf(B > c, B, c)
If D < a Then PtInLine = True
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function QuadrantFix(ByVal MyAngle As Single) As Single
MyAngle = dAngle(MyAngle)
Dim MySin As Single
MySin = Format(Sin(MyAngle * Pi / 180), "0.000")
If MySin = 0 Then QuadrantFix = -1: Exit Function
If MySin = 1 Then QuadrantFix = -2: Exit Function
If MyAngle > 0 And MyAngle < 90 Then QuadrantFix = 0: Exit Function
If MyAngle > 90 And MyAngle < 180 Then QuadrantFix = 90: Exit Function
If MyAngle > 180 And MyAngle < 270 Then QuadrantFix = 180: Exit Function
If MyAngle > 270 And MyAngle < 360 Then QuadrantFix = 270: Exit Function
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub RelimitLine(ByRef MyGeo As Geometry, SelNum As Integer, EndNum As Integer, MovePoint As CadPoint)
Select Case EndNum
    Case 0: MyGeo.Lines(SelNum).P1 = MovePoint
    Case 1: MyGeo.Lines(SelNum).P2 = MovePoint
End Select
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub RelimitPoint(ByRef MyGeo As Geometry, SelNum As Integer, EndNum As Integer, MovePoint As CadPoint)
MyGeo.Points(SelNum).x = MovePoint.x
MyGeo.Points(SelNum).y = MovePoint.y
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub RelimitArc(ByRef MyGeo As Geometry, SelNum As Integer, EndNum As Integer, MovePoint As CadPoint)
Select Case EndNum
    Case 0: MyGeo.Arcs(SelNum).Angle1 = ArcPtAngle(MyGeo.Arcs(SelNum), MovePoint)
    Case 1: MyGeo.Arcs(SelNum).Angle2 = ArcPtAngle(MyGeo.Arcs(SelNum), MovePoint)
End Select
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub RelimitEllipse(ByRef MyGeo As Geometry, SelNum As Integer, EndNum As Integer, MovePoint As CadPoint)
Select Case EndNum
    Case 0: MyGeo.Ellipses(SelNum).Angle1 = EllipsePtAngle(MyGeo.Ellipses(SelNum), MovePoint)
    Case 1: MyGeo.Ellipses(SelNum).Angle2 = EllipsePtAngle(MyGeo.Ellipses(SelNum), MovePoint)
End Select
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub RelimitSpline(ByRef MyGeo As Geometry, SelNum As Integer, EndNum As Integer, MovePoint As CadPoint)
On Error GoTo eTrap
Dim tPolyLine As CadPolyLine
Dim K As Integer
tPolyLine = SplineToPolyLine(MyGeo.Splines(SelNum))
RemoveGeo MyGeo, "Spline", SelNum
K = UBound(MyGeo.PolyLines) + 1
ReDim Preserve MyGeo.PolyLines(K) As CadPolyLine
MyGeo.PolyLines(K) = tPolyLine
RelimitPolyLine MyGeo, K, EndNum, MovePoint
Exit Sub
eTrap:
    K = 0
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub RelimitPolyLine(ByRef MyGeo As Geometry, SelNum As Integer, EndNum As Integer, MovePoint As CadPoint)
Dim iPts() As CadPoint
Dim i As Integer
Dim isTrim As Boolean
Dim tLine As CadLine
For i = 1 To UBound(MyGeo.PolyLines(SelNum).Vertex)
    tLine.P1 = MyGeo.PolyLines(SelNum).Vertex(i - 1)
    tLine.P2 = MyGeo.PolyLines(SelNum).Vertex(i)
    If PtInLine(tLine, MovePoint) Then isTrim = True: Exit For
Next i
If isTrim Then
    'If the Relimit is a TRIM then we have to break it an remove one of the "ends" of the broken polyline
    BreakPolyLine MyGeo, SelNum, MovePoint
    Select Case EndNum
        Case 0: RemoveGeo MyGeo, "PolyLine", SelNum ': SelNum = UBound(MyGeo.PolyLines)
        Case 1: RemoveGeo MyGeo, "PolyLine", UBound(MyGeo.PolyLines)
    End Select
Else
    'If the Relimit is an EXTEND then we either add a point to the end or the beginning
    Select Case EndNum
        Case 0: MyGeo.PolyLines(SelNum).Vertex(0) = MovePoint
        Case 1: MyGeo.PolyLines(SelNum).Vertex(UBound(MyGeo.PolyLines(SelNum).Vertex)) = MovePoint
    End Select
End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function RotatePoint(MyPoint As CadPoint, Pivot As CadPoint, Angle As Single) As CadPoint
RotatePoint = MyPoint
RotatePoint.x = RotX(MyPoint.x - Pivot.x, MyPoint.y - Pivot.y, Angle) + Pivot.x
RotatePoint.y = RotY(MyPoint.x - Pivot.x, MyPoint.y - Pivot.y, Angle) + Pivot.y
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function RotateLine(MyLine As CadLine, Pivot As CadPoint, Angle As Single) As CadLine
RotateLine = MyLine
RotateLine.P1 = RotatePoint(MyLine.P1, Pivot, Angle)
RotateLine.P2 = RotatePoint(MyLine.P2, Pivot, Angle)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function RotateFace(MyFace As CadFace, Pivot As CadPoint, Angle As Single) As CadFace
RotateFace = MyFace
RotateFace.Vertex(0) = RotatePoint(MyFace.Vertex(0), Pivot, Angle)
RotateFace.Vertex(1) = RotatePoint(MyFace.Vertex(1), Pivot, Angle)
RotateFace.Vertex(2) = RotatePoint(MyFace.Vertex(2), Pivot, Angle)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function RotateText(MyText As CadText, Pivot As CadPoint, Angle As Single) As CadText
RotateText = MyText
RotateText.Start = RotatePoint(MyText.Start, Pivot, Angle)
RotateText.Angle = MyText.Angle + Angle
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function RotateSpline(MySPline As CadSpline, Pivot As CadPoint, Angle As Single) As CadSpline
Dim i As Integer
RotateSpline = MySPline
For i = 0 To UBound(MySPline.Vertex)
    RotateSpline.Vertex(i) = RotatePoint(MySPline.Vertex(i), Pivot, Angle)
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function RotatePolyLine(MyPolyLine As CadPolyLine, Pivot As CadPoint, Angle As Single) As CadPolyLine
Dim i As Integer
RotatePolyLine = MyPolyLine
For i = 0 To UBound(MyPolyLine.Vertex)
    RotatePolyLine.Vertex(i) = RotatePoint(MyPolyLine.Vertex(i), Pivot, Angle)
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function RotateArc(MyArc As CadArc, Pivot As CadPoint, Angle As Single) As CadArc
RotateArc = MyArc
RotateArc.Center = RotatePoint(MyArc.Center, Pivot, Angle)
RotateArc.Angle1 = dAngle(RotateArc.Angle1 + Angle)
RotateArc.Angle2 = dAngle(RotateArc.Angle2 + Angle)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function RotateEllipse(MyEllipse As CadEllipse, Pivot As CadPoint, Angle As Single) As CadEllipse
RotateEllipse = MyEllipse
RotateEllipse.F1 = RotatePoint(MyEllipse.F1, Pivot, Angle)
RotateEllipse.F2 = RotatePoint(MyEllipse.F2, Pivot, Angle)
RotateEllipse.P1 = RotatePoint(MyEllipse.P1, Pivot, Angle)
'RotateEllipse.Theta = dAngle(RotateEllipse.Theta + Angle)
'RotateEllipse.Hyp = RotatePoint(Myellipse.Hyp, Myellipse.Center, angle)
'RotateEllipse.Angle1 = (EllipseAngle(RotateEllipse.Angle1 + angle, RotateEllipse))
'RotateEllipse.Angle2 = (EllipseAngle(RotateEllipse.Angle2 + angle, RotateEllipse))
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function RotX(X1 As Single, Y1 As Single, Angle As Single) As Single
RotX = cHyp(X1, Y1) * Cos((PtAng(X1, Y1) + Angle) * Pi / 180)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function RotY(X1 As Single, Y1 As Single, Angle As Single) As Single
RotY = cHyp(X1, Y1) * Sin((PtAng(X1, Y1) + Angle) * Pi / 180)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Sub SaveGeo(zLine As CadLine, MyGeo() As Geometry, FName As String)
On Error GoTo eTrap
Dim FF As Integer
Dim i As Integer
Dim J As Integer
Dim g As Integer
Dim gCount As Integer
FF = FreeFile
Open FName For Output As #FF
Write #FF, "Zoom"
Write #FF, zLine.P1.x, zLine.P1.y, zLine.P2.x, zLine.P2.y
For gCount = 0 To UBound(MyGeo)
    Write #FF, "GEO"
    Write #FF, MyGeo(gCount).Name
    Write #FF, gCount
    g = UBound(MyGeo(gCount).Points)
    For i = 0 To g
        Write #FF, "Point"
        With MyGeo(gCount).Points(i)
            Write #FF, .x, .y, .Layer.Color, .Layer.style, .Layer.Width
        End With
    Next i
    g = UBound(MyGeo(gCount).Lines)
    For i = 0 To g
        Write #FF, "Line"
        With MyGeo(gCount).Lines(i)
            Write #FF, .P1.x, .P1.y, .P2.x, .P2.y, .Layer.Color, .Layer.style, .Layer.Width
        End With
    Next i
    g = UBound(MyGeo(gCount).Arcs)
    For i = 0 To g
        Write #FF, "Arc"
        With MyGeo(gCount).Arcs(i)
            Write #FF, .Center.x, .Center.y, .Radius, .Angle1, .Angle2, .Layer.Color, .Layer.style, .Layer.Width
        End With
    Next i
    g = UBound(MyGeo(gCount).Ellipses)
    For i = 0 To g
        Write #FF, "Ellipse"
        With MyGeo(gCount).Ellipses(i)
            Write #FF, .F1.x, .F1.y, .F2.x, .F2.y, .P1.x, .P1.y, .Angle1, .Angle2, .NumPoints, .Layer.Color, .Layer.style, .Layer.Width
        End With
    Next i
    g = UBound(MyGeo(gCount).Splines)
    For i = 0 To g
        Write #FF, "Spline"
        With MyGeo(gCount).Splines(i)
            Write #FF, UBound(.Vertex), .Layer.Color, .Layer.style, .Layer.Width
            For J = 0 To UBound(.Vertex)
                With .Vertex(J)
                    Write #FF, .x, .y
                End With
            Next J
        End With
    Next i
    g = UBound(MyGeo(gCount).PolyLines)
    For i = 0 To g
        Write #FF, "PolyLine"
        With MyGeo(gCount).PolyLines(i)
            Write #FF, UBound(.Vertex), .Layer.Color, .Layer.style, .Layer.Width
            For J = 0 To UBound(.Vertex)
                With .Vertex(J)
                    Write #FF, .x, .y
                End With
            Next J
        End With
    Next i
    g = UBound(MyGeo(gCount).Faces)
    For i = 0 To g
        Write #FF, "Face"
        With MyGeo(gCount).Faces(i)
            Write #FF, .Layer.Color, .Layer.style, .Layer.Width
            For J = 0 To 2
                With .Vertex(J)
                    Write #FF, .x, .y
                End With
            Next J
        End With
    Next i
    g = UBound(MyGeo(gCount).Text)
    For i = 0 To g
        Write #FF, "Text"
        With MyGeo(gCount).Text(i)
            Write #FF, .Start.x, .Start.y, .Angle, .Length, .Size, CStr(.Layer.FontName), CStr(.Text), .Layer.Color
        End With
    Next i
    g = UBound(MyGeo(gCount).Inserts)
    For i = 0 To g
        Write #FF, "Insert"
        With MyGeo(gCount).Inserts(i)
            Write #FF, .Name, .Base.x, .Base.y, .ScaleX, .ScaleY, .Angle
        End With
    Next i
Next gCount
Close #FF
eTrap:
    g = -1
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function ScaleLine(MyLine As CadLine, ScaleX As Single, ScaleY As Single) As CadLine
ScaleLine = MyLine
ScaleLine.P1 = ScalePoint(MyLine.P1, ScaleX, ScaleY)
ScaleLine.P2 = ScalePoint(MyLine.P2, ScaleX, ScaleY)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ScaleFace(MyFace As CadFace, ScaleX As Single, ScaleY As Single) As CadFace
ScaleFace = MyFace
ScaleFace.Vertex(0) = ScalePoint(MyFace.Vertex(0), ScaleX, ScaleY)
ScaleFace.Vertex(1) = ScalePoint(MyFace.Vertex(1), ScaleX, ScaleY)
ScaleFace.Vertex(2) = ScalePoint(MyFace.Vertex(2), ScaleX, ScaleY)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function ScalePolyLine(MyPolyLine As CadPolyLine, ScaleX As Single, ScaleY As Single) As CadPolyLine
Dim i As Integer
ScalePolyLine = MyPolyLine
For i = 0 To UBound(MyPolyLine.Vertex)
    ScalePolyLine.Vertex(i) = ScalePoint(MyPolyLine.Vertex(i), ScaleX, ScaleY)
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function ScaleText(MyText As CadText, ScaleX As Single, ScaleY As Single) As CadText
ScaleText = MyText
ScaleText.Start = ScalePoint(MyText.Start, ScaleX, ScaleY)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Sub RemoveSelection(MyArray() As SelSet, Index As Integer)
On Error GoTo eTrap
Dim Count As Integer
Dim i As Integer
Count = UBound(MyArray)
For i = Index To Count - 1
    MyArray(i) = MyArray(i + 1)
Next i
If Count = 0 Then Erase MyArray Else ReDim Preserve MyArray(Count - 1)
eTrap:
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub RemoveGeo(MyGeo As Geometry, GeoType As String, Index As Integer)
On Error GoTo eTrap
Dim Count As Integer
Dim i As Integer
Select Case GeoType
    Case "Point"
        Count = UBound(MyGeo.Points)
        For i = Index To Count - 1
            MyGeo.Points(i) = MyGeo.Points(i + 1)
        Next i
        If Count = 0 Then Erase MyGeo.Points Else ReDim Preserve MyGeo.Points(Count - 1)
    Case "Line"
        Count = UBound(MyGeo.Lines)
        For i = Index To Count - 1
            MyGeo.Lines(i) = MyGeo.Lines(i + 1)
        Next i
        If Count = 0 Then Erase MyGeo.Lines Else ReDim Preserve MyGeo.Lines(Count - 1)
    Case "Arc"
        Count = UBound(MyGeo.Arcs)
        For i = Index To Count - 1
            MyGeo.Arcs(i) = MyGeo.Arcs(i + 1)
        Next i
        If Count = 0 Then Erase MyGeo.Arcs Else ReDim Preserve MyGeo.Arcs(Count - 1)
    Case "Ellipse"
        Count = UBound(MyGeo.Ellipses)
        For i = Index To Count - 1
            MyGeo.Ellipses(i) = MyGeo.Ellipses(i + 1)
        Next i
        If Count = 0 Then Erase MyGeo.Ellipses Else ReDim Preserve MyGeo.Ellipses(Count - 1)
    Case "Spline"
        Count = UBound(MyGeo.Splines)
        For i = Index To Count - 1
            MyGeo.Splines(i) = MyGeo.Splines(i + 1)
        Next i
        If Count = 0 Then Erase MyGeo.Splines Else ReDim Preserve MyGeo.Splines(Count - 1)
    Case "PolyLine"
        Count = UBound(MyGeo.PolyLines)
        For i = Index To Count - 1
            MyGeo.PolyLines(i) = MyGeo.PolyLines(i + 1)
        Next i
        If Count = 0 Then Erase MyGeo.PolyLines Else ReDim Preserve MyGeo.PolyLines(Count - 1)
    Case "Text"
        Count = UBound(MyGeo.Text)
        For i = Index To Count - 1
            MyGeo.Text(i) = MyGeo.Text(i + 1)
        Next i
        If Count = 0 Then Erase MyGeo.Text Else ReDim Preserve MyGeo.Text(Count - 1)
    Case "Insert"
        Count = UBound(MyGeo.Inserts)
        For i = Index To Count - 1
            MyGeo.Inserts(i) = MyGeo.Inserts(i + 1)
        Next i
        If Count = 0 Then Erase MyGeo.Inserts Else ReDim Preserve MyGeo.Inserts(Count - 1)
End Select
eTrap:
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Function ScalePoint(MyPoint As CadPoint, ScaleX As Single, ScaleY As Single) As CadPoint
ScalePoint = MyPoint
ScalePoint.x = MyPoint.x * ScaleX
ScalePoint.y = MyPoint.y * ScaleY
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ScaleSpline(MySPline As CadSpline, ScaleX As Single, ScaleY As Single) As CadSpline
Dim i As Integer
ScaleSpline = MySPline
For i = 0 To UBound(MySPline.Vertex)
    ScaleSpline.Vertex(i) = ScalePoint(MySPline.Vertex(i), ScaleX, ScaleY)
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ZZ_OLD_SelectFromPoint(SelRad As Single, MyGeo As Geometry, SelPoint As CadPoint, MySel As SelSet, ByRef ActualPoint As CadPoint) As Boolean
On Error GoTo eTrap
Dim SelArc As CadArc
Dim i As Integer
Dim J As Integer
Dim rad As Single
Dim iPts() As CadPoint
Dim iRes As Integer
Dim Found As Boolean
Dim gCount As Integer
Dim tLine As CadLine
SelArc.Angle1 = 0
SelArc.Angle2 = 360
SelArc.Center = SelPoint
SelArc.Radius = SelRad
gCount = UBound(MyGeo.Points)
For i = 0 To gCount
    If PtLen(SelPoint, MyGeo.Points(i)) < SelArc.Radius Then
        ZZ_OLD_SelectFromPoint = True
        MySel.Type = "Point"
        MySel.Index = i
        ActualPoint = SelPoint
        Exit Function
    End If
Next i
gCount = UBound(MyGeo.Lines)
For i = 0 To gCount
    iRes = LineArcIntersect(MyGeo.Lines(i), SelArc, iPts())
    For J = 0 To iRes
        If iPts(J).Layer.Color = vbBlue Then
            ZZ_OLD_SelectFromPoint = True
            MySel.Type = "Line"
            MySel.Index = i
            ActualPoint = iPts(J)
            Exit Function
        End If
    Next J
Next i
gCount = UBound(MyGeo.Arcs)
For i = 0 To gCount
    iRes = ArcArcIntersect(MyGeo.Arcs(i), SelArc, iPts())
    For J = 0 To iRes
        If iPts(J).Layer.Color = vbBlue Then
            ZZ_OLD_SelectFromPoint = True
            MySel.Type = "Arc"
            MySel.Index = i
            ActualPoint = iPts(J)
            Exit Function
        End If
    Next J
Next i
gCount = UBound(MyGeo.Ellipses)
For i = 0 To gCount
    iRes = EllipseArcIntersect(MyGeo.Ellipses(i), SelArc, iPts())
    For J = 0 To iRes
        If iPts(J).Layer.Color = vbBlue Then
            ZZ_OLD_SelectFromPoint = True
            MySel.Type = "Ellipse"
            MySel.Index = i
            ActualPoint = iPts(J)
            Exit Function
        End If
    Next J
Next i
gCount = UBound(MyGeo.Splines)
For i = 0 To gCount
    iRes = SplineArcIntersect(MyGeo.Splines(i), SelArc, iPts())
    For J = 0 To iRes
        If iPts(J).Layer.Color = vbBlue Then
            ZZ_OLD_SelectFromPoint = True
            MySel.Type = "Spline"
            MySel.Index = i
            ActualPoint = iPts(J)
            Exit Function
        End If
    Next J
Next i
gCount = UBound(MyGeo.PolyLines)
For i = 0 To gCount
    iRes = PolyLineArcIntersect(MyGeo.PolyLines(i), SelArc, iPts())
    For J = 0 To iRes
        If iPts(J).Layer.Color = vbBlue Then
            ZZ_OLD_SelectFromPoint = True
            MySel.Type = "PolyLine"
            MySel.Index = i
            ActualPoint = iPts(J)
            Exit Function
        End If
    Next J
Next i
gCount = UBound(MyGeo.Text)
For i = 0 To gCount
    SelArc.Radius = (MyGeo.Text(i).Size + 1) * 2
    tLine = cAngLine(MyGeo.Text(i).Angle, MyGeo.Text(i).Start, MyGeo.Text(i).Length)
    'tLine.Width = 3
    'DrawCadLine frmDraw.picDraw, tLine, 13, vbGreen
    iRes = LineArcIntersect(tLine, SelArc, iPts())
    For J = 0 To iRes
        If iPts(J).Layer.Color = vbBlue Then
            ZZ_OLD_SelectFromPoint = True
            MySel.Type = "Text"
            MySel.Index = i
            ActualPoint = iPts(J)
            Exit Function
       End If
    Next J
Next i
Erase iPts()
    
'------------ERROR HANDLER------------\
ZZ_OLD_SelectFromPoint = False
Exit Function
eTrap:
    gCount = -1
    If Err.Number <> 9 Then
        MsgBox Err.Description
        DoEvents
    End If
    Resume Next
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function ZZ_OLD_SplineLineIntersect(MySPline As CadSpline, MyLine As CadLine, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim du As Single
Dim vX As Single
Dim vY As Single
Dim bv As Single
Dim K As Single
Dim u As Single
Dim vCount As Integer
Dim iRes As Integer
Dim tLine As CadLine
Dim iPt() As CadPoint
Dim iCount As Integer
Dim sCount As Integer
tLine.P1 = MySPline.Vertex(0)
vCount = UBound(MySPline.Vertex)
du = 0.025 'SplineSmooth
For u = 0 To 1 Step du
    vX = 0: vY = 0
    For K = 0 To vCount ' For Each control point
        bv = sBlend(K, vCount, u) ' Calculate blending Function
        vX = vX + MySPline.Vertex(K).x * bv
        vY = vY + MySPline.Vertex(K).y * bv
    Next K
    tLine.P2.x = vX
    tLine.P2.y = vY
    If LineLineIntersect(MyLine, tLine, iPt()) = 0 Then
        If iPt(0).Layer.Color = vbGreen Or iPt(0).Layer.Color = vbBlue Or sCount = 1 Then
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPt(0)
            iCount = iCount + 1
        End If
    End If
    sCount = sCount + 1
    tLine.P1 = tLine.P2
Next u
tLine.P2 = MySPline.Vertex(vCount)
If LineLineIntersect(MyLine, tLine, iPt()) = 0 Then
    'If iPt(0).layer.color = vbGreen Or iPt(0).layer.color = vbBlue Then
        ReDim Preserve iPoints(iCount) As CadPoint
        iPoints(iCount) = iPt(0)
        iCount = iCount + 1
    'End If
End If
ZZ_OLD_SplineLineIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function SplineLineIntersect(MySPline As CadSpline, MyLine As CadLine, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim i As Integer
Dim tPts() As CadPoint
Dim iPts() As CadPoint
Dim tLine As CadLine
Dim iCount As Integer
SplinePoints MySPline, tPts()
tLine.P1 = tPts(0)
For i = 1 To UBound(tPts)
    tLine.P2 = tPts(i)
    If LineLineIntersect(MyLine, tLine, iPts()) >= 0 Then
        If iPts(0).Layer.Color = vbGreen Or iPts(0).Layer.Color = vbBlue Or i = 1 Or i = UBound(tPts) Then
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPts(0)
            iCount = iCount + 1
        End If
    End If
    tLine.P1 = tLine.P2
Next i
SplineLineIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ZZ_OLD_SplineInBox(MySPline As CadSpline, SelLine As CadLine) As Boolean
Dim du As Single
Dim vX As Single
Dim vY As Single
Dim bv As Single
Dim K As Single
Dim u As Single
Dim vCount As Integer
Dim iPt As CadPoint
vCount = UBound(MySPline.Vertex)
If PtInBox(SelLine, MySPline.Vertex(0)) Then ZZ_OLD_SplineInBox = True: Exit Function
If PtInBox(SelLine, MySPline.Vertex(vCount)) Then ZZ_OLD_SplineInBox = True: Exit Function
du = 0.025 'SplineSmooth
For u = 0 To 1 Step du
    vX = 0: vY = 0
    For K = 0 To vCount ' For Each control point
        bv = sBlend(K, vCount, u) ' Calculate blending Function
        vX = vX + MySPline.Vertex(K).x * bv
        vY = vY + MySPline.Vertex(K).y * bv
    Next K
    iPt.x = vX
    iPt.y = vY
    If PtInBox(SelLine, iPt) Then ZZ_OLD_SplineInBox = True: Exit Function
Next u
ZZ_OLD_SplineInBox = False
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function SplineInBox(MySPline As CadSpline, SelLine As CadLine) As Boolean
Dim i As Integer
Dim tPts() As CadPoint
SplinePoints MySPline, tPts()
For i = 0 To UBound(tPts)
    If PtInBox(SelLine, tPts(i)) Then SplineInBox = True: Exit Function
Next i
SplineInBox = False
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ZZ_OLD_SplineSplineIntersect(SplineA As CadSpline, SplineB As CadSpline, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim iCount As Integer
Dim iPt() As CadPoint
Dim iRes As Integer
'-------------------------------
Dim duA As Single
Dim vXA As Single
Dim vYA As Single
Dim bvA As Single
Dim ka As Single
Dim uA As Single
Dim VCountA As Integer
Dim iResA As Integer
Dim tLineA As CadLine
Dim iPtA As CadPoint
'------------------------------
Dim duB As Single
Dim vXB As Single
Dim vYB As Single
Dim bvB As Single
Dim kB As Single
Dim uB As Single
Dim VCountB As Integer
Dim tLineB As CadLine
'------------------------------
tLineA.P1 = SplineA.Vertex(0)
VCountA = UBound(SplineA.Vertex)
duA = 0.025 'SplineSmooth
For uA = 0 To 1 Step duA
    vXA = 0: vYA = 0
    For ka = 0 To VCountA ' For Each control point
        bvA = sBlend(ka, VCountA, uA) ' Calculate blending Function
        vXA = vXA + SplineA.Vertex(ka).x * bvA
        vYA = vYA + SplineA.Vertex(ka).y * bvA
    Next ka
    tLineA.P2.x = vXA
    tLineA.P2.y = vYA
    '----------------------------
    tLineB.P1 = SplineB.Vertex(0)
    VCountB = UBound(SplineB.Vertex)
    duB = 0.025 'SplineSmooth
    For uB = 0 To 1 Step duB
        vXB = 0: vYB = 0
        For kB = 0 To VCountB ' For Each control point
            bvB = sBlend(kB, VCountB, uB) ' Calculate blending Function
            vXB = vXB + SplineB.Vertex(kB).x * bvB
            vYB = vYB + SplineB.Vertex(kB).y * bvB
        Next kB
        tLineB.P2.x = vXB
        tLineB.P2.y = vYB
        '----------------------------
        If LineLineIntersect(tLineA, tLineB, iPt()) = 0 Then
            If iPt(0).Layer.Color = vbBlue Then
                ReDim Preserve iPoints(iCount) As CadPoint
                iPoints(iCount) = iPt(0)
                iCount = iCount + 1
            End If
        End If
        '----------------------------
        tLineB.P1 = tLineB.P2
    Next uB
    tLineB.P2 = SplineB.Vertex(VCountB)
    If LineLineIntersect(tLineA, tLineB, iPt()) = 0 Then
        If iPt(0).Layer.Color = vbBlue Then
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPt(0)
            iCount = iCount + 1
        End If
    End If
    '----------------------------
    tLineA.P1 = tLineA.P2
Next uA
tLineA.P2 = SplineA.Vertex(VCountA)
'----------------------------
tLineB.P1 = SplineB.Vertex(0)
VCountB = UBound(SplineB.Vertex)
duB = 0.025 'SplineSmooth
For uB = 0 To 1 Step duB
    vXB = 0: vYB = 0
    For kB = 0 To VCountB ' For Each control point
        bvB = sBlend(kB, VCountB, uB) ' Calculate blending Function
        vXB = vXB + SplineB.Vertex(kB).x * bvB
        vYB = vYB + SplineB.Vertex(kB).y * bvB
    Next kB
    tLineB.P2.x = vXB
    tLineB.P2.y = vYB
    '----------------------------
    If LineLineIntersect(tLineA, tLineB, iPt()) = 0 Then
        If iPt(0).Layer.Color = vbBlue Then
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPt(0)
            iCount = iCount + 1
        End If
    End If
    '----------------------------
    tLineB.P1 = tLineB.P2
Next uB
tLineB.P2 = SplineB.Vertex(VCountB)
If LineLineIntersect(tLineA, tLineB, iPt()) = 0 Then
    If iPt(0).Layer.Color = vbBlue Then
        ReDim Preserve iPoints(iCount) As CadPoint
        iPoints(iCount) = iPt(0)
        iCount = iCount + 1
    End If
End If
'----------------------------
ZZ_OLD_SplineSplineIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function SplineSplineIntersect(SplineA As CadSpline, SplineB As CadSpline, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim iCount As Integer
Dim iPts() As CadPoint
Dim tLineA  As CadLine
Dim tLineB As CadLine
Dim APts() As CadPoint
Dim BPts() As CadPoint
Dim a As Integer
Dim B As Integer
SplinePoints SplineA, APts()
SplinePoints SplineB, BPts()

tLineA.P1 = APts(0)
For a = 1 To UBound(APts)
    tLineA.P2 = APts(a)
    tLineB.P1 = BPts(0)
    For B = 1 To UBound(BPts)
        tLineB.P2 = BPts(B)
        If LineLineIntersect(tLineA, tLineB, iPts()) >= 0 Then
            If iPts(0).Layer.Color = vbBlue Then
                ReDim Preserve iPoints(iCount) As CadPoint
                iPoints(iCount) = iPts(0)
                iCount = iCount + 1
            End If
        End If
        tLineB.P1 = tLineB.P2
    Next B
    tLineA.P1 = tLineA.P2
Next a
'----------------------------
SplineSplineIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function


Function ZZ_OLD_SplineArcIntersect(MySPline As CadSpline, MyArc As CadArc, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim du As Single
Dim vX As Single
Dim vY As Single
Dim bv As Single
Dim K As Single
Dim u As Single
Dim vCount As Integer
Dim iRes As Integer
Dim tLine As CadLine
Dim iPts() As CadPoint
Dim iCount As Integer
Dim i As Integer
Dim sCount As Integer
tLine.P1 = MySPline.Vertex(0)
vCount = UBound(MySPline.Vertex)
du = 0.025 'SplineSmooth
For u = 0 To 1 Step du
    vX = 0: vY = 0
    For K = 0 To vCount ' For Each control point
        bv = sBlend(K, vCount, u) ' Calculate blending Function
        vX = vX + MySPline.Vertex(K).x * bv
        vY = vY + MySPline.Vertex(K).y * bv
    Next K
    tLine.P2.x = vX
    tLine.P2.y = vY
    iRes = LineArcIntersect(tLine, MyArc, iPts())
    For i = 0 To iRes
        If iPts(i).Layer.Color = vbBlue Or iPts(i).Layer.Color = vbRed Or sCount = 1 Then
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPts(i)
            iCount = iCount + 1
        End If
    Next i
    sCount = sCount + 1
    tLine.P1 = tLine.P2
Next u
tLine.P2 = MySPline.Vertex(vCount)
iRes = LineArcIntersect(tLine, MyArc, iPts())
For i = 0 To iRes
    'If ipts(i).layer.color = vbBlue Then
        ReDim Preserve iPoints(iCount) As CadPoint
        iPoints(iCount) = iPts(i)
        iCount = iCount + 1
    'End If
Next i
ZZ_OLD_SplineArcIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function SplineArcIntersect(MySPline As CadSpline, MyArc As CadArc, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim iRes As Integer
Dim tLine As CadLine
Dim iPts() As CadPoint
Dim iCount As Integer
Dim i As Integer
Dim K As Integer
Dim tPts() As CadPoint
SplinePoints MySPline, tPts()
tLine.P1 = tPts(0)
For K = 1 To UBound(tPts)
    tLine.P2 = tPts(K)
    iRes = LineArcIntersect(tLine, MyArc, iPts())
    For i = 0 To iRes
        If iPts(i).Layer.Color = vbBlue Or iPts(i).Layer.Color = vbRed Or K = 1 Or K = UBound(tPts) Then
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPts(i)
            iCount = iCount + 1
        End If
    Next i
    tLine.P1 = tLine.P2
Next K
SplineArcIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Function ZZ_OLD_SplineEllipseIntersect(MySPline As CadSpline, MyEllipse As CadEllipse, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim du As Single
Dim vX As Single
Dim vY As Single
Dim bv As Single
Dim K As Single
Dim u As Single
Dim vCount As Integer
Dim iRes As Integer
Dim tLine As CadLine
Dim iPts() As CadPoint
Dim iCount As Integer
Dim i As Integer
Dim sCount As Integer
tLine.P1 = MySPline.Vertex(0)
vCount = UBound(MySPline.Vertex)
du = 0.025 'SplineSmooth
For u = 0 To 1 Step du
    vX = 0: vY = 0
    For K = 0 To vCount ' For Each control point
        bv = sBlend(K, vCount, u) ' Calculate blending Function
        vX = vX + MySPline.Vertex(K).x * bv
        vY = vY + MySPline.Vertex(K).y * bv
    Next K
    tLine.P2.x = vX
    tLine.P2.y = vY
    iRes = LineEllipseIntersect(tLine, MyEllipse, iPts())
    For i = 0 To iRes
        If iPts(i).Layer.Color = vbBlue Or iPts(i).Layer.Color = vbRed Or sCount = 1 Then
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPts(i)
            iCount = iCount + 1
        End If
    Next i
    sCount = sCount + 1
    tLine.P1 = tLine.P2
Next u
tLine.P2 = MySPline.Vertex(vCount)
iRes = LineEllipseIntersect(tLine, MyEllipse, iPts())
For i = 0 To iRes
    'If ipts(i).layer.color = vbBlue Or ipts(i).layer.color = vbRed Then
        ReDim Preserve iPoints(iCount) As CadPoint
        iPoints(iCount) = iPts(i)
        iCount = iCount + 1
    'End If
Next i
ZZ_OLD_SplineEllipseIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function SplineEllipseIntersect(MySPline As CadSpline, MyEllipse As CadEllipse, ByRef iPoints() As CadPoint) As Integer
Erase iPoints()
Dim iRes As Integer
Dim tLine As CadLine
Dim iPts() As CadPoint
Dim iCount As Integer
Dim i As Integer
Dim tPts() As CadPoint
Dim K As Integer
SplinePoints MySPline, tPts()
tLine.P1 = tPts(0)
For K = 1 To UBound(tPts)
    tLine.P2 = tPts(K)
    iRes = LineEllipseIntersect(tLine, MyEllipse, iPts())
    For i = 0 To iRes
        If iPts(i).Layer.Color = vbBlue Or iPts(i).Layer.Color = vbRed Or K = 1 Or K = UBound(tPts) Then
            ReDim Preserve iPoints(iCount) As CadPoint
            iPoints(iCount) = iPts(i)
            iCount = iCount + 1
        End If
    Next i
    tLine.P1 = tLine.P2
Next K
SplineEllipseIntersect = iCount - 1
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function

Sub Swap(ByRef a As Variant, ByRef B As Variant)
Dim c As Variant
c = a
a = B
B = c
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub DrawCadLine(Canvas As PictureBox, MyLine As CadLine, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
If Width = 0 Then Width = MyLine.Layer.Width
Canvas.DrawWidth = Width
Canvas.DrawMode = Mode
Canvas.DrawStyle = MyLine.Layer.style
If Color = -1 Then Color = MyLine.Layer.Color
Canvas.Line (MyLine.P1.x, -MyLine.P1.y)-(MyLine.P2.x, -MyLine.P2.y), Color
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub DrawCadFace(Canvas As PictureBox, MyFace As CadFace, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
If Width = 0 Then Width = MyFace.Layer.Width
Canvas.DrawWidth = Width
Canvas.DrawMode = Mode
Canvas.DrawStyle = MyFace.Layer.style
If Color = -1 Then Color = MyFace.Layer.Color
Canvas.Line (MyFace.Vertex(0).x, -MyFace.Vertex(0).y)-(MyFace.Vertex(1).x, -MyFace.Vertex(1).y), Color
Canvas.Line (MyFace.Vertex(1).x, -MyFace.Vertex(1).y)-(MyFace.Vertex(2).x, -MyFace.Vertex(2).y), Color
Canvas.Line (MyFace.Vertex(2).x, -MyFace.Vertex(2).y)-(MyFace.Vertex(0).x, -MyFace.Vertex(0).y), Color
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub DrawCadPolyLine(Canvas As PictureBox, MyPolyLine As CadPolyLine, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
Dim i As Integer
If Width = 0 Then Width = MyPolyLine.Layer.Width
Canvas.DrawWidth = Width
Canvas.DrawMode = Mode
Canvas.DrawStyle = MyPolyLine.Layer.style
If Color = -1 Then Color = MyPolyLine.Layer.Color
Canvas.PSet (MyPolyLine.Vertex(0).x, -MyPolyLine.Vertex(0).y), Color
For i = 1 To UBound(MyPolyLine.Vertex)
    Canvas.Line -(MyPolyLine.Vertex(i).x, -MyPolyLine.Vertex(i).y), Color
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub

Sub ZZ_OLD_DrawCadSpline(Canvas As PictureBox, MySPline As CadSpline, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
If Width = 0 Then Width = MySPline.Layer.Width
Canvas.DrawWidth = Width
Canvas.DrawMode = Mode
Canvas.DrawStyle = MySPline.Layer.style
If Color = -1 Then Color = MySPline.Layer.Color
Dim du As Single
Dim vX As Single
Dim vY As Single
Dim bv As Single
Dim K As Single
Dim u As Single
Dim vCount As Integer
vCount = UBound(MySPline.Vertex)
If vCount < 2 Then Exit Sub
'If (MySpline.Vertex(VCount).x = MySpline.Vertex(VCount - 1).x) And (MySpline.Vertex(VCount).y = MySpline.Vertex(VCount - 1).y) Then Exit Sub
du = 0.025 'SplineSmooth
Canvas.PSet (MySPline.Vertex(0).x, -MySPline.Vertex(0).y), Color
For u = 0 To 1 Step du
    vX = 0: vY = 0
    For K = 0 To vCount ' For Each control point
        bv = sBlend(K, vCount, u) ' Calculate blending Function
        vX = vX + MySPline.Vertex(K).x * bv
        vY = vY + MySPline.Vertex(K).y * bv
    Next K
    Canvas.Line -(vX, -vY), Color ' Draw To the point
Next u
Canvas.Line -(MySPline.Vertex(vCount).x, -MySPline.Vertex(vCount).y), Color

'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Sub DrawCadSpline(Canvas As PictureBox, MySPline As CadSpline, Optional ByVal Mode As Integer = 13, Optional ByVal Color As Long = -1, Optional ByVal Width As Integer)
Dim tPts() As CadPoint
Dim i As Integer
If Width = 0 Then Width = MySPline.Layer.Width
Canvas.DrawWidth = Width
Canvas.DrawMode = Mode
Canvas.DrawStyle = MySPline.Layer.style
If Color = -1 Then Color = MySPline.Layer.Color
SplinePoints MySPline, tPts()
Canvas.CurrentX = tPts(0).x: Canvas.CurrentY = -tPts(0).y
For i = 1 To UBound(tPts)
    Canvas.Line -(tPts(i).x, -tPts(i).y), Color
Next i
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Sub
Function sBlend(K, n, u)
    'Bezier blending function
    sBlend = sFunct(n, K) * (u ^ K) * (1 - u) ^ (n - K)
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function Factorial(n)
    ' Recursive factorial fucntion
    If n = 1 Or n = 0 Then
        Factorial = 1
    Else
        Factorial = n * Factorial(n - 1)
    End If
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function
Function sFunct(n, r)
    ' Implements c!/r!*(n-r)!
    sFunct = Factorial(n) / (Factorial(r) * Factorial(n - r))
'***************************************************************************
'All code researched and developed by Dave Andrews unless otherwise noted.
'Feel free to use or recycle this code in any way you like.


'***************************************************************************
End Function





