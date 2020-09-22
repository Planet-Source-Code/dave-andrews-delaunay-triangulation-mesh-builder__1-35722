Attribute VB_Name = "modDeLaunay"
Option Explicit
'This is my Delaunay module which took a VERY long time
'It will even detect "HOLES" and build a mesh around those holes
'Please give me credit if you use this code.
'Dave Andrews, 2000-2002
Const FloodColor = vbMagenta
Private Type OType
    ID As Long
    Start As CadLine
    End As CadLine
    Reverse As Boolean
    Order As Long
    Count As Long
End Type
Type RefinedEdge
    P(2) As CadPoint
    Use As Boolean
    Flipped As Boolean
End Type
Type Edge
    StartPt As Integer
    EndPt As Integer
    Right As Integer
    left As Integer
    Boundary As Boolean
    OutMost As Boolean
    Que As Integer
End Type
Type Element
    Vertices(2) As Integer
    Edges(2) As Integer
    Que As Integer
End Type
Type PointArray
    P() As CadPoint
End Type
Dim LGrad As Single
Dim cx As Single ' center user coords for flood
Dim cy As Single
Dim pX As Long 'pixel coords for flood
Dim pY As Long
Dim dPoints() As CadPoint
Dim dEdges() As Edge
Dim dTriangles() As Element
Dim Refined() As RefinedEdge
Dim SkipSplit As Boolean

Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Function CheckTriAngle(P1 As CadPoint, P2 As CadPoint, P3 As CadPoint) As Single
Dim ang1 As Single
ang1 = TriPtAngle(P1, P2, P3)
If ang1 > 180 Then ang1 = 360 - ang1
CheckTriAngle = ang1
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function

Sub CompleteSplit(Canvas As PictureBox, ByRef Ref() As RefinedEdge, I As Integer, a As Integer, B As Integer, c As Integer, Min As Single)
Dim K As Integer
Dim MPt As CadPoint
'A-B is too long
MPt = MidPoint(PtLine(Ref(I).P(a), Ref(I).P(B)))
MPt.Layer.ID = 0
K = UBound(Ref) + 1
ReDim Preserve Ref(K + 1)
Ref(K) = Ref(I)
Ref(K + 1) = Ref(I)
Ref(K).P(a) = MPt
Ref(K + 1).P(B) = MPt
Ref(K).Use = True
Ref(K + 1).Use = True
Ref(I).Use = False

DrawCadLine Canvas, PtLine(Ref(K).P(0), Ref(K).P(1)), , vbRed, 3
DrawCadLine Canvas, PtLine(Ref(K).P(1), Ref(K).P(2)), , vbRed, 3
DrawCadLine Canvas, PtLine(Ref(K).P(2), Ref(K).P(0)), , vbRed, 3
DrawCadLine Canvas, PtLine(Ref(K + 1).P(0), Ref(K + 1).P(1)), , vbGreen, 1
DrawCadLine Canvas, PtLine(Ref(K + 1).P(1), Ref(K + 1).P(2)), , vbGreen, 1
DrawCadLine Canvas, PtLine(Ref(K + 1).P(2), Ref(K + 1).P(0)), , vbGreen, 1
DoEvents

SplitTriangle Canvas, Ref(), K, Min, False
SplitTriangle Canvas, Ref(), K + 1, Min, False
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Function DeLaunayGeo(Canvas As PictureBox, ByRef MyGeo As Geometry, Center As CadPoint, ByRef MyFaces() As CadFace, ByRef MyOutline() As CadLine, DFactor As Single, Optional UseLines As Boolean, Optional SplitT As Boolean = True) As Geometry
On Local Error GoTo eTrap
Dim I As Integer
Dim J As Integer
Dim XT As Long
Dim tPoint As Integer
Erase MyOutline()
Erase dPoints()
Erase dEdges()
Erase dTriangles()
XT = Timer
tPoint = Canvas.MousePointer
Canvas.MousePointer = 11
SkipSplit = UseLines
cx = Center.x
cy = Center.y
Canvas.AutoRedraw = True ' This is needed to properly flood-fill for adjutment of normals
If Not UseLines Then
    GeometryToLines MyGeo, MyOutline()
    RemoveNullLines MyOutline()
    OrderLines MyOutline(), "0.00000"
Else
    AddLines MyGeo.Lines(), MyOutline()
End If
LGrad = DFactor
I = UBound(MyGeo.Arcs)
If I <> -1 Then
    For I = 0 To UBound(MyGeo.Arcs)
        ReDim Preserve dPoints(J) As CadPoint
        dPoints(J) = MyGeo.Arcs(I).Center
        J = J + 1
    Next I
End If
FullDelaunay Canvas, MyOutline()
'RefineEdges Canvas, dTriangles(), dEdges(), MyFaces(), Sqr(DFactor ^ 2 + DFactor ^ 2)
RefineEdges Canvas, dTriangles(), dEdges(), MyFaces(), 2 * DFactor, SplitT
'UnifyNormals dPoints(), dTriangles()
'GenSimpleFaces MyFaces()
For I = 0 To UBound(MyOutline)
    MyOutline(I).Layer.Color = vbBlue
    MyOutline(I).Layer.Mode = 13
    MyOutline(I).Layer.style = 0
    MyOutline(I).Layer.Width = 2
Next I
UnifyFaces MyFaces(), False
Canvas.MousePointer = tPoint
MsgBox "Mesh Completed: " & Long2Time((Timer - XT) * 1000) & Chr$(10) & UBound(MyFaces) & " Faces"
Exit Function
eTrap:
    I = -1
    Resume Next
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function


Function FlipTriangle(Canvas As PictureBox, ByRef Ref() As RefinedEdge, R1 As Integer, P1 As Integer, P2 As Integer, P3 As Integer, Min As Single) As Boolean
Dim I As Integer
Dim J As Integer
Dim a As Integer
Dim B As Integer
Dim c As Integer
Dim K As Integer
For I = 0 To UBound(Ref) 'cycle through all our triangles
    If Ref(I).Use And I <> R1 Then  'this triangle is not locked
        For J = 0 To 5
            Select Case J
                Case 0: a = 0: B = 1: c = 2
                Case 1: a = 0: B = 2: c = 1
                Case 2: a = 1: B = 0: c = 2
                Case 3: a = 1: B = 2: c = 0
                Case 4: a = 2: B = 0: c = 1
                Case 5: a = 2: B = 1: c = 0
            End Select
            If MPts(Ref(R1).P(P1), Ref(I).P(a)) And MPts(Ref(R1).P(P2), Ref(I).P(B)) Then
                'we found our match
                
                Ref(R1).P(P1) = Ref(I).P(c)
                Ref(R1).P(P1).Layer.ID = 0
                
                Ref(I).P(B) = Ref(R1).P(P3)
                Ref(I).P(B).Layer.ID = 0
                
                Ref(R1).Flipped = True
                Ref(I).Flipped = True
                
                'DrawCadLine Canvas, PtLine(Ref(R1).P(0), Ref(R1).P(1)), , vbWhite, 3
                'DrawCadLine Canvas, PtLine(Ref(R1).P(1), Ref(R1).P(2)), , vbWhite, 3
                'DrawCadLine Canvas, PtLine(Ref(R1).P(2), Ref(R1).P(0)), , vbWhite, 3
                'DrawCadLine Canvas, PtLine(Ref(i).P(0), Ref(i).P(1)), , vbBlue, 1
                'DrawCadLine Canvas, PtLine(Ref(i).P(1), Ref(i).P(2)), , vbBlue, 1
                'DrawCadLine Canvas, PtLine(Ref(i).P(2), Ref(i).P(0)), , vbBlue, 1
                'DoEvents
                
                FlipTriangle = True
                
                SplitTriangle Canvas, Ref(), R1, Min, True
                SplitTriangle Canvas, Ref(), I, Min, True
                
                Exit Function
            End If
        Next J
    End If
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Sub GenMeshPoints(Canvas As PictureBox, DLines() As CadLine)
On Local Error GoTo eTrap
Dim I As Single
Dim J As Single
Dim K As Integer
Dim gLeft As Single
Dim gRight As Single
Dim gTop As Single
Dim gBottom As Single
Dim tPt As CadPoint
Dim Limit As Single
'ReDim Preserve dPoints(I) As CadPoint
ReDim dEdges(0) As Edge
ReDim dTriangles(0) As Element
AdjustEdgeNormals Canvas, DLines()
gLeft = GetLeft(DLines()) - (LGrad / 2)
gRight = GetRight(DLines()) + (LGrad / 2)
gTop = GetUpper(DLines()) + (LGrad / 2)
gBottom = GetLower(DLines()) - (LGrad / 2)
Limit = LGrad
For I = 0 To UBound(DLines)
    DrawCadLine Canvas, DLines(I), , vbRed, Canvas.ScaleX(LGrad, 0, 3)
Next I
LineArrayToEdges Canvas, DLines(), dPoints(), dEdges()
K = UBound(dPoints) + 1
'For i = gLeft To gRight Step Limit
'    For j = gBottom To gTop Step Limit
'        If Canvas.Point(i, -j) = FloodColor Then
'            tPt.x = i
'            tPt.y = j
'            ReDim Preserve dPoints(k) As CadPoint
'            dPoints(k) = tPt
'            k = k + 1
'            DrawCadPoint Canvas, tPt, , vbBlue, 3
'        End If
'    Next j
'Next i
DoEvents
Canvas.Picture = Canvas.Image
ReversePSLG DLines() ' This is for the edges of the 3d lug . . the normals must point outward
Exit Sub
eTrap:
    I = 0
    Resume Next
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub


Sub GenSimpleFaces(ByRef MyFaces() As CadFace)
Dim I As Integer
Dim J As Integer
For I = 0 To UBound(dTriangles)
    If dTriangles(I).Que > -1 Then
        ReDim Preserve MyFaces(J)
        MyFaces(J).Vertex(0) = dPoints(dEdges(dTriangles(I).Edges(0)).StartPt)
        MyFaces(J).Vertex(1) = dPoints(dEdges(dTriangles(I).Edges(1)).StartPt)
        MyFaces(J).Vertex(2) = dPoints(dEdges(dTriangles(I).Edges(2)).StartPt)
        MyFaces(J).Layer.Color = vbMagenta
        MyFaces(J).Layer.Width = 1
        MyFaces(J).Layer.Mode = 13
        MyFaces(J).Layer.style = 0
        J = J + 1
    End If
Next I

'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Sub RefineEdges(Canvas As PictureBox, ByRef MyTriangles() As Element, ByRef MyEdges() As Edge, ByRef MyFaces() As CadFace, Min As Single, SplitT As Boolean)
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim a As Integer
Dim B As Integer
Dim c As Integer
Dim P As Integer
Dim MPt As CadPoint
Dim Count As Integer
Dim TooBig As Boolean
For I = 0 To UBound(MyTriangles)
    If MyTriangles(I).Que >= 0 Then
        For J = 0 To 2
            MyEdges(MyTriangles(I).Edges(J)).StartPt = MyTriangles(I).Vertices(J)
            MyEdges(MyTriangles(I).Edges(J)).EndPt = MyTriangles(I).Vertices(IIf(J < 2, J + 1, 0))
            ReDim Preserve Refined(K)
            Refined(K).P(J) = dPoints(MyEdges(MyTriangles(I).Edges(J)).StartPt)
            If MyEdges(MyTriangles(I).Edges(J)).Boundary = True Then
                Refined(K).P(J).Layer.ID = -100
            Else
                Refined(K).P(J).Layer.ID = 0
            End If
            Refined(K).Use = True
            Refined(K).Flipped = False
        Next J
        K = K + 1
    End If
Next I
If SplitT Then
    a = UBound(Refined)
    For I = 0 To a
        SplitTriangle Canvas, Refined(), I, Min, False
    Next I
End If
'a = UBound(Refined)
'For i = 0 To a
'    SplitTriangle Canvas, Refined(), i, Min, True
'Next i
K = 0
For I = 0 To UBound(Refined)
    If Refined(I).Use Then
        ReDim Preserve MyFaces(K)
        For J = 0 To 2
            MyFaces(K).Vertex(J) = Refined(I).P(J)
            MyFaces(K).Vertex(J).Layer.Locked = False
        Next J
        MyFaces(K).Layer.Color = vbMagenta
        MyFaces(K).Layer.Width = 1
        MyFaces(K).Layer.Mode = 13
        MyFaces(K).Layer.style = 0
        MyFaces(K).Layer.ID = I
        K = K + 1
    End If
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Sub ReversePSLG(ByRef LineAry() As CadLine)
Dim I As Integer
Dim J As Integer
Dim P1 As CadPoint
ReDim AryCopy(0) As CadLine
For I = UBound(LineAry) To 0 Step -1
    P1 = LineAry(I).P1
    LineAry(I).P1 = LineAry(I).P2
    LineAry(I).P2 = P1
    ReDim Preserve AryCopy(J) As CadLine
    AryCopy(J) = LineAry(I)
    J = J + 1
Next I
ReDim LineAry(0) As CadLine
For I = 0 To UBound(AryCopy)
    ReDim Preserve LineAry(I) As CadLine
    LineAry(I) = AryCopy(I)
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Sub AdjustEdgeNormals(Canvas As PictureBox, DLines() As CadLine)
Dim I As Single
Dim tLine As CadLine
Dim P1 As CadPoint
Canvas.Picture = LoadPicture()
For I = 0 To UBound(DLines)
    DrawCadLine Canvas, DLines(I), , vbRed, 2
Next I
'Flood Center
pX = Canvas.ScaleX(cx, 0, 3) - Canvas.ScaleX(Canvas.ScaleLeft, 0, 3)
pY = Canvas.ScaleY(-cy, 0, 3) - Canvas.ScaleY(Canvas.ScaleTop, 0, 3)
Canvas.FillStyle = 0
Canvas.FillColor = FloodColor
ExtFloodFill Canvas.hDC, pX, pY, Canvas.Point(cx, -cy), 1
Canvas.FillStyle = 1
'--------------Fix Normals----------
For I = 0 To UBound(DLines)
    tLine = cNormal(DLines(I), Canvas.ScaleX(3, 3, 0))
    If Canvas.Point(tLine.P2.x, -tLine.P2.y) <> FloodColor Then
        SwapPt DLines(I).P1, DLines(I).P2
    End If
Next I

If Not SkipSplit Then SplitLineApi DLines(), LGrad
'SplitLines DLines(), LGrad
For I = 0 To UBound(DLines)
    DrawCadLine Canvas, DLines(I), , vbBlue, 1
Next I
For I = 0 To UBound(DLines)
        tLine.P1 = DLines(I).P1
        tLine.P2 = DLines(I).P2
        DrawCadLine Canvas, cNormal(tLine, 0.02), , vbGreen, 1
        DrawCadPoint Canvas, DLines(I).P1, , vbWhite, 3
        DrawCadPoint Canvas, DLines(I).P2, , vbWhite, 3
Next I
DoEvents
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub






Sub SplitLineApi(ByRef LineAry() As CadLine, sLen As Single)
Dim I As Integer
Dim J As Integer
Dim CurLen As Single
Dim CutLine As CadLine
Dim P1 As CadPoint
Dim IntPt As CadPoint
ReDim WorkLine(0) As CadLine
'-------X's First----------
For I = 0 To UBound(LineAry)
    If cXLen(LineAry(I)) > sLen Then
        If LineAry(I).P1.x < LineAry(I).P2.x Then
            P1 = LineAry(I).P1
            For CurLen = LineAry(I).P1.x To LineAry(I).P2.x - sLen Step sLen
                CutLine.P1.x = CurLen + sLen
                CutLine.P2.x = CurLen + sLen
                CutLine.P1.y = -50
                CutLine.P2.y = 50
                CheckIntersect LineAry(I), CutLine, IntPt
                ReDim Preserve WorkLine(J) As CadLine
                WorkLine(J).P1 = P1
                WorkLine(J).P2 = IntPt
                J = J + 1
                P1 = IntPt
            Next CurLen
            If Not MPts(P1, LineAry(I).P2) Then
                ReDim Preserve WorkLine(J) As CadLine
                WorkLine(J).P1 = P1
                WorkLine(J).P2 = LineAry(I).P2
                J = J + 1
            End If
        ElseIf LineAry(I).P1.x > LineAry(I).P2.x Then
            P1 = LineAry(I).P1
            For CurLen = LineAry(I).P1.x To LineAry(I).P2.x + sLen Step -sLen
                CutLine.P1.x = CurLen - sLen
                CutLine.P2.x = CurLen - sLen
                CutLine.P1.y = -50
                CutLine.P2.y = 50
                CheckIntersect LineAry(I), CutLine, IntPt
                ReDim Preserve WorkLine(J) As CadLine
                WorkLine(J).P1 = P1
                WorkLine(J).P2 = IntPt
                J = J + 1
                P1 = IntPt
            Next CurLen
            If Not MPts(P1, LineAry(I).P2) Then
                ReDim Preserve WorkLine(J) As CadLine
                WorkLine(J).P1 = P1
                WorkLine(J).P2 = LineAry(I).P2
                J = J + 1
            End If
        Else
            ReDim Preserve WorkLine(J) As CadLine
            WorkLine(J) = LineAry(I)
            J = J + 1
        End If
    Else
        ReDim Preserve WorkLine(J) As CadLine
        WorkLine(J) = LineAry(I)
        J = J + 1
    End If
Next I
'Copy The Array
ReDim LineAry(0) As CadLine
For I = 0 To UBound(WorkLine)
    ReDim Preserve LineAry(I) As CadLine
    LineAry(I) = WorkLine(I)
Next I
ReDim WorkLine(0) As CadLine
J = 0
'-------Now Y's----------
For I = 0 To UBound(LineAry)
    If cYLen(LineAry(I)) > sLen Then
        If LineAry(I).P1.y < LineAry(I).P2.y Then
            P1 = LineAry(I).P1
            For CurLen = LineAry(I).P1.y To LineAry(I).P2.y - sLen Step sLen
                CutLine.P1.y = CurLen + sLen
                CutLine.P2.y = CurLen + sLen
                CutLine.P1.x = -50
                CutLine.P2.x = 50
                CheckIntersect LineAry(I), CutLine, IntPt
                ReDim Preserve WorkLine(J) As CadLine
                WorkLine(J).P1 = P1
                WorkLine(J).P2 = IntPt
                J = J + 1
                P1 = IntPt
            Next CurLen
            If Not MPts(P1, LineAry(I).P2) Then
                ReDim Preserve WorkLine(J) As CadLine
                WorkLine(J).P1 = P1
                WorkLine(J).P2 = LineAry(I).P2
                J = J + 1
            End If
        ElseIf LineAry(I).P1.y > LineAry(I).P2.y Then
            P1 = LineAry(I).P1
            For CurLen = LineAry(I).P1.y To LineAry(I).P2.y + sLen Step -sLen
                CutLine.P1.y = CurLen - sLen
                CutLine.P2.y = CurLen - sLen
                CutLine.P1.x = -50
                CutLine.P2.x = 50
                CheckIntersect LineAry(I), CutLine, IntPt
                ReDim Preserve WorkLine(J) As CadLine
                WorkLine(J).P1 = P1
                WorkLine(J).P2 = IntPt
                J = J + 1
                P1 = IntPt
            Next CurLen
            If Not MPts(P1, LineAry(I).P2) Then
                ReDim Preserve WorkLine(J) As CadLine
                WorkLine(J).P1 = P1
                WorkLine(J).P2 = LineAry(I).P2
                J = J + 1
            End If
        Else
            ReDim Preserve WorkLine(J) As CadLine
            WorkLine(J) = LineAry(I)
            J = J + 1
        End If
    Else
        ReDim Preserve WorkLine(J) As CadLine
        WorkLine(J) = LineAry(I)
        J = J + 1
    End If
Next I
'Copy The Array
ReDim LineAry(0) As CadLine
For I = 0 To UBound(WorkLine)
    ReDim Preserve LineAry(I) As CadLine
    LineAry(I) = WorkLine(I)
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Function cXLen(ChkLine As CadLine) As Single
    cXLen = Abs(ChkLine.P2.x - ChkLine.P1.x)
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Function cYLen(ChkLine As CadLine) As Single
    cYLen = Abs(ChkLine.P2.y - ChkLine.P1.y)
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Sub LineArrayToEdges(Canvas As PictureBox, LineAry() As CadLine, ByRef PointAry() As CadPoint, ByRef EdgeAry() As Edge)
On Local Error GoTo eTrap
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim x As Integer
I = UBound(PointAry) + 1
x = I
'If UBound(PointAry) > 0 Then x = UBound(PointAry) + 1
For I = 0 To UBound(LineAry)
    ReDim Preserve PointAry(x) As CadPoint
    PointAry(x) = LineAry(I).P1
    x = x + 1
    'ReDim Preserve PointAry(x) As POINTAPI
    'PointAry(x) = LineAry(i).P2
    'x = x + 1
Next I
'RemovePointDups PointAry()
For I = 0 To UBound(LineAry)
    ReDim Preserve EdgeAry(K) As Edge
    EdgeAry(K).Boundary = True
    EdgeAry(K).StartPt = FindPoint(PointAry(), LineAry(I).P1)
    EdgeAry(K).EndPt = FindPoint(PointAry(), LineAry(I).P2)
    '-*-DrawEdgeAPI Canvas, PointAry(), EdgeAry(k), vbBlue, 3
    '-*-DrawCadPoint Canvas, PointAry(EdgeAry(k).StartPt), , vbWhite, 2
    '-*-DrawCadPoint Canvas, PointAry(EdgeAry(k).EndPt), , vbRed, 2
    K = K + 1
Next I
Exit Sub
eTrap:
    I = 0
    Resume Next
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Sub DrawEdgeAPI(MyPic As PictureBox, PointAry() As CadPoint, MyEdge As Edge, Color As Long, Width As Integer, Optional LabelIt As Boolean)
Dim tW As Integer
Dim MPt As CadPoint
Dim L1 As CadLine
tW = MyPic.DrawWidth
MyPic.DrawWidth = Width
MyPic.DrawMode = 13
MyPic.Line (PointAry(MyEdge.StartPt).x, -PointAry(MyEdge.StartPt).y)-(PointAry(MyEdge.EndPt).x, -PointAry(MyEdge.EndPt).y), Color
MyPic.DrawWidth = tW
If LabelIt Then
    L1.P1 = PointAry(MyEdge.StartPt)
    L1.P2 = PointAry(MyEdge.EndPt)
    MPt = MidPoint(L1)
    MyPic.CurrentX = MPt.x
    MyPic.CurrentY = -MPt.y
    MyPic.Print "( " & MyEdge.left & " , " & MyEdge.Right & " )"
End If
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Function FindPoint(PointAry() As CadPoint, P1 As CadPoint) As Integer
Dim I As Integer
Dim P2 As CadPoint
Dim P3 As CadPoint
Dim Prec As String
For I = 0 To UBound(PointAry)
    If MPts(PointAry(I), P1) Then
        FindPoint = I
        Exit Function
    End If
Next I
Prec = "0.0000"
Do While Len(Prec) > 2
    P3.x = Format(P1.x, Prec)
    P3.y = Format(P1.y, Prec)
    For I = 0 To UBound(PointAry)
        P2.x = Format(PointAry(I).x, Prec)
        P2.y = Format(PointAry(I).y, Prec)
        If MPts(P2, P3) Then
            FindPoint = I
            Exit Function
        End If
    Next I
    Prec = left(Prec, Len(Prec) - 1)
Loop
'MsgBox "NotFound!!!"
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Function GetLeft(Points() As CadLine)
Dim I As Integer
Dim Xmin As Single
Xmin = 32000
For I = 0 To UBound(Points)
    If Points(I).P1.x < Xmin Then Xmin = Points(I).P1.x
    If Points(I).P2.x < Xmin Then Xmin = Points(I).P2.x
Next I
GetLeft = Xmin
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Sub FullDelaunay(Canvas As PictureBox, DLines() As CadLine)
GenMeshPoints Canvas, DLines()
BuildConformedMesh Canvas, DLines(), dPoints(), dEdges(), dTriangles()
Dim I As Integer
Dim J As Integer
Dim K As Integer

'UnifyNormals dPoints(), dTriangles()

'WHAT IS THIS FOR?
'ReDim Edges(0) As PointArray
'For i = 0 To UBound(dTriangles)
'    If dTriangles(i).Que > -1 Then
'        ReDim Preserve Edges(k) As PointArray
'        ReDim Edges(k).P(2) As CadPoint
'        For j = 0 To 2
'            Edges(k).P(j) = dPoints(dTriangles(i).Vertices(j))
'        Next j
'        k = k + 1
'    End If
'Next i
Canvas.Picture = LoadPicture()
For I = 0 To UBound(dTriangles)
    If dTriangles(I).Que > -1 Then
        DrawElementAPI Canvas, dPoints(), dTriangles(I), vbBlack, 1
    End If
Next I
'Canvas.Picture = Canvas.Image
'SavePointArray Edges()
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Sub DrawElementAPI(MyPic As PictureBox, PointAry() As CadPoint, ChkElement As Element, Color As Long, Width As Integer)
Dim tW As Integer
tW = MyPic.DrawWidth
MyPic.DrawWidth = Width
MyPic.Line (PointAry(ChkElement.Vertices(0)).x, -PointAry(ChkElement.Vertices(0)).y)-(PointAry(ChkElement.Vertices(1)).x, -PointAry(ChkElement.Vertices(1)).y), Color
MyPic.Line (PointAry(ChkElement.Vertices(1)).x, -PointAry(ChkElement.Vertices(1)).y)-(PointAry(ChkElement.Vertices(2)).x, -PointAry(ChkElement.Vertices(2)).y), Color
MyPic.Line (PointAry(ChkElement.Vertices(2)).x, -PointAry(ChkElement.Vertices(2)).y)-(PointAry(ChkElement.Vertices(0)).x, -PointAry(ChkElement.Vertices(0)).y), Color
MyPic.DrawWidth = tW
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Sub SplitTriangle(Canvas As PictureBox, ByRef Ref() As RefinedEdge, I As Integer, Min As Single, Force As Boolean)
Dim J As Integer
Dim a As Integer
Dim B As Integer
Dim c As Integer
Dim Tmod As Boolean
If I = 34 Then
    DoEvents
End If
For J = 0 To 2
    If Ref(I).Use = False Then Exit Sub
    Select Case J
        Case 0: a = 0: B = 1: c = 2
        Case 1: a = 1: B = 2: c = 0
        Case 2: a = 2: B = 0: c = 1
    End Select
    DrawCadLine Canvas, PtLine(Ref(I).P(a), Ref(I).P(B)), , vbYellow, 3
    DrawCadLine Canvas, PtLine(Ref(I).P(B), Ref(I).P(c)), , vbBlue, 3
    DrawCadLine Canvas, PtLine(Ref(I).P(c), Ref(I).P(a)), , vbBlue, 3
    If CheckTriAngle(Ref(I).P(a), Ref(I).P(B), Ref(I).P(c)) < 30 Then 'angle is too small  - we want to flip
        If Ref(I).Flipped = False Then
            If FlipTriangle(Canvas, Ref(), I, a, B, c, Min) Then GoTo ExitMe
        End If
    End If
    If (PtLen(Ref(I).P(a), Ref(I).P(B)) >= PtLen(Ref(I).P(a), Ref(I).P(c))) And (PtLen(Ref(I).P(a), Ref(I).P(B)) >= PtLen(Ref(I).P(c), Ref(I).P(B))) Then 'A-B is the longest edge
        If PtLen(Ref(I).P(a), Ref(I).P(B)) > Min Then 'the edge is too long and should be split
            If Ref(I).P(a).Layer.ID <> -100 Or Ref(I).P(B).Layer.ID <> -100 Then 'A-B is not the border
                CompleteSplit Canvas, Ref(), I, a, B, c, Min
                GoTo ExitMe
            End If
        End If
    End If
    DrawCadLine Canvas, PtLine(Ref(I).P(a), Ref(I).P(B)), , vbBlack, 3
    DrawCadLine Canvas, PtLine(Ref(I).P(B), Ref(I).P(c)), , vbBlack, 3
    DrawCadLine Canvas, PtLine(Ref(I).P(c), Ref(I).P(a)), , vbBlack, 3
Next J
Exit Sub
ExitMe:
DrawCadLine Canvas, PtLine(Ref(I).P(a), Ref(I).P(B)), , vbBlack, 3
DrawCadLine Canvas, PtLine(Ref(I).P(B), Ref(I).P(c)), , vbBlack, 3
DrawCadLine Canvas, PtLine(Ref(I).P(c), Ref(I).P(a)), , vbBlack, 3
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Sub UnifyNormals(PointAry() As CadPoint, ByRef ElementAry() As Element)
Dim I As Integer
Dim tPt As Integer
For I = 0 To UBound(ElementAry)
    If TriPtAngle(PointAry(ElementAry(I).Vertices(0)), PointAry(ElementAry(I).Vertices(1)), PointAry(ElementAry(I).Vertices(2))) < 180 Then
        tPt = ElementAry(I).Vertices(1)
        ElementAry(I).Vertices(1) = ElementAry(I).Vertices(2)
        ElementAry(I).Vertices(2) = tPt
    End If
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Sub BuildConformedMesh(Canvas As PictureBox, ByRef LineAry() As CadLine, ByRef PointAry() As CadPoint, ByRef EdgeAry() As Edge, ByRef ElementAry() As Element)
Dim I As Integer
Dim J As Integer
Dim ECount As Integer
ElementAry(0).Que = -1 ' we do this because we are skipping the first one (speed optimiztion)
I = 0
Do While I <= UBound(EdgeAry)
    If EdgeAry(0).Que <> -1 Then
        If EdgeAry(I).Right = 0 Then BuildElement Canvas, PointAry(), EdgeAry(), I, ElementAry(), 1, 0
        If EdgeAry(I).left = 0 Then BuildElement Canvas, PointAry(), EdgeAry(), I, ElementAry(), -1, 0
    End If
    I = I + 1
Loop
''ReDelaunay PointAry(), EdgeAry(), ElementAry(), 0
'frmEndCap.PicCap.Picture = LoadPicture()
For I = 0 To UBound(ElementAry)
    If ElementAry(I).Que > -1 Then DrawElementAPI Canvas, PointAry(), ElementAry(I), vbBlack, 1
Next I
DoEvents
CheckPSLGSegments Canvas, PointAry(), EdgeAry(), ElementAry()
DoEvents
RemoveUnwantedTriangles Canvas, PointAry(), EdgeAry(), ElementAry()
DoEvents
RemoveIntersectingEdges Canvas, PointAry(), EdgeAry(), ElementAry()
DoEvents
RemoveOutsideTriangles Canvas, PointAry(), EdgeAry(), ElementAry()
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Sub RemoveOutsideTriangles(Canvas As PictureBox, PointAry() As CadPoint, EdgeAry() As Edge, ByRef ElementAry() As Element)
Dim I As Integer
Dim J As Integer
Dim K As Integer
Canvas.Cls
For I = 1 To UBound(ElementAry)
    If ElementAry(I).Que > -1 Then
        For J = 0 To 2
            If EdgeAry(ElementAry(I).Edges(J)).Boundary = True Then
                '-*-DrawEdgeAPI Canvas, PointAry(), EdgeAry(ElementAry(i).Edges(j)), vbWhite, 1
                For K = 0 To 2
                    If EdgeAry(ElementAry(I).Edges(J)).StartPt <> ElementAry(I).Vertices(K) And EdgeAry(ElementAry(I).Edges(J)).EndPt <> ElementAry(I).Vertices(K) Then
                        If PtSide(PointAry(), EdgeAry(ElementAry(I).Edges(J)), PointAry(ElementAry(I).Vertices(K))) = 1 Then
                            ElementAry(I).Que = -1
                            GoTo NextElement
                        End If
                    End If
                Next K
            End If
        Next J
    End If
NextElement:
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Function PtSide(PointAry() As CadPoint, chkEdge As Edge, ChkPt As CadPoint) As Integer
Dim Angle As Single
Angle = TriPtAngle(PointAry(chkEdge.StartPt), PointAry(chkEdge.EndPt), ChkPt)
If Angle > 0 And Angle < 180 Then
    PtSide = 1
ElseIf Angle > 180 And Angle < 360 Then
    PtSide = -1
Else
    PtSide = 0
End If
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Sub RemoveIntersectingEdges(Canvas As PictureBox, ByRef PointAry() As CadPoint, ByRef EdgeAry() As Edge, ByRef ElementAry() As Element)
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim M As Integer
Dim P As Integer
Dim s As Integer
Dim x As Integer
Dim CPt As CadPoint
Dim XPt As CadPoint
Dim rad As Single
x = UBound(ElementAry)
Dim IsBorder As Boolean
Dim FixHole As Boolean
Dim IsInt As Boolean
x = UBound(EdgeAry)
For I = 0 To UBound(EdgeAry)
'For i = UBound(EdgeAry) To 0 Step -1
    If EdgeAry(I).Que > -1 And EdgeAry(I).Boundary = False Then
    'If EdgeAry(i).Que > -1 And EdgeAry(i).Boundary = False And cEdgeSlope(PointAry(), EdgeAry(i)) <> 0 And cEdgeSlope(PointAry(), EdgeAry(i)) <> 32000 And cEdgeSlope(PointAry(), EdgeAry(i)) <> -32000 Then
        EdgeAry(I).left = 0
        EdgeAry(I).Right = 0
        EdgeAry(I).Que = 0
        DrawEdgeAPI Canvas, PointAry(), EdgeAry(I), vbBlue, 2
        DoEvents
        For J = UBound(EdgeAry) To I + 1 Step -1 ' the latest edges will most likely intersect earlier ones
        'For j = i + 1 To UBound(EdgeAry)
        'For j = i - 1 To 0 Step -1
            'If EdgeAry(j).Que > -1 Then
                'We might as well test them all  . . the IF statment takes more time than the check!
                If IntersectTest(PointAry(), EdgeAry(), I, J) Then
                    EdgeAry(I).Que = -1
                    DrawEdgeAPI Canvas, PointAry(), EdgeAry(I), vbRed, 3
                    Exit For
                End If
            'End If
        Next J
    End If
Next I
x = UBound(ElementAry)
For I = 1 To UBound(ElementAry)
    If ElementAry(I).Que > -1 Then
        For J = 0 To 2
            If EdgeAry(ElementAry(I).Edges(J)).Que = -1 Then
                ElementAry(I).Que = -1
            Else
                DefineElementEdges PointAry(), EdgeAry(), ElementAry(I), I
            End If
        Next J
    ElseIf ElementAry(I).Que = -99 Then 'an 'outside element'
        For J = 0 To 2
            If EdgeAry(ElementAry(I).Edges(J)).Boundary = False Then
                EdgeAry(ElementAry(I).Edges(J)).Que = -1
            End If
        Next J
    End If
Next I
For I = 1 To UBound(ElementAry)
    If ElementAry(I).Que > -1 Then
        For J = 0 To 2
            EdgeAry(ElementAry(I).Edges(J)).Que = EdgeAry(ElementAry(I).Edges(J)).Que + 1
        Next J
    End If
Next I
ClearPointID PointAry()
x = UBound(EdgeAry)
For I = 0 To UBound(EdgeAry)
    If EdgeAry(I).Boundary Then
        If EdgeAry(I).Que < 1 Then
            PointAry(EdgeAry(I).StartPt).Layer.ID = 1
            PointAry(EdgeAry(I).EndPt).Layer.ID = 1
            EdgeAry(I).Que = 99
            EdgeAry(I).left = 0
            EdgeAry(I).Right = 0
        End If
    ElseIf EdgeAry(I).Que > -1 And EdgeAry(I).Que < 2 Then
        PointAry(EdgeAry(I).StartPt).Layer.ID = 1
        PointAry(EdgeAry(I).EndPt).Layer.ID = 1
        EdgeAry(I).Que = 99
        EdgeAry(I).left = 0
        EdgeAry(I).Right = 0
    End If
    If EdgeAry(I).Que <> 99 Then EdgeAry(I).Que = -1
Next I
For I = 0 To UBound(PointAry)
    If PointAry(I).Layer.ID = 1 Then DrawCadPoint Canvas, PointAry(I), , vbBlack, 4
Next I
ReDelaunay Canvas, PointAry(), EdgeAry(), ElementAry(), 1
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Sub ReDelaunay(Canvas As PictureBox, ByRef PointAry() As CadPoint, ByRef EdgeAry() As Edge, ByRef ElementAry() As Element, PtID As Integer)
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim M As Integer
Dim ECount As Integer
Dim E1 As Integer
Dim E2 As Integer
'Start The First one outside the loop
I = 0
Do While I <= UBound(EdgeAry)
    E1 = -1
    E2 = -1
    If EdgeAry(I).Que > -1 Then
        If EdgeAry(I).Right = 0 Then E1 = BuildElement(Canvas, PointAry(), EdgeAry(), I, ElementAry(), 1, PtID)
        If EdgeAry(I).left = 0 Then E2 = BuildElement(Canvas, PointAry(), EdgeAry(), I, ElementAry(), -1, PtID)
    End If
    If E1 > -1 Then
        DrawElementAPI Canvas, PointAry(), ElementAry(E1), vbMagenta, 1
        For J = 0 To UBound(ElementAry)
            If ElementAry(J).Que > -1 And J <> E1 Then
                For K = 0 To 2
                    For M = 0 To 2
                        If IntersectTest(PointAry(), EdgeAry(), ElementAry(E1).Edges(K), ElementAry(J).Edges(M)) Then
                            ElementAry(E1).Que = -1
                            GoTo Skip1
                        End If
                    Next M
                Next K
            End If
        Next J
        For J = 0 To UBound(EdgeAry)
            If EdgeAry(J).Boundary Then
                For K = 0 To 2
                    If IntersectTest(PointAry(), EdgeAry(), J, ElementAry(E1).Edges(K)) Then
                        ElementAry(E1).Que = -1
                        GoTo Skip1
                    End If
                Next K
            End If
        Next J
    End If
Skip1:
    If E2 > -1 Then
        DrawElementAPI Canvas, PointAry(), ElementAry(E2), vbMagenta, 1
        For J = 0 To UBound(ElementAry)
            If ElementAry(J).Que > -1 And J <> E2 Then
                For K = 0 To 2
                    For M = 0 To 2
                        If IntersectTest(PointAry(), EdgeAry(), ElementAry(E2).Edges(K), ElementAry(J).Edges(M)) Then
                            ElementAry(E2).Que = -1
                            GoTo Skip2
                        End If
                    Next M
                Next K
            End If
        Next J
        For J = 0 To UBound(EdgeAry)
            If EdgeAry(J).Boundary Then
                For K = 0 To 2
                    If IntersectTest(PointAry(), EdgeAry(), J, ElementAry(E2).Edges(K)) Then
                        ElementAry(E2).Que = -1
                        GoTo Skip2
                    End If
                Next K
            End If
        Next J
    End If
Skip2:
    I = I + 1
    '-*-DoEvents
Loop
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Sub ClearPointID(ByRef PointAry() As CadPoint)
Dim I As Integer
For I = 0 To UBound(PointAry)
    PointAry(I).Layer.ID = 0
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Sub DefineElementEdges(ByRef PointAry() As CadPoint, ByRef EdgeAry() As Edge, ByRef ChkElement As Element, EID As Integer)
Dim I As Integer
Dim J As Integer
Dim PtVal As Integer
For I = 0 To 2
    For J = 0 To 2
        If EdgeAry(ChkElement.Edges(I)).StartPt <> ChkElement.Vertices(J) And EdgeAry(ChkElement.Edges(I)).EndPt <> ChkElement.Vertices(J) Then
            PtVal = PtSide(PointAry(), EdgeAry(ChkElement.Edges(I)), PointAry(ChkElement.Vertices(J)))
            If PtVal = 1 Then
                If EdgeAry(ChkElement.Edges(I)).Right = 0 Then EdgeAry(ChkElement.Edges(I)).Right = EID
                If ChkElement.Que = -1 Then EdgeAry(ChkElement.Edges(I)).Right = 0
            ElseIf PtVal = -1 Then
                If EdgeAry(ChkElement.Edges(I)).left = 0 Then EdgeAry(ChkElement.Edges(I)).left = EID
                If ChkElement.Que = -1 Then EdgeAry(ChkElement.Edges(I)).left = 0
            End If
        End If
    Next J
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Function IntersectTest(PointAry() As CadPoint, EdgeAry() As Edge, Edge1 As Integer, Edge2 As Integer) As Boolean
If EdgeAry(Edge1).StartPt = EdgeAry(Edge2).StartPt Or EdgeAry(Edge1).StartPt = EdgeAry(Edge2).EndPt Or EdgeAry(Edge1).EndPt = EdgeAry(Edge2).EndPt Or EdgeAry(Edge1).EndPt = EdgeAry(Edge2).StartPt Then
    IntersectTest = False
    Exit Function
End If
Dim IntVal As Integer
Dim IntPt As CadPoint
Dim L1 As CadLine
Dim L2 As CadLine
L1.P1 = PointAry(EdgeAry(Edge1).StartPt)
L1.P2 = PointAry(EdgeAry(Edge1).EndPt)
L2.P1 = PointAry(EdgeAry(Edge2).StartPt)
L2.P2 = PointAry(EdgeAry(Edge2).EndPt)
IntVal = CheckIntersect(L1, L2, IntPt)
If IntVal = 3 Then
    IntersectTest = True
Else
    IntersectTest = False
End If
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Function CheckIntersect(Line1 As CadLine, Line2 As CadLine, ByRef ptIntersect As CadPoint) As Integer

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
            ptIntersect.x = Line1.P1.x + (dPctDelta1 * Delta(1).x)
            ptIntersect.y = Line1.P1.y + (dPctDelta1 * Delta(1).y)
        
        End If
        
        'Return the results.
        CheckIntersect = iReturn
        
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Function FaceCenter(MyFace As CadFace) As CadPoint
On Error Resume Next
Dim x As Single
Dim y As Single
Dim I As Integer
Dim L1 As CadLine
Dim L2 As CadLine
Dim L3 As CadLine
Dim IL1 As CadLine
Dim IL2 As CadLine
Dim IntPt As CadPoint
L1.P1 = MyFace.Vertex(0)
L1.P2 = MyFace.Vertex(1)
L2.P1 = MyFace.Vertex(0)
L2.P2 = MyFace.Vertex(2)
L3.P1 = MyFace.Vertex(1)
L3.P2 = MyFace.Vertex(2)
IL1.P1 = MidPoint(L1)
IL1.P2 = MidPoint(L2)
IL2.P1 = MyFace.Vertex(0)
IL2.P2 = MidPoint(L3)
CheckIntersect IL1, IL2, IntPt
FaceCenter = IntPt
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function 'All code researched and developed by Dave Andrews unless otherwise noted.


Sub RemoveUnwantedTriangles(Canvas As PictureBox, PointAry() As CadPoint, EdgeAry() As Edge, ByRef ElementAry() As Element)
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim x As Integer
Dim Bound  As Boolean
Dim Adjusted As Boolean
'frmEndCap.PicCap.Cls
x = UBound(ElementAry)
For I = 1 To UBound(ElementAry)
    If ElementAry(I).Que > -1 Then
        For J = 0 To 2
            If EdgeAry(ElementAry(I).Edges(J)).Boundary = True Then
                For K = 0 To 2
                    If EdgeAry(ElementAry(I).Edges(J)).StartPt <> ElementAry(I).Vertices(K) And EdgeAry(ElementAry(I).Edges(J)).EndPt <> ElementAry(I).Vertices(K) Then
                        If PtSide(PointAry(), EdgeAry(ElementAry(I).Edges(J)), PointAry(ElementAry(I).Vertices(K))) = 1 Then
                            'DrawElementAPI Canvas, PointAry(), ElementAry(i), vbWhite, 1
                            ElementAry(I).Que = -99
                        End If
                    End If
                Next K
            End If
        Next J
    End If
Next I
'-*-DoEvents
Adjusted = True
'''Exit Sub'''
'-----Virus------------
Do While Adjusted
    Adjusted = False
    For I = 1 To UBound(ElementAry)
        If ElementAry(I).Que = -99 Then ' we're looking at one of the 'outside' elements
            For J = 0 To 2
                If EdgeAry(ElementAry(I).Edges(J)).Boundary = False Then
                    Bound = BoundElement(EdgeAry(), ElementAry(EdgeAry(ElementAry(I).Edges(J)).left))
                    If Not Bound And ElementAry(EdgeAry(ElementAry(I).Edges(J)).left).Que > -1 Then
                        ElementAry(EdgeAry(ElementAry(I).Edges(J)).left).Que = -99
                        'DrawElementAPI frmEndCap.PicCap, PointAry(), ElementAry(EdgeAry(ElementAry(i).Edges(j)).left), vbRed, 2
                        Adjusted = True
                    End If
                    Bound = BoundElement(EdgeAry(), ElementAry(EdgeAry(ElementAry(I).Edges(J)).Right))
                    If Not Bound And ElementAry(EdgeAry(ElementAry(I).Edges(J)).Right).Que > -1 Then
                        ElementAry(EdgeAry(ElementAry(I).Edges(J)).Right).Que = -99
                        'DrawElementAPI frmEndCap.PicCap, PointAry(), ElementAry(EdgeAry(ElementAry(i).Edges(j)).Right), vbRed, 2
                        Adjusted = True
                    End If
                End If
            Next J
        End If
    Next I
Loop
'-*-DoEvents
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Function BoundElement(EdgeAry() As Edge, ChkElement As Element) As Boolean
Dim I As Integer
For I = 0 To 2
    If EdgeAry(ChkElement.Edges(I)).Boundary Then
        BoundElement = True
        Exit Function
    End If
Next I
BoundElement = False
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Sub CheckPSLGSegments(Canvas As PictureBox, ByRef PointAry() As CadPoint, ByRef EdgeAry() As Edge, ByRef ElementAry() As Element)
Dim I As Integer
Dim J As Integer
Dim x As Integer
Dim FoundInt As Boolean
Dim tColor As Long
ClearPointID PointAry()
'frmEndCap.PicCap.Cls
x = UBound(EdgeAry)
For I = 0 To UBound(EdgeAry)
    FoundInt = False
    If EdgeAry(I).Boundary = True Then  'and EdgeAry(i).Que <> -1 Then
        DrawEdgeAPI Canvas, PointAry(), EdgeAry(I), vbYellow, 3
        DoEvents
        For J = 0 To UBound(EdgeAry)
            If EdgeAry(J).Que <> -1 And J <> I And IntersectTest(PointAry(), EdgeAry(), I, J) = True Then
                'we found edges that properly intersect our PSGL border
                EdgeAry(J).Que = -1
                '''PointAry(EdgeAry(j).StartPt).ID = -(PtSide(PointAry(), EdgeAry(i), PointAry(EdgeAry(j).StartPt)))
                '''PointAry(EdgeAry(j).EndPt).ID = -(PtSide(PointAry(), EdgeAry(i), PointAry(EdgeAry(j).EndPt)))
                PointAry(EdgeAry(J).StartPt).Layer.ID = 1
                PointAry(EdgeAry(J).EndPt).Layer.ID = 1
                '-*-DrawEdgeAPI Canvas, PointAry(), EdgeAry(j), vbMagenta, 1
                FoundInt = True
            End If
        Next J
        If FoundInt Then
            ClearPointID PointAry()
            PointAry(EdgeAry(I).StartPt).Layer.ID = 1
            PointAry(EdgeAry(I).EndPt).Layer.ID = 1
            EdgeAry(I).Right = 0
            'ReDelaunay PointAry(), EdgeAry(), ElementAry(), 1
            '------------------
            '''PointAry(EdgeAry(i).StartPt).ID = -1
            '''pointAry(EdgeAry(i).EndPt).ID = -1
            EdgeAry(I).left = 0
            '''ReDelaunay PointAry(), EdgeAry(), ElementAry(), -1
        Else
            If EdgeAry(I).Que <> -1 Then EdgeAry(I).Que = -50
        End If
    Else
        If EdgeAry(I).Que <> -1 Then EdgeAry(I).Que = -50
    End If
Next I
'Now we have to remove the elements that use those 'marked edges'
For I = 1 To UBound(ElementAry)
    If EdgeAry(ElementAry(I).Edges(0)).Que = -1 Then ElementAry(I).Que = -1
    If EdgeAry(ElementAry(I).Edges(1)).Que = -1 Then ElementAry(I).Que = -1
    If EdgeAry(ElementAry(I).Edges(2)).Que = -1 Then ElementAry(I).Que = -1
    'DefineElementEdges PointAry(), EdgeAry(), ElementAry(i), i
Next I
'Regenerate the mesh
ReDelaunay Canvas, PointAry(), EdgeAry(), ElementAry(), 1
For I = 0 To UBound(EdgeAry)
    If EdgeAry(I).Que = -50 Then EdgeAry(I).Que = 0
Next I

'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Function BuildElement(Canvas As PictureBox, ByRef PointAry() As CadPoint, ByRef EdgeAry() As Edge, EdgeNum As Integer, ByRef ElementAry() As Element, Side As Integer, PtID As Integer) As Integer
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim PIndex As Integer
Dim StartPt As Integer
Dim EndPt As Integer
Dim CVal As Integer
Dim CPt As CadPoint
Dim rad As Single
Dim GoodCircle As Boolean
Dim GoodPoint As Boolean
Dim NewEdge As Integer
Dim LExt As CadLine
Dim tArc As CadArc

GoodCircle = False

For I = 0 To UBound(PointAry)
    If Side = PtSide(PointAry(), EdgeAry(EdgeNum), PointAry(I)) And PointAry(I).Layer.ID = PtID Then GoodPoint = True Else GoodPoint = False
CheckDelaunay:        'we are looking at a point other than the two in the specified edge
        If I <> EdgeAry(EdgeNum).StartPt And I <> EdgeAry(EdgeNum).EndPt And GoodPoint Then
        CVal = IsDelaunayCenter(PointAry(EdgeAry(EdgeNum).StartPt), PointAry(EdgeAry(EdgeNum).EndPt), PointAry(I), CPt)
        If CVal > 0 Then 'we have a circumcircle!
            rad = PtLen(PointAry(I), CPt)
            'now we check to see if there are any points in our circumcircle
            GoodCircle = True
            For J = I + 1 To UBound(PointAry)
                'we only have to start at i + 1 because everyhting beforehand was already checked
                If PtLen(PointAry(J), CPt) <= rad Then
                    'we found a point inside our circumcircle
                        '----------------
                    'we should check this one to see if it's on the correct side of the edge
                    If Side = PtSide(PointAry(), EdgeAry(EdgeNum), PointAry(J)) And PointAry(J).Layer.ID = PtID Then GoodPoint = True Else GoodPoint = False
                    If J <> EdgeAry(EdgeNum).StartPt And J <> EdgeAry(EdgeNum).EndPt And GoodPoint Then
                        'we are looking at a point, not part of our potential element
                        GoodCircle = False
                        I = J
                        Exit For
                    End If
                End If
            Next J
            If GoodCircle Then
                PIndex = I
                Exit For ' we got our circumcircle
            Else
                If I <= UBound(PointAry) Then
                    GoTo CheckDelaunay
                Else
                    Exit For
                End If
            End If
        End If
    End If
Next I
If GoodCircle Then
    'we got our circumcircle
    'otherwise, if you get this far, and have no goodcircle, there must be
    'no points found on that side of the given edge
    '----------------------------
    Canvas.Cls
    Canvas.DrawWidth = 3
    Canvas.PSet (PointAry(EdgeAry(EdgeNum).StartPt).x, -PointAry(EdgeAry(EdgeNum).StartPt).y), vbRed
    Canvas.PSet (PointAry(EdgeAry(EdgeNum).EndPt).x, -PointAry(EdgeAry(EdgeNum).EndPt).y), vbWhite
    Canvas.PSet (PointAry(I).x, -PointAry(I).y), vbBlue
    tArc.Angle1 = 0
    tArc.Angle2 = 360
    tArc.Center = CPt
    tArc.Radius = rad
    DrawCadArc Canvas, tArc, , vbGreen, 2
    DoEvents
    '------------------------------
    'MsgBox EdgeNum & " , " & Side
    '------Add The Points to the element array
    '----------although, we are base0, the real start of this array is at 1
    K = UBound(ElementAry) + 1
    ReDim Preserve ElementAry(K) As Element
    ElementAry(K).Vertices(0) = EdgeAry(EdgeNum).StartPt
    ElementAry(K).Vertices(1) = EdgeAry(EdgeNum).EndPt
    ElementAry(K).Vertices(2) = PIndex 'new point
    'Now we have so look through the array of edegs to see if the two
    'new edges to be created are already in the database
    ElementAry(K).Edges(0) = EdgeNum
    EdgeAry(EdgeNum).Que = 0
    StartPt = EdgeAry(EdgeNum).StartPt
    EndPt = EdgeAry(EdgeNum).EndPt
    '---------------EndPt of orig edge---------
    NewEdge = EdgeInArray(EdgeAry(), EndPt, PIndex)
    If NewEdge = -1 Then 'it needs to be added / isn't in the array
        NewEdge = AddEdge(EdgeAry(), EndPt, PIndex)
    End If
    ElementAry(K).Edges(1) = NewEdge
    EdgeAry(NewEdge).Que = 0
    '--------StartPt of orig edge---------
    NewEdge = EdgeInArray(EdgeAry(), PIndex, StartPt)
    If NewEdge = -1 Then 'it needs to be added / isn't in the array
        NewEdge = AddEdge(EdgeAry(), PIndex, StartPt)
    End If
    ElementAry(K).Edges(2) = NewEdge
    EdgeAry(NewEdge).Que = 0
    '-----------------------------------------------
    'Now we set the right/left triangle properties of the edges of the new element
    DefineElementEdges PointAry(), EdgeAry(), ElementAry(K), K
    'If CheckElementIntersect(PointAry(), EdgeAry(), ElementAry(k)) Then
    '    ElementAry(k).Que = -1
    '    BuildElement = -1
    '    Exit Function
    'Else
        ElementAry(K).Que = 0
    'End If
    BuildElement = K 'pass back the new count of the elementary
Else
    BuildElement = -1
End If
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Function IsDelaunayCenter(P1 As CadPoint, P2 As CadPoint, P3 As CadPoint, ByRef MyPoint As CadPoint) As Integer
Dim L1 As CadLine
Dim L2 As CadLine
Dim L3 As CadLine
Dim Mid1 As CadLine
Dim Mid2 As CadLine
Dim Mid3 As CadLine
L1.P1 = P1
L1.P2 = P2
L2.P1 = P2
L2.P2 = P3
L3.P1 = P3
L3.P2 = P1
Mid1 = cAngLine(cAngle(L1) + 90, MidPoint(L1), 50, True)
Mid2 = cAngLine(cAngle(L2) + 90, MidPoint(L2), 50, True)
Mid3 = cAngLine(cAngle(L3) + 90, MidPoint(L3), 50, True)
IsDelaunayCenter = CheckIntersect(Mid1, Mid2, MyPoint)
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Function AddEdge(ByRef EdgeAry() As Edge, Index1 As Integer, Index2 As Integer) As Integer
Dim I As Integer
I = UBound(EdgeAry) + 1
ReDim Preserve EdgeAry(I) As Edge
EdgeAry(I).StartPt = Index1
EdgeAry(I).EndPt = Index2
AddEdge = I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Function EdgeInArray(EdgeAry() As Edge, Index1 As Integer, Index2 As Integer) As Integer
Dim I As Integer
For I = 0 To UBound(EdgeAry)
    If EdgeAry(I).StartPt = Index1 And EdgeAry(I).EndPt = Index2 Then
        EdgeInArray = I
        Exit Function
    End If
    If EdgeAry(I).StartPt = Index2 And EdgeAry(I).EndPt = Index1 Then
        EdgeInArray = I
        Exit Function
    End If
Next I
EdgeInArray = -1
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Function GetLower(Points() As CadLine)
Dim I As Integer
Dim Ymin As Single
Ymin = 32000
For I = 0 To UBound(Points)
    If Points(I).P1.y < Ymin Then Ymin = Points(I).P1.y
    If Points(I).P2.y < Ymin Then Ymin = Points(I).P2.y
Next I
GetLower = Ymin
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Function GetRight(Points() As CadLine)
Dim I As Integer
Dim Xmax As Single
Xmax = -32000
For I = 0 To UBound(Points)
    If Points(I).P1.x > Xmax Then Xmax = Points(I).P1.x
    If Points(I).P2.x > Xmax Then Xmax = Points(I).P2.x
Next I
GetRight = Xmax
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Sub UnifyFaces(ByRef FaceAry() As CadFace, CCW As Boolean)
Dim I As Integer
Dim tPt As CadPoint
For I = 0 To UBound(FaceAry)
    If CCW Then
        If TriPtAngle(FaceAry(I).Vertex(0), FaceAry(I).Vertex(1), FaceAry(I).Vertex(2)) > 180 Then
            tPt = FaceAry(I).Vertex(1)
            FaceAry(I).Vertex(1) = FaceAry(I).Vertex(2)
            FaceAry(I).Vertex(2) = tPt
        End If
    Else
        If TriPtAngle(FaceAry(I).Vertex(0), FaceAry(I).Vertex(1), FaceAry(I).Vertex(2)) < 180 Then
            tPt = FaceAry(I).Vertex(1)
            FaceAry(I).Vertex(1) = FaceAry(I).Vertex(2)
            FaceAry(I).Vertex(2) = tPt
        End If
    End If
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub 'All code researched and developed by Dave Andrews unless otherwise noted.

Function GetUpper(Points() As CadLine)
Dim I As Integer
Dim Ymax As Single
Ymax = -32000
For I = 0 To UBound(Points)
    If Points(I).P1.y > Ymax Then Ymax = Points(I).P1.y
    If Points(I).P2.y > Ymax Then Ymax = Points(I).P2.y
Next I
GetUpper = Ymax
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function

Function FindMin(DLines() As CadLine)
Dim I As Integer
Dim Min As Single
Min = 32000
For I = 0 To UBound(DLines)
    If LineLen(DLines(I)) < Min Then Min = LineLen(DLines(I))
Next I
FindMin = Min
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function


Function FindMax(DLines() As CadLine)
Dim I As Integer
Dim Max As Single
For I = 0 To UBound(DLines)
    If LineLen(DLines(I)) > Max Then Max = LineLen(DLines(I))
Next I
FindMax = Max
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Sub GeometryToLines(MyGeo As Geometry, DLines() As CadLine)
On Local Error GoTo eTrap
Dim I As Integer
Dim J As Integer
Dim tPoly As CadPolyLine
Erase DLines()
'------------------ARCS-----------------------
J = UBound(MyGeo.Arcs)
For I = 0 To J
    tPoly = ArcToPolyLine(MyGeo.Arcs(I))
    PolyLineToLines tPoly, DLines()
Next I
'------------------ELLIPSES-------------------
J = UBound(MyGeo.Ellipses)
For I = 0 To J
    tPoly = EllipseToPolyLine(MyGeo.Ellipses(I))
    PolyLineToLines tPoly, DLines()
Next I
'------------------SPLINES---------------------
J = UBound(MyGeo.Splines)
For I = 0 To J
    tPoly = SplineToPolyLine(MyGeo.Splines(I))
    PolyLineToLines tPoly, DLines()
Next I
'-------------------POLYLINES------------------
J = UBound(MyGeo.PolyLines)
For I = 0 To J
    PolyLineToLines MyGeo.PolyLines(I), DLines()
Next I
'-------------------LINES----------------------
AddLines MyGeo.Lines(), DLines()
'----------------------------------------------
Exit Sub
eTrap:
    J = -1
    Resume Next
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Sub OrderLines(DLines() As CadLine, Precision As String)
Dim I As Long
Dim J As Long
Dim Count As Long
Dim K As Integer
Dim junk As Single
Dim LTemp As CadLine
Dim Xtemp As Single
Dim Ytemp As Single
Dim ILoop As Integer
Dim Done As Boolean
Dim NPre As String
NPre = Precision
ReStart: 'we come here if we need to adjust the precision
K = UBound(DLines)
ReDim OArray(0) As OType
J = 0: I = 0: Count = 0: ILoop = 0
'get starts and ends
For I = 0 To K
    ReDim Preserve OArray(J) As OType
    OArray(J).ID = DLines(I).Layer.ID
    OArray(J).Start = DLines(I)
    Count = 0
    Do While OArray(J).ID = DLines(I).Layer.ID
        Count = Count + 1
        I = I + 1
        If I > UBound(DLines) Then Exit Do
    Loop
    I = I - 1
    OArray(J).End = DLines(I)
    OArray(J).Reverse = False
    OArray(J).Count = Count
    J = J + 1
Next I
K = UBound(OArray)
I = 0
Count = 1
Do While Not Done
    For J = 1 To K
        If OArray(J).Order = 0 Then
            If (Format(OArray(I).End.P2.x, NPre) = Format(OArray(J).Start.P1.x, NPre)) And (Format(OArray(I).End.P2.y, NPre) = Format(OArray(J).Start.P1.y, NPre)) Then
                OArray(J).Reverse = False
                OArray(J).Order = Count
                I = J
                Count = Count + 1
                ILoop = 0
                Exit For
            ElseIf (Format(OArray(I).End.P2.x, NPre) = Format(OArray(J).End.P2.x, NPre)) And (Format(OArray(I).End.P2.y, NPre) = Format(OArray(J).End.P2.y, NPre)) Then
                OArray(J).Reverse = True
                LTemp = OArray(J).Start
                OArray(J).Start = RevLine(OArray(J).End)
                OArray(J).End = RevLine(LTemp)
                OArray(J).Order = Count
                I = J
                Count = Count + 1
                ILoop = 0
                Exit For
            Else
                If ILoop > K Then
                    'Lower the precision
                    If Len(NPre) < 2 Then
                        'MsgBox "Can't order the array" & Chr$(10) & "Please adjust the precsision."
                        Exit Sub
                    Else
                        NPre = left(NPre, Len(NPre) - 1)
                        GoTo ReStart
                    End If
                End If
                ILoop = ILoop + 1
            End If
        End If
    Next J
    'Done = True
    If Count > K Then Done = True
Loop
ReDim AryCopy(0) As CadLine
Done = False
Count = 0
K = 0
Do While Not Done
    For I = 0 To UBound(OArray)
        If OArray(I).Order = Count Then Exit For
    Next I
    If OArray(I).Reverse Then
        For J = OArray(I).ID + (OArray(I).Count - 1) To OArray(I).ID Step -1
            ReDim Preserve AryCopy(K) As CadLine
            AryCopy(K) = RevLine(DLines(J))
            K = K + 1
        Next J
    Else
        For J = OArray(I).ID To OArray(I).ID + (OArray(I).Count - 1)
            ReDim Preserve AryCopy(K) As CadLine
            AryCopy(K) = DLines(J)
            K = K + 1
        Next J
    End If
    Count = Count + 1
    If Count > UBound(OArray) Then Done = True
Loop
For I = 0 To UBound(DLines)
    DLines(I) = AryCopy(I)
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub

Sub RemoveNullLines(ByRef DLines() As CadLine)
Dim I As Integer
Dim K As Integer
ReDim AryCopy(0) As CadLine
For I = 0 To UBound(DLines)
    If Not MPts(DLines(I).P1, DLines(I).P2) Then
        ReDim Preserve AryCopy(K) As CadLine
        AryCopy(K) = DLines(I)
        K = K + 1
    End If
Next I
ReDim DLines(0) As CadLine
For I = 0 To UBound(AryCopy)
    ReDim Preserve DLines(I) As CadLine
    DLines(I) = AryCopy(I)
Next I
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub
Function MPts(P1 As CadPoint, P2 As CadPoint) As Boolean
If P1.x = P2.x And P1.y = P2.y Then
    MPts = True
Else
    MPts = False
End If
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Function
Sub PolyLineToLines(MyPoly As CadPolyLine, MyLines() As CadLine)
On Local Error GoTo eTrap
Dim I As Integer
Dim J As Integer
J = UBound(MyLines) + 1
For I = 0 To UBound(MyPoly.Vertex) - 1
    ReDim Preserve MyLines(J + I)
    MyLines(J + I).P1 = MyPoly.Vertex(I)
    MyLines(J + I).P2 = MyPoly.Vertex(I + 1)
Next I

Exit Sub
eTrap:
    J = 0
    Resume Next
'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub


Sub SplitLines(DLines() As CadLine, Max As Single)
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim tArray() As CadLine
Dim sPts() As CadPoint
Dim sAngles() As Single
Dim sCount As Integer
For I = 0 To UBound(DLines)
    sCount = Fix(LineLen(DLines(I)) / Max)
    If sCount > 0 Then
        SpaceLine DLines(I), DLines(I).P1, Max, sCount, DLines(I).P2, sPts(), sAngles()
        For J = 0 To UBound(sPts) - 1
            ReDim Preserve tArray(K)
            tArray(K).Layer = DLines(I).Layer
            tArray(K).P1 = sPts(J)
            tArray(K).P2 = sPts(J + 1)
            K = K + 1
        Next J
        ReDim Preserve tArray(K)
        tArray(K).Layer = DLines(I).Layer
        tArray(K).P1 = sPts(UBound(sPts))
        tArray(K).P2 = DLines(I).P2
        K = K + 1
    Else
        ReDim Preserve tArray(K)
        tArray(K) = DLines(I)
        K = K + 1
    End If
Next I
ReDim DLines(UBound(tArray))
For I = 0 To UBound(tArray)
    DLines(I) = tArray(I)
Next I

'All Code Researched and Developed by Dave Andrews unless otherwise noted
End Sub


