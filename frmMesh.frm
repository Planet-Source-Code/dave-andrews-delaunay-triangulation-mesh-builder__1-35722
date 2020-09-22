VERSION 5.00
Begin VB.Form frmMesh 
   Caption         =   "MeshMaker"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9225
   Icon            =   "frmMesh.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   615
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   6120
      Left            =   7485
      ScaleHeight     =   406
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   114
      TabIndex        =   1
      Top             =   0
      Width           =   1740
      Begin VB.CommandButton cmdClearMesh 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Clear Mesh"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton cmdSaveMesh 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save Mesh"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdImport 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DXF Import"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoad 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Load"
         Height          =   255
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H00FFC0C0&
         Caption         =   "?"
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFC0C0&
         Caption         =   "New / Clear"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Delete Selection"
         Height          =   255
         Index           =   9
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Delete"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select Geometry"
         Height          =   255
         Index           =   8
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Select"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Set Mesh Center"
         Height          =   255
         Index           =   10
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Set Mesh Center"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtGridY 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Text            =   "5"
         Top             =   5760
         Width           =   495
      End
      Begin VB.TextBox txtGridX 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   600
         TabIndex        =   32
         Text            =   "5"
         Top             =   5760
         Width           =   495
      End
      Begin VB.CheckBox chkGrid 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Grid"
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5760
         Value           =   1  'Checked
         Width           =   540
      End
      Begin VB.CheckBox chkSnapGrid 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Snap Grid"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CheckBox chkSnapEdit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Snap Edit Points"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox chkSnapIntersect 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Snap Intersection"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CheckBox chkSnapEndMid 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Snap End/Mid"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   4800
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton cmdZoomOut 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1200
         Picture         =   "frmMesh.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdZoomIn 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   840
         Picture         =   "frmMesh.frx":0371
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1920
         Width           =   375
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   12
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":03D9
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Magnify"
         Top             =   1920
         Width           =   375
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   11
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":043F
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Pan"
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtExtrude 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Text            =   "5"
         Top             =   3960
         Width           =   615
      End
      Begin VB.CheckBox chkChain 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         DownPicture     =   "frmMesh.frx":04AE
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   285
         Picture         =   "frmMesh.frx":04FA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1365
         Width           =   195
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   7
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":0546
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Open Ellipse"
         Top             =   1560
         Width           =   375
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   6
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":0606
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ellipse"
         Top             =   1560
         Width           =   375
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   5
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":069A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Arc"
         Top             =   1560
         Width           =   375
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   4
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":076B
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Circle"
         Top             =   1560
         Width           =   375
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   3
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":084A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "3 Point Arc"
         Top             =   1200
         Width           =   375
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   2
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Spline"
         Top             =   1200
         Width           =   375
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   1
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":0A3B
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Polyline"
         Top             =   1200
         Width           =   375
      End
      Begin VB.OptionButton optCommand 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   0
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMesh.frx":0AA9
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Line"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtDTol 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "5"
         Top             =   3360
         Width           =   615
      End
      Begin VB.CommandButton cmdMesh 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Show Mesh"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelaunay 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Triangulate"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Extrude:"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tolerance:"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   855
      End
   End
   Begin VB.PictureBox picDraw 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      Height          =   6120
      Left            =   0
      MouseIcon       =   "frmMesh.frx":0B31
      MousePointer    =   2  'Cross
      ScaleHeight     =   121.322
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type FHArray
    pts() As CadPoint
End Type
Dim DrawGeo() As Geometry
Dim FreeHand() As FHArray
Dim FloodCtr As CadPoint
Dim SnapPts() As CadPoint
Dim Command As String
Dim Stage As Integer
Dim StageMax As Integer
Dim StartStage As Boolean
Dim Layers() As CADLayer
Dim cView As Integer
Dim cLayer As Integer
Dim CurFileName As String
Dim ChainMode As Boolean
Dim OutLines() As CadLine

'------------Temporary Vars for Creation
Dim tAngle As Single
Dim tDist As Single
Dim VisPoint As CadPoint
Dim tLine As CadLine
Dim vCount As Integer
Dim CurX As Single
Dim CurY As Single
Dim StartMove As Boolean
Dim Editing As Boolean
Dim tBool As Boolean
Dim DPoint As CadPoint
Dim PivotPt As CadPoint
Dim FStartX As Single
Dim FStartY As Single
Dim FScaleX As Single
Dim FScaleY As Single
Dim FEdge As Integer
'------------Creation Elements------
Dim CurPoint As CadPoint
Dim CurLine As CadLine
Dim CurArc As CadArc
Dim CurEllipse As CadEllipse
Dim CurOpenEllipse As CadEllipse
Dim CurSpline As CadSpline
Dim CurPolyLine As CadPolyLine
Dim CurText As CadText
Dim Cur3Pt(2) As CadPoint
Dim curInsert As CadInsert
Dim CurLayer As CADLayer
Dim CurFace As CadFace
'-----------Selection,Grids, Snap and ZOom--------------------------
Dim Selection() As SelSet
Dim Closest As SelSet
Dim SelLine As CadLine
Dim ZoomLine As CadLine
Dim SelOn As Boolean
Dim GridStart As CadPoint
Dim gScaleX As Single
Dim gScaleY As Single
Dim StartGrid As Boolean
Dim sX As Single
Dim sY As Single
Dim UseSnap As Boolean
Dim DoSel As Boolean

'--------------MOdification Vars---------------------
Dim ModSel(2) As SelSet
Dim ModPt(3) As CadPoint
Dim ModEnd(2) As Integer
Dim TempPts() As CadPoint
Dim SelPts() As CadPoint
Sub CreateMesh(Optional SaveName As String)
Dim MForm As New frmFull3D
MForm.Show
Dim Center As CadPoint
Dim f As Direct3DRMFace2
Dim Part As Direct3DRMMeshBuilder3
Dim I As Integer
Dim cx As Single
Dim cy As Single
Dim ZDepth As Single
Dim FCount As Integer
ZDepth = Val(txtExtrude.Text)
Center = MidPoint(GetExtents(DrawGeo(), cView))
cx = Center.X
cy = Center.Y
Set Part = MForm.RMC.mDrm.CreateMeshBuilder
For I = 0 To UBound(DrawGeo(cView).Faces)
    'Frontside
    Set f = MForm.RMC.mDrm.CreateFace
    f.AddVertex DrawGeo(cView).Faces(I).Vertex(0).X - cx, DrawGeo(cView).Faces(I).Vertex(0).Y - cy, 0
    f.AddVertex DrawGeo(cView).Faces(I).Vertex(2).X - cx, DrawGeo(cView).Faces(I).Vertex(2).Y - cy, 0
    f.AddVertex DrawGeo(cView).Faces(I).Vertex(1).X - cx, DrawGeo(cView).Faces(I).Vertex(1).Y - cy, 0
    f.SetColorRGB 1, 0, 0
    Part.AddFace f
    'Backside
    Set f = MForm.RMC.mDrm.CreateFace
    f.AddVertex DrawGeo(cView).Faces(I).Vertex(0).X - cx, DrawGeo(cView).Faces(I).Vertex(0).Y - cy, ZDepth
    f.AddVertex DrawGeo(cView).Faces(I).Vertex(1).X - cx, DrawGeo(cView).Faces(I).Vertex(1).Y - cy, ZDepth
    f.AddVertex DrawGeo(cView).Faces(I).Vertex(2).X - cx, DrawGeo(cView).Faces(I).Vertex(2).Y - cy, ZDepth
    f.SetColorRGB 0, 0, 1
    Part.AddFace f
    MForm.Caption = FCount & " / " & (UBound(DrawGeo(cView).Faces) * 2) + UBound(OutLines)
    FCount = FCount + 2
Next I
'Sides
For I = 0 To UBound(OutLines)
    Set f = MForm.RMC.mDrm.CreateFace
    f.AddVertex OutLines(I).P1.X - cx, OutLines(I).P1.Y - cy, 0
    f.AddVertex OutLines(I).P1.X - cx, OutLines(I).P1.Y - cy, ZDepth
    f.AddVertex OutLines(I).P2.X - cx, OutLines(I).P2.Y - cy, ZDepth
    f.AddVertex OutLines(I).P2.X - cx, OutLines(I).P2.Y - cy, 0
    f.SetColorRGB 1, 1, 0
    Part.AddFace f
    MForm.Caption = FCount & " / " & (UBound(DrawGeo(cView).Faces) * 2) + UBound(OutLines)
    FCount = FCount + 1
Next I
'-------------------------
MForm.RMC.mFrO.AddVisual Part
Set f = Nothing
MForm.Show
'-----------------Zoom Out---------------
Dim xVect As D3DVECTOR
MForm.RMC.mFrO.GetPosition Nothing, xVect
xVect.z = 40
MForm.RMC.mVpt.SetBack 500
'MForm.RMC.mVpt.SetFront -500
MForm.RMC.mFrO.SetPosition Nothing, xVect.X, xVect.Y, xVect.z
MForm.RMC.Update
MForm.Caption = (UBound(DrawGeo(cView).Faces) * 2) + UBound(OutLines) & " Faces"
MForm.SetZoom
If SaveName <> "" Then
    Part.Save SaveName, D3DRMXOF_BINARY, D3DRMXOFSAVE_ALL
    MsgBox "Saved"
End If
Set Part = Nothing
End Sub

Sub Redraw()
On Error Resume Next
Dim I As Integer
picDraw.Cls
DrawCadGeo picDraw, DrawGeo(), 0
DrawCadPoint picDraw, DPoint, 13, vbGreen, 5
For I = 0 To UBound(Selection)
    ShowSelection picDraw, DrawGeo(), cView, Selection(I), vbWhite
Next I
If chkGrid.Value = vbChecked Then DrawGrid
GetSnap
End Sub

Private Sub chkChain_Click()
If chkChain.Value = vbChecked Then
    ChainMode = True
Else
    ChainMode = False
End If
End Sub

Private Sub chkGrid_Click()
Redraw
End Sub

Sub DrawGrid()
Dim I As Single
Dim J As Single
picDraw.DrawWidth = 1
picDraw.DrawMode = 6
picDraw.DrawStyle = 0
picDraw.PSet (0, 0), vbBlack
gScaleX = CSng(txtGridX.Text)
gScaleY = CSng(txtGridY.Text)
If picDraw.ScaleX(gScaleX, vbUser, vbPixels) < 2 Then Exit Sub
If picDraw.ScaleY(gScaleY, vbUser, vbPixels) < 2 Then Exit Sub
picDraw.DrawWidth = 1
picDraw.DrawMode = 13
picDraw.DrawStyle = 0

For I = GridStart.X To picDraw.ScaleWidth + picDraw.ScaleLeft Step gScaleX
    For J = GridStart.Y To picDraw.ScaleHeight + picDraw.ScaleTop Step gScaleY
        picDraw.PSet (I, J), vbBlack
    Next J
Next I
For I = GridStart.X To picDraw.ScaleLeft Step -gScaleX
    For J = GridStart.Y To picDraw.ScaleTop Step -gScaleY
        picDraw.PSet (I, J), vbBlack
    Next J
Next I
For I = GridStart.X + CSng(txtGridX) To picDraw.ScaleWidth + picDraw.ScaleLeft Step gScaleX
    For J = GridStart.Y - CSng(txtGridY) To picDraw.ScaleTop Step -gScaleY
        picDraw.PSet (I, J), vbBlack
    Next J
Next I
For I = GridStart.X - CSng(txtGridX) To picDraw.ScaleLeft Step -gScaleX
    For J = GridStart.Y + CSng(txtGridY) To picDraw.ScaleHeight + picDraw.ScaleTop Step gScaleY
        picDraw.PSet (I, J), vbBlack
    Next J
Next I
End Sub

Private Sub chkSnapEdit_Click()
Redraw
End Sub

Private Sub chkSnapEndMid_Click()
Redraw
End Sub

Private Sub chkSnapGrid_Click()
Redraw
End Sub

Private Sub chkSnapIntersect_Click()
Redraw
End Sub


Private Sub cmdClear_Click()
Erase DrawGeo()
Erase OutLines
ReDim DrawGeo(0) As Geometry
DrawGeo(0).Name = "0"
Redraw
End Sub

Private Sub cmdClearMesh_Click()
Erase DrawGeo(0).Faces
Erase OutLines()
Redraw
End Sub

Private Sub cmdDelaunay_Click()
On Local Error GoTo eTrap
Dim DFactor As Single
Dim tmpGeo() As Geometry
Dim tmpFaces() As CadFace
Dim tmpLines() As CadLine
Dim I As Integer
Dim J As Integer
ReDim tmpGeo(0)
I = UBound(Selection)
If I > -1 Then
    For J = 0 To I
        AddSelectionToGeo DrawGeo(0), Selection(J), tmpGeo(0)
    Next J
Else
    AddGeo DrawGeo(), tmpGeo()
End If
tmpFaces() = DrawGeo(0).Faces
tmpLines() = OutLines()
DFactor = Val(txtDTol.Text)
If DFactor < 0 Then MsgBox "The mesh tolerance must be greater than 0": Exit Sub
DoEvents
DeLaunayGeo picDraw, tmpGeo(0), DPoint, DrawGeo(0).Faces(), OutLines(), DFactor, False, False
AddLines tmpLines(), OutLines()
AddFaces tmpFaces(), DrawGeo(0).Faces()
Redraw
Exit Sub
eTrap:
    I = -1
    Resume Next
End Sub



Private Sub cmdHelp_Click()
Dim msg As String
msg = "STEP 1:   Create 'closed' geometry." & vbNewLine
msg = msg & "STEP 2:   Set mesh center (as if you were to 'fill)." & vbNewLine
msg = msg & "STEP 3:   Select geometry. (hold CTRL to ADD to selections)" & vbNewLine
msg = msg & "STEP 4:   Triangulate geometry." & vbNewLine
msg = msg & "STEP 5:   Show mesh." & vbNewLine
MsgBox msg, vbInformation, "*BASIC* Instructions"
End Sub

Private Sub cmdImport_Click()
Dim Res As String
Res = InputBox("Enter Filename", "DXF Import")
If Res = "" Then Exit Sub
If Dir(Res) = "" Then Exit Sub
Erase DrawGeo()
Erase OutLines
ImportDXF Me, DrawGeo(), Layers(), Res
Me.Caption = "Drawing DXF"
Redraw
Me.Caption = "MeshMaker"
End Sub

Private Sub cmdLoad_Click()
Dim Res As String
Dim tLine As CadLine
Res = InputBox("Enter Filename", "LOAD GEOMETRY")
If Res = "" Then Exit Sub
If Dir(Res) = "" Then Exit Sub
Erase DrawGeo()
ReDim DrawGeo(0) As Geometry
DrawGeo(0).Name = "0"
LoadGeo tLine, DrawGeo(), Res
Redraw
End Sub

Private Sub cmdMesh_Click()
CreateMesh
End Sub





Private Sub cmdSave_Click()
Dim Res As String
Res = InputBox("Enter Filename")
If Res = "" Then Exit Sub
SaveGeo GetExtents(DrawGeo(), 0), DrawGeo(), Res
End Sub

Private Sub cmdSaveMesh_Click()
Dim Res As String
Res = InputBox("Enter Filename", "Export Mesh For DirectX")
If Res = "" Then Exit Sub
CreateMesh Res
End Sub

Private Sub cmdZoomIn_Click()
ZoomPic 2
End Sub

Sub ZoomPic(Level As Single)
Dim xInt As Single
Dim yInt As Single
ZoomLine = GetExtents(DrawGeo(), 0)
xInt = picDraw.ScaleWidth / Level
yInt = picDraw.ScaleHeight / Level
ZoomLine.P1.X = ZoomLine.P1.X + xInt
ZoomLine.P2.X = ZoomLine.P2.X - xInt
ZoomLine.P1.Y = ZoomLine.P1.Y + yInt
ZoomLine.P2.Y = ZoomLine.P2.Y - yInt
Zoom ZoomLine, picDraw
Redraw
End Sub

Private Sub cmdZoomOut_Click()
ZoomPic -1.5
End Sub

Private Sub Form_Load()

DoEvents
Dim tLine As CadLine
ReDim DrawGeo(0) As Geometry
DrawGeo(0).Name = "0"
'LoadGeo tLine, DrawGeo(), "psc"
End Sub

Sub CreateMDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = Format(X, "0.000")
Y = -Format(Y, "0.000")
Snap X, Y
CurX = X
CurY = Y
VisPoint.X = X
VisPoint.Y = Y
If Stage > StageMax Then Stage = 0
If Stage > 0 Then Exit Sub
StartStage = True
CurLayer.Name = 0
CurLayer.Color = vbBlue
CurLayer.style = 0
CurLayer.Width = 1
CurLayer.FontName = "Arial"
Select Case Command
    Case "Pan"
        sX = X
        sY = Y
    Case "Zoom"
        ZoomLine.P1.X = X
        ZoomLine.P1.Y = -Y
        ZoomLine.P2 = ZoomLine.P1
    Case "Box Select"
        SelLine.P1.X = X
        SelLine.P1.Y = Y
        SelLine.P2 = SelLine.P1
    Case "Line", "Move", "Rotate", "Offset", "Array"
        CurLine.Layer = CurLayer
    Case "Circle"
        CurArc.Layer = CurLayer
        CurArc.Angle1 = 0
        CurArc.Angle2 = 360
        CurArc.Radius = 0
        CurLine.Layer = CurLayer
    Case "Arc"
        CurArc.Layer = CurLayer
        CurArc.Angle1 = 0
        CurArc.Angle2 = 360
        CurArc.Radius = 0
        CurLine.Layer = CurLayer
    Case "Closed Ellipse"
        CurEllipse.Layer = CurLayer
        CurEllipse.Angle1 = 0
        CurEllipse.Angle2 = 360
        CurEllipse.NumPoints = 32
        CurLine.Layer = CurLayer
    Case "Ellipse"
        CurEllipse.Layer = CurLayer
        CurEllipse.NumPoints = 32
        CurEllipse.Angle1 = 0
        CurEllipse.Angle2 = 360
        CurLine.Layer = CurLayer
    Case "Spline"
        vCount = 0
        StartMove = False
        CurLine.P1 = VisPoint
        CurLine.P2 = VisPoint
        CurLine.Layer = CurLayer
        CurSpline.Layer.Color = vbBlue
        CurSpline.Layer.style = 0
        CurSpline.Layer.Width = 1
        ReDim CurSpline.Vertex(vCount)
    Case "PolyLine"
        vCount = 0
        StartMove = False
        CurLine.P1 = VisPoint
        CurLine.P2 = VisPoint
        CurLine.Layer = CurLayer
        CurPolyLine.Layer = CurLayer
        ReDim CurPolyLine.Vertex(vCount)
    Case "3 Point Arc"
        Cur3Pt(0) = VisPoint
        Cur3Pt(1) = VisPoint
        Cur3Pt(2) = VisPoint
        CurArc.Layer = CurLayer
        CurArc.Angle1 = 0
        CurArc.Angle2 = 360
        CurArc.Radius = 0
End Select

End Sub

Sub CreateMMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
X = Format(X, "0.000")
Y = -Format(Y, "0.000")
Snap X, Y
CurX = X
CurY = Y

If Button And 2 Then Exit Sub
'----------Visual Aid----------
If Command <> "Pan" Then
    VisPoint.Layer.Width = 3
    VisPoint.Layer.Color = vbBlue
    VisPoint.Layer.style = 0
    If StartStage Then DrawCadPoint picDraw, VisPoint, 6
    VisPoint.X = X
    VisPoint.Y = Y
    DrawCadPoint picDraw, VisPoint, 6
End If
StartStage = True
'Show Data
Select Case Command
Case "Pan"
        If Button = 0 Then Exit Sub
        picDraw.ScaleTop = picDraw.ScaleTop + (Y - sY)
        picDraw.ScaleLeft = picDraw.ScaleLeft + (sX - X)
        ZoomLine.P1.X = picDraw.ScaleLeft
        ZoomLine.P1.Y = picDraw.ScaleTop
        ZoomLine.P2.X = picDraw.ScaleLeft + picDraw.ScaleWidth
        ZoomLine.P2.Y = picDraw.ScaleTop + picDraw.ScaleHeight
        picDraw.Cls
        DrawCadGeo picDraw, DrawGeo(), cView
    Case "Zoom"
        If Button = 0 Then Exit Sub
        picDraw.DrawMode = 6
        picDraw.DrawWidth = 1
        picDraw.DrawStyle = 2
        picDraw.Line (ZoomLine.P1.X, ZoomLine.P1.Y)-(ZoomLine.P2.X, ZoomLine.P2.Y), vbBlack, B
        ZoomLine.P2.X = X
        ZoomLine.P2.Y = -Y
        picDraw.Line (ZoomLine.P1.X, ZoomLine.P1.Y)-(ZoomLine.P2.X, ZoomLine.P2.Y), vbBlack, B
    Case "Box Select"
        If Button = 0 Then Exit Sub
        picDraw.DrawMode = 6
        picDraw.DrawWidth = 1
        picDraw.DrawStyle = 2
        picDraw.Line (SelLine.P1.X, -SelLine.P1.Y)-(SelLine.P2.X, -SelLine.P2.Y), vbBlack, B
        SelLine.P2.X = X
        SelLine.P2.Y = Y
        picDraw.Line (SelLine.P1.X, -SelLine.P1.Y)-(SelLine.P2.X, -SelLine.P2.Y), vbBlack, B
    Case "Line"
        Select Case Stage
            Case 0
                CurLine.P1 = VisPoint
                CurLine.P2 = VisPoint
            Case 1
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
        End Select
    Case "Circle"
        Select Case Stage
            Case 0
                CurLine.P1 = VisPoint
            Case 1
                CurLine.P1 = CurArc.Center
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadArc picDraw, CurArc, 6
                CurArc.Radius = LineLen(CurLine)
                DrawCadArc picDraw, CurArc, 6
        End Select
    Case "Arc"
        Select Case Stage
            Case 0
                CurLine.P1 = VisPoint
            Case 1
                CurLine.P1 = CurArc.Center
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadArc picDraw, CurArc, 6
                CurArc.Radius = LineLen(CurLine)
                DrawCadArc picDraw, CurArc, 6
            Case 2
                CurLine.P1 = CurArc.Center
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadArc picDraw, CurArc, 6
                CurArc.Angle1 = cAngle(CurLine)
                DrawCadArc picDraw, CurArc, 6
            Case 3
                CurLine.P1 = CurArc.Center
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadArc picDraw, CurArc, 6
                CurArc.Angle2 = cAngle(CurLine)
                DrawCadArc picDraw, CurArc, 6
        End Select
    Case "Closed Ellipse"
        Select Case Stage
            Case 0
                CurLine.P1 = VisPoint
            Case 1
                CurLine.P1 = CurEllipse.F1
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadEllipse picDraw, CurEllipse, 6
                CurEllipse.F2 = VisPoint
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 2
                CurLine.P1 = EllipseCenter(CurEllipse)
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadEllipse picDraw, CurEllipse, 6
                CurEllipse.P1 = VisPoint
                DrawCadEllipse picDraw, CurEllipse, 6
        End Select
    Case "Ellipse"
        Select Case Stage
            Case 0
                CurLine.P1 = VisPoint
            Case 1
                CurLine.P1 = CurEllipse.F1
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadEllipse picDraw, CurEllipse, 6
                CurEllipse.F2 = VisPoint
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 2
                CurLine.P1 = EllipseCenter(CurEllipse)
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadEllipse picDraw, CurEllipse, 6
                CurEllipse.P1 = VisPoint
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 3
                CurLine.P1 = EllipseCenter(CurEllipse)
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadEllipse picDraw, CurEllipse, 6
                CurEllipse.Angle1 = EllipseAngle(cAngle(CurLine), CurEllipse)
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 4
                CurLine.P1 = EllipseCenter(CurEllipse)
                DrawCadLine picDraw, CurLine, 6
                CurLine.P2 = VisPoint
                DrawCadLine picDraw, CurLine, 6
                DrawCadEllipse picDraw, CurEllipse, 6
                CurEllipse.Angle2 = EllipseAngle(cAngle(CurLine), CurEllipse)
                DrawCadEllipse picDraw, CurEllipse, 6
        End Select
    Case "3 Point Arc"
        If Stage = 2 Then
            Cur3Pt(2) = VisPoint
            DrawCadArc picDraw, CurArc, 6
            DelaunayArc Cur3Pt(), CurArc
            DrawCadArc picDraw, CurArc, 6
        End If
    Case "Spline"
        If Stage = 0 Then
            CurSpline.Vertex(0) = VisPoint
            Exit Sub
        End If
        DrawCadLine picDraw, CurLine, 6
        CurLine.P2 = VisPoint
        DrawCadLine picDraw, CurLine, 6
        If StartMove Then DrawCadSpline picDraw, CurSpline, 6
        CurSpline.Vertex(vCount).X = X
        CurSpline.Vertex(vCount).Y = Y
        DrawCadSpline picDraw, CurSpline, 6
        StartMove = True
    Case "PolyLine"
        If Stage = 0 Then
            CurPolyLine.Vertex(0) = VisPoint
        Else
            DrawCadLine picDraw, CurLine, 6
            CurLine.P2 = VisPoint
            DrawCadLine picDraw, CurLine, 6
        End If
End Select
End Sub
Sub CreateMUp(Button As Integer, Shift As Integer, X As Single, Y As Single, Optional KeyEntry As Boolean)
On Local Error Resume Next
Dim I As Integer
Dim J As Integer
X = Format(X, "0.000")
Y = -Format(Y, "0.000")
Snap X, Y
CurX = X
CurY = Y
VisPoint.X = X
VisPoint.Y = Y
VisPoint.Layer.Width = 3
VisPoint.Layer.Color = vbBlue
DrawCadPoint picDraw, VisPoint, 6
If Button And 2 Then Exit Sub
sX = X
sY = Y
Select Case Command
    Case "Zoom"
        ZoomLine.P1.Y = -ZoomLine.P1.Y
        ZoomLine.P2.Y = -ZoomLine.P2.Y
        Zoom ZoomLine, picDraw
        Redraw
    Case "Pan"
        Redraw
    Case "Box Select"
        picDraw.Line (SelLine.P1.X, -SelLine.P1.Y)-(SelLine.P2.X, -SelLine.P2.Y), vbBlack, B
        SelLine.P2 = VisPoint
        Dim TempSel() As SelSet
        If Shift = 2 Or Shift = 4 Then AddSelection Selection(), TempSel()
        SelectFromBox SelLine, DrawGeo(), cView, Selection()
        If isSelected(Selection()) Then
            If Shift = 4 And isSelected(TempSel()) Then
                For I = 0 To UBound(Selection)
                    For J = 0 To UBound(TempSel)
                        If Selection(I).Type = TempSel(J).Type And Selection(I).Index = TempSel(J).Index Then TempSel(J).Type = "BLANK"
                    Next J
                    Selection(I).Type = "BLANK"
                Next I
            Else
                Closest = Selection(0)
                MakeSelCurrent Closest
            End If
        End If
        AddSelection TempSel(), Selection()
        ShowSelCount
        Redraw
    Case "DCenter"
        DPoint = VisPoint
        Redraw
    Case "Line"
        '----------Visual Aid----------
        Select Case Stage
            Case 0
                DrawCadLine picDraw, CurLine, 6
                CurLine.P1 = VisPoint
                Stage = Stage + 1
                DrawCadLine picDraw, CurLine, 6
            Case 1
                CurLine.P2 = VisPoint
                AddElement "Line"
                If chkChain.Value Then
                    CurLine.P1 = CurLine.P2
                Else
                    Stage = 0
                    StartStage = False
                End If
                Redraw
        End Select
    Case "Circle"
        Select Case Stage
            Case 0
                DrawCadArc picDraw, CurArc, 6
                CurLine.P1 = VisPoint
                CurLine.P2 = VisPoint
                CurArc.Center = VisPoint
                Stage = Stage + 1
                DrawCadArc picDraw, CurArc, 6
            Case 1
                CurLine.P2 = VisPoint
                CurArc.Radius = LineLen(CurLine)
                AddElement "Arc"
                Stage = 0
                StartStage = False
                Redraw
        End Select
    Case "Arc"
        Select Case Stage
            Case 0
                DrawCadArc picDraw, CurArc, 6
                CurLine.P1 = VisPoint
                CurLine.P2 = VisPoint
                CurArc.Center = VisPoint
                Stage = Stage + 1
                DrawCadArc picDraw, CurArc, 6
            Case 1
                DrawCadArc picDraw, CurArc, 6
                CurLine.P2 = VisPoint
                CurArc.Radius = LineLen(CurLine)
                Stage = Stage + 1
                DrawCadArc picDraw, CurArc, 6
            Case 2
                DrawCadArc picDraw, CurArc, 6
                CurLine.P2 = VisPoint
                CurArc.Angle1 = cAngle(CurLine)
                Stage = Stage + 1
                DrawCadArc picDraw, CurArc, 6
            Case 3
                CurLine.P2 = VisPoint
                CurArc.Angle2 = cAngle(CurLine)
                StartStage = False
                AddElement "Arc"
                Stage = 0
                Redraw
        End Select
    Case "Closed Ellipse"
        Select Case Stage
            Case 0
                'DrawCadEllipse picDraw, CurEllipse, 6
                CurLine.P1 = VisPoint
                CurLine.P2 = VisPoint
                CurEllipse.F1 = VisPoint
                CurEllipse.F2 = VisPoint
                CurEllipse.P1 = VisPoint
                Stage = Stage + 1
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 1
                DrawCadEllipse picDraw, CurEllipse, 6
                CurLine.P2 = VisPoint
                CurEllipse.F2 = VisPoint
                CurLine.P1 = EllipseCenter(CurEllipse)
                DrawCadLine picDraw, CurLine
                DrawCadEllipse picDraw, CurEllipse, 6
                Stage = Stage + 1
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 2
                CurLine.P2 = VisPoint
                CurEllipse.P1 = VisPoint
                CurEllipse.NumPoints = 32
                StartStage = False
                AddElement "Ellipse"
                Stage = 0
                Redraw
        End Select
    Case "Ellipse"
        Select Case Stage
            Case 0
                'DrawCadEllipse picDraw, CurEllipse, 6
                CurLine.P1 = VisPoint
                CurLine.P2 = VisPoint
                CurEllipse.F1 = VisPoint
                CurEllipse.F2 = VisPoint
                CurEllipse.P1 = VisPoint
                Stage = Stage + 1
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 1
                DrawCadEllipse picDraw, CurEllipse, 6
                CurLine.P2 = VisPoint
                CurEllipse.F2 = VisPoint
                CurLine.P1 = EllipseCenter(CurEllipse)
                DrawCadLine picDraw, CurLine
                DrawCadEllipse picDraw, CurEllipse, 6
                Stage = Stage + 1
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 2
                DrawCadEllipse picDraw, CurEllipse, 6
                CurLine.P2 = VisPoint
                CurEllipse.P1 = VisPoint
                Stage = Stage + 1
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 3
                DrawCadEllipse picDraw, CurEllipse, 6
                CurLine.P2 = VisPoint
                CurEllipse.Angle1 = EllipseAngle(cAngle(CurLine), CurEllipse)
                Stage = Stage + 1
                DrawCadEllipse picDraw, CurEllipse, 6
            Case 4
                CurLine.P2 = VisPoint
                CurEllipse.Angle2 = EllipseAngle(cAngle(CurLine), CurEllipse)
                CurEllipse.NumPoints = 32
                StartStage = False
                AddElement "Ellipse"
                Stage = 0
                Redraw
        End Select
    Case "3 Point Arc"
        Select Case Stage
            Case 0
                Cur3Pt(0) = VisPoint
                DrawCadPoint picDraw, VisPoint
                Stage = Stage + 1
            Case 1
                Cur3Pt(1) = VisPoint
                Cur3Pt(2) = VisPoint
                DrawCadPoint picDraw, VisPoint
                DelaunayArc Cur3Pt(), CurArc
                DrawCadArc picDraw, CurArc, 6
                Stage = Stage + 1
            Case 2
                Cur3Pt(2) = VisPoint
                DrawCadPoint picDraw, VisPoint
                DelaunayArc Cur3Pt(), CurArc
                AddElement "Arc"
                Stage = 0
                Redraw
        End Select
    Case "Spline"
        DrawCadPoint picDraw, VisPoint
        CurSpline.Vertex(vCount) = VisPoint
        CurLine.P1 = CurLine.P2
        vCount = vCount + 1
        DrawCadSpline picDraw, CurSpline, 6
        ReDim Preserve CurSpline.Vertex(vCount)
        CurSpline.Vertex(vCount) = VisPoint
        Stage = Stage + 1
        StartMove = False
        StageMax = StageMax + 1
    Case "PolyLine"
        DrawCadPoint picDraw, VisPoint
        CurPolyLine.Vertex(vCount) = VisPoint
        CurLine.P1 = CurLine.P2
        vCount = vCount + 1
        ReDim Preserve CurPolyLine.Vertex(vCount)
        CurPolyLine.Vertex(vCount) = VisPoint
        Stage = Stage + 1
        StartMove = False
        StageMax = StageMax + 1
End Select
GetSnap
End Sub
Sub ShowSelCount()
On Local Error GoTo eTrap
Dim I As Integer
Dim J As Integer
Dim K As Integer
K = UBound(Selection) + 1
For I = 0 To UBound(Selection) - 1
    For J = I + 1 To UBound(Selection)
        If Selection(I).Index = Selection(J).Index And Selection(I).Type = Selection(J).Type Then
            K = K - 1
        End If
    Next J
Next I
Me.Caption = "VB CAD (" & K & " Selected)"
eTrap:
End Sub
Private Sub MakeSelCurrent(MySel As SelSet)
Dim I As Integer
Dim J As Integer
Dim K As Integer
J = MySel.Index
Select Case MySel.Type
    Case "Point"
        CurPoint = DrawGeo(cView).Points(J)
    Case "Line"
        CurLine = DrawGeo(cView).Lines(J)
    Case "Arc"
        CurArc = DrawGeo(cView).Arcs(J)
    Case "Ellipse"
        CurEllipse = DrawGeo(cView).Ellipses(J)
    Case "Spline"
        CurSpline = DrawGeo(cView).Splines(J)
    Case "PolyLine"
        CurPolyLine = DrawGeo(cView).PolyLines(J)
    Case "Text"
        CurText = DrawGeo(cView).Text(J)
    Case "Insert"
        curInsert = DrawGeo(cView).Inserts(J)
    Case "Face"
        CurFace = DrawGeo(cView).Faces(J)
End Select
Closest = MySel
End Sub
Sub AddElement(EType As String)
On Error GoTo eTrap
Dim Count As Integer
Select Case EType
    Case "Point"
        Count = UBound(DrawGeo(cView).Points) + 1
        ReDim Preserve DrawGeo(cView).Points(Count)
        DrawGeo(cView).Points(Count) = CurPoint
    Case "Line"
        Count = UBound(DrawGeo(cView).Lines) + 1
        ReDim Preserve DrawGeo(cView).Lines(Count)
        DrawGeo(cView).Lines(Count) = CurLine
    Case "Arc"
        Count = UBound(DrawGeo(cView).Arcs) + 1
        ReDim Preserve DrawGeo(cView).Arcs(Count)
        DrawGeo(cView).Arcs(Count) = CurArc
    Case "Ellipse"
        Count = UBound(DrawGeo(cView).Ellipses) + 1
        ReDim Preserve DrawGeo(cView).Ellipses(Count)
        DrawGeo(cView).Ellipses(Count) = CurEllipse
    Case "Spline"
        ReDim Preserve CurSpline.Vertex(UBound(CurSpline.Vertex) - 1)
        Count = UBound(DrawGeo(cView).Splines) + 1
        ReDim Preserve DrawGeo(cView).Splines(Count)
        DrawGeo(cView).Splines(Count) = CurSpline
    Case "PolyLine"
        ReDim Preserve CurPolyLine.Vertex(UBound(CurPolyLine.Vertex) - 1)
        Count = UBound(DrawGeo(cView).PolyLines) + 1
        ReDim Preserve DrawGeo(cView).PolyLines(Count)
        DrawGeo(cView).PolyLines(Count) = CurPolyLine
    Case "Text"
        Count = UBound(DrawGeo(cView).Text) + 1
        ReDim Preserve DrawGeo(cView).Text(Count)
        DrawGeo(cView).Text(Count) = CurText
    Case "Insert"
        Count = UBound(DrawGeo(cView).Inserts) + 1
        ReDim Preserve DrawGeo(cView).Inserts(Count)
        DrawGeo(cView).Inserts(Count) = curInsert
End Select
Exit Sub
eTrap:
Count = 0
Resume Next

End Sub

Sub GetSnap()
On Error GoTo eTrap
Erase SnapPts()
Dim SnapSel() As SelSet
Dim tPts() As CadPoint
Dim I As Integer
Dim K As Integer
'-------------------
SelectAllGeo SnapSel(), DrawGeo(0)
K = UBound(SnapSel)
'---------------------
'Get Ends and Mid
If chkSnapEndMid.Value = vbChecked Then
    For I = 0 To K
        GeoSelEndMidPoints DrawGeo(cView), SnapSel(I), tPts()
        AddPoints tPts(), SnapPts()
    Next I
End If
'Get editing points
If chkSnapEdit.Value = vbChecked Then
    For I = 0 To K
        GeoSelEditPoints DrawGeo(0), SnapSel(I), tPts()
        AddPoints tPts(), SnapPts()
    Next I
End If
'Get intersections
If chkSnapIntersect.Value = vbChecked Then
    SelectionIntersect DrawGeo(cView), SnapSel(), tPts(), True
    AddPoints tPts(), SnapPts()
End If
DrawPointArray picDraw, SnapPts(), , RGB(255, 127, 127), 3
Exit Sub
eTrap:
End Sub
Sub Snap(ByRef X As Single, ByRef Y As Single)
On Error GoTo eTrap
Dim I As Integer
Dim K As Integer
Dim Min As Single
Dim tPt As CadPoint
Select Case Command
    Case "PolyLine"
        AddPoints CurPolyLine.Vertex(), SnapPts()
    Case "Spline"
        AddPoints CurSpline.Vertex(), SnapPts()
End Select
If chkSnapGrid.Value = vbChecked Then
    gScaleX = CSng(txtGridX.Text)
    gScaleY = CSng(txtGridY.Text)
    Dim Xmin As Single
    Dim Ymin As Single
    Dim Xmax As Single
    Dim Ymax As Single
    Xmin = GridStart.X + (((X - GridStart.X) \ gScaleX) * gScaleX)  'number of Xgrid points between grid origin and point
    Ymin = GridStart.Y + (((Y - GridStart.Y) \ gScaleY) * gScaleY)  'number of Ygrid points between grid origin and point
    Xmax = Xmin + gScaleX
    Ymax = Ymin + gScaleY
    X = IIf(Abs(X - Xmin) < Abs(X - Xmax), Xmin, Xmax)
    Y = IIf(Abs(Y - Ymin) < Abs(Y - Ymax), Ymin, Ymax)
End If
K = UBound(SnapPts)
Min = picDraw.ScaleX(8, vbPixels, vbUser)
tPt.X = X
tPt.Y = Y
For I = 0 To K
    If PtLen(tPt, SnapPts(I)) <= Min Then
        Min = PtLen(tPt, SnapPts(I))
        tPt.X = SnapPts(I).X
        tPt.Y = SnapPts(I).Y
    End If
Next I
X = tPt.X
Y = tPt.Y
Exit Sub
eTrap:
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
picDraw.Width = Me.ScaleWidth - 115
End Sub



Private Sub optCommand_Click(Index As Integer)
Redraw
StartCommand Index
End Sub
Sub StartCommand(Index As Integer)
Stage = 0
picDraw.MousePointer = 2
Select Case Index
    Case 0: Command = "Line": StageMax = 1
    Case 1: Command = "PolyLine": StageMax = 1
    Case 2: Command = "Spline": StageMax = 1
    Case 3: Command = "3 Point Arc": StageMax = 2
    Case 4: Command = "Circle": StageMax = 1
    Case 5: Command = "Arc": StageMax = 3
    Case 6: Command = "Closed Ellipse": StageMax = 2
    Case 7: Command = "Ellipse": StageMax = 4
    Case 8: Command = "Box Select": StageMax = 1
    Case 9
        Command = ""
        DeleteSelection Selection(), DrawGeo(0)
        Redraw
        Erase Selection()
        optCommand(8).Value = True
    Case 10: Command = "DCenter": StageMax = 0
    Case 11
        picDraw.MousePointer = 99
        Command = "Pan"
        StageMax = 0
    Case 12: Command = "Zoom": StageMax = 0
End Select
GetSnap
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CreateMDown Button, Shift, X, Y
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CreateMMove Button, Shift, X, Y
End Sub


Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    AddElement Command
    Stage = 0
    Redraw
Else
    CreateMUp Button, Shift, X, Y
End If
End Sub


Private Sub picDraw_Resize()
On Error Resume Next
Redraw
End Sub


