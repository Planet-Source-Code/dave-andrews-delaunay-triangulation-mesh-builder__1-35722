VERSION 5.00
Begin VB.Form frmFull3D 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4785
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   3120
   End
   Begin VB.PictureBox picDX 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2445
      Left            =   -15
      ScaleHeight     =   159
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   1
      Top             =   255
      Width           =   2895
      Begin VB.CommandButton cmdQual 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   150
         TabIndex        =   7
         Top             =   30
         Width           =   180
      End
      Begin VB.CommandButton cmdCenter 
         Caption         =   "X"
         Height          =   135
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   135
      End
      Begin VB.CommandButton cmdMag20 
         Caption         =   "X"
         Height          =   135
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   135
      End
      Begin VB.VScrollBar vscrMag 
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Value           =   40
         Width           =   135
      End
      Begin VB.HScrollBar hscrX 
         Height          =   135
         Left            =   120
         Min             =   -32767
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.VScrollBar vscrY 
         Height          =   615
         Left            =   360
         Max             =   -32767
         Min             =   32767
         TabIndex        =   5
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmFull3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RMC As New clsDX3D

Dim Spin As Boolean
Dim wType As Integer
Sub CenterObject()
Dim Part As Direct3DRMMeshBuilder3
Dim xBox As D3DRMBOX
Dim cx As Single
Dim cy As Single
Dim cZ As Single
Set Part = frmFull3D.RMC.mFrO.GetVisual(0)
Part.GetBox xBox
cx = -(xBox.Min.X + xBox.Max.X) / 2
cy = -(xBox.Min.Y + xBox.Max.Y) / 2
cZ = -(xBox.Min.z + xBox.Max.z) / 2
Part.Translate cx, cy, cZ
frmFull3D.RMC.Update
End Sub


Private Sub GetPyramid()
Dim f As Direct3DRMFace2
Set f = RMC.mDrm.CreateFace
f.AddVertex 0, 0.5, 0
f.AddVertex -0.5, -0.5, 0
f.AddVertex 0.5, -0.5, 0
mMesh.AddFace f
Set f = RMC.mDrm.CreateFace
f.AddVertex 0, 0, -1
f.AddVertex 0.5, -0.5, 0
f.AddVertex -0.5, -0.5, 0
mMesh.AddFace f
Set f = RMC.mDrm.CreateFace
f.AddVertex 0, 0, -1
f.AddVertex 0, 0.5, 0
f.AddVertex 0.5, -0.5, 0
mMesh.AddFace f
Set f = RMC.mDrm.CreateFace
f.AddVertex 0, 0, -1
f.AddVertex -0.5, -0.5, 0
f.AddVertex 0, 0.5, 0
mMesh.AddFace f
Set f = Nothing
End Sub

Private Sub cmdCenter_Click()
hscrX = 0
vscrY = 0
End Sub



Function ChkStr(InString As String) As String
Dim I As Integer
For I = 1 To Len(InString)
    If Mid(InString, I, 1) = Chr$(34) Then
        ChkStr = ChkStr & "'"
    Else
        ChkStr = ChkStr & Mid(InString, I, 1)
    End If
Next I
End Function


Private Sub cmdMag20_Click()
vscrMag = 40
End Sub


Private Sub cmdQual_Click()
Dim Part As Direct3DRMMeshBuilder3
Dim I As Integer
wType = wType + 1
If wType > 4 Then wType = 0
For I = 0 To RMC.mFrO.GetVisualCount - 1
    Set Part = RMC.mFrO.GetVisual(I)
    
    Select Case wType
        Case 0
            Part.SetQuality D3DRMRENDER_FLAT
            cmdQual.Caption = "F"
        Case 1
            Part.SetQuality D3DRMRENDER_GOURAUD
            cmdQual.Caption = "G"
        Case 2
            Part.SetQuality D3DRMRENDER_PHONG
            cmdQual.Caption = "P"
        Case 3
            Part.SetQuality D3DRMRENDER_UNLITFLAT
            cmdQual.Caption = "U"
        Case 4
            Part.SetQuality D3DRMRENDER_WIREFRAME
            cmdQual.Caption = "W"
    End Select
Next I
RMC.Update
End Sub

Private Sub Form_Load()
RMC.InitDx picDX
End Sub

Sub SetZoom()
Dim B As D3DRMBOX
Dim cx As Long
Dim cy As Long
Dim MaxZ As Long
RMC.mFrO.GetHierarchyBox B
cx = (B.Max.X + B.Min.X) / 2
cy = (B.Max.Y + B.Min.Y) / 2
MaxZ = Sqr(Abs(B.Max.X - B.Min.X) ^ 2 + Abs(B.Max.Y - B.Min.Y) ^ 2)
If MaxZ > 300 Then
    MaxZ = 300
    cx = 0
    cy = 0
End If
RMC.mVpt.SetBack MaxZ * 2.5
RMC.mFrO.SetPosition Nothing, -cx, -cy, MaxZ
vscrMag.Value = MaxZ
End Sub

Private Sub Form_Resize()
picDX.Width = Me.ScaleWidth
picDX.Height = Me.ScaleHeight
lblResult.Width = Me.ScaleWidth
End Sub


Private Sub hscrX_Change()
Dim xVect As D3DVECTOR
RMC.mFrO.GetPosition Nothing, xVect
xVect.X = hscrX
RMC.mFrO.SetPosition Nothing, xVect.X, xVect.Y, xVect.z
RMC.Update
End Sub




Private Sub picDX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RMC.MouseDown X, Y
'lblResult = RMC.Pick(x, y)
'If Button = 1 And Shift = 1 Then Spin = True Else Spin = False
End Sub


Private Sub picDX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And 1 Then RMC.MouseMove X, Y
End Sub


Private Sub picDX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Spin Then RMC.MouseUp
End Sub


Private Sub picDX_Resize()
RMC.Resize picDX
RMC.mVpt.SetBack 500
End Sub


Private Sub Timer1_Timer()
RMC.Update
DoEvents
End Sub

Private Sub vscrMag_Change()
Dim xVect As D3DVECTOR
RMC.mFrO.GetPosition Nothing, xVect
xVect.z = vscrMag
RMC.mFrO.SetPosition Nothing, xVect.X, xVect.Y, xVect.z
RMC.Update
End Sub


Private Sub vscrMag_Scroll()
Dim xVect As D3DVECTOR
RMC.mFrO.GetPosition Nothing, xVect
xVect.z = vscrMag
RMC.mFrO.SetPosition Nothing, xVect.X, xVect.Y, xVect.z
RMC.Update
End Sub


Private Sub vscrY_Change()
Dim xVect As D3DVECTOR
RMC.mFrO.GetPosition Nothing, xVect
xVect.Y = vscrY
RMC.mFrO.SetPosition Nothing, xVect.X, xVect.Y, xVect.z
RMC.Update
End Sub


Private Sub vscrY_Scroll()
Dim xVect As D3DVECTOR
RMC.mFrO.GetPosition Nothing, xVect
xVect.Y = vscrY
RMC.mFrO.SetPosition Nothing, xVect.X, xVect.Y, xVect.z
RMC.Update
End Sub


