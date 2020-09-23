VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ucStatusBar"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fmTheme 
      Caption         =   "Theme:"
      Height          =   855
      Left            =   3120
      TabIndex        =   21
      Top             =   2880
      Width           =   2775
      Begin VB.ComboBox cmbTheme 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fmBound 
      Caption         =   "Object Binding:"
      Height          =   2655
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   2775
      Begin VB.ComboBox cmbSizeMethod 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Height          =   255
         Index           =   1
         Left            =   1860
         TabIndex        =   20
         Top             =   700
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Height          =   255
         Index           =   0
         Left            =   660
         TabIndex        =   19
         Top             =   700
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.ComboBox cmbBound 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   2040
         Width           =   1215
      End
      Begin prjucStatusBar.ucProgressBar ucProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   390
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   12563634
      End
      Begin VB.Label Label1 
         Caption         =   "Sizing:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1350
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text Align:"
      Height          =   855
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cmbAlign 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fmPanel 
      Caption         =   "Active Panel:"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cmbPanelIndex 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblPanel 
         Caption         =   "Index:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   390
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Icons:"
      Height          =   1695
      Left            =   3120
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   4
         Left            =   2220
         TabIndex        =   8
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   3
         Left            =   1740
         TabIndex        =   7
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   4
         Left            =   2160
         Picture         =   "frmMain.frx":0000
         Top             =   360
         Width           =   360
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   3
         Left            =   1680
         Picture         =   "frmMain.frx":076A
         Top             =   360
         Width           =   360
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   2
         Left            =   1140
         Picture         =   "frmMain.frx":0ED4
         Top             =   360
         Width           =   360
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   1
         Left            =   660
         Picture         =   "frmMain.frx":163E
         Top             =   360
         Width           =   360
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":1DA8
         Top             =   360
         Width           =   360
      End
   End
   Begin prjucStatusBar.ucStatusBar ucStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   3825
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GripShape       =   2
      Theme           =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAlign_Click()
    With Me
        .ucStatusBar1.PanelAlignment(.cmbPanelIndex.ListIndex + 1) = .cmbAlign.ListIndex
    End With
End Sub

Private Sub cmbPanelIndex_Click()
    With Me
        .cmbAlign.ListIndex = .ucStatusBar1.PanelAlignment(cmbPanelIndex.ListIndex + 1)
    End With
End Sub

Private Sub cmbTheme_Click()
    With Me
        .ucStatusBar1.Theme = .cmbTheme.ListIndex
    End With
End Sub

Private Sub Command1_Click()
    With Me
        Set .ucStatusBar1.PanelIcon(.cmbPanelIndex.ListIndex + 1) = .imgIcon(GetOption).Picture
    End With
End Sub

Private Sub Command2_Click()
    With Me
        Set .ucStatusBar1.PanelIcon(.cmbPanelIndex.ListIndex + 1) = Nothing
    End With
End Sub

Private Sub Command3_Click()
    With Me
        If .Option2(0).Value Then
            Call .ucStatusBar1.BoundControl(.cmbPanelIndex.ListIndex + 1, .ucProgressBar1, .cmbSizeMethod.ListIndex)
        Else
            Call .ucStatusBar1.BoundControl(.cmbPanelIndex.ListIndex + 1, .cmbBound, .cmbSizeMethod.ListIndex)
        End If
        .Command3.Enabled = False
        .Command4.Enabled = True
    End With
End Sub

Private Sub Command4_Click()
    With Me
        Call .ucStatusBar1.BoundControl(.cmbPanelIndex.ListIndex + 1, Nothing, usbAutoSize)
        If Option2(0).Value Then
            .ucProgressBar1.Left = 240
            .ucProgressBar1.Top = 390
            .ucProgressBar1.Width = 1095
            .ucProgressBar1.Height = 255
        Else
            .cmbBound.Left = 1440
            .cmbBound.Top = 360
            .cmbBound.Width = 1095
            'cmbBound.Height = 255  'Read-Only!!
        End If
        .Command3.Enabled = True
        .Command4.Enabled = False
    End With
End Sub

Private Sub Form_Load()

    With Me
        .Caption = "ucStatusBar - v" & .ucStatusBar1.Version
        With .ucStatusBar1
            .AddPanel "Left Align", usbLeft, True, True, , , , , , , , 0
            .AddPanel "Left Align + Icon", usbLeft, True, , imgIcon(0).Picture, , , , , , 0
            .AddPanel "Fixed Size", usbCenter, False, , , , , , , , 120
            .AddPanel "Try resizing the control to see more...!!", usbLeft, True, , , , , , , , 0
            .AddPanel "Right Align + Icon", usbRight, True, True, imgIcon(1).Picture, , , , , , , 0
        End With
        With .cmbAlign
            .AddItem "usbLeft"
            .AddItem "usbCenter"
            .AddItem "usbRight"
            .ListIndex = 0
        End With
        With .cmbBound
            .AddItem 1
            .AddItem 2
            .AddItem 3
            .AddItem 4
            .AddItem 5
            .ListIndex = 0
        End With
        With .cmbPanelIndex
            .AddItem 1
            .AddItem 2
            .AddItem 3
            .AddItem 4
            .AddItem 5
            .ListIndex = 0
        End With
        With .cmbSizeMethod
            .AddItem "usbNoSize"
            .AddItem "usbAutoSize"
            .ListIndex = 1
        End With
        With .cmbTheme
            .AddItem "usbAuto"
            .AddItem "usbClassic"
            .AddItem "usbBlue"
            .AddItem "usbHomeStead"
            .AddItem "usbMetallic"
            .ListIndex = 0
        End With
    End With
End Sub

Private Function GetOption() As Long
    Dim i As Long
    
    With Me
        For i = 1 To .Option1.UBound
            If .Option1(i).Value Then
                GetOption = i
                Exit For
            End If
        Next
    End With
End Function

Private Sub ucStatusBar1_PanelClick(index As Long)
    cmbPanelIndex.ListIndex = index - 1
    Debug.Print "PanelClick:", index
End Sub

Private Sub ucStatusBar1_PanelDblClick(index As Long)
    Debug.Print "PanelDblClick:", index
End Sub

Private Sub ucStatusBar1_PanelMouseDown(index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "PanelMouseDown:", index, Button, Shift, X, Y
End Sub

Private Sub ucStatusBar1_PanelMouseMove(index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "PanelMouseMove:", index, Button, Shift, X, Y
End Sub

Private Sub ucStatusBar1_PanelMouseUp(index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "PanelMouseUp:", index, Button, Shift, X, Y
End Sub
