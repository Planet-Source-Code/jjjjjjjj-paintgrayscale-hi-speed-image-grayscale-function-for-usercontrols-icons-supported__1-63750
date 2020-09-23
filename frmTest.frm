VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paint GrayScale !! - Test Form             "
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   761
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTimer 
      Caption         =   "Timer Enabled"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   5160
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Timer tmrDraw 
      Interval        =   1
      Left            =   1560
      Top             =   4800
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      Height          =   6615
      Left            =   5640
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   373
      TabIndex        =   15
      Top             =   360
      Width           =   5655
   End
   Begin VB.PictureBox picSrc 
      AutoSize        =   -1  'True
      Height          =   3705
      Left            =   120
      Picture         =   "frmTest.frx":000C
      ScaleHeight     =   3645
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   360
      Width           =   5445
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   120
      Picture         =   "frmTest.frx":3FE4
      ScaleHeight     =   840
      ScaleWidth      =   1125
      TabIndex        =   11
      Top             =   4320
      Width           =   1185
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Icon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton optBitmap 
      Caption         =   "Bitmap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   5415
      Begin VB.TextBox txtValue 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   9
         Text            =   "-1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   7
         Text            =   "-1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Text            =   "25"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtValue 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Text            =   "10"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "Draw Now"
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lbTime 
         AutoSize        =   -1  'True
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   195
         Left            =   3480
         TabIndex        =   10
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Top"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Left"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   270
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private cTimer  As New cTiming

Private Sub cmdDraw_Click()
    cTimer.Reset
    If optBitmap Then
        PaintGrayScale picDraw.hdc, picSrc.Picture, Val(txtValue(0)), Val(txtValue(1)), Val(txtValue(2)), Val(txtValue(3))
    Else
        PaintGrayScale picDraw.hdc, picIcon.Picture, Val(txtValue(0)), Val(txtValue(1)), Val(txtValue(2)), Val(txtValue(3))
    End If
    lbTime = "Process Time = " & Format$(cTimer.Elapsed / 1000, "0.0000 sec")
    picDraw.Refresh
End Sub





Private Sub chkTimer_Click()
    tmrDraw.Enabled = chkTimer
End Sub

Private Sub Form_Load()
    If Not FileExists(App.Path & "\" & App.EXEName & ".exe") Then
        MsgBox "Please compile before use.... There is about 400% of increase in speed when compiled!!", vbInformation, "Please compile !!"
        End
    End If
End Sub

Private Sub tmrDraw_Timer()
Dim lLeft As Long
Dim lTop As Long

    Randomize
    cTimer.Reset
    If optBitmap Then
        lLeft = Rnd * picDraw.Width - picSrc.Width / 2
        lTop = Rnd * picDraw.Height - picSrc.Width / 2
        PaintGrayScale picDraw.hdc, picSrc.Picture, lLeft, lTop
    Else
        lLeft = Rnd * picDraw.Width - picIcon.Width / 2
        lTop = Rnd * picDraw.Height - picIcon.Width / 2
        PaintGrayScale picDraw.hdc, picIcon.Picture, lLeft, lTop
    End If
    lbTime = "Process Time = " & Format$(cTimer.Elapsed / 1000, "0.0000 sec")
    picDraw.Refresh

End Sub


' Checks the existance of a file
Private Function FileExists(sFile As String) As Boolean
On Error GoTo Check
    If Trim(sFile) = "" Then
        FileExists = False
        Exit Function
    End If
    If Dir$(sFile, vbNormal) = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
Exit Function
Check:
    FileExists = False
End Function


