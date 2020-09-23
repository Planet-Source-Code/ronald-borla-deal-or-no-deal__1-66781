VERSION 5.00
Object = "{835C98B9-0E80-4239-851F-55937A7C92ED}#2.0#0"; "Lvbutton.ocx"
Object = "{C53EC27C-82C0-4D47-8B53-DCFED34236B0}#1.0#0"; "volControl.ocx"
Begin VB.Form frmOptions 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H bNew 
      Default         =   -1  'True
      Height          =   660
      Index           =   0
      Left            =   6825
      TabIndex        =   0
      Top             =   5115
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1164
      Caption         =   "Done"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16512
      cGradient       =   16512
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   8421631
   End
   Begin lvButton.lvButtons_H bNew 
      Cancel          =   -1  'True
      Height          =   660
      Index           =   1
      Left            =   405
      TabIndex        =   1
      Top             =   5115
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1164
      Caption         =   "Back"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16512
      cGradient       =   16512
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   8421631
   End
   Begin lvButton.lvButtons_H bOpt 
      Height          =   465
      Index           =   0
      Left            =   285
      TabIndex        =   2
      Top             =   300
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   820
      Caption         =   "Player Options"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483624
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   16512
      Gradient        =   3
      Mode            =   2
      Value           =   -1  'True
      cBack           =   8421631
   End
   Begin lvButton.lvButtons_H bOpt 
      Height          =   465
      Index           =   1
      Left            =   285
      TabIndex        =   3
      Top             =   960
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   820
      Caption         =   "High Scores"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483624
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   16512
      Gradient        =   3
      Mode            =   2
      Value           =   0   'False
      cBack           =   8421631
   End
   Begin lvButton.lvButtons_H bOpt 
      Height          =   465
      Index           =   2
      Left            =   285
      TabIndex        =   4
      Top             =   1605
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   820
      Caption         =   "Sound"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483624
      Focus           =   0   'False
      LockHover       =   1
      cGradient       =   16512
      Gradient        =   3
      Mode            =   2
      Value           =   0   'False
      cBack           =   8421631
   End
   Begin VB.Frame fmeOpt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4575
      Index           =   0
      Left            =   2985
      TabIndex        =   5
      Top             =   195
      Width           =   5700
      Begin VB.ListBox lstOffers 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   2310
         Left            =   555
         TabIndex        =   9
         Top             =   1770
         Width           =   4575
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   600
         Left            =   1845
         MaxLength       =   8
         TabIndex        =   6
         Text            =   "Player 1"
         Top             =   390
         Width           =   3540
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Offers:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   1
         Left            =   510
         TabIndex        =   8
         Top             =   1050
         Width           =   2820
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   0
         Left            =   510
         TabIndex        =   7
         Top             =   450
         Width           =   1140
      End
   End
   Begin VB.Frame fmeOpt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4575
      Index           =   2
      Left            =   2985
      TabIndex        =   18
      Top             =   195
      Width           =   5700
      Begin lvButton.lvButtons_H bVol 
         Height          =   150
         Index           =   0
         Left            =   660
         TabIndex        =   19
         Top             =   1830
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   265
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   210
         Index           =   1
         Left            =   1065
         TabIndex        =   20
         Top             =   1770
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   370
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   285
         Index           =   2
         Left            =   1470
         TabIndex        =   21
         Top             =   1695
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   345
         Index           =   3
         Left            =   1875
         TabIndex        =   22
         Top             =   1635
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   609
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   420
         Index           =   4
         Left            =   2280
         TabIndex        =   23
         Top             =   1560
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   741
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   495
         Index           =   5
         Left            =   2685
         TabIndex        =   24
         Top             =   1485
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   570
         Index           =   6
         Left            =   3090
         TabIndex        =   25
         Top             =   1410
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   1005
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   660
         Index           =   7
         Left            =   3495
         TabIndex        =   26
         Top             =   1320
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   1164
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   750
         Index           =   8
         Left            =   3900
         TabIndex        =   27
         Top             =   1230
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   1323
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   840
         Index           =   9
         Left            =   4305
         TabIndex        =   28
         Top             =   1140
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   1482
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bVol 
         Height          =   960
         Index           =   10
         Left            =   4710
         TabIndex        =   29
         Top             =   1020
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   1693
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
      Begin lvButton.lvButtons_H bMute 
         Height          =   510
         Left            =   1680
         TabIndex        =   30
         Top             =   2370
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         Caption         =   "Disable Sound"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   -2147483624
         Focus           =   0   'False
         LockHover       =   1
         cGradient       =   16512
         Gradient        =   3
         Mode            =   1
         Value           =   0   'False
         cBack           =   8421631
      End
   End
   Begin VB.Frame fmeOpt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4575
      Index           =   1
      Left            =   2985
      TabIndex        =   10
      Top             =   195
      Width           =   5700
      Begin VB.Label lblHS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Index           =   2
         Left            =   4785
         TabIndex        =   17
         Top             =   3105
         Width           =   225
      End
      Begin VB.Label lblHS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Index           =   1
         Left            =   4785
         TabIndex        =   16
         Top             =   2205
         Width           =   225
      End
      Begin VB.Label lblHS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Index           =   0
         Left            =   4785
         TabIndex        =   15
         Top             =   1305
         Width           =   225
      End
      Begin VB.Label lblHSN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player 3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   2
         Left            =   795
         TabIndex        =   14
         Top             =   3105
         Width           =   1440
      End
      Begin VB.Label lblHSN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   1
         Left            =   795
         TabIndex        =   13
         Top             =   2205
         Width           =   1440
      End
      Begin VB.Label lblHSN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   0
         Left            =   795
         TabIndex        =   12
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top Dealers:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   2
         Left            =   375
         TabIndex        =   11
         Top             =   510
         Width           =   2190
      End
   End
   Begin VB.Shape shpBor 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   6000
      Left            =   30
      Top             =   30
      Width           =   8940
   End
   Begin VB.Image imgLogo 
      Height          =   1830
      Left            =   285
      Picture         =   "frmOptions.frx":0000
      Top             =   2505
      Width           =   2430
   End
   Begin volControl.VolumeControl vcDeal 
      Left            =   330
      Top             =   4410
      _ExtentX        =   847
      _ExtentY        =   847
      Volume          =   100
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bMute_Click()
bMute.Caption = IIf(Not bMute.Value, "Disable ", "Enable ") & "Sound"
vcDeal.Volume = IIf(Not bMute.Value, Vol * 10, 0)
Mute = bMute.Value
End Sub

Private Sub bNew_Click(Index As Integer)
Select Case Index
    Case 1
        Unload Me
    Case 0
        PName = txtName.Text
        Unload Me
End Select
End Sub

Private Sub bOpt_Click(Index As Integer)
Dim i As Integer
fmeOpt(Index).ZOrder 0
LoadHighScores
For i = 0 To 2
    lblHSN(i).Caption = HSName(i)
    lblHS(i).Caption = HSScore(i)
Next i
End Sub

Private Sub bVol_Click(Index As Integer)
Dim i As Integer
For i = 0 To Index
    bVol(i).Value = False
Next i
If Index + 1 <> 11 Then
    For i = Index + 1 To 10
        bVol(i).Value = True
    Next i
End If
vcDeal.Volume = Index * 10
Vol = Index
End Sub

Private Sub Form_Load()
Dim i As Integer
vcDeal.DeviceToControl = mWave
Vol = vcDeal.Volume \ 10
bVol_Click Vol
txtName.Text = PName
bMute.Value = Mute
If Offered.Count = 0 Then Exit Sub
For i = 1 To Offered.Count
    lstOffers.AddItem Offered.Item(i)
Next i
LoadHighScores
For i = 0 To 2
    lblHSN(i).Caption = HSName(i)
    lblHS(i).Caption = HSScore(i)
Next i
End Sub
