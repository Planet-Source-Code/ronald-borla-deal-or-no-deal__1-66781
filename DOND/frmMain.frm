VERSION 5.00
Object = "{835C98B9-0E80-4239-851F-55937A7C92ED}#2.0#0"; "Lvbutton.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Deal Or No Deal"
   ClientHeight    =   11520
   ClientLeft      =   1980
   ClientTop       =   915
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fmeHide 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   1845
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   16
      Left            =   2985
      TabIndex        =   17
      Top             =   1935
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":08CA
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   11
      Left            =   2985
      TabIndex        =   12
      Top             =   3405
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":2408
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bNew 
      Height          =   660
      Index           =   0
      Left            =   585
      TabIndex        =   31
      Top             =   10605
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1164
      Caption         =   "New Game"
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
   Begin lvButton.lvButtons_H bDeal 
      Height          =   915
      Index           =   0
      Left            =   3390
      TabIndex        =   28
      Top             =   9195
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   1614
      Caption         =   "Deal"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   20.25
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
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   0
      Left            =   2070
      TabIndex        =   1
      Top             =   6345
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":3F46
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bAmt 
      Height          =   465
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   820
      Caption         =   "lvButtons_H1"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   8421631
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   1
      Left            =   3945
      TabIndex        =   2
      Top             =   6360
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":5A84
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   2
      Left            =   5820
      TabIndex        =   3
      Top             =   6345
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":75C2
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   3
      Left            =   7695
      TabIndex        =   4
      Top             =   6345
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":9100
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   4
      Left            =   9570
      TabIndex        =   5
      Top             =   6345
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":AC3E
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   5
      Left            =   11445
      TabIndex        =   6
      Top             =   6345
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":C77C
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   6
      Left            =   2985
      TabIndex        =   7
      Top             =   4875
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":E2BA
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   7
      Left            =   4860
      TabIndex        =   8
      Top             =   4875
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":FDF8
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   8
      Left            =   6735
      TabIndex        =   9
      Top             =   4875
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":11936
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   9
      Left            =   8610
      TabIndex        =   10
      Top             =   4875
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":13474
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   10
      Left            =   10485
      TabIndex        =   11
      Top             =   4875
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":14FB2
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   12
      Left            =   4860
      TabIndex        =   13
      Top             =   3405
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":16AF0
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   13
      Left            =   6735
      TabIndex        =   14
      Top             =   3405
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":1862E
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   14
      Left            =   8610
      TabIndex        =   15
      Top             =   3405
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":1A16C
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   15
      Left            =   10485
      TabIndex        =   16
      Top             =   3405
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":1BCAA
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   17
      Left            =   4860
      TabIndex        =   18
      Top             =   1935
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":1D7E8
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   18
      Left            =   6735
      TabIndex        =   19
      Top             =   1935
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":1F326
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   19
      Left            =   8610
      TabIndex        =   20
      Top             =   1935
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":20E64
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   20
      Left            =   10485
      TabIndex        =   21
      Top             =   1935
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":229A2
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   21
      Left            =   2985
      TabIndex        =   22
      Top             =   465
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":244E0
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   22
      Left            =   4860
      TabIndex        =   23
      Top             =   465
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":2601E
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   23
      Left            =   6735
      TabIndex        =   24
      Top             =   465
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":27B5C
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   24
      Left            =   8610
      TabIndex        =   25
      Top             =   465
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":2969A
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bCase 
      Height          =   1440
      Index           =   25
      Left            =   10485
      TabIndex        =   26
      Top             =   465
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   2540
      CapAlign        =   2
      BackStyle       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMain.frx":2B1D8
      ImgSize         =   48
      cBack           =   -2147483630
   End
   Begin lvButton.lvButtons_H bDeal 
      Height          =   915
      Index           =   1
      Left            =   9390
      TabIndex        =   29
      Top             =   9195
      Visible         =   0   'False
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   1614
      Caption         =   "No Deal"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   20.25
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
      Height          =   660
      Index           =   1
      Left            =   4545
      TabIndex        =   32
      Top             =   10605
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1164
      Caption         =   "Exit Game"
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
      Index           =   2
      Left            =   2550
      TabIndex        =   34
      Top             =   10605
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1164
      Caption         =   "Options"
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
   Begin VB.Label lblOffer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1,100,029"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   570
      Left            =   6315
      TabIndex        =   30
      Top             =   9330
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Briefcase"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   480
      Left            =   5310
      TabIndex        =   27
      Top             =   8085
      Width           =   4830
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const Amounts As String = "1,5,10,25,50,75,100,150,200,300,400,500,750," & _
                                   "1000 ,2500,5000,10000,25000,50000,100000,200000,300000," & _
                                   "400000,500000,1000000,2000000" 'refers to the amounts present in the game

'------------------------------------------------------------------------------------
'source: "http://mathforum.org/kb/servlet/JiveServlet/download/67-1388391-4739858-286755/att1.html"
Private Const lfPerRound As String = ".07,.09,.13,.17,.2,.25,.33,.5" 'refers to the low factors per round
Private Const hfPerRound As String = ".11,.16,.21,.26,.31,.41,.51,.61" 'refers to the high factors per round
Private Const lBase As Long = 50000
Private Const hBase As Long = 100000
'------------------------------------------------------------------------------------

Private Const genFactor As String = "1.18,1.28,1.34,1.38,1.4,1.47,1.56,1.67"
Dim CurAmounts As New Collection, CurRound As Integer, CurGFP As Double, ani As Boolean
Dim cc As Integer, ongame As Boolean, prevr As Integer, sel As Boolean
Dim bLeft As Double, bTop As Double, bIndex As Integer, Stun As Boolean, sMes As String
Dim Counter As Byte, gCount As Integer, Done As Boolean, cased As Integer
Dim nW As Double, nH As Double, pIndex As Byte, selecting As Boolean
Dim Loaded As Boolean

Private Sub ShowGFP()
Dim sP As Integer, GF As Double
sP = CurAmounts.Count
Select Case sP
    Case Is <= 13
        GF = ((sP * 2 - 2) / 100) + 1
    Case Is >= 14
        GF = ((Choose(sP - 13, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1) * 2 - 2) / 100) + 1
End Select
CurGFP = GF
End Sub

Private Sub NewGame()
If ani Then Exit Sub
If selecting Then Exit Sub
'If Not sel Then Exit Sub
lblStatus.Caption = "Working..."
lblStatus.Visible = True
fmeHide.Width = Me.ScaleWidth
fmeHide.Height = Me.ScaleHeight
fmeHide.Visible = True
If gCount <> 0 Then
    BringBCaseBack
End If
'If gCount = 0 Then Me.Hide
SetAmounts
ReloadAmounts
ongame = True
prevr = 0
Done = False
CurRound = 0
cc = 0
sMes = "The banker's offer is..."
bDeal(0).Visible = False
bDeal(1).Visible = False
'bDeal(0).Caption = "Deal"
'bDeal(1).Caption = "No Deal"
selecting = True
sel = False
bIndex = -1
Stun = False
ani = False
Loaded = True
Me.ZOrder 0
Me.Show
If gCount <> 0 Then
    PlayMusic 2
End If
lblStatus.Caption = "Select Your Briefcase"
lblOffer.Visible = False
fmeHide.Visible = False
If Offered.Count <> 0 Then
    While Offered.Count <> 0
        Offered.Remove 1
    Wend
End If
gCount = Count + 1
End Sub

Private Sub ReloadAmounts() 'to reload all amounts
Dim Nums() As String, i As Integer
While CurAmounts.Count <> 0
    CurAmounts.Remove 1
Wend
Nums = Split(Amounts, ",")
For i = LBound(Nums) To UBound(Nums)
    CurAmounts.Add Nums(i)
Next i
End Sub

Private Sub RemoveAmount(ByVal Number As String, ByRef Index As Integer) 'to remove _
                                a certain amount from list and return an index value
Dim i As Integer
For i = 1 To CurAmounts.Count
    If CurAmounts.Item(i) = Number Then
        Index = i
        CurAmounts.Remove Index
        Exit Sub
    End If
Next i
End Sub

Private Sub DoSelect()
selecting = True
lblStatus.Caption = "You chose briefcase number " & bCase(bIndex).Caption & "?"
bDeal(0).Caption = "Yes"
bDeal(1).Caption = "No"
DoStun
lblOffer.Visible = False
End Sub

Private Sub bCase_Click(Index As Integer)
If sel And Index = bIndex And cc <> 25 Then Exit Sub
If Stun Then Exit Sub
If ani Then Exit Sub
If Not ongame Then Exit Sub
Dim i As Integer, a As Long, b As Integer
'bNew(0).Caption = cc
'RemoveAmount bCase(Index).Tag, i
'txtOffer.Text = CurAmounts.Count
PlaySE 0
If cc = 0 Then
    bIndex = Index
    DoSelect
    Exit Sub
    'If selecting Then Exit Sub
    'SelectBCase (index)
    'Counter = 6
    'lblStatus.Caption = "Select " & Counter & " " & IIf(Counter = 1, "Briefcase", "Briefcases")
Else
    ani = True
    bCase(Index).Font.Size = 15
    bCase(Index).ForeColor = vbRed
    bCase(Index).Caption = Format(bCase(Index).Tag, "#,##0")
    PlaySE Result(bCase(Index).Tag)
    Pause 2000
    bCase(Index).Font.Size = 12
    bCase(Index).ForeColor = vbBlack
    
    For i = 1 To CurAmounts.Count
        If CurAmounts.Item(i) = bCase(Index).Tag Then
            CurAmounts.Remove i
            Exit For
        End If
    Next i
    For i = 0 To 25
        If bAmt(i).Tag = bCase(Index).Tag Then
            'bAmt(i).Enabled = False
            bAmt(i).Visible = False
        End If
    Next i
    bCase(Index).Visible = False
    ani = False
    
    Select Case cc
        Case 6
            ShowOffer
            Counter = 5
        Case 11
            ShowOffer
            Counter = 4
        Case 15
            ShowOffer
            Counter = 3
        Case 18
            ShowOffer
            Counter = 2
        Case 20, 21, 22, 23 ', 24
            ShowOffer
            Counter = 1
        'Case 24
            'ShowOffer
            'Exit Sub
        'Case 24
         '   ShowOffer
          '  Counter = 1
            'DoStun
        Case 24
            ShowOffer
            GoTo 1
            ani = True
            lblStatus.Caption = "We will now open your own briefcase"
            Pause 5000
            bCase(cased).Width = 300
            bCase(cased).Height = 250
            bCase(cased).Left = Me.ScaleWidth / 2 - bCase(cased).Width / 2
            bCase(cased).Top = Me.ScaleHeight / 2 - bCase(cased).Height / 2
            lblStatus.Caption = "And..."
            Pause 5000
            bCase(cased).Font.Size = 20
            bCase(cased).ForeColor = vbRed
            pIndex = bIndex
            bCase(cased).Caption = Format(bCase(cased).Tag, "#,##0")
            PlaySE 3
            Pause 2000
            For i = 0 To 25
                If i <> cased And bCase(i).Visible = True Then
                    bCase(i).ForeColor = vbRed
                    bCase(i).Caption = Format(bCase(i).Tag, "#,##0")
                End If
            Next i
            lblStatus.Caption = IIf(InStr(sMes, "been") <> 0, "You could have won " & Format(bCase(bIndex).Tag, "#,##0"), "You won " & Format(bCase(bIndex).Tag, "#,##0"))
            If InStr(sMes, "been") = 0 Then
                CurScore = Val(bCase(bIndex).Tag)
                SaveHighScores
            End If
            ongame = False
            ani = False
    End Select
    
    Counter = Counter - 1
    If Not Stun And cc <> 24 And Counter <> 0 Then lblStatus.Caption = "Select " & IIf(cc = 23, "Final", Counter) & " " & IIf(Counter = 1, "Briefcase", "Briefcases")
End If
1:
cc = cc + 1
End Sub

Private Sub Dealt()
        'Case 24
            Dim i As Integer
            ani = True
            lblStatus.Caption = "We will now open your own briefcase"
            Pause 5000
            'bCase(cased).Width = 300
            'bCase(cased).Height = 250
            bCase(cased).Left = Me.ScaleWidth / 2 - bCase(cased).Width / 2
            bCase(cased).Top = Me.ScaleHeight / 2 - bCase(cased).Height / 2
            lblStatus.Caption = "And..."
            Pause 5000
            'bCase(cased).Font.Size = 20
            bCase(cased).ForeColor = vbRed
            pIndex = bIndex
            bCase(cased).Caption = Format(bCase(cased).Tag, "#,##0")
            PlaySE Result(bCase(cased).Tag)
            Pause 2000
            For i = 0 To 25
                If i <> cased And bCase(i).Visible = True Then
                    bCase(i).ForeColor = vbRed
                    bCase(i).Caption = Format(bCase(i).Tag, "#,##0")
                End If
            Next i
            lblStatus.Caption = IIf(InStr(sMes, "been") <> 0, "You could have won " & Format(bCase(bIndex).Tag, "#,##0"), "You won " & Format(bCase(bIndex).Tag, "#,##0"))
            If InStr(sMes, "been") = 0 Then
                CurScore = Val(bCase(bIndex).Tag)
                SaveHighScores
            End If
            ongame = False
            ani = False
End Sub

Private Sub DoStun()
Stun = True
lblOffer.Visible = True
bDeal(0).Visible = True
bDeal(1).Visible = True
End Sub

Private Sub ShowOffer()
lblStatus.Caption = sMes
lblOffer.Caption = Format(Offer, "#,##0")
Offered.Add lblOffer.Caption
PlayMusic 0
PlaySE 4
DoStun
End Sub

Private Sub Pause(pLen As Long)
Sleep pLen
End Sub

Private Sub bDeal_Click(Index As Integer)
If ani Then Exit Sub
If selecting Then
    If Index = 0 Then
        selecting = False
        SelectBCase (bIndex)
        Counter = 6
        bDeal(0).Visible = False
        bDeal(1).Visible = False
        bDeal(0).Caption = "Deal"
        bDeal(1).Caption = "No Deal"
        lblStatus.Caption = "Select " & Counter & " " & IIf(Counter = 1, "Briefcase", "Briefcases")
        cc = 1
        Stun = False
        PlayMusic 1
    Else
        lblStatus.Caption = "Select Your Briefcase"
        bDeal(0).Visible = False
        bDeal(1).Visible = False
        selecting = True
        Stun = False
    End If
    Exit Sub
End If
If Not ongame Then
    Counter = Counter + 1
End If
Select Case Index
    Case 1
1:
        If cc = 25 Then
            lblOffer.Visible = False
            bDeal(0).Visible = False
            bDeal(1).Visible = False
            Dealt
            Exit Sub
        End If
        If Not Done Then sMes = "The banker's offer is..."
        Stun = False
        Counter = Counter + 1
        PlayMusic 1
        lblStatus.Caption = "Select " & Counter & " " & IIf(Counter = 1, "Briefcase", "Briefcases")
        lblOffer.Visible = False
        bDeal(0).Visible = False
        bDeal(1).Visible = False
    Case 0
        If cc = 25 And Not Done Then GoTo 2
        If Done Then GoTo 1
2:
        ani = True
        sMes = "The banker's offer could have been..."
        lblStatus.Caption = "You have won " & lblOffer.Caption
        CurScore = Val(Replace(lblOffer.Caption, ",", ""))
        SaveHighScores
        PlaySE 3
        Pause 2000
        lblStatus.Caption = "Let's see how much you could have won when you continue..."
        Pause 2000
        PlayMusic 1
        ani = False
        Stun = False
        If cc = 25 And Not Done Then
            lblOffer.Visible = False
            bDeal(0).Visible = False
            bDeal(1).Visible = False
            Dealt
            Exit Sub
        End If
        Counter = Counter + 1
        lblStatus.Caption = "Select " & Counter & " " & IIf(Counter = 1, "Briefcase", "Briefcases")
        lblOffer.Visible = False
        bDeal(0).Visible = False
        bDeal(1).Visible = False
        bDeal(0).Caption = "Continue"
        bDeal(1).Caption = "Continue"
        Done = True
End Select
End Sub

Private Sub bNew_Click(Index As Integer)
Select Case Index
    Case 0
        NewGame
    Case 1
        StopMusic
        Unload frmMain
        End
    Case 2
        frmOptions.Show vbModal
End Select
End Sub

Private Sub Form_Click()
'Unload Me
End Sub

Private Sub Form_Initialize()
Dim i As Integer, sLen As Integer
sLen = (Me.ScaleHeight - 100) / 13
bAmt(0).Top = sLen - sLen / 2 - bAmt(0).Height / 2
For i = 1 To 12
    Load bAmt(i)
    With bAmt(i)
        .Left = bAmt(0).Left
        .Top = sLen * (i + 1) - sLen / 2 - bAmt(i).Height / 2
        .Visible = True
    End With
    'frmSplash.Label1.Caption = "Loading Game... " & Int(((i + 1) / 26) * 100) & "%"
Next i
For i = 13 To 25
    Load bAmt(i)
    With bAmt(i)
        If i = 13 Then
            .Left = Me.ScaleWidth - bAmt(0).Left - .Width
            .Top = bAmt(0).Top
            .Visible = True
        Else
            .Left = bAmt(13).Left
            .Top = bAmt(i - 13).Top
            .Visible = True
        End If
    End With
    'frmSplash.Label1.Caption = "Loading Game... " & Int(((i + 1) / 26) * 100) & "%"
Next i
PName = "Player 1"
NewGame
frmSplash.Hide
Me.Show
End Sub

Private Sub SetAmounts()
Randomize
Dim i As Integer, sTemp() As String, nCol As New Collection, rndVal As Integer
Dim dSt As Double, dTop As Integer
dTop = 20
sTemp = Split(Amounts, ",")
For i = 0 To 25
    nCol.Add sTemp(i)
Next i
dSt = bAmt(25).Left - bAmt(1).Left + bAmt(1).Width
For i = 0 To 25
    'If i <> 0 Then Load bCase(i)
    'Select Case i
    '    Case 0 To 5
    '        bCase(i).Left = dSt / 6 - bCase(i).Width / 2
    '        bCase(i).Top = bCase(i).Height * 4 + dTop
    '    Case 6 To 10
    '        bCase(i).Left = dSt / 6 - bCase(i).Width / 2
    '        bCase(i).Top = bCase(i).Height * 4 + dTop
    '    Case 11 To 15
    '        bCase(i).Left = dSt / 6 - bCase(i).Width / 2
    '        bCase(i).Top = bCase(i).Height * 4 + dTop
    '    Case 16 To 20
    '        bCase(i).Left = dSt / 6 - bCase(i).Width / 2
    '        bCase(i).Top = bCase(i).Height * 4 + dTop
    '    Case 21 To 25
    '        bCase(i).Left = dSt / 6 - bCase(i).Width / 2
    '        bCase(i).Top = bCase(i).Height * 4 + dTop
    'End Select
    bAmt(i).Tag = sTemp(i)
    bAmt(i).Caption = Format(sTemp(i), "#,##0")
    bAmt(i).Visible = True
    bCase(i).Caption = i + 1
    bCase(i).ForeColor = vbBlack
    bCase(i).Font.Size = 12
    rndVal = Int(Rnd() * nCol.Count)
    bCase(i).Tag = nCol.Item(rndVal + 1)
    nCol.Remove rndVal + 1
    bCase(i).Visible = True
    'frmSplash.Label1.Caption = "Loading Game... " & Int(((i + 1) / 26) * 100) & "%"
Next i
End Sub

Private Sub Form_Load()
'Dim Disp As Long
'Disp = GetDisplaySettings
'CurrentIndex = lookupCurrent
'ChangeScreenResolution Disp
Me.ScaleMode = vbPixels
nW = bCase(0).Width
nH = bCase(0).Height
LoadHighScores
End Sub

Private Sub Form_Unload(Cancel As Integer)
ChangeScreenResolution CurrentIndex
End Sub

'----------COMPUTATIONS-----------
'------------------------------------------------------------------------------------
'source: "http://mathforum.org/kb/servlet/JiveServlet/download/67-1388391-4739858-286755/att1.html"
Private Function cmpFactor() As Double
Dim LowFactor As Double, HighFactor As Double, fTemp() As String, i As Integer, _
    lSum As Double, hSum As Double
fTemp = Split(lfPerRound, ",")
LowFactor = CDbl(fTemp(CurRound - 1))
Erase fTemp
fTemp = Split(hfPerRound, ",")
HighFactor = CDbl(fTemp(CurRound - 1))
'mali
For i = 1 To CurAmounts.Count
    If CurAmounts.Item(i) <= lBase Then lSum = lSum + Val(CurAmounts.Item(i))
    If CurAmounts.Item(i) = lBase Then Exit For
Next i
lSum = lSum * LowFactor

For i = CurAmounts.Count To 1 Step -1
    If CurAmounts.Item(i) >= hBase Then hSum = hSum + Val(CurAmounts.Item(i))
    If CurAmounts.Item(i) = hBase Then Exit For
Next i
hSum = hSum * LowFactor * HighFactor

cmpFactor = lSum + hSum
End Function
'------------------------------------------------------------------------------------

Private Function cmpOthers() As Double
Randomize
Dim rPick As Byte, Sum As Double, i As Integer, Factor As Double, fTemp() As String
rPick = Int(Rnd() * 3 + 1)
rPick = 2
Select Case rPick
    Case 1
        For i = 1 To CurAmounts.Count
            Sum = Sum + Val(CurAmounts.Item(i))
        Next i
        Sum = Sum / CurAmounts.Count
    Case 2
        For i = 1 To CurAmounts.Count
            Sum = Sum + Sqr(Val(CurAmounts.Item(i)))
        Next i
        Sum = (Sum / CurAmounts.Count) ^ 2
    Case 3
        Sum = CurAmounts.Item(1)
        For i = 2 To CurAmounts.Count
            Sum = (Sum + CurAmounts.Item(i)) / 2
        Next i
End Select
'fTemp = Split(genFactor, ",")
'Factor = CDbl(fTemp(CurRound - 1))
cmpOthers = Sum * CurGFP
End Function

Private Function Offer() As Double
ShowGFP
Offer = Int(cmpOthers)
End Function
'----------COMPUTATIONS-----------

Private Sub SelectBCase(ByVal Index As Integer)
bIndex = Index
bLeft = bCase(Index).Left
bTop = bCase(Index).Top
sel = True
BringBCase
cased = bIndex
End Sub

Private Sub BringBCaseBack()
bCase(bIndex).Left = bLeft
bCase(bIndex).Top = bTop
bCase(bIndex).Font.Size = 12
bCase(bIndex).ForeColor = vbBlack
bCase(bIndex).Width = nW
bCase(bIndex).Height = nH
End Sub

Private Sub BringBCase()
bCase(bIndex).Left = bDeal(1).Left + bDeal(1).Width + 50
bCase(bIndex).Top = bDeal(1).Top + bCase(bIndex).Height / 2
End Sub

Private Function Result(ByVal Value As Long) As Integer
Result = IIf(Value >= 100000, 2, 1)
If Done Then Result = IIf(Result = 2, 1, 2)
End Function

