VERSION 5.00
Object = "{AF14FA88-F8B7-4E1F-9B2E-726C257104DE}#1.0#0"; "SNDPlayer.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1845
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   1845
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   9000
      Left            =   1920
      Top             =   675
   End
   Begin SNDPlayer.SoundPlayer spSound 
      Left            =   165
      Top             =   165
      _ExtentX        =   873
      _ExtentY        =   873
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Dim Disp As Long
Disp = GetDisplaySettings
CurrentIndex = lookupCurrent
ChangeScreenResolution Disp
PlayMusic 2
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Load frmMain
End Sub
