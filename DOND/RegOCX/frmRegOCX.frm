VERSION 5.00
Begin VB.Form frmRegOCX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register OCX"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   915
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DROP OCX HERE TO REGISTER IT IN YOUR SYSTEM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   315
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   300
      Width           =   8265
   End
End
Attribute VB_Name = "frmRegOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Load()
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 1 To Data.Files.Count
    If LCase(Right(Data.Files.Item(i), 4)) = ".ocx" Then
        FileCopy Data.Files.Item(i), Dir(Data.Files.Item(i))
        Shell "regsvr32 " & Dir(Data.Files.Item(i))
    End If
Next i
End Sub

Private Sub Label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 1 To Data.Files.Count
    If LCase(Right(Data.Files.Item(i), 4)) = ".ocx" Then
        FileCopy Data.Files.Item(i), Dir(Data.Files.Item(i))
        Shell "regsvr32 " & Dir(Data.Files.Item(i))
    End If
Next i
End Sub
