VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Start Button"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   2895
   End
   Begin VB.CommandButton btnStatus 
      Caption         =   "Disable Start button"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
' if fEnable=0 then you disable the window. if 1, you enable it.

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private xText(1) As String
Private Toggle As Long
Private StartButtonHandle As Long

Private Sub btnStatus_Click()
Dim Funk As Long

Toggle = (Toggle + 1) Mod 2

btnStatus.Caption = xText(Toggle)
Funk = EnableWindow(StartButtonHandle, Toggle)
End Sub

Private Sub Form_Load()
Dim hWParent As Long
            
Toggle = 1
                       
hWParent = FindWindow("Shell_TrayWnd", vbNullString)
StartButtonHandle = FindWindowEx(hWParent, 0&, "Button", vbNullString)
                                  
List1.AddItem "Start button handle : " + Trim(Str(StartButtonHandle))
            
xText(1) = "Disable Start button"
xText(0) = "Enable Start button"
            
End Sub
