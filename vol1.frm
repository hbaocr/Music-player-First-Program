VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vol"
   ClientHeight    =   1665
   ClientLeft      =   7740
   ClientTop       =   1485
   ClientWidth     =   390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   390
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider scoll3 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   2990
      _Version        =   393216
      OLEDropMode     =   1
      Orientation     =   1
      Min             =   1
      Max             =   5000
      SelStart        =   5000
      TickStyle       =   3
      Value           =   5000
      TextPosition    =   1
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
scoll3.Value = -gtc
End Sub

Private Sub scoll3_Scroll()
giatri = scoll3.Value
End Sub
