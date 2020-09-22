VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTest.frx":0000
      Left            =   60
      List            =   "frmTest.frx":000D
      TabIndex        =   2
      Text            =   "http://www.microsoft.com"
      Top             =   0
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get"
      Height          =   255
      Left            =   5700
      TabIndex        =   1
      Top             =   60
      Width           =   435
   End
   Begin VB.TextBox Text2 
      Height          =   3795
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   6075
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MousePointer = vbHourglass
    Text2.Text = GetUrlSource(Combo1.Text)
    MousePointer = vbDefault
End Sub

