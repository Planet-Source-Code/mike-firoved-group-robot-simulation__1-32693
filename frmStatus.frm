VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Status"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   2145
      TabIndex        =   1
      Top             =   3480
      Width           =   2235
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "frmStatus.frx":0000
      Left            =   135
      List            =   "frmStatus.frx":0002
      TabIndex        =   0
      Top             =   780
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Updates every five seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1005
      TabIndex        =   2
      Top             =   120
      Width           =   4515
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub


