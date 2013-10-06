VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7695
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame WelcomeMsgFrame 
      Caption         =   "Some heading like thing here."
      Height          =   1215
      Left            =   960
      TabIndex        =   7
      Top             =   4560
      Width           =   7095
      Begin VB.Label WelcomeMsg 
         Caption         =   "Some random welcome message here."
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton JetsButton 
      Caption         =   "Jets"
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton TanksButton 
      Caption         =   "Tanks"
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton TransportButton 
      Caption         =   "Transport"
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton GearButton 
      Caption         =   "Gear"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton FirearmsButton 
      Caption         =   "Firearms"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label Heading 
      Alignment       =   2  'Center
      Caption         =   "Military Database"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Width           =   4815
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FirearmsButton_Click()

End Sub
