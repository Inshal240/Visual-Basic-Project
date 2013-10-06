VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form3"
   ScaleHeight     =   8370
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ItemName 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
      DataField       =   "Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   840
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   3585
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      DataField       =   "Type"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton SearchButton 
      Caption         =   "Go"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox SearchBox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Text            =   "Search Gear..."
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CommandButton DeleteButton 
      Caption         =   "Delete"
      Height          =   735
      Left            =   3120
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "Add"
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Browse"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\XPMUser\My Documents\GitHub\Visual-Basic-Project\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Gear"
      Top             =   2640
      Width           =   4095
   End
   Begin VB.PictureBox Picture 
      Height          =   3015
      Left            =   5400
      ScaleHeight     =   2955
      ScaleWidth      =   3675
      TabIndex        =   7
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Frame Description 
      Caption         =   "Description"
      Height          =   2055
      Left            =   5400
      TabIndex        =   6
      Top             =   4320
      Width           =   3735
      Begin VB.TextBox DescriptionBox 
         DataField       =   "Description"
         DataSource      =   "Data1"
         Height          =   1455
         Left            =   240
         TabIndex        =   13
         Text            =   "Description goes here."
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CommandButton FirearmsButton 
      Caption         =   "Firearms"
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton GearButton 
      Caption         =   "Gear"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton TransportButton 
      Caption         =   "Transport"
      Height          =   615
      Left            =   6240
      TabIndex        =   3
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton TanksButton 
      Caption         =   "Tanks"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton JetsButton 
      Caption         =   "Jets"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6240
      TabIndex        =   0
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Type"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   15
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label TableName 
      Alignment       =   2  'Center
      Caption         =   "Gear"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExitButton_Click()
End
End Sub

Private Sub FirearmsButton_Click()
Form2.Show
Form3.Hide
End Sub

Private Sub JetsButton_Click()
Form6.Show
Form3.Hide
End Sub

Private Sub TanksButton_Click()
Form5.Show
Form3.Hide
End Sub

Private Sub TransportButton_Click()
Form4.Show
Form3.Hide
End Sub

