VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form2"
   ScaleHeight     =   8370
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton NextButton1 
      Caption         =   "Next >>"
      Height          =   495
      Left            =   3120
      TabIndex        =   29
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton PreviousButton1 
      Caption         =   "<< Previous"
      Height          =   495
      Left            =   840
      TabIndex        =   28
      Top             =   2520
      Width           =   1935
   End
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
      TabIndex        =   26
      Text            =   "Text7"
      Top             =   3585
      Width           =   4095
   End
   Begin VB.CommandButton SearchButton 
      Caption         =   "Go"
      Height          =   375
      Left            =   4440
      TabIndex        =   23
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
      TabIndex        =   22
      Text            =   "Search Firearms..."
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CommandButton DeleteButton 
      Caption         =   "Delete"
      Height          =   615
      Left            =   3120
      TabIndex        =   21
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "Add"
      Height          =   615
      Left            =   840
      TabIndex        =   20
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Data Data1 
      BOFAction       =   1  'BOF
      Caption         =   "Browse"
      Connect         =   "Access"
      DatabaseName    =   "C:\Visual-Basic-Project\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   1  'EOF
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
      RecordSource    =   "Firearms"
      Top             =   360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text6 
      DataField       =   "Units Available"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text5 
      DataField       =   "Recoil"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text4 
      DataField       =   "Capacity"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text3 
      DataField       =   "Fire Rate"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text2 
      DataField       =   "Fire Power"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataField       =   "Type"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4440
      Width           =   975
   End
   Begin VB.PictureBox Preview 
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
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "FirearmsForm.frx":0000
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CommandButton FirearmsButton 
      Caption         =   "Firearms"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton GearButton 
      Caption         =   "Gear"
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
      TabIndex        =   27
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label TableName 
      Alignment       =   2  'Center
      Caption         =   "Firearms"
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
      TabIndex        =   24
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label6 
      Caption         =   "Units Avaialable"
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
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Recoil"
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
      Left            =   2880
      TabIndex        =   18
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Capacity"
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
      Left            =   2880
      TabIndex        =   17
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Fire Rate"
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
      TabIndex        =   16
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Fire Power"
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
      Height          =   495
      Left            =   720
      TabIndex        =   15
      Top             =   5040
      Width           =   855
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
      TabIndex        =   14
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddButton_Click()

Data1.Recordset.AddNew

End Sub

Private Sub DeleteButton_Click()
Data1.Recordset.Delete
MsgBox "Current Record has been deleted."
End Sub

Private Sub ExitButton_Click()
End
End Sub

Private Sub GearButton_Click()
Form3.Show
Form2.Hide
End Sub

Private Sub JetsButton_Click()
Form6.Show
Form2.Hide
End Sub

Private Sub NextButton1_Click()
    
    Dim imagePath As String
    Dim imageIsNull As Boolean
    
     
    Data1.Recordset.MoveNext
    
    If Data1.Recordset.EOF = True Then
        Data1.Recordset.MoveFirst
    End If
    
    'Image null check code
    
    
    'If path is empty
    imagePath = Data1.Recordset.Fields("Image")
    
    If imagePath = "" Then
        imageIsNull = True
    End If
    
    
    If imageIsNull Then
        Preview.Picture = LoadPicture("C:\Visual-Basic-Project\Images\default.jpg")
    End If
    
    ' if path is not empty BUT the mentioned file doesn't exist
    
    If Not imageIsNull And Dir(imagePath) <> "" Then
       Preview.Picture = LoadPicture(imagePath)
    Else
        Preview.Picture = LoadPicture("C:\Visual-Basic-Project\Images\default.jpg")
    End If
       
End Sub

Private Sub PreviousButton1_Click()

    Dim imagePath As String
    Dim imageIsNull As Boolean
    
    Data1.Recordset.MovePrevious
    
    If Data1.Recordset.BOF Then
        Data1.Recordset.MoveLast
    End If
    
    'Image null check code
    
    
    'If path is empty
    imagePath = Data1.Recordset.Fields("Image")
    
    If imagePath = "" Then
        imageIsNull = True
    End If
    
    
    If imageIsNull Then
        Preview.Picture = LoadPicture("C:\Visual-Basic-Project\Images\default.jpg")
    End If
    
    ' if path is not empty BUT the mentioned file doesn't exist
    
    If Not imageIsNull And Dir(imagePath) <> "" Then
       Preview.Picture = LoadPicture(imagePath)
    Else
        Preview.Picture = LoadPicture("C:\Visual-Basic-Project\Images\default.jpg")
    End If
    
End Sub

Private Sub TanksButton_Click()
Form5.Show
Form2.Hide
End Sub

Private Sub TransportButton_Click()
Form4.Show
Form2.Hide
End Sub
