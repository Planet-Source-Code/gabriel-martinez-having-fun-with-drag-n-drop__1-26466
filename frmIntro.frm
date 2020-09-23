VERSION 5.00
Begin VB.Form frmIntro 
   Caption         =   "Drag & Drop"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3330
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1733
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "*  Type in the text box and drag and drop to desired list. (Won't let you add if element exists in either list)."
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label Label5 
      Caption         =   "*  To permanently remove items from trash can, choose     element(s) and hit delete key"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "*  To restore from deleted items, drag and drop to desired list"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "* To view Deletem Items, double click on trash can"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "* To remove from list drag and drop to the trash can"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "*  Drag and Drop from one list to another. (Multiselect)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me
    Form1.Show
    

End Sub


