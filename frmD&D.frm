VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drag & Drop Sample"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   6300
   Icon            =   "frmD&D.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      OLEDragMode     =   1  'Automatic
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox lstDeleted 
      Height          =   1620
      Left            =   3720
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2400
      OLEDropMode     =   1  'Manual
      Picture         =   "frmD&D.frx":0442
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   2880
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   1920
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Gabo"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "List 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "List 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Drag & Drop to Lists"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Deleted Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim deleted As Boolean
Dim lst1 As ListBox
Dim lst2 As ListBox
Dim dbList As Database
Dim rstList As Recordset
Dim cnt As Integer
Dim obj As String
Dim vList As String

Private Sub Form_Load()

    Set dbList = OpenDatabase(App.Path & "\list.mdb")
    Set rstList = dbList.OpenRecordset("list")
    rstList.Index = "primarykey"
    Me.Width = 3690
    For i = 1 To 7
        List1.AddItem "Element " & i
    Next
    deleted = False
    dbList.Execute ("delete * from list")
    
End Sub
Private Sub List1_OLEDragDrop(DATA As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    addList List1, DATA

End Sub

Private Sub List1_OLEStartDrag(DATA As DataObject, AllowedEffects As Long)

    Set lst1 = List1
    List1.OLEDropMode = 0
    List2.OLEDropMode = 1
    Picture1.OLEDropMode = 1
    obj = "1"

End Sub
Private Sub List2_OLEDragDrop(DATA As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    addList List2, DATA

End Sub
Private Sub List2_OLEStartDrag(DATA As DataObject, AllowedEffects As Long)

    Set lst1 = List2
    List1.OLEDropMode = 1
    List2.OLEDropMode = 0
    Picture1.OLEDropMode = 1
    obj = "2"

End Sub

Private Sub lstDeleted_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 Then
        For i = lstDeleted.ListCount - 1 To 0 Step -1
            If lstDeleted.Selected(i) Then
                lstDeleted.RemoveItem (i)
            End If
        Next
        checkDeleted lstDeleted
    End If

End Sub
Private Sub lstDeleted_OLEStartDrag(DATA As DataObject, AllowedEffects As Long)

    Set lst1 = lstDeleted
    List1.OLEDropMode = 1
    List2.OLEDropMode = 1
    Picture1.OLEDropMode = 0
    obj = "D"

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuHelp_Click()

    frmIntro.Show vbModal

End Sub

Private Sub Picture1_DblClick()

    If Not deleted Then
        Me.Width = 5910
        lstDeleted.Visible = True
        deleted = True
    ElseIf deleted Then
        Me.Width = 3690
        lstDeleted.Visible = False
        deleted = False
    End If

End Sub
Private Sub Picture1_OLEDragDrop(DATA As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    addList lstDeleted, DATA
    Picture1.Picture = LoadPicture(App.Path & "\trash02b.ico")

End Sub
Private Function reOrder(nList As ListBox)

    With rstList
        For i = 0 To nList.ListCount - 1
            nList.ListIndex = i
            .AddNew
            !List = nList.Text
            .Update
        Next
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        nList.Clear
        Do While Not .EOF
            nList.AddItem !List
            .MoveNext
        Loop
    End With
    dbList.Execute ("delete * from list")

End Function
Private Sub checkDeleted(nList As ListBox)
    
    If nList.Name = "lstDeleted" Then
        If lstDeleted.ListCount < 1 Then
            Picture1.Picture = LoadPicture(App.Path & "\trash02a.ico")
            Me.Width = 3690
        End If
    End If

End Sub
Private Sub removeList()

    For i = lst1.ListCount - 1 To 0 Step -1
        If lst1.Selected(i) Then
            lst1.RemoveItem (i)
        End If
    Next

End Sub

Private Sub Text1_OLEStartDrag(DATA As DataObject, AllowedEffects As Long)

    Set lst1 = Nothing
    List1.OLEDropMode = 1
    List2.OLEDropMode = 1
    Picture1.OLEDropMode = 0
    obj = "T"

End Sub
Private Sub addList(nList As ListBox, dataS As DataObject)

    
    Select Case obj
        Case "1", "2"
            For i = 0 To lst1.ListCount - 1
                If lst1.Selected(i) Then
                    If Not checkExist(lst1.List(i), nList) Or nList.Name = "lstDeleted" Then
                        nList.AddItem lst1.List(i)
                    Else
                        MsgBox "Element  " & lst1.List(i) & " already exists in " & vList, vbInformation + vbOKOnly
                    End If
                End If
            Next
            removeList
            checkDeleted lst1
        Case "D"
            For i = 0 To lst1.ListCount - 1
                If lst1.Selected(i) Then
                    If Not checkExist(lst1.List(i), List1) Then
                        If Not checkExist(lst1.List(i), List2) Then
                            nList.AddItem lst1.List(i)
                        Else
                            MsgBox "Element  " & lst1.List(i) & " already exists in " & vList, vbInformation + vbOKOnly
                        End If
                    Else
                        MsgBox "Element  " & lst1.List(i) & " already exists in " & vList, vbInformation + vbOKOnly
                    End If
                End If
            Next
            removeList
            checkDeleted lst1
        Case "T"
            If Not checkExist(Text1.Text, List1) And Not checkExist(Text1.Text, List2) Then
                nList.AddItem Text1.Text
            Else
                MsgBox "Element already exists in  " & vList, vbInformation + vbOKOnly
                Exit Sub
            End If
    End Select
    If nList <> lstDeleted Then reOrder nList

End Sub
Private Function checkExist(ByVal DATA, nList As ListBox)

    If nList.Name <> "lstDeleted" Then
        
        For i = 0 To nList.ListCount - 1
            If DATA = nList.List(i) Then
                checkExist = True
                vList = nList.Name
                Exit Function
            End If
        Next
    End If
    checkExist = False

End Function
