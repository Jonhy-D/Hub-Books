VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormLibrary 
   Caption         =   "formLibrary"
   ClientHeight    =   7164
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   11256
   LinkTopic       =   "Form1"
   ScaleHeight     =   7164
   ScaleWidth      =   11256
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textSinopsis 
      Height          =   1332
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2760
      Width           =   4572
   End
   Begin MSComctlLib.ListView listBooks 
      Height          =   2652
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   9972
      _ExtentX        =   17590
      _ExtentY        =   4678
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CommandSearchBook 
      Caption         =   "Search"
      Height          =   372
      Left            =   5160
      TabIndex        =   8
      Top             =   2040
      Width           =   2052
   End
   Begin VB.CommandButton CommandDeleteBook 
      Caption         =   "Delete Book"
      Height          =   372
      Left            =   5160
      TabIndex        =   7
      Top             =   1440
      Width           =   2052
   End
   Begin VB.CommandButton CommandAddBook 
      Caption         =   "Add Book"
      Height          =   612
      Left            =   5160
      TabIndex        =   6
      Top             =   600
      Width           =   2052
   End
   Begin VB.TextBox textGenre 
      Height          =   372
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   4572
   End
   Begin VB.TextBox textAuthor 
      Height          =   372
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   4572
   End
   Begin VB.TextBox textTitle 
      Height          =   372
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   4572
   End
   Begin VB.Label labelAvailable 
      Alignment       =   2  'Center
      Caption         =   "Available Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8520
      TabIndex        =   13
      Top             =   4200
      Width           =   2052
   End
   Begin VB.Image Image1 
      Height          =   2532
      Left            =   7800
      Top             =   1440
      Width           =   2892
   End
   Begin VB.Label labelCatalog 
      Alignment       =   2  'Center
      Caption         =   "Catalog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4920
      TabIndex        =   12
      Top             =   0
      Width           =   1572
   End
   Begin VB.Label labelSinopsis 
      Caption         =   "Sinopsis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   732
   End
   Begin VB.Image frontPage 
      Height          =   732
      Left            =   7800
      Top             =   600
      Width           =   2892
   End
   Begin VB.Label labelGenre 
      Caption         =   "Genre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   612
   End
   Begin VB.Label labelAuthor 
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   612
   End
   Begin VB.Label labelTitle 
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   372
   End
   Begin VB.Menu profile 
      Caption         =   "Profile"
   End
End
Attribute VB_Name = "formLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandAddBook_Click()
    Dim newBook As New Book
    newBook.Title = textTitle.Text
    newBook.Author = textAuthor.Text
    newBook.Genre = textGenre.Text
    newBook.Sinopsis = textSinopsis.Text
        
    Call Connect
    conn.Execute "INSERT INTO books(title, author, genre, sinopsis) values('" & newBook.Title & "', '" & newBook.Author & "', '" & newBook.Genre & "', '" & newBook.Sinopsis & "')"
    Call Disconnect
    MsgBox "Book added successfully"
    
    textTitle.Text = ""
    textAuthor.Text = ""
    textGenre.Text = ""

End Sub

Private Sub CommandDeleteBook_Click()
    If Not listBooks.SelectedItem Is Nothing Then
        listBooks.ListItems.Remove listBooks.SelectedItem.Index
        MsgBox "Book removed with success", vbInformation
    Else
        MsgBox "You need to select a book for remove", vbExclamation
    End If
End Sub

Private Sub CommandSearchBook_Click()
    Dim searchTitle As String
    Dim item As ListItem
    searchTitle = InputBox("Enter the title of the book to search for:", "Search Book")
    
    For Each item In listBooks.ListItems
        If LCase(item.Text) = LCase(searchTitle) Then
            item.Selected = True
            item.EnsureVisible
            Exit For
        End If
    Next
    
    If listBooks.SelectedItem Is Nothing Then
        MsgBox "Book couldn't found", vbExclamation
    End If
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Call Connect
        
    rs.open "SELECT * FROM books", conn, adOpenStatic, adLockReadOnly
    
    listBooks.ListItems.Clear
    
    Do Until rs.EOF
        Dim item As ListItem
        Set item = listBooks.ListItems.Add(, , rs!bookId)
        item.ListSubItems.Add , , rs!Title
        item.ListSubItems.Add , , rs!Author
        item.ListSubItems.Add , , rs!Genre
        item.ListSubItems.Add , , rs!Sinopsis
        
        rs.MoveNext
    Loop
    Call Disconnect
    
    With listBooks
        .View = lvwReport
            .ColumnHeaders.Add , , "Id", 500
            .ColumnHeaders.Add , , "Title", 3000
            .ColumnHeaders.Add , , "Author", 3000
            .ColumnHeaders.Add , , "Genre", 3000
            .ColumnHeaders.Add , , "Sinopsis", 3000
    End With
    
End Sub


Private Sub profile_Click()
    FormUser.Show
End Sub
