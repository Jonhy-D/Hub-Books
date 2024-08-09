VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormUser 
   Caption         =   "formUser"
   ClientHeight    =   7440
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   12540
   LinkTopic       =   "Form2"
   ScaleHeight     =   7440
   ScaleWidth      =   12540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bookRead 
      Caption         =   "Book Read"
      Height          =   732
      Left            =   10440
      TabIndex        =   12
      Top             =   2400
      Width           =   1452
   End
   Begin MSComctlLib.ListView listBooks 
      Height          =   3492
      Left            =   6240
      TabIndex        =   11
      Top             =   960
      Width           =   4092
      _ExtentX        =   7218
      _ExtentY        =   6160
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
   Begin MSComctlLib.ListView listFavorites 
      Height          =   3372
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   5948
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
   Begin VB.CommandButton addFavorite 
      Caption         =   "Add Favorite"
      Height          =   612
      Left            =   10440
      TabIndex        =   9
      Top             =   960
      Width           =   1452
   End
   Begin VB.CommandButton unLiked 
      Caption         =   "Unliked Book"
      Height          =   612
      Left            =   10440
      TabIndex        =   8
      Top             =   1680
      Width           =   1452
   End
   Begin MSComctlLib.ListView listUnliked 
      Height          =   2172
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   5772
      _ExtentX        =   10181
      _ExtentY        =   3831
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
   Begin MSComctlLib.ListView listRead 
      Height          =   2292
      Left            =   6840
      TabIndex        =   4
      Top             =   5040
      Width           =   5292
      _ExtentX        =   9335
      _ExtentY        =   4043
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
   Begin VB.CommandButton unFavorite 
      Caption         =   "Unfavorite"
      Height          =   612
      Left            =   3960
      TabIndex        =   2
      Top             =   1680
      Width           =   1452
   End
   Begin VB.CommandButton seeBook 
      Caption         =   "See Book"
      Height          =   612
      Left            =   3960
      TabIndex        =   1
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label libraryBooks 
      Alignment       =   2  'Center
      Caption         =   "Library of Books"
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
      Left            =   7560
      TabIndex        =   13
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label labelBadBooks 
      Alignment       =   2  'Center
      Caption         =   "Not My Favorite Reads"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   4680
      Width           =   2172
   End
   Begin VB.Label labelBooksRead 
      Alignment       =   2  'Center
      Caption         =   "Books Read"
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
      Left            =   8640
      TabIndex        =   5
      Top             =   4680
      Width           =   1452
   End
   Begin VB.Label labelTitle 
      Alignment       =   2  'Center
      Caption         =   "Profile"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   1452
   End
   Begin VB.Label labelFavorites 
      Caption         =   "Favorites Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   1452
   End
   Begin VB.Menu profile 
      Caption         =   "Profile"
   End
   Begin VB.Menu library 
      Caption         =   "Library"
   End
End
Attribute VB_Name = "FormUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addFavorite_Click()
    If Not listBooks.SelectedItem Is Nothing Then
        Dim bookId As Integer
        Dim userId As Integer
        bookId = listBooks.SelectedItem.Text
        userId = 1
        
        Call Connect
        conn.Execute "INSERT INTO bookFavorite(bookId, userId) values(" & bookId & ", " & userId & ")"
        Call Disconnect
        MsgBox "Book added successfully"
    End If
End Sub

Private Sub bookRead_Click()
    If Not listBooks.SelectedItem Is Nothing Then
        Dim bookId As Integer
        Dim userId As Integer
        bookId = listBooks.SelectedItem.Text
        userId = 1
        
        Call Connect
        conn.Execute "INSERT INTO readBooks(bookId, userId) values(" & bookId & ", " & userId & ")"
        Call Disconnect
        MsgBox "Book added successfully"
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim rsf As ADODB.Recordset
    Set rsf = New ADODB.Recordset
    
    Dim rsuF As ADODB.Recordset
    Set rsuF = New ADODB.Recordset
    
    Dim rsr As ADODB.Recordset
    Set rsr = New ADODB.Recordset
    
    Call Connect
    
    rs.open "SELECT * FROM books", conn, adOpenStatic, adLockReadOnly
    rsf.open "SELECT bk.bookId ,bk.title,bk.author,bk.genre,bk.sinopsis FROM bookFavorite as bf inner join books as bk on bf.bookId = bk.bookId WHERE bf.userId = 1", conn, adOpenStatic, adLockReadOnly
    rsuF.open "SELECT bk.bookId ,bk.title,bk.author,bk.genre,bk.sinopsis FROM unlikedBooks as ub inner join books as bk on ub.bookId = bk.bookId WHERE ub.userId = 1", conn, adOpenStatic, adLockReadOnly
    rsr.open "SELECT bk.bookId ,bk.title,bk.author,bk.genre,bk.sinopsis FROM readBooks as rb inner join books as bk on rb.bookId = bk.bookId WHERE rb.userId = 1", conn, adOpenStatic, adLockReadOnly
    
    listBooks.ListItems.Clear
    listFavorites.ListItems.Clear
    listUnliked.ListItems.Clear
    listRead.ListItems.Clear
    
    Do Until rs.EOF
        Dim item As ListItem
        Set item = listBooks.ListItems.Add(, , rs!bookId)
        item.ListSubItems.Add , , rs!Title
        item.ListSubItems.Add , , rs!Author
        item.ListSubItems.Add , , rs!Genre
        item.ListSubItems.Add , , rs!Sinopsis
        rs.MoveNext
    Loop
    
    Do Until rsf.EOF
        Dim itemf As ListItem
        Set itemf = listFavorites.ListItems.Add(, , rsf!bookId)
        itemf.ListSubItems.Add , , rsf!Title
        itemf.ListSubItems.Add , , rsf!Author
        itemf.ListSubItems.Add , , rsf!Genre
        itemf.ListSubItems.Add , , rsf!Sinopsis
        rsf.MoveNext
    Loop
    
    Do Until rsuF.EOF
        Dim itemuF As ListItem
        Set itemuF = listUnliked.ListItems.Add(, , rsuF!bookId)
        itemuF.ListSubItems.Add , , rsuF!Title
        itemuF.ListSubItems.Add , , rsuF!Author
        itemuF.ListSubItems.Add , , rsuF!Genre
        itemuF.ListSubItems.Add , , rsuF!Sinopsis
        rsuF.MoveNext
    Loop
    
    Do Until rsr.EOF
        Dim itemR As ListItem
        Set itemR = listRead.ListItems.Add(, , rsr!bookId)
        itemR.ListSubItems.Add , , rsr!Title
        itemR.ListSubItems.Add , , rsr!Author
        itemR.ListSubItems.Add , , rsr!Genre
        itemR.ListSubItems.Add , , rsr!Sinopsis
        rsr.MoveNext
    Loop
    
    Call Disconnect

    With listBooks
        .View = lvwReport
            .ColumnHeaders.Add , , "Id", 500
            .ColumnHeaders.Add , , "Title", 3000
            .ColumnHeaders.Add , , "Author", 3000
            .ColumnHeaders.Add , , "Genre", 1500
            .ColumnHeaders.Add , , "Sinopsis", 3000
    End With
    
    With listFavorites
        .View = lvwReport
            .ColumnHeaders.Add , , "Id", 500
            .ColumnHeaders.Add , , "Title", 3000
            .ColumnHeaders.Add , , "Author", 3000
            .ColumnHeaders.Add , , "Genre", 1500
            .ColumnHeaders.Add , , "Sinopsis", 3000
    End With
    
    With listUnliked
        .View = lvwReport
            .ColumnHeaders.Add , , "Id", 500
            .ColumnHeaders.Add , , "Title", 3000
            .ColumnHeaders.Add , , "Author", 3000
            .ColumnHeaders.Add , , "Genre", 1500
            .ColumnHeaders.Add , , "Sinopsis", 3000
    End With
    
    With listRead
        .View = lvwReport
            .ColumnHeaders.Add , , "Id", 500
            .ColumnHeaders.Add , , "Title", 3000
            .ColumnHeaders.Add , , "Author", 3000
            .ColumnHeaders.Add , , "Genre", 1500
            .ColumnHeaders.Add , , "Sinopsis", 3000
    End With
End Sub

Private Sub library_Click()
    formLibrary.Show
End Sub
Private Sub profile_Click()
    FormUser.Show
End Sub

Private Sub seeBook_Click()
    formBook.Show
End Sub

Private Sub unFavorite_Click()
    If Not listFavorites.SelectedItem Is Nothing Then
        Dim bookId As Integer
        Dim userId As Integer
        bookId = listBooks.SelectedItem.Text
        userId = 1
        
        Call Connect
        conn.Execute "DELETE from bookFavorite WHERE bookId = " & bookId & " and userId = " & userId & ""
        Call Disconnect
        MsgBox "Book Unfavorite successfully"
    End If
End Sub

Private Sub unLiked_Click()
    If Not listBooks.SelectedItem Is Nothing Then
        Dim bookId As Integer
        Dim userId As Integer
        bookId = listBooks.SelectedItem.Text
        userId = 1
        
        Call Connect
        conn.Execute "INSERT INTO unlikedBooks(bookId, userId) values(" & bookId & ", " & userId & ")"
        Call Disconnect
        MsgBox "Book added successfully"
    End If
End Sub
