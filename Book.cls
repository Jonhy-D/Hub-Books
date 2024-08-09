VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTitle As String
Private mAuthor As String
Private mGenre As String
Private mSinopsis As String

Public Property Get Title() As String
    Title = mTitle
End Property

Public Property Let Title(ByVal value As String)
    mTitle = value
End Property

Public Property Get Author() As String
    Author = mAuthor
End Property

Public Property Let Author(ByVal value As String)
    mAuthor = value
End Property

Public Property Get Genre() As String
    Genre = mGenre
End Property

Public Property Let Genre(ByVal value As String)
    mGenre = value
End Property

Public Property Get Sinopsis() As String
    Sinopsis = mSinopsis
End Property

Public Property Let Sinopsis(ByVal value As String)
    mSinopsis = value
End Property
