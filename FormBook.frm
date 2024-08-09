VERSION 5.00
Begin VB.Form FormBook 
   Caption         =   "Form Book"
   ClientHeight    =   6504
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   7152
   LinkTopic       =   "Form1"
   ScaleHeight     =   6504
   ScaleWidth      =   7152
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   2652
      Left            =   4680
      Top             =   720
      Width           =   1932
   End
   Begin VB.Label Label7 
      Height          =   2052
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   6252
   End
   Begin VB.Label Label6 
      Height          =   492
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   3732
   End
   Begin VB.Label Label5 
      Height          =   372
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   3732
   End
   Begin VB.Label bookTitle 
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   3732
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   732
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   492
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   612
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   492
   End
   Begin VB.Label labelDetails 
      Alignment       =   2  'Center
      Caption         =   "Details of Book"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   2412
   End
End
Attribute VB_Name = "formBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
