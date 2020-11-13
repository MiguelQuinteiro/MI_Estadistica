VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Ingresar Datos"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form2"
   ScaleHeight     =   7995
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   4350
      Left            =   5280
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   4350
      Left            =   3960
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   2640
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Frecuencia Absoluta ( fi )"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Limite Superior"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Limite Inferior"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
