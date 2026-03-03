VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MiEstadistica 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ESTADÍSTICA I.- CÁLCULO DE TABLAS DE DISTRIBUCIÓN DE FRECUENCIAS."
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13590
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   13590
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text6 
      BackColor       =   &H80000003&
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
      Left            =   9360
      TabIndex        =   91
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000003&
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
      Left            =   8040
      TabIndex        =   90
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "COMANDOS"
      Height          =   2655
      Left            =   120
      TabIndex        =   89
      Top             =   6960
      Width           =   13335
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "T. D. D. F."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10560
         MaskColor       =   &H0000FFFF&
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Datos 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         MaskColor       =   &H0000FFFF&
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         MaskColor       =   &H0000FFFF&
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Caption         =   "Datos 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         MaskColor       =   &H0000FFFF&
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000FFFF&
         Caption         =   "Datos 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         MaskColor       =   &H0000FFFF&
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0000FFFF&
         Caption         =   "Datos 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         MaskColor       =   &H0000FFFF&
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0000FFFF&
         Caption         =   "Datos 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         MaskColor       =   &H0000FFFF&
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0000FFFF&
         Caption         =   "Datos 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         MaskColor       =   &H0000FFFF&
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0000FFFF&
         Caption         =   "Datos 7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         MaskColor       =   &H0000FFFF&
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0000FFFF&
         Caption         =   "Datos 8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         MaskColor       =   &H0000FFFF&
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0000FFFF&
         Caption         =   "Datos 9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6840
         MaskColor       =   &H0000FFFF&
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H0000FFFF&
         Caption         =   "Sin Ordenar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         MaskColor       =   &H0000FFFF&
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H0000FFFF&
         Caption         =   "Iniciales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10560
         MaskColor       =   &H0000FFFF&
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H0000FFFF&
         Caption         =   "Reportar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11880
         MaskColor       =   &H0000FFFF&
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H0000FFFF&
         Caption         =   "Histograma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6840
         MaskColor       =   &H0000FFFF&
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H0000FFFF&
         Caption         =   "B M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         MaskColor       =   &H0000FFFF&
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H0000FFFF&
         Caption         =   "Listados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         MaskColor       =   &H0000FFFF&
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H0000FFFF&
         Caption         =   "Lími. Inter."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11880
         MaskColor       =   &H0000FFFF&
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.PictureBox Slider1 
         Enabled         =   0   'False
         Height          =   555
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   12795
         TabIndex        =   19
         Top             =   1920
         Width           =   12855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "TABLA DE DISTRIBUCIÓN DE FRECUENCIAS"
      Height          =   3015
      Left            =   120
      TabIndex        =   68
      Top             =   3840
      Width           =   13335
      Begin VB.ListBox Fa 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   120
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox Xm 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   5400
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox XmFi 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   6720
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox FiXmX2 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   10680
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox XmX2 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   9360
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox XmX 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   8040
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox Fr 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   1440
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox FrP 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   2760
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox Percentiles 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   12000
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox FaP 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   4080
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "MiEstadistica.frx":0000
         TabIndex        =   79
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   495
         Left            =   1440
         OleObjectBlob   =   "MiEstadistica.frx":0084
         TabIndex        =   80
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   495
         Left            =   2760
         OleObjectBlob   =   "MiEstadistica.frx":010C
         TabIndex        =   81
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   495
         Left            =   4080
         OleObjectBlob   =   "MiEstadistica.frx":019A
         TabIndex        =   82
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   495
         Left            =   5400
         OleObjectBlob   =   "MiEstadistica.frx":0222
         TabIndex        =   83
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   495
         Left            =   6720
         OleObjectBlob   =   "MiEstadistica.frx":02A8
         TabIndex        =   84
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   495
         Left            =   8040
         OleObjectBlob   =   "MiEstadistica.frx":0314
         TabIndex        =   85
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   495
         Left            =   9360
         OleObjectBlob   =   "MiEstadistica.frx":037E
         TabIndex        =   86
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   495
         Left            =   10680
         OleObjectBlob   =   "MiEstadistica.frx":03F4
         TabIndex        =   87
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   495
         Left            =   12000
         OleObjectBlob   =   "MiEstadistica.frx":0470
         TabIndex        =   88
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "RESULTADOS DE LOS CÁLCULOS ESTADÍSTICOS"
      Height          =   3615
      Left            =   6120
      TabIndex        =   38
      Top             =   120
      Width           =   7335
      Begin VB.TextBox MediaAritmetica 
         Height          =   495
         Left            =   360
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Moda 
         Height          =   495
         Left            =   1680
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox SumaXmFi 
         Height          =   495
         Left            =   3000
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox SumaFiXmX2 
         Height          =   495
         Left            =   4320
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Mediana 
         Height          =   495
         Left            =   5640
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox TamañoMuestra 
         Height          =   495
         Left            =   360
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox DesviacionTipica 
         Height          =   495
         Left            =   1680
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox IntervaloInferior 
         Height          =   495
         Left            =   3000
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox IntervaloSuperior 
         Height          =   495
         Left            =   4320
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Percentil75 
         Height          =   495
         Left            =   5640
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox MiNPercentil 
         Height          =   495
         Left            =   3000
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox MiVPercentil 
         Height          =   495
         Left            =   4320
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox SumaFr 
         Height          =   495
         Left            =   360
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox SumaFrP 
         Height          =   495
         Left            =   1680
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Percentil25 
         Height          =   495
         Left            =   5640
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   495
         Left            =   360
         OleObjectBlob   =   "MiEstadistica.frx":04F8
         TabIndex        =   54
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   495
         Left            =   1680
         OleObjectBlob   =   "MiEstadistica.frx":057A
         TabIndex        =   55
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   495
         Left            =   5640
         OleObjectBlob   =   "MiEstadistica.frx":0600
         TabIndex        =   56
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   495
         Left            =   360
         OleObjectBlob   =   "MiEstadistica.frx":066C
         TabIndex        =   57
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   495
         Left            =   1680
         OleObjectBlob   =   "MiEstadistica.frx":06EA
         TabIndex        =   58
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   495
         Left            =   3000
         OleObjectBlob   =   "MiEstadistica.frx":0750
         TabIndex        =   59
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   495
         Left            =   4320
         OleObjectBlob   =   "MiEstadistica.frx":07C6
         TabIndex        =   60
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   495
         Left            =   5640
         OleObjectBlob   =   "MiEstadistica.frx":0842
         TabIndex        =   61
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   495
         Left            =   360
         OleObjectBlob   =   "MiEstadistica.frx":08C8
         TabIndex        =   62
         Top             =   2520
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
         Height          =   495
         Left            =   1680
         OleObjectBlob   =   "MiEstadistica.frx":094E
         TabIndex        =   63
         Top             =   2520
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
         Height          =   495
         Left            =   3000
         OleObjectBlob   =   "MiEstadistica.frx":09CE
         TabIndex        =   64
         Top             =   2520
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
         Height          =   495
         Left            =   5640
         OleObjectBlob   =   "MiEstadistica.frx":0A62
         TabIndex        =   65
         Top             =   2520
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
         Height          =   495
         Left            =   3000
         OleObjectBlob   =   "MiEstadistica.frx":0ACE
         TabIndex        =   66
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
         Height          =   495
         Left            =   4320
         OleObjectBlob   =   "MiEstadistica.frx":0B4A
         TabIndex        =   67
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATOS INICIALES PARA EL CÁLCULO"
      Height          =   3615
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   5775
      Begin VB.ListBox LimiteInferior 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   1680
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ListBox LimiteSuperior 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   3000
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ListBox Fi 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   4320
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000003&
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
         Left            =   1680
         TabIndex        =   31
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000003&
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
         Left            =   3000
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000003&
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
         Left            =   4320
         TabIndex        =   29
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox Datos 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   360
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Left            =   360
         OleObjectBlob   =   "MiEstadistica.frx":0BCA
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   495
         Left            =   1680
         OleObjectBlob   =   "MiEstadistica.frx":0C46
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   495
         Left            =   3000
         OleObjectBlob   =   "MiEstadistica.frx":0CD0
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Left            =   4320
         OleObjectBlob   =   "MiEstadistica.frx":0D5A
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6120
      OleObjectBlob   =   "MiEstadistica.frx":0DE4
      Top             =   -120
   End
   Begin VB.ListBox DatosModa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox DatosAcumulado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox DatosUnico 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox DatosOrdenado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox FiOrdenado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label NumeroCasos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6720
      TabIndex        =   24
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Menu mnuEjercicios 
      Caption         =   "EJERCICIOS"
      Begin VB.Menu mnuExamen 
         Caption         =   "EXAMEN"
      End
      Begin VB.Menu mnuExamen02 
         Caption         =   "EXAMEN 02"
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio Clase"
         Index           =   0
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 01"
         Index           =   1
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 02"
         Index           =   2
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 03"
         Index           =   3
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 04"
         Index           =   4
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 05"
         Index           =   5
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 06"
         Index           =   6
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 07"
         Index           =   7
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 08"
         Index           =   8
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 09"
         Index           =   9
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 10"
         Index           =   10
      End
      Begin VB.Menu mnuEjercicioBlanca 
         Caption         =   "Ejercicio 11"
         Index           =   11
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "OPCIONES"
      Begin VB.Menu mnuSkin 
         Caption         =   "Skin"
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "B-Studio"
            Index           =   0
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Comander"
            Index           =   1
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Cool Breeze"
            Index           =   2
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Copper"
            Index           =   3
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Corona"
            Index           =   4
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "DogmaX"
            Index           =   5
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Droid"
            Index           =   6
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Green"
            Index           =   7
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Gris"
            Index           =   8
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "KOZ"
            Index           =   9
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "LaST v1-2"
            Index           =   10
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "LongHorn"
            Index           =   11
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Mac"
            Index           =   12
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Media"
            Index           =   13
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Messenger"
            Index           =   14
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Metallic"
            Index           =   15
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "MMD"
            Index           =   16
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Neo"
            Index           =   17
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Office"
            Index           =   18
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Office2007"
            Index           =   19
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Orange_Graf"
            Index           =   20
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Paper"
            Index           =   21
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "SknR"
            Index           =   22
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "SoftCrystal"
            Index           =   23
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "St"
            Index           =   24
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "TopSecret"
            Index           =   25
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Tp"
            Index           =   26
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Web-II"
            Index           =   27
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Winamp 5"
            Index           =   28
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Zega"
            Index           =   29
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Zhelezo"
            Index           =   30
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Zippo"
            Index           =   31
         End
      End
   End
End
Attribute VB_Name = "MiEstadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *****************************************************************************
' * PROYECTO   : SISTEMA DE ESTADÍSTICA
' * FORMULARIO : Formulario Principal
' * AUTORES    : Miguel Quinteiro
' * FECHA      : 25 de Abril de 2008
' * ***************************************************************************


'AL CARGAR EL FORMULARIO
Private Sub Form_Load()
  Open "Skin.txt" For Input As #1
  Line Input #1, A$
  Close #1
  Mi_Skin = "\Skins\" + A$
  Aplicar_skin Me
End Sub


' CÁLCULO DE TODA LA ESTADÍSTICA
Private Sub Command1_Click()
  Text5.Enabled = True
  Text6.Enabled = True

  ' Imprime Cálculos
  Open "Calculos.txt" For Output As #1
  Print #1, "  "
  Print #1, "   ****** REPORTE DEL SISTEMA DE ESTADÍSTICA ******"

  '*********************************************************************************************
  ' PRIMER NIVEL DE CALCULO (PRIMER CICLO)
  '*********************************************************************************************
  Acumulador = 0
  N = 0
  For k = 0 To Fi.ListCount - 1
    LimiteInferior.ListIndex = k
    LimiteSuperior.ListIndex = k
    Fi.ListIndex = k

    ' Cálculo de N
    N = N + Val(Fi.Text)
    TamañoMuestra = N

    ' Cálculo de la Frecuencia Acumulada (Fa)
    Acumulador = Acumulador + Val(Fi.Text)
    Fa.AddItem Acumulador

    ' Cálculo del Punto Medio (Xm)
    PuntoMedio = (Val(Replace(LimiteInferior.Text, ",", ".")) + Val(Replace(LimiteSuperior.Text, ",", "."))) / 2
    PuntoMedio = Replace(PuntoMedio, ",", ".")
    Xm.AddItem PuntoMedio
  Next k

  '*********************************************************************************************
  ' SEGUNDO NIVEL DE CALCULO (SEGUNDO CICLO)
  '*********************************************************************************************
  SumaXmporFi = 0
  SumaFrPor = 0
  SumaFre = 0
  For k = 0 To Fi.ListCount - 1
    LimiteInferior.ListIndex = k
    LimiteSuperior.ListIndex = k
    Fi.ListIndex = k
    Fa.ListIndex = k
    Xm.ListIndex = k

    ' Cálculo de Xm * Fi
    XmPorFi = Val(Fi.Text) * Val(Xm.Text)
    SumaXmporFi = SumaXmporFi + XmPorFi
    XmPorFi = Replace(XmPorFi, ",", ".")
    XmFi.AddItem XmPorFi

    ' Cálculo del Fr
    FrecuenciaRelativa = Val(Fi.Text) / N
    FrecuenciaRelativa = Round(FrecuenciaRelativa, 3)
    SumaFre = SumaFre + FrecuenciaRelativa

    FrecuenciaRelativa = Replace(FrecuenciaRelativa, ",", ".")
    Fr.AddItem FrecuenciaRelativa

    ' Cálculo del Fr%
    FrecuenciaRelativaPor = (Val(Fi.Text) / N) * 100
    FrecuenciaRelativaPor = Round(FrecuenciaRelativaPor, 3)
    SumaFrPor = SumaFrPor + FrecuenciaRelativaPor

    FrecuenciaRelativaPor = Replace(FrecuenciaRelativaPor, ",", ".")
    FrP.AddItem FrecuenciaRelativaPor

    ' Cálculo del Fa%
    FaPor = (Val(Fa.Text) / N) * 100
    FaPor = Round(FaPor, 3)
    FaPor = Replace(FaPor, ",", ".")
    FaP.AddItem FaPor
  Next k

  SumaXmFi = Replace(SumaXmporFi, ",", ".")
  SumaFrP = Replace(SumaFrPor, ",", ".")
  SumaFr = Replace(SumaFre, ",", ".")

  MA = SumaXmporFi / N
  MA = Round(MA, 6)
  MediaAritmetica = Replace(MA, ",", ".")

  'Impresión del Cálculo Media Aritmética
  Print #1, "  "
  Print #1, "  * MEDIA ARITMÉTICA (X)"
  Print #1, "  ============================================================="
  Print #1, "        Sumatoria (Xm * Fi)"
  Print #1, "  X =  ______________________"
  Print #1, "                N"
  Print #1, "  "
  Print #1, "           "; SumaXmFi
  Print #1, "  X =  _______________"
  Print #1, "          "; N
  Print #1, "  "
  Print #1, "  X =  "; MediaAritmetica

  '*********************************************************************************************
  ' TERCER NIVEL DE CALCULO (TERCER CICLO)
  '*********************************************************************************************
  SumaFiXmMenosX2 = 0
  For k = 0 To Fi.ListCount - 1
    LimiteInferior.ListIndex = k
    LimiteSuperior.ListIndex = k
    Fi.ListIndex = k
    Xm.ListIndex = k
    XmFi.ListIndex = k

    ' Cálculo de Xm-X , (Xm-X)2 , Fi*(Xm-X)2
    XmMenosX = Val(Xm.Text) - MA
    XmMenosX2 = XmMenosX * XmMenosX
    FiXmMenosX2 = Val(Fi.Text) * XmMenosX2

    XmMenosX = Round(XmMenosX, 6)
    XmMenosX2 = Round(XmMenosX2, 6)
    FiXmMenosX2 = Round(FiXmMenosX2, 6)

    SumaFiXmMenosX2 = SumaFiXmMenosX2 + FiXmMenosX2
    SumaFiXmMenosX2 = Round(SumaFiXmMenosX2, 6)
    SumaFiXmX2 = Replace(SumaFiXmMenosX2, ",", ".")

    XmMenosX = Replace(XmMenosX, ",", ".")
    XmMenosX2 = Replace(XmMenosX2, ",", ".")
    FiXmMenosX2 = Replace(FiXmMenosX2, ",", ".")

    XmX.AddItem XmMenosX
    XmX2.AddItem XmMenosX2
    FiXmX2.AddItem FiXmMenosX2
  Next k

  ' CALCULO DE LA MODA
  'Busqueda del Mayor Fi
  Fi.ListIndex = 0
  IndiceMayorFi = 0
  MayorFi = Val(Fi.Text)
  For k = 0 To Fi.ListCount - 1
    Fi.ListIndex = k
    If Val(Fi.Text) > MayorFi Then
      MayorFi = Val(Fi.Text)
      IndiceMayorFi = k
    End If
  Next k
  'Calculo Moda
  LimiteInferior.ListIndex = IndiceMayorFi
  LimiteSuperior.ListIndex = IndiceMayorFi
  i = Val(Replace(LimiteSuperior.Text, ",", ".")) - Val(Replace(LimiteInferior.Text, ",", ".")) + 1

  If IndiceMayorFi > 0 Then
    LaModa = Val(Replace(LimiteInferior.Text, ",", ".")) + (((Val(Fi.List(IndiceMayorFi + 1)) * i) / (Val(Fi.List(IndiceMayorFi - 1)) + Val(Fi.List(IndiceMayorFi + 1)))))
  Else
    LaModa = Val(Replace(LimiteInferior.Text, ",", ".")) + (((Val(Fi.List(IndiceMayorFi + 1)) * i) / (0 + Val(Fi.List(IndiceMayorFi + 1)))))
  End If
  LaModa = Round(LaModa, 6)
  Moda = Replace(LaModa, ",", ".")

  'Impresión del Cálculo de la Moda
  Print #1, "  "
  Print #1, "  "
  Print #1, "  * MODA (Mo)"
  Print #1, "  ============================================================="
  Print #1, "  Primero se ubica la clase de Mayor Frecuencia Absoluta "
  Print #1, "  La Frecuencia Absoluta (Fi) mayor es  : "; Fi.List(IndiceMayorFi)
  Print #1, "  Y se encuentra en la Fila número  : "; IndiceMayorFi + 1; " de la T.D.D.F."
  Print #1, "  "
  Print #1, "  * DATOS QUE RECORREN LA CLASE (I)"
  Print #1, "  ============================================================="
  Print #1, "  I =  Ls - Li + 1"
  Print #1, "  "
  Print #1, "  I = "; Val(Replace(LimiteSuperior.Text, ",", ".")); " - "; Val(Replace(LimiteInferior.Text, ",", ".")); " + 1"
  Print #1, "  "
  Print #1, "  I = "; Val(Replace(LimiteSuperior.Text, ",", ".")) - Val(Replace(LimiteInferior.Text, ",", ".")) + 1
  Print #1, "  "
  Print #1, "  ============================================================="
  Print #1, "                [     F2            ]"
  Print #1, "  Mo =  Li  +   | _________  *  (I) |"
  Print #1, "                [  F1 + F2          ]"
  Print #1, "  "
  If IndiceMayorFi > 0 Then
    LaModa = Val(Replace(LimiteInferior.Text, ",", ".")) + (((Val(Fi.List(IndiceMayorFi + 1)) * i) / (Val(Fi.List(IndiceMayorFi - 1)) + Val(Fi.List(IndiceMayorFi + 1)))))
    MiRelleno = Space(Len(LimiteInferior.Text))
    Print #1, "  "; MiRelleno; "             [     "; Val(Fi.List(IndiceMayorFi + 1)); "            ]"
    Print #1, "  Mo = "; Val(Replace(LimiteInferior.Text, ",", ".")); "  +   | _________  *  "; i; " |"
    Print #1, "  "; MiRelleno; "             [  "; Val(Fi.List(IndiceMayorFi - 1)); " + "; Val(Fi.List(IndiceMayorFi + 1)); "          ]"

    Print #1, "  "
    Print #1, "  "; MiRelleno; "               "; Val(Fi.List(IndiceMayorFi + 1)) * i; " "
    Print #1, "  Mo = "; Val(Replace(LimiteInferior.Text, ",", ".")); "  +    ________   "
    Print #1, "  "; MiRelleno; "               "; Val(Fi.List(IndiceMayorFi - 1)) + Val(Fi.List(IndiceMayorFi + 1)); " "

    Print #1, "  "
    Print #1, "  Mo = "; Val(Replace(LimiteInferior.Text, ",", ".")); "  +  "; Round((Val(Fi.List(IndiceMayorFi + 1)) * i) / (Val(Fi.List(IndiceMayorFi - 1)) + Val(Fi.List(IndiceMayorFi + 1))), 6)
    Print #1, "  "

  Else
    LaModa = Val(Replace(LimiteInferior.Text, ",", ".")) + (((Val(Fi.List(IndiceMayorFi + 1)) * i) / (0 + Val(Fi.List(IndiceMayorFi + 1)))))
    MiRelleno = Space(Len(LimiteInferior.Text))
    Print #1, "  "; MiRelleno; "             [     "; Val(Fi.List(IndiceMayorFi + 1)); "            ]"
    Print #1, "  Mo = "; Val(Replace(LimiteInferior.Text, ",", ".")); "  +   | _________  *  "; i; " |"
    Print #1, "  "; MiRelleno; "             [  "; 0; " + "; Val(Fi.List(IndiceMayorFi + 1)); "          ]"

    Print #1, "  "
    Print #1, "  "; MiRelleno; "               "; Val(Fi.List(IndiceMayorFi + 1)) * i; " "
    Print #1, "  Mo = "; Val(Replace(LimiteInferior.Text, ",", ".")); "  +    ________  "
    Print #1, "  "; MiRelleno; "               "; 0 + Val(Fi.List(IndiceMayorFi + 1)); " "

    Print #1, "  "
    Print #1, "  Mo = "; Val(Replace(LimiteInferior.Text, ",", ".")); "  +  "; Round((Val(Fi.List(IndiceMayorFi + 1)) * i) / (0 + Val(Fi.List(IndiceMayorFi + 1))), 6)
    Print #1, "  "
  End If
  Print #1, "  Mo =  "; Moda
  Print #1, "  "

  ' CALCULO DE LA MEDIANA
  ' Busqueda de la Linea Medianal
  Medianal = N / 2
  If Medianal <= Val(Fa.List(0)) Then
    IndiceMedianal = 0
  Else
    For k = 1 To Fa.ListCount - 1
      Fa.ListIndex = k
      If (Medianal > Val(Fa.List(k - 1))) And (Medianal <= Val(Fa.List(k))) Then
        IndiceMedianal = k
      End If
    Next k
  End If

  ' Calculo Mediana
  LimiteInferior.ListIndex = IndiceMedianal
  LimiteSuperior.ListIndex = IndiceMedianal
  i = Val(Replace(LimiteSuperior.Text, ",", ".")) - Val(Replace(LimiteInferior.Text, ",", ".")) + 1
  Lir = ((Val(Replace(LimiteInferior.List(IndiceMedianal), ",", ".")) + Val(Replace(LimiteSuperior.List(IndiceMedianal - 1), ",", "."))) / 2)
  WFa = Val(Fa.List(IndiceMedianal - 1))
  WFi = Val(Fi.List(IndiceMedianal))
  LaMediana = Lir + ((Medianal - WFa) * i / WFi)
  LaMediana = Round(LaMediana, 6)
  Mediana = Replace(LaMediana, ",", ".")

  'Impresión del Cálculo de la Mediana
  Print #1, "  "
  Print #1, "  * MEDIANA (Me)"
  Print #1, "  ============================================================="
  Print #1, "  Primero se ubica (N/2) --> ("; N; " / 2)= "; N / 2; " dentro de las "
  Print #1, "  Frecuencias Acumuladas para localizar la Línea Referencial"
  Print #1, "  Y se encuentra en la Fila número  : "; IndiceMedianal + 1; " de la T.D.D.F."
  Print #1, "  "
  Print #1, "  * LÍMITE REAL INFERIOR (Lir)"
  Print #1, "  ============================================================="
  Print #1, "           Li + Ls(anterior)   "
  Print #1, "  Lir =  _____________________ "
  Print #1, "                  2            "
  Print #1, "  "
  Print #1, "            "; Replace(LimiteInferior.List(IndiceMedianal), ",", "."); " + "; Replace(LimiteSuperior.List(IndiceMedianal - 1), ",", ".")
  Print #1, "  Lir =  _____________ "
  Print #1, "              2            "
  Print #1, "  "
  Print #1, "             "; Val(Replace(LimiteInferior.List(IndiceMedianal), ",", ".")) + Val(Replace(LimiteSuperior.List(IndiceMedianal - 1), ",", "."))
  Print #1, "  Lir =  _____________ "
  Print #1, "              2            "
  Print #1, "  "
  Print #1, "  Lir = "; Lir
  Print #1, "  "
  Print #1, "  * DATOS QUE RECORREN LA CLASE (I)"
  Print #1, "  ============================================================="
  Print #1, "  I =  Ls - Li + 1"
  Print #1, "  "
  Print #1, "  I = "; Val(Replace(LimiteSuperior.Text, ",", ".")); " - "; Val(Replace(LimiteInferior.Text, ",", ".")); " + 1"
  Print #1, "  "
  Print #1, "  I = "; Val(Replace(LimiteSuperior.Text, ",", ".")) - Val(Replace(LimiteInferior.Text, ",", ".")) + 1
  Print #1, "  "
  Print #1, "  ============================================================="

  MiRelleno = Space(Len(Str(Lir)))
  Print #1, "                [ ( N/2 ) - Fa (anterior)        ]"
  Print #1, "  Me =  Lir  +  | _______________________  * (I) |"
  Print #1, "                [          Fi                    ]"
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [   ( "; N; "/2 ) - "; WFa; "                ]"
  Print #1, "  Me = "; Lir; "  +  | _______________________  * ("; i; ") |"
  Print #1, "   "; MiRelleno; "          [          "; WFi; "                    ]"
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [   ( "; N / 2; " ) - "; WFa; "              ]"
  Print #1, "  Me = "; Lir; "  +  | _______________________  * ("; i; ") |"
  Print #1, "   "; MiRelleno; "          [          "; WFi; "                    ]"
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [    "; (N / 2) - WFa; "             ]"
  Print #1, "  Me = "; Lir; "  +  | ___________  * ("; i; ") |"
  Print #1, "   "; MiRelleno; "          [     "; WFi; "             ]"
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [     "; ((N / 2) - WFa) * i; "  ]"
  Print #1, "  Me = "; Lir; "  +  | ___________ |"
  Print #1, "   "; MiRelleno; "          [     "; WFi; "    ]"
  Print #1, "  "
  Print #1, "  Me = "; Lir; "  +  "; Round(((((N / 2) - WFa) * i) / WFi), 6)
  Print #1, "  "
  Print #1, "  Me = "; LaMediana

  ' Desviación Típica
  DT = Sqr((SumaFiXmMenosX2 / N))
  DT = Round(DT, 6)
  DesviacionTipica = Replace(DT, ",", ".")

  ' Impresión de la Desviación Típica
  Print #1, "  "
  Print #1, "  "
  Print #1, "* DESVIACIÓN TÍPICA (S)"
  Print #1, "  ============================================================="
  Print #1, "               __________________________________"
  Print #1, "       _      /                             2"
  Print #1, "  S =   \    /  Sumatoria [  Fi * ( Xm - X )   ]"
  Print #1, "         \  /  __________________________________"
  Print #1, "          \/                  N"
  Print #1, "  "
  Print #1, "       _      ___________________"
  Print #1, "        \    /     "; SumaFiXmMenosX2
  Print #1, "  S =    \  /  _________________"
  Print #1, "          \/        "; N
  Print #1, "  "
  Print #1, "        _    ________________"
  Print #1, "         \  /                "
  Print #1, "  S =     \/ "; Round(SumaFiXmMenosX2 / N, 6)
  Print #1, "  "
  Print #1, "  S = "; DT
  Print #1, "  "
  Print #1, "  "

  ' Intervalo de Confianza
  IntInfe = MA - DT
  IntInfe = Round(IntInfe, 6)
  IntSupe = MA + DT
  IntSupe = Round(IntSupe, 6)
  IntervaloInferior = Replace(IntInfe, ",", ".")
  IntervaloSuperior = Replace(IntSupe, ",", ".")

  ' Impresión del Intervalo de Confianza
  Print #1, "  * INTERVALO DE CONFIANZA (Ic)"
  Print #1, "  ============================================================="
  Print #1, "  Ic =  [ X - S ; X + S ]"
  Print #1, "  "
  Print #1, "  Ic =  [ "; MA; " - "; DT; " ; "; MA; " + "; DT; " ]"
  Print #1, "  "
  Print #1, "  Ic =  [ "; IntervaloInferior; " ; "; IntervaloSuperior; " ]"
  Print #1, "  "

  ' CALCULO DEL PERCENTIL 25
  ' Busqueda de la Linea Referencial
  SubI = 25
  Referencial = SubI * N / 100
  If Referencial <= Val(Fa.List(0)) Then
    IndiceReferencial = 0
  Else
    For k = 1 To Fa.ListCount - 1
      Fa.ListIndex = k
      If (Referencial > Val(Fa.List(k - 1))) And (Referencial <= Val(Fa.List(k))) Then
        IndiceReferencial = k
      End If
    Next k
  End If

  ' Calculo Percentil
  LimiteInferior.ListIndex = IndiceReferencial
  LimiteSuperior.ListIndex = IndiceReferencial
  i = Val(LimiteSuperior.Text) - Val(LimiteInferior.Text) + 1
  Lir = ((Val(Replace(LimiteInferior.List(IndiceReferencial), ",", ".")) + Val(Replace(LimiteSuperior.List(IndiceReferencial - 1), ",", "."))) / 2)
  WFa = Val(Fa.List(IndiceReferencial - 1))
  WFi = Val(Fi.List(IndiceReferencial))
  ElPercentil25 = Lir + ((Referencial - WFa) * i / WFi)
  ElPercentil25 = Round(ElPercentil25, 6)
  Percentil25 = Replace(ElPercentil25, ",", ".")

  ' Impresión del Percentil 25
  Print #1, "  "
  Print #1, "  * POSICIÓN DEL PERCENTIL (P)"
  Print #1, "  ============================================================="
  Print #1, "  Se Busca el Percentil 25 (P25) Equivalente al Cuartil 1 (Q1)"
  Print #1, "  número del percentil i = 25"
  Print #1, "  "
  Print #1, "  P = (i * N) / 100"
  Print #1, "  "
  Print #1, "  P = ("; SubI; " * "; N; ") / 100"
  Print #1, "  "
  Print #1, "  P = ("; SubI * N; ") / 100"
  Print #1, "  "
  Print #1, "  P =  "; (SubI * N) / 100
  Print #1, "  "
  Print #1, "  Primero se ubica P = "; Referencial; " dentro de las "
  Print #1, "  Frecuencias Acumuladas para localizar la Línea Referencial"
  Print #1, "  Y se encuentra en la Fila número  : "; IndiceReferencial + 1; " de la T.D.D.F."
  Print #1, "  "

  Print #1, "  "
  Print #1, "  * LÍMITE REAL INFERIOR (Lir)"
  Print #1, "  ============================================================="
  Print #1, "           Li + Ls(anterior)   "
  Print #1, "  Lir =  _____________________ "
  Print #1, "                  2            "
  Print #1, "  "
  Print #1, "            "; Replace(LimiteInferior.List(IndiceReferencial), ",", "."); " + "; Replace(LimiteSuperior.List(IndiceReferencial - 1), ",", ".")
  Print #1, "  Lir =  _____________ "
  Print #1, "              2            "
  Print #1, "  "
  Print #1, "             "; Val(Replace(LimiteInferior.List(IndiceReferencial), ",", ".")) + Val(Replace(LimiteSuperior.List(IndiceReferencial - 1), ",", "."))
  Print #1, "  Lir =  _____________ "
  Print #1, "              2            "
  Print #1, "  "
  Print #1, "  Lir = "; Lir
  Print #1, "  "
  Print #1, "  * DATOS QUE RECORREN LA CLASE (I)"
  Print #1, "  ============================================================="
  Print #1, "  I =  Ls - Li + 1"
  Print #1, "  "
  Print #1, "  I = "; Val(Replace(LimiteSuperior.Text, ",", ".")); " - "; Val(Replace(LimiteInferior.Text, ",", ".")); " + 1"
  Print #1, "  "
  Print #1, "  I = "; Val(Replace(LimiteSuperior.Text, ",", ".")) - Val(Replace(LimiteInferior.Text, ",", ".")) + 1
  Print #1, "  "
  Print #1, "  ============================================================="

  Print #1, "  * PERCENTIL (Pi)"
  Print #1, "  ============================================================="
  Print #1, "               [ P - Fa (anterior)        ]"
  Print #1, "  Pi = Lir  +  | _________________  * (I) |"
  Print #1, "               [        Fi                ]"

  MiRelleno = Space(Len(Str(Lir)) + Len(Str(SubI)))
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [    "; Referencial; " - "; WFa; "                ]"
  Print #1, "  P"; SubI; " = "; Lir; "  +  | _______________________  * ("; i; ") |"
  Print #1, "   "; MiRelleno; "          [          "; WFi; "                    ]"
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [    "; Referencial - WFa; "            ]"
  Print #1, "  P"; SubI; " = "; Lir; "  +  | ___________  * ("; i; ") |"
  Print #1, "   "; MiRelleno; "          [     "; WFi; "             ]"
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [     "; (Referencial - WFa) * i; "  ]"
  Print #1, "  P"; SubI; " = "; Lir; "  +  | ___________ |"
  Print #1, "   "; MiRelleno; "          [     "; WFi; "    ]"
  Print #1, "  "
  Print #1, "  P"; SubI; " = "; Lir; "  +  "; Round((((Referencial - WFa) * i) / WFi), 6)
  Print #1, "  "
  Print #1, "  P"; SubI; " = "; Percentil25
  Print #1, "  "
  Print #1, "  P25 = Q1 = "; Percentil25

  ' CALCULO DEL PERCENTIL 75
  ' Busqueda de la Linea Referencial
  SubI = 75
  Referencial = SubI * N / 100
  If Referencial <= Val(Fa.List(0)) Then
    IndiceReferencial = 0
  Else
    For k = 1 To Fa.ListCount - 1
      Fa.ListIndex = k
      If (Referencial > Val(Fa.List(k - 1))) And (Referencial <= Val(Fa.List(k))) Then
        IndiceReferencial = k
      End If
    Next k
  End If

  ' Calculo Percentil
  LimiteInferior.ListIndex = IndiceReferencial
  LimiteSuperior.ListIndex = IndiceReferencial
  i = Val(LimiteSuperior.Text) - Val(LimiteInferior.Text) + 1
  Lir = ((Val(Replace(LimiteInferior.List(IndiceReferencial), ",", ".")) + Val(Replace(LimiteSuperior.List(IndiceReferencial - 1), ",", "."))) / 2)
  WFa = Val(Fa.List(IndiceReferencial - 1))
  WFi = Val(Fi.List(IndiceReferencial))
  ElPercentil75 = Lir + ((Referencial - WFa) * i / WFi)
  ElPercentil75 = Round(ElPercentil75, 6)
  Percentil75 = Replace(ElPercentil75, ",", ".")

  ' Impresión del Percentil 75
  Print #1, "  "
  Print #1, "  * POSICIÓN DEL PERCENTIL (P)"
  Print #1, "  ============================================================="
  Print #1, "  Se Busca el Percentil 75 (P75) Equivalente al Cuartil 3 (Q3)"
  Print #1, "  número del percentil i = 75"
  Print #1, "  "
  Print #1, "  P = (i * N) / 100"
  Print #1, "  "
  Print #1, "  P = ("; SubI; " * "; N; ") / 100"
  Print #1, "  "
  Print #1, "  P = ("; SubI * N; ") / 100"
  Print #1, "  "
  Print #1, "  P =  "; (SubI * N) / 100
  Print #1, "  "
  Print #1, "  Primero se ubica P = "; Referencial; " dentro de las "
  Print #1, "  Frecuencias Acumuladas para localizar la Línea Referencial"
  Print #1, "  Y se encuentra en la Fila número  : "; IndiceReferencial + 1; " de la T.D.D.F."
  Print #1, "  "

  Print #1, "  "
  Print #1, "  * LÍMITE REAL INFERIOR (Lir)"
  Print #1, "  ============================================================="
  Print #1, "           Li + Ls(anterior)   "
  Print #1, "  Lir =  _____________________ "
  Print #1, "                  2            "
  Print #1, "  "
  Print #1, "            "; Replace(LimiteInferior.List(IndiceReferencial), ",", "."); " + "; Replace(LimiteSuperior.List(IndiceReferencial - 1), ",", ".")
  Print #1, "  Lir =  _____________ "
  Print #1, "              2            "
  Print #1, "  "
  Print #1, "             "; Val(Replace(LimiteInferior.List(IndiceReferencial), ",", ".")) + Val(Replace(LimiteSuperior.List(IndiceReferencial - 1), ",", "."))
  Print #1, "  Lir =  _____________ "
  Print #1, "              2            "
  Print #1, "  "
  Print #1, "  Lir = "; Lir
  Print #1, "  "
  Print #1, "  * DATOS QUE RECORREN LA CLASE (I)"
  Print #1, "  ============================================================="
  Print #1, "  I =  Ls - Li + 1"
  Print #1, "  "
  Print #1, "  I = "; Val(Replace(LimiteSuperior.Text, ",", ".")); " - "; Val(Replace(LimiteInferior.Text, ",", ".")); " + 1"
  Print #1, "  "
  Print #1, "  I = "; Val(Replace(LimiteSuperior.Text, ",", ".")) - Val(Replace(LimiteInferior.Text, ",", ".")) + 1
  Print #1, "  "
  Print #1, "  ============================================================="

  Print #1, "  * PERCENTIL (Pi)"
  Print #1, "  ============================================================="
  Print #1, "               [ P - Fa (anterior)        ]"
  Print #1, "  Pi = Lir  +  | _________________  * (I) |"
  Print #1, "               [        Fi                ]"

  MiRelleno = Space(Len(Str(Lir)) + Len(Str(SubI)))
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [    "; Referencial; " - "; WFa; "              ]"
  Print #1, "  P"; SubI; " = "; Lir; "  +  | _______________________  * ("; i; ") |"
  Print #1, "   "; MiRelleno; "          [          "; WFi; "                    ]"
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [    "; Referencial - WFa; "             ]"
  Print #1, "  P"; SubI; " = "; Lir; "  +  | ___________  * ("; i; ") |"
  Print #1, "   "; MiRelleno; "          [     "; WFi; "             ]"
  Print #1, "  "
  Print #1, "   "; MiRelleno; "          [     "; (Referencial - WFa) * i; "  ]"
  Print #1, "  P"; SubI; " = "; Lir; "  +  | ___________ |"
  Print #1, "   "; MiRelleno; "          [     "; WFi; "    ]"
  Print #1, "  "
  Print #1, "  P"; SubI; " = "; Lir; "  +  "; Round((((Referencial - WFa) * i) / WFi), 6)
  Print #1, "  "
  Print #1, "  P"; SubI; " = "; Percentil75
  Print #1, "  "
  Print #1, "  P75 = Q3 = "; Percentil75

  ' CALCULO DE TODOS LOS PERCENTILES
  ' Busqueda de la Linea Referencial
  For P = 1 To 99
    SubI = P
    Referencial = SubI * N / 100
    If Referencial <= Val(Fa.List(0)) Then
      IndiceReferencial = 0
    Else
      For k = 1 To Fa.ListCount - 1
        Fa.ListIndex = k
        If (Referencial > Val(Fa.List(k - 1))) And (Referencial <= Val(Fa.List(k))) Then
          IndiceReferencial = k
        End If
      Next k
    End If

    ' Calculo Percentil
    LimiteInferior.ListIndex = IndiceReferencial
    LimiteSuperior.ListIndex = IndiceReferencial
    i = Val(LimiteSuperior.Text) - Val(LimiteInferior.Text) + 1
    If IndiceReferencial <> 0 Then
      Lir = ((Val(Replace(LimiteInferior.List(IndiceReferencial), ",", ".")) + Val(Replace(LimiteSuperior.List(IndiceReferencial - 1), ",", "."))) / 2)
      WFa = Val(Fa.List(IndiceReferencial - 1))
    Else
      Lir = ((Val(Replace(LimiteInferior.List(IndiceReferencial), ",", ".")) + 0) / 2)
      WFa = 0
    End If
    WFi = Val(Fi.List(IndiceReferencial))
    ElPercentil = Lir + ((Referencial - WFa) * i / WFi)
    ElPercentil = Round(ElPercentil, 6)
    Percentiles.AddItem Replace(ElPercentil, ",", ".")
  Next P

  ' Cierra los Archivos
  Close #1

  ' Activación y Desactivación de Botones
  Command1.Enabled = False
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command13.Enabled = False
  Command14.Enabled = True
  Command15.Enabled = False
  Command16.Enabled = False
  Slider1.Enabled = True
End Sub

' DATOS INICIALES PARA LA T.D.D.F.
Private Sub Command13_Click()
' Imprime Sturges
  Open "Sturges.txt" For Output As #1
  Print #1, "  "
  Print #1, "   ****** REPORTE DEL SISTEMA DE ESTADÍSTICA ******"
  Print #1, "  "

  ' Busqueda de los Valores Máximo y Mínimo
  For k = 0 To Datos.ListCount - 1
    Datos.ListIndex = k
    If k = 0 Then
      ValorMaximo = Val(Replace(Datos.Text, ",", "."))
      ValorMinimo = Val(Replace(Datos.Text, ",", "."))
    End If
    If Val(Replace(Datos.Text, ",", ".")) > ValorMaximo Then
      ValorMaximo = Val(Replace(Datos.Text, ",", "."))
    End If
    If Val(Replace(Datos.Text, ",", ".")) < ValorMinimo Then
      ValorMinimo = Val(Replace(Datos.Text, ",", "."))
    End If
  Next k

  ' Cálculo de la Amplitud Total
  AmplitudTotal = ValorMaximo - ValorMinimo

  ' Impresión de la Amplitud Total
  Print #1, "  * AMPLITUD TOTAL (At)"
  Print #1, "  ============================================================="
  Print #1, "  At = Vmax - Vmin"
  Print #1, "  "
  Print #1, "  At = "; ValorMaximo; " - "; ValorMinimo
  Print #1, "  "
  Print #1, "  At = "; AmplitudTotal
  Print #1, "  "
  Print #1, "  "


  ' Cálculo de Sturges
  Sturges = Round(AmplitudTotal / (1 + (3.322 * (Log(Datos.ListCount) / Log(10)))), 0)
  If Sturges = 1 Then
    Sturges = 2
  End If

  ' Impresión de Sturges
  Print #1, "  * FÓRMULA DE STURGES (St)"
  Print #1, "  ============================================================="
  Print #1, "                 At"
  Print #1, "  St = ______________________"
  Print #1, "       1 + [ 3.322 * log(N) ]"
  Print #1, "  "
  Print #1, "             "; AmplitudTotal; ""
  Print #1, "  St = _________________________"
  Print #1, "       1 + [ 3.322 * log("; Datos.ListCount; ") ]"
  Print #1, "  "
  Print #1, "  "
  Print #1, "             "; AmplitudTotal; ""
  Print #1, "  St = _________________________"
  Print #1, "       1 + [ 3.322 * "; Round(Log(Datos.ListCount) / Log(10), 6); " ]"
  Print #1, "  "
  Print #1, "  "
  Print #1, "         "; AmplitudTotal; ""
  Print #1, "  St = ___________"
  Print #1, "       "; Round((1 + (3.322 * (Log(Datos.ListCount) / Log(10)))), 6)
  Print #1, "  "
  Print #1, "  "
  Print #1, "  St = "; Round(AmplitudTotal / (1 + (3.322 * (Log(Datos.ListCount) / Log(10)))), 6)
  Print #1, "  "
  Print #1, "  "
  Print #1, "  St = "; Sturges

  ' Ordenación creciente de los Datos
  For k = 0 To Datos.ListCount - 1
    Datos.ListIndex = k
    DatosOrdenado.List(k) = Replace(Datos.List(k), ",", ".")
  Next k
  For k = 1 To DatosOrdenado.ListCount - 1
    For m = 0 To DatosOrdenado.ListCount - 1
      If Val(DatosOrdenado.List(k)) < Val(DatosOrdenado.List(m)) Then
        Tempo = DatosOrdenado.List(k)
        DatosOrdenado.List(k) = DatosOrdenado.List(m)
        DatosOrdenado.List(m) = Tempo
      End If
    Next m
  Next k

  ' Impresión de Los Datos Ordenados
  Print #1, "  "
  Print #1, "  "
  Print #1, "  * DATOS ORDENADOS EN FORMA CRECIENTE"
  Print #1, "  ============================================================="
  Fila = 1
  MiEspacio = Space(10)
  For k = 0 To DatosOrdenado.ListCount - 1

    MiEspacio = Space(10 - Len(DatosOrdenado.List(k)))

    If Fila < 5 Then
      If Val(DatosOrdenado.List(k)) = Int(Val(DatosOrdenado.List(k))) Then
        MiEspacio = Space(8 - Len(DatosOrdenado.List(k)))
        Print #1, MiEspacio + DatosOrdenado.List(k) + ".0" + "  ";
        Fila = Fila + 1
      Else
        Print #1, MiEspacio + DatosOrdenado.List(k) + "  ";
        Fila = Fila + 1
      End If
    Else
      If Val(DatosOrdenado.List(k)) = Int(Val(DatosOrdenado.List(k))) Then
        MiEspacio = Space(8 - Len(DatosOrdenado.List(k)))
        Print #1, MiEspacio + DatosOrdenado.List(k) + ".0" + "  "
        Fila = 1
      Else
        Print #1, MiEspacio + DatosOrdenado.List(k) + "  "
        Fila = 1
      End If
    End If


  Next k
  Print #1, "  "


  ' Cálculo e Impresión de la Mediana para datos no agrupados
  Print #1, "  "
  Print #1, "  "
  Print #1, "  * MEDIANA PARA DATOS NO AGRUPADOS"
  Print #1, "  ============================================================="
  CantidadDatos = DatosOrdenado.ListCount
  If (-1) ^ CantidadDatos > 0 Then
    Mitad01 = Int(CantidadDatos / 2) - 1
    Mitad02 = Mitad01 + 1
    MedianaDatos = (Val(DatosOrdenado.List(Mitad01)) + Val(DatosOrdenado.List(Mitad02))) / 2
    Print #1, "              "; Val(DatosOrdenado.List(Mitad01)); " + "; Val(DatosOrdenado.List(Mitad02))
    Print #1, "  Mediana = __________________"
    Print #1, "                    2         "
    Print #1, "  "
    Print #1, "             "; Val(DatosOrdenado.List(Mitad01)) + Val(DatosOrdenado.List(Mitad02))
    Print #1, "  Mediana = ________"
    Print #1, "               2   "
    Print #1, "  "
  Else
    Mitad01 = Int(CantidadDatos / 2)
    MedianaDatos = Round(Val(DatosOrdenado.List(Mitad01)), 6)
  End If
  Print #1, "  Mediana = "; MedianaDatos

  ' Cálculo e Impresión de la Moda para datos no agrupados
  Print #1, "  "
  Print #1, "  "
  Print #1, "  * MODA PARA DATOS NO AGRUPADOS"
  Print #1, "  ============================================================="

  ' Recorrer datos ordenados para contar cada caso
  DatosUnico.Clear
  DatosAcumulado.Clear
  For k = 0 To DatosOrdenado.ListCount - 1
    If k = 0 Then
      ElDato = DatosOrdenado.List(k)
      Acumulador = 1
    Else
      If ElDato = DatosOrdenado.List(k) Then
        Acumulador = Acumulador + 1
      Else
        DatosUnico.AddItem ElDato
        DatosAcumulado.AddItem Acumulador

        ElDato = DatosOrdenado.List(k)
        Acumulador = 1
      End If
    End If
    If (k = DatosOrdenado.ListCount - 1) Then
      DatosUnico.AddItem ElDato
      DatosAcumulado.AddItem Acumulador
    End If
  Next k
  ' Buscar el caso con el acumulador mas grande
  For k = 0 To DatosAcumulado.ListCount - 1
    DatosAcumulado.ListIndex = k
    If k = 0 Then
      MiMaximo = Val(Replace(DatosAcumulado.Text, ",", "."))
      MiMinimo = Val(Replace(DatosAcumulado.Text, ",", "."))
    End If
    If Val(Replace(DatosAcumulado.Text, ",", ".")) > MiMaximo Then
      MiMaximo = Val(Replace(DatosAcumulado.Text, ",", "."))
    End If
    If Val(Replace(DatosAcumulado.Text, ",", ".")) < MiMinimo Then
      MiMinimo = Val(Replace(DatosAcumulado.Text, ",", "."))
    End If
  Next k
  ' Decir cual es la moda
  DatosModa.Clear
  If MiMaximo <> MiMinimo Then
    For k = 0 To DatosAcumulado.ListCount - 1
      DatosAcumulado.ListIndex = k
      If Val(DatosAcumulado.Text) = MiMaximo Then
        DatosModa.AddItem DatosUnico.List(k)
      End If
    Next k
  End If
  ' Impresión
  If DatosModa.ListCount > 0 Then
    For k = 0 To DatosModa.ListCount - 1
      DatosModa.ListIndex = k
      Print #1, "  Moda = "; DatosModa.Text
    Next k
  Else
    Print #1, "  No hay ninguna Moda"
  End If
  Print #1, "  "

  ' Carga de Limite Inferior y Limite Superior
  LimiteInferior.AddItem ValorMinimo
  LimiteSuperior.AddItem ValorMinimo + Sturges - 1
  NuevoDato = ValorMinimo + Sturges
  While NuevoDato <= ValorMaximo
    LimiteInferior.AddItem NuevoDato
    LimiteSuperior.AddItem NuevoDato + Sturges - 1
    NuevoDato = NuevoDato + Sturges
  Wend

  ' Cálculo de las Frecuencias Absolutas
  For k = 0 To LimiteInferior.ListCount - 1
    LimiteInferior.ListIndex = k
    LimiteSuperior.ListIndex = k
    Contador = 0
    For m = 0 To Datos.ListCount - 1
      Datos.ListIndex = m
      If (Val(Datos.Text) >= Val(LimiteInferior.Text)) And (Val(Datos.Text) <= Val(LimiteSuperior.Text)) Then
        Contador = Contador + 1
      End If
    Next m
    Fi.AddItem Contador
  Next k

  ' Cierra los Archivos
  Close #1

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command13.Enabled = False
  Command14.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' IMPRIME LOS RESULTADOS
Private Sub Command14_Click()
' Imprime las Reporte
  Open "Reporte.txt" For Output As #1
  Print #1, "  "
  Print #1, "   ****** REPORTE DEL SISTEMA DE ESTADÍSTICA ******"

  '*********************************************************************************************
  ' Imprime la Tabla de Distribución de Frecuencias
  '*********************************************************************************************
  Print #1, "  "
  Print #1, "   * TABLA DE DISTRIBUCIÓN DE FRECUENCIAS *"

  ' Calcula en Ancho Mayor de cada campo
  CadenaLimiteInferior = Len(LimiteInferior.List(0))
  CadenaLimiteSuperior = Len(LimiteSuperior.List(0))
  CadenaFi = Len(Fi.List(0))
  CadenaFa = Len(Fa.List(0))
  CadenaFr = Len(Fr.List(0))
  CadenaFrP = Len(FrP.List(0))
  CadenaFaP = Len(FaP.List(0))
  CadenaXm = Len(Xm.List(0))
  CadenaXmFi = Len(XmFi.List(0))
  CadenaXmX = Len(XmX.List(0))
  CadenaXmX2 = Len(XmX2.List(0))
  CadenaFiXmX2 = Len(FiXmX2.List(0))

  For k = 0 To LimiteInferior.ListCount - 1
    LimiteInferior.ListIndex = k
    LimiteSuperior.ListIndex = k
    Fi.ListIndex = k
    Fa.ListIndex = k
    Fr.ListIndex = k
    FrP.ListIndex = k
    FaP.ListIndex = k
    Xm.ListIndex = k
    XmFi.ListIndex = k
    XmX.ListIndex = k
    XmX2.ListIndex = k
    FiXmX2.ListIndex = k

    If CadenaLimiteInferior < Len(LimiteInferior.Text) Then
      CadenaLimiteInferior = Len(LimiteInferior.Text)
    End If
    If CadenaLimiteSuperior < Len(LimiteSuperior.Text) Then
      CadenaLimiteSuperior = Len(LimiteSuperior.Text)
    End If
    If CadenaFi < Len(Fi.Text) Then
      CadenaFi = Len(Fi.Text)
    End If
    If CadenaFa < Len(Fa.Text) Then
      CadenaFa = Len(Fa.Text)
    End If
    If CadenaFr < Len(Fr.Text) Then
      CadenaFr = Len(Fr.Text)
    End If
    If CadenaFrP < Len(FrP.Text) Then
      CadenaFrP = Len(FrP.Text)
    End If
    If CadenaFaP < Len(FaP.Text) Then
      CadenaFaP = Len(FaP.Text)
    End If
    If CadenaXm < Len(Xm.Text) Then
      CadenaXm = Len(Xm.Text)
    End If
    If CadenaXmFi < Len(XmFi.Text) Then
      CadenaXmFi = Len(XmFi.Text)
    End If
    If CadenaXmX < Len(XmX.Text) Then
      CadenaXmX = Len(XmX.Text)
    End If
    If CadenaXmX2 < Len(XmX2.Text) Then
      CadenaXmX2 = Len(XmX2.Text)
    End If
    If CadenaFiXmX2 < Len(FiXmX2.Text) Then
      CadenaFiXmX2 = Len(FiXmX2.Text)
    End If
  Next k

  ' Imprime los encabezados OJO OJO OJO
  MiEspacio01 = Space(4 + CadenaLimiteInferior)
  MiEspacio02 = Space(4 + CadenaLimiteSuperior)
  MiEspacio03 = Space(2 + CadenaFi)
  MiEspacio04 = Space(2 + CadenaFa)
  MiEspacio05 = Space(2 + CadenaFr)
  MiEspacio06 = Space(2 + CadenaFrP)
  MiEspacio07 = Space(2 + CadenaFaP)
  MiEspacio08 = Space(2 + CadenaXm)
  MiEspacio09 = Space(2 + CadenaXmFi)
  MiEspacio10 = Space(2 + CadenaXmX)
  MiEspacio11 = Space(2 + CadenaXmX2)
  MiEspacio12 = Space(2 + CadenaFiXmX2)

  MiLinea = Len(MiEspacio01) + _
            Len(MiEspacio02) + _
            Len(MiEspacio03) + _
            Len(MiEspacio04) + _
            Len(MiEspacio05) + _
            Len(MiEspacio06) + _
            Len(MiEspacio07) + _
            Len(MiEspacio08) + _
            Len(MiEspacio09) + _
            Len(MiEspacio10) + _
            Len(MiEspacio11) + _
            Len(MiEspacio12)

  MiEspacio01 = Space(Abs(Len(MiEspacio01) - Len("Li")))
  MiEspacio02 = Space(Abs(Len(MiEspacio02) - Len("Ls")))
  MiEspacio03 = Space(Abs(Len(MiEspacio03) - Len("Fi")))
  MiEspacio04 = Space(Abs(Len(MiEspacio04) - Len("Fa")))
  MiEspacio05 = Space(Abs(Len(MiEspacio05) - Len("Fr")))
  MiEspacio06 = Space(Abs(Len(MiEspacio06) - Len("Fr%")))
  MiEspacio07 = Space(Abs(Len(MiEspacio07) - Len("Fa%")))
  MiEspacio08 = Space(Abs(Len(MiEspacio08) - Len("Xm")))
  MiEspacio09 = Space(Abs(Len(MiEspacio09) - Len("XmFi")))
  MiEspacio10 = Space(Abs(Len(MiEspacio10) - Len("XmX")))
  MiEspacio11 = Space(Abs(Len(MiEspacio11) - Len("XmX2")))
  MiEspacio12 = Space(Abs(Len(MiEspacio12) - Len("FiXmX2")))

  Print #1, " "; String(MiLinea, "=")
  Print #1, " "; String(MiLinea, "_")

  Print #1, MiEspacio01; "Li" _
            ; MiEspacio02; "Ls" _
            ; MiEspacio03; "Fi" _
            ; MiEspacio04; "Fa" _
            ; MiEspacio05; "Fr" _
            ; MiEspacio06; "Fr%" _
            ; MiEspacio07; "Fa%" _
            ; MiEspacio08; "Xm" _
            ; MiEspacio09; "XmFi" _
            ; MiEspacio10; "XmX" _
            ; MiEspacio11; "XmX2" _
            ; MiEspacio12; "FiXmX2"
  Print #1, " "; String(MiLinea, "_")

  ' Recorre los Datos para Buscar la información
  For k = 0 To LimiteInferior.ListCount - 1
    LimiteInferior.ListIndex = k
    LimiteSuperior.ListIndex = k
    Fi.ListIndex = k
    Fa.ListIndex = k
    Fr.ListIndex = k
    FrP.ListIndex = k
    FaP.ListIndex = k
    Xm.ListIndex = k
    XmFi.ListIndex = k
    XmX.ListIndex = k
    XmX2.ListIndex = k
    FiXmX2.ListIndex = k

    MiEspacio01 = Space(2 + CadenaLimiteInferior)
    MiEspacio02 = Space(2 + CadenaLimiteSuperior)
    MiEspacio03 = Space(2 + CadenaFi)
    MiEspacio04 = Space(2 + CadenaFa)
    MiEspacio05 = Space(2 + CadenaFr)
    MiEspacio06 = Space(2 + CadenaFrP)
    MiEspacio07 = Space(2 + CadenaFaP)
    MiEspacio08 = Space(2 + CadenaXm)
    MiEspacio09 = Space(2 + CadenaXmFi)
    MiEspacio10 = Space(2 + CadenaXmX)
    MiEspacio11 = Space(2 + CadenaXmX2)
    MiEspacio12 = Space(2 + CadenaFiXmX2)

    MiEspacio01 = Space(Len(MiEspacio01) - Len(LimiteInferior.Text))
    MiEspacio02 = Space(Len(MiEspacio02) - Len(LimiteSuperior.Text))
    MiEspacio03 = Space(Len(MiEspacio03) - Len(Fi.Text))
    MiEspacio04 = Space(Len(MiEspacio04) - Len(Fa.Text))
    MiEspacio05 = Space(Len(MiEspacio05) - Len(Fr.Text))
    MiEspacio06 = Space(Len(MiEspacio06) - Len(FrP.Text))
    MiEspacio07 = Space(Len(MiEspacio07) - Len(FaP.Text))
    MiEspacio08 = Space(Len(MiEspacio08) - Len(Xm.Text))
    MiEspacio09 = Space(Len(MiEspacio09) - Len(XmFi.Text))
    MiEspacio10 = Space(Len(MiEspacio10) - Len(XmX.Text))
    MiEspacio11 = Space(Len(MiEspacio11) - Len(XmX2.Text))
    MiEspacio12 = Space(Len(MiEspacio12) - Len(FiXmX2.Text))

    'Print #1, MiEspacio01; Format(Val(Replace(LimiteInferior.Text, ",", ".")), "##,##0.00") _

     ' OJO OJO OJO
    Print #1, MiEspacio01; Trim(Str(Val(LimiteInferior.Text) - 0.5)) _
              ; MiEspacio02; Trim(Str(Val(LimiteSuperior.Text) + 0.5)) _
              ; MiEspacio03; Fi.Text _
              ; MiEspacio04; Fa.Text _
              ; MiEspacio05; Fr.Text _
              ; MiEspacio06; FrP.Text _
              ; MiEspacio07; FaP.Text _
              ; MiEspacio08; Xm.Text _
              ; MiEspacio09; XmFi.Text _
              ; MiEspacio10; XmX.Text _
              ; MiEspacio11; XmX2.Text _
              ; MiEspacio12; FiXmX2.Text
  Next k
  Print #1, " "; String(MiLinea, "_")

  '*********************************************************************************************
  ' Imprime los Valores Estadísticos Obtenidos
  '*********************************************************************************************
  Print #1, "  "
  Print #1, "   * VALORES ESTADÍSTICOS OBTENIDOS *"
  Print #1, " "; String(MiLinea, "=")
  Print #1, "  "; "Tamaño de la Muestra (N)............: "; TamañoMuestra
  Print #1, "  "
  Print #1, "  "; "Suma Frecuencia Relativa............: "; SumaFr
  Print #1, "  "; "Suma Frecuencia Relativa %..........: "; SumaFrP
  Print #1, "  "; "Suma (Xm * Fi)......................: "; SumaXmFi
  Print #1, "  "; "Suma Fi*(Xm-X)2.....................: "; SumaFiXmX2
  Print #1, "  "
  Print #1, "  "; "Media Aritmética (X)................: "; MediaAritmetica
  Print #1, "  "; "Moda (Mo)...........................: "; Moda
  Print #1, "  "; "Mediana (Me), (P50), (D5), (Q2).....: "; Mediana
  Print #1, "  "
  Print #1, "  "; "Dsviación Típica (S)................: "; DesviacionTipica
  Print #1, "  "; "Intervalo de Confianza (Ic).........: [ "; IntervaloInferior; " ; "; IntervaloSuperior; " ]"
  Print #1, "  "
  Print #1, "  "; "Percentil 25 (P25), Cuartil 1 (Q1)..: "; Percentil25
  Print #1, "  "; "Percentil 75 (P75), Cuartil 3 (Q3)..: "; Percentil75

  '*********************************************************************************************
  ' Imprime el Histograma
  '*********************************************************************************************
  Print #1, "  "
  Print #1, "   * HISTOGRAMA *"
  Print #1, " "; String(MiLinea, "=")
  Print #1, "  "

  ' Ordenación decreciente de los Fi
  For k = 0 To Fi.ListCount - 1
    Fi.ListIndex = k
    FiOrdenado.List(k) = Fi.List(k)
  Next k
  For k = 1 To FiOrdenado.ListCount - 1
    For m = 0 To FiOrdenado.ListCount - 1
      If Val(FiOrdenado.List(k)) > Val(FiOrdenado.List(m)) Then
        Tempo = FiOrdenado.List(k)
        FiOrdenado.List(k) = FiOrdenado.List(m)
        FiOrdenado.List(m) = Tempo
      End If
    Next m
  Next k

  ' Calculo de la Escala
  MiEscala = (Val(FiOrdenado.List(0)) - Val(FiOrdenado.List(Fi.ListCount - 1))) / (Fi.ListCount - 1)
  MiEscala = Round(MiEscala)
  If MiEscala = 0 Then
    MiEscala = 1
  End If
  MiMaximo = Round(Val(FiOrdenado.List(0)) / MiEscala) * MiEscala

  ' Impresión de la Línea del Histograma
  MiLongitud = 4
  For m = MiMaximo To MiEscala Step -MiEscala
    MiVacio = Len(Str(m))
    MiEspacio = Space(MiLongitud - Len(Str(m)))
    MiHistograma = MiEspacio + Str(m) + "   |  "
    For k = 0 To Fi.ListCount - 1
      LimiteInferior.ListIndex = k
      LimiteSuperior.ListIndex = k
      Fi.ListIndex = k

      If Val(Fi.Text) < m Then
        MiHistograma = MiHistograma + "      "
      Else
        MiHistograma = MiHistograma + "  []  "
      End If
    Next k
    Print #1, MiHistograma
  Next m

  ' Imprime el EJE inferior
  MiHistograma = MiEspacio + Space(MiVacio) + "   |"
  For h = 0 To LimiteInferior.ListCount - 1
    LimiteInferior.ListIndex = h
    LimiteSuperior.ListIndex = h
    MiHistograma = MiHistograma + "______"
  Next h
  MiHistograma = MiHistograma + "______"
  Print #1, MiHistograma

  ' Imprime los Valores Inferiores
  MiLongitud = 8
  MiHistograma = MiEspacio + Space(MiVacio)  ' + "   |  "
  For h = 0 To LimiteInferior.ListCount - 1
    LimiteInferior.ListIndex = h
    LimiteSuperior.ListIndex = h
    MiEspacio = Space(MiLongitud - Len(LimiteInferior.Text))
    MiHistograma = MiHistograma + MiEspacio + LimiteInferior.Text
  Next h
  MiHistograma = MiHistograma + MiEspacio + LimiteSuperior.Text
  Print #1, MiHistograma

  '*********************************************************************************************
  ' Imprime el Polígono
  '*********************************************************************************************
  Print #1, "  "
  Print #1, "   * POLÍGONO DE FRECUENCIAS *"
  Print #1, " "; String(MiLinea, "=")
  Print #1, "  "

  ' Ordenación decreciente de los Fi
  For k = 0 To Fi.ListCount - 1
    Fi.ListIndex = k
    FiOrdenado.List(k) = Fi.List(k)
  Next k

  For k = 1 To FiOrdenado.ListCount - 1
    For m = 0 To FiOrdenado.ListCount - 1
      If Val(FiOrdenado.List(k)) > Val(FiOrdenado.List(m)) Then
        Tempo = FiOrdenado.List(k)
        FiOrdenado.List(k) = FiOrdenado.List(m)
        FiOrdenado.List(m) = Tempo
      End If
    Next m
  Next k

  ' Calculo de la Escala
  MiEscala = (Val(FiOrdenado.List(0)) - Val(FiOrdenado.List(Fi.ListCount - 1))) / (Fi.ListCount - 1)
  MiEscala = Round(MiEscala)
  If MiEscala = 0 Then
    MiEscala = 1
  End If
  MiMaximo = Round(Val(FiOrdenado.List(0)) / MiEscala) * MiEscala

  ' Impresión de la Línea del Poligono
  MiLongitud = 4
  For m = MiMaximo To MiEscala Step -MiEscala
    MiVacio = Len(Str(m))
    MiEspacio = Space(MiLongitud - Len(Str(m)))
    MiPoligono = MiEspacio + Str(m) + "   |  "
    For k = 0 To Fi.ListCount - 1
      LimiteInferior.ListIndex = k
      LimiteSuperior.ListIndex = k
      Fi.ListIndex = k
      If Val(Fi.Text) < m Then
        ' If Abs((Val(Fi.Text) - m)) > 0 Then
        MiPoligono = MiPoligono + "      "
      Else
        MiPoligono = MiPoligono + "  *   "
      End If
    Next k
    Print #1, MiPoligono
  Next m

  ' Imprime el EJE inferior
  MiPoligono = MiEspacio + Space(MiVacio) + "   |"
  For h = 0 To LimiteInferior.ListCount - 1
    LimiteInferior.ListIndex = h
    LimiteSuperior.ListIndex = h
    MiPoligono = MiPoligono + "______"
  Next h
  MiPoligono = MiPoligono + "_________"
  Print #1, MiPoligono

  ' Imprime los Valores Inferiores
  MiNumero = Val(Xm.List(1)) - Val(Xm.List(0))
  MiInicio = Val(Xm.List(0)) - MiNumero
  MiFin = Val(Xm.List(Xm.ListCount - 1)) + MiNumero
  MiLongitud = 8
  MiPoligono = MiEspacio + Space(MiVacio) + Str(MiInicio)
  For h = 0 To Xm.ListCount - 1
    Xm.ListIndex = h
    MiEspacio = Space(MiLongitud - Len(Xm.Text))
    MiPoligono = MiPoligono + MiEspacio + Xm.Text
  Next h
  MiPoligono = MiPoligono + MiEspacio + Str(MiFin)
  Print #1, MiPoligono

  ' Cierra los Archivos
  Close #1

  ' Activación y Desactivación de Botones
  Command1.Enabled = False
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command13.Enabled = False
  Command14.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

Private Sub Command17_Click()
  Documento.Show

End Sub

Private Sub Command18_Click()
' Ingresar y comprobar valores para límites Intermedios

' Ubicar la Clase Inferior
' Ubicar la Clase Superior
' Ubicar las Clases intermedias

' Cálculo de I en la clase Inferior Ampliada
' Cálculo de I en la clase Inferior Reducida
' Regla de tres para la clase inferior

' Cálculo de I en la clase Superior Ampliada
' Cálculo de I en la clase Superior Reducida
' Regla de tres para la clase superior

' Sumatoria de las respuestas
End Sub

' BORRA TODOS LOS DATOS
Private Sub Command3_Click()
  Slider1.Value = 1
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  Text4.Text = ""
  MiNPercentil.Text = ""
  MiVPercentil.Text = ""

  LimiteInferior.Clear
  LimiteSuperior.Clear
  Fi.Clear
  Fa.Clear
  Fr.Clear
  FrP.Clear
  FaP.Clear
  Xm.Clear
  XmFi.Clear
  XmX.Clear
  XmX2.Clear
  FiXmX2.Clear
  Percentiles.Clear
  Datos.Clear
  FiOrdenado.Clear
  DatosOrdenado.Clear
  DatosUnico.Clear
  DatosAcumulado.Clear
  DatosModa.Clear

  TamañoMuestra.Text = ""
  MediaAritmetica = ""
  DesviacionTipica = ""
  IntervaloInferior = ""
  IntervaloSuperior = ""
  Mediana = ""
  Moda = ""
  Percentil25 = ""
  Percentil75 = ""
  SumaXmFi = ""
  SumaFr = ""
  SumaFrP = ""
  SumaFiXmX2 = ""

  ' Activación y Desactivación de Botones
  Command1.Enabled = False
  Command2.Enabled = True
  Command4.Enabled = True
  Command5.Enabled = True
  Command6.Enabled = True
  Command7.Enabled = True
  Command8.Enabled = True
  Command9.Enabled = True
  Command10.Enabled = True
  Command11.Enabled = True
  Command12.Enabled = True
  Command13.Enabled = False
  Command14.Enabled = False
  Command15.Enabled = True
  Command16.Enabled = True
  Slider1.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA 1
Private Sub Command2_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 1.5
  LimiteInferior.AddItem 5.5
  LimiteInferior.AddItem 9.5
  LimiteInferior.AddItem 13.5
  LimiteInferior.AddItem 17.5
  LimiteInferior.AddItem 21.5

  LimiteSuperior.AddItem 5.5
  LimiteSuperior.AddItem 9.5
  LimiteSuperior.AddItem 13.5
  LimiteSuperior.AddItem 17.5
  LimiteSuperior.AddItem 21.5
  LimiteSuperior.AddItem 25

  Fi.AddItem 2
  Fi.AddItem 4
  Fi.AddItem 8
  Fi.AddItem 7
  Fi.AddItem 3
  Fi.AddItem 1

  'LimiteInferior.AddItem 200
  'LimiteInferior.AddItem 203
  'LimiteInferior.AddItem 206
  'LimiteInferior.AddItem 209
  'LimiteInferior.AddItem 212
  'LimiteInferior.AddItem 215
  'LimiteInferior.AddItem 218

  'LimiteSuperior.AddItem 202
  'LimiteSuperior.AddItem 205
  'LimiteSuperior.AddItem 208
  'LimiteSuperior.AddItem 211
  'LimiteSuperior.AddItem 214
  'LimiteSuperior.AddItem 217
  'LimiteSuperior.AddItem 220

  'Fi.AddItem 2
  'Fi.AddItem 4
  'Fi.AddItem 2
  'Fi.AddItem 1
  'Fi.AddItem 5
  'Fi.AddItem 6
  'Fi.AddItem 4

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command13.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA 2
Private Sub Command4_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 0
  LimiteInferior.AddItem 30
  LimiteInferior.AddItem 60
  LimiteInferior.AddItem 90
  LimiteInferior.AddItem 120
  LimiteInferior.AddItem 150

  LimiteSuperior.AddItem 29
  LimiteSuperior.AddItem 59
  LimiteSuperior.AddItem 89
  LimiteSuperior.AddItem 119
  LimiteSuperior.AddItem 149
  LimiteSuperior.AddItem 179

  Fi.AddItem 14
  Fi.AddItem 36
  Fi.AddItem 56
  Fi.AddItem 42
  Fi.AddItem 17
  Fi.AddItem 15

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA 3
Private Sub Command5_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 130
  LimiteInferior.AddItem 139
  LimiteInferior.AddItem 148
  LimiteInferior.AddItem 157
  LimiteInferior.AddItem 166
  LimiteInferior.AddItem 175
  LimiteInferior.AddItem 184

  LimiteSuperior.AddItem 138
  LimiteSuperior.AddItem 147
  LimiteSuperior.AddItem 156
  LimiteSuperior.AddItem 165
  LimiteSuperior.AddItem 174
  LimiteSuperior.AddItem 183
  LimiteSuperior.AddItem 192

  Fi.AddItem 9
  Fi.AddItem 6
  Fi.AddItem 8
  Fi.AddItem 2
  Fi.AddItem 11
  Fi.AddItem 2
  Fi.AddItem 2

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA 4
Private Sub Command6_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 4
  LimiteInferior.AddItem 7
  LimiteInferior.AddItem 10
  LimiteInferior.AddItem 13
  LimiteInferior.AddItem 16
  LimiteInferior.AddItem 19
  LimiteInferior.AddItem 22

  LimiteSuperior.AddItem 6
  LimiteSuperior.AddItem 9
  LimiteSuperior.AddItem 12
  LimiteSuperior.AddItem 15
  LimiteSuperior.AddItem 18
  LimiteSuperior.AddItem 21
  LimiteSuperior.AddItem 24

  Fi.AddItem 4
  Fi.AddItem 5
  Fi.AddItem 11
  Fi.AddItem 10
  Fi.AddItem 20
  Fi.AddItem 13
  Fi.AddItem 8

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA 5
Private Sub Command7_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 5
  LimiteInferior.AddItem 9
  LimiteInferior.AddItem 13
  LimiteInferior.AddItem 17
  LimiteInferior.AddItem 21
  LimiteInferior.AddItem 25
  LimiteInferior.AddItem 29

  LimiteSuperior.AddItem 8
  LimiteSuperior.AddItem 12
  LimiteSuperior.AddItem 16
  LimiteSuperior.AddItem 20
  LimiteSuperior.AddItem 24
  LimiteSuperior.AddItem 28
  LimiteSuperior.AddItem 32

  Fi.AddItem 5
  Fi.AddItem 2
  Fi.AddItem 4
  Fi.AddItem 4
  Fi.AddItem 8
  Fi.AddItem 3
  Fi.AddItem 4

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA 6
Private Sub Command8_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 890
  LimiteInferior.AddItem 895
  LimiteInferior.AddItem 900
  LimiteInferior.AddItem 905
  LimiteInferior.AddItem 910
  LimiteInferior.AddItem 915
  LimiteInferior.AddItem 920

  LimiteSuperior.AddItem 894
  LimiteSuperior.AddItem 899
  LimiteSuperior.AddItem 904
  LimiteSuperior.AddItem 909
  LimiteSuperior.AddItem 914
  LimiteSuperior.AddItem 919
  LimiteSuperior.AddItem 924

  Fi.AddItem 6
  Fi.AddItem 10
  Fi.AddItem 6
  Fi.AddItem 7
  Fi.AddItem 5
  Fi.AddItem 11
  Fi.AddItem 5

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA 7
Private Sub Command9_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 0
  LimiteInferior.AddItem 2
  LimiteInferior.AddItem 4
  LimiteInferior.AddItem 6
  LimiteInferior.AddItem 8

  LimiteSuperior.AddItem Replace("1.99", ",", ".")
  LimiteSuperior.AddItem Replace("3.99", ",", ".")
  LimiteSuperior.AddItem Replace("5.99", ",", ".")
  LimiteSuperior.AddItem Replace("7.99", ",", ".")
  LimiteSuperior.AddItem Replace("10", ",", ".")

  Fi.AddItem 10
  Fi.AddItem 48
  Fi.AddItem 20
  Fi.AddItem 20
  Fi.AddItem 10

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA 8
Private Sub Command10_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 5
  LimiteInferior.AddItem 9
  LimiteInferior.AddItem 13
  LimiteInferior.AddItem 17
  LimiteInferior.AddItem 21
  LimiteInferior.AddItem 25
  LimiteInferior.AddItem 29

  LimiteSuperior.AddItem 8
  LimiteSuperior.AddItem 12
  LimiteSuperior.AddItem 16
  LimiteSuperior.AddItem 20
  LimiteSuperior.AddItem 24
  LimiteSuperior.AddItem 28
  LimiteSuperior.AddItem 32

  Fi.AddItem 5
  Fi.AddItem 2
  Fi.AddItem 4
  Fi.AddItem 4
  Fi.AddItem 8
  Fi.AddItem 3
  Fi.AddItem 2

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA 9
Private Sub Command11_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 100
  LimiteInferior.AddItem 151
  LimiteInferior.AddItem 201
  LimiteInferior.AddItem 252
  LimiteInferior.AddItem 304
  LimiteInferior.AddItem 354

  LimiteSuperior.AddItem 150
  LimiteSuperior.AddItem 200
  LimiteSuperior.AddItem 251
  LimiteSuperior.AddItem 303
  LimiteSuperior.AddItem 353
  LimiteSuperior.AddItem 404

  Fi.AddItem 3
  Fi.AddItem 9
  Fi.AddItem 2
  Fi.AddItem 7
  Fi.AddItem 4
  Fi.AddItem 3

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA PARA HISTOGRAMA
Private Sub Command15_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  LimiteInferior.AddItem 20
  LimiteInferior.AddItem 31
  LimiteInferior.AddItem 42
  LimiteInferior.AddItem 53
  LimiteInferior.AddItem 64

  LimiteSuperior.AddItem 30
  LimiteSuperior.AddItem 41
  LimiteSuperior.AddItem 52
  LimiteSuperior.AddItem 63
  LimiteSuperior.AddItem 74

  Fi.AddItem 8
  Fi.AddItem 12
  Fi.AddItem 10
  Fi.AddItem 4
  Fi.AddItem 2

  ' Activación y Desactivación de Botones
  Command1.Enabled = True
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA A DATOS
Private Sub Command12_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  'MIGUEL
  Datos.AddItem 130
  Datos.AddItem 137
  Datos.AddItem 166
  Datos.AddItem 143
  Datos.AddItem 138
  Datos.AddItem 143
  Datos.AddItem 148
  Datos.AddItem 166
  Datos.AddItem 140
  Datos.AddItem 148
  Datos.AddItem 168
  Datos.AddItem 131
  Datos.AddItem 168
  Datos.AddItem 170
  Datos.AddItem 152
  Datos.AddItem 132
  Datos.AddItem 160
  Datos.AddItem 173
  Datos.AddItem 135
  Datos.AddItem 160
  Datos.AddItem 135
  Datos.AddItem 143
  Datos.AddItem 130
  Datos.AddItem 148
  Datos.AddItem 150
  Datos.AddItem 143
  Datos.AddItem 180
  Datos.AddItem 172
  Datos.AddItem 140
  Datos.AddItem 180
  Datos.AddItem 150
  Datos.AddItem 135
  Datos.AddItem 170
  Datos.AddItem 152
  Datos.AddItem 186
  Datos.AddItem 152
  Datos.AddItem 173
  Datos.AddItem 168
  Datos.AddItem 186
  Datos.AddItem 173

  ' Activación y Desactivación de Botones
  Command1.Enabled = False
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command13.Enabled = True
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

' INGRESA DATOS DE PRUEBA DE BM
Private Sub Command16_Click()
  Open "Sturges.txt" For Output As #1
  Close #1

  Datos.AddItem 0.1
  Datos.AddItem 1.5
  Datos.AddItem 2.2
  Datos.AddItem 2.7
  Datos.AddItem 4.3
  Datos.AddItem 5
  Datos.AddItem 4.6
  Datos.AddItem 4.6
  Datos.AddItem 5
  Datos.AddItem 3
  Datos.AddItem 4.5
  Datos.AddItem 3.5
  Datos.AddItem 4.1
  Datos.AddItem 2.9
  Datos.AddItem 4.5
  Datos.AddItem 4.8
  Datos.AddItem 4.6
  Datos.AddItem 4
  Datos.AddItem 3.2
  Datos.AddItem 3.5
  Datos.AddItem 4.2
  Datos.AddItem 3.2
  Datos.AddItem 3.2
  Datos.AddItem 3.6
  Datos.AddItem 3.8
  Datos.AddItem 2.9
  Datos.AddItem 2.5
  Datos.AddItem 3.6
  Datos.AddItem 2.7
  Datos.AddItem 3.7
  Datos.AddItem 2.5
  Datos.AddItem 2.9
  Datos.AddItem 2.7
  Datos.AddItem 3
  Datos.AddItem 3.5
  Datos.AddItem 3.2
  Datos.AddItem 2.9
  Datos.AddItem 2.4
  Datos.AddItem 2.7
  Datos.AddItem 3
  Datos.AddItem 3.5
  Datos.AddItem 2.9
  Datos.AddItem 2.7
  Datos.AddItem 2.9
  Datos.AddItem 4.1
  Datos.AddItem 3.2
  Datos.AddItem 3.7
  Datos.AddItem 3.8
  Datos.AddItem 7.9
  Datos.AddItem 4.3
  Datos.AddItem 4.7
  Datos.AddItem 3.8
  Datos.AddItem 4.6
  Datos.AddItem 5.3
  Datos.AddItem 4.1
  Datos.AddItem 4.9
  Datos.AddItem 3
  Datos.AddItem 2.8
  Datos.AddItem 4.8
  Datos.AddItem 3
  Datos.AddItem 2.8
  Datos.AddItem 6
  Datos.AddItem 5.4
  Datos.AddItem 5.7
  Datos.AddItem 6.8
  Datos.AddItem 10.1
  Datos.AddItem 6.1
  Datos.AddItem 5.8
  Datos.AddItem 7
  Datos.AddItem 5
  Datos.AddItem 7.2
  Datos.AddItem 6.9
  Datos.AddItem 7.7
  Datos.AddItem 6.2
  Datos.AddItem 7.2
  Datos.AddItem 6.9
  Datos.AddItem 7.1
  Datos.AddItem 7.1
  Datos.AddItem 7.4
  Datos.AddItem 12.5
  Datos.AddItem 10
  Datos.AddItem 8.8
  Datos.AddItem 9.8
  Datos.AddItem 9.3
  Datos.AddItem 12.2
  Datos.AddItem 8.1
  Datos.AddItem 10.8
  Datos.AddItem 11.3
  Datos.AddItem 9.6
  Datos.AddItem 10.7
  Datos.AddItem 10.4
  Datos.AddItem 7.4
  Datos.AddItem 6.4
  Datos.AddItem 4.9
  Datos.AddItem 16.2
  Datos.AddItem 16.6
  Datos.AddItem 13.3
  Datos.AddItem 16.7
  Datos.AddItem 15.9
  Datos.AddItem 15.8
  Datos.AddItem 20.9
  Datos.AddItem 16.2
  Datos.AddItem 18.3
  Datos.AddItem 19
  Datos.AddItem 18.8
  Datos.AddItem 17
  Datos.AddItem 21.9
  Datos.AddItem 16.4
  Datos.AddItem 23.6
  Datos.AddItem 15.8
  Datos.AddItem 21.4
  Datos.AddItem 20.5
  Datos.AddItem 23.2
  Datos.AddItem 20.2
  Datos.AddItem 22.9
  Datos.AddItem 27.4
  Datos.AddItem 43.6
  Datos.AddItem 32.6
  Datos.AddItem 37
  Datos.AddItem 34.3
  Datos.AddItem 43.1
  Datos.AddItem 35
  Datos.AddItem 40.7
  Datos.AddItem 48.3
  Datos.AddItem 50.8
  Datos.AddItem 49.5
  Datos.AddItem 45.1
  Datos.AddItem 50
  Datos.AddItem 46.4
  Datos.AddItem 27.4
  Datos.AddItem 35.9
  Datos.AddItem 41.1
  Datos.AddItem 41
  Datos.AddItem 37
  Datos.AddItem 45.2
  Datos.AddItem 48.6
  Datos.AddItem 51.3
  Datos.AddItem 54
  Datos.AddItem 59.9
  Datos.AddItem 59.9
  Datos.AddItem 56.4
  Datos.AddItem 29.6
  Datos.AddItem 52.6
  Datos.AddItem 48
  Datos.AddItem 56
  Datos.AddItem 51.7
  Datos.AddItem 54.1
  Datos.AddItem 56.4
  Datos.AddItem 50.8
  Datos.AddItem 55.2
  Datos.AddItem 53.1
  Datos.AddItem 51.8
  Datos.AddItem 52.5
  Datos.AddItem 30.3
  Datos.AddItem 70.7
  Datos.AddItem 63
  Datos.AddItem 61.3
  Datos.AddItem 54.6
  Datos.AddItem 66.8
  Datos.AddItem 84
  Datos.AddItem 105.7
  Datos.AddItem 117.6
  Datos.AddItem 104.1
  Datos.AddItem 148.7
  Datos.AddItem 148.4
  Datos.AddItem 93.9
  Datos.AddItem 104.8
  Datos.AddItem 101.1
  Datos.AddItem 109.2
  Datos.AddItem 100.1
  Datos.AddItem 147.8
  Datos.AddItem 112.7
  Datos.AddItem 149
  Datos.AddItem 160.1
  Datos.AddItem 161.7
  Datos.AddItem 180.7
  Datos.AddItem 156.3
  Datos.AddItem 103.6
  Datos.AddItem 153
  Datos.AddItem 175.9
  Datos.AddItem 186.1
  Datos.AddItem 217.6
  Datos.AddItem 197.5
  Datos.AddItem 232.1
  Datos.AddItem 291.5
  Datos.AddItem 305.6
  Datos.AddItem 332.5
  Datos.AddItem 384.8
  Datos.AddItem 395.3
  Datos.AddItem 403.4
  Datos.AddItem 534.4
  Datos.AddItem 536.8
  Datos.AddItem 634.4
  Datos.AddItem 609.5
  Datos.AddItem 973.6
  Datos.AddItem 696.3
  Datos.AddItem 1141.1
  Datos.AddItem 1163.9
  Datos.AddItem 1060.5
  Datos.AddItem 1151.4
  Datos.AddItem 1488.4
  Datos.AddItem 1426.1
  Datos.AddItem 1723.7
  Datos.AddItem 2151.8
  Datos.AddItem 3832.9
  Datos.AddItem 2397.7
  Datos.AddItem 3613.1
  Datos.AddItem 3993.8
  Datos.AddItem 2506.1
  Datos.AddItem 3203.6
  Datos.AddItem 8865.2
  Datos.AddItem 11965
  Datos.AddItem 18859.3
  Datos.AddItem 24100.3

  '    Datos.AddItem 0.1
  '    Datos.AddItem 1.5
  '    Datos.AddItem 2.2
  '    Datos.AddItem 2.7
  '    Datos.AddItem 4.3
  '    Datos.AddItem 5
  '    Datos.AddItem 4.6
  '    Datos.AddItem 4.6
  '    Datos.AddItem 5
  '    Datos.AddItem 3
  '    Datos.AddItem 4.5
  '    Datos.AddItem 3.5
  '    Datos.AddItem 4.1
  '    Datos.AddItem 2.9
  '    Datos.AddItem 4.5
  '    Datos.AddItem 4.8
  '    Datos.AddItem 4.6
  '    Datos.AddItem 4
  '    Datos.AddItem 3.2
  '    Datos.AddItem 3.5
  '    Datos.AddItem 4.2
  '    Datos.AddItem 3.2
  '    Datos.AddItem 3.2
  '    Datos.AddItem 3.6
  '    Datos.AddItem 3.8
  '    Datos.AddItem 2.9
  '    Datos.AddItem 2.5
  '    Datos.AddItem 3.6
  '    Datos.AddItem 2.7
  '    Datos.AddItem 3.7
  '    Datos.AddItem 2.5
  '    Datos.AddItem 2.9
  '    Datos.AddItem 2.7
  '    Datos.AddItem 3
  '    Datos.AddItem 3.5
  '    Datos.AddItem 3.2
  '    Datos.AddItem 2.9
  '    Datos.AddItem 2.4
  '    Datos.AddItem 2.7
  '    Datos.AddItem 3
  '    Datos.AddItem 3.5
  '    Datos.AddItem 2.9
  '    Datos.AddItem 2.7
  '    Datos.AddItem 2.9
  '    Datos.AddItem 4.1
  '    Datos.AddItem 3.2
  '    Datos.AddItem 3.7
  '    Datos.AddItem 3.8
  '    Datos.AddItem 7.9
  '    Datos.AddItem 4.3
  '    Datos.AddItem 4.7
  '    Datos.AddItem 3.8
  '    Datos.AddItem 4.6
  '    Datos.AddItem 5.3
  '    Datos.AddItem 4.1
  '    Datos.AddItem 4.9
  '    Datos.AddItem 3
  '    Datos.AddItem 2.8
  '    Datos.AddItem 4.8
  '    Datos.AddItem 3
  '    Datos.AddItem 2.8
  '    Datos.AddItem 6
  '    Datos.AddItem 5.4
  '    Datos.AddItem 5.7
  '    Datos.AddItem 6.8
  '    Datos.AddItem 10.1
  '    Datos.AddItem 6.1
  '    Datos.AddItem 5.8
  '    Datos.AddItem 7
  '    Datos.AddItem 5
  '    Datos.AddItem 7.2
  '    Datos.AddItem 6.9
  '    Datos.AddItem 7.7
  '    Datos.AddItem 6.2
  '    Datos.AddItem 7.2
  '    Datos.AddItem 6.9
  '    Datos.AddItem 7.1
  '    Datos.AddItem 7.1
  '    Datos.AddItem 7.4
  '    Datos.AddItem 12.5
  '    Datos.AddItem 10
  '    Datos.AddItem 8.8
  '    Datos.AddItem 9.8
  '    Datos.AddItem 9.3
  '    Datos.AddItem 12.2
  '    Datos.AddItem 8.1
  '    Datos.AddItem 10.8
  '    Datos.AddItem 11.3
  '    Datos.AddItem 9.6
  '    Datos.AddItem 10.7
  '    Datos.AddItem 10.4
  '    Datos.AddItem 7.4
  '    Datos.AddItem 6.4
  '    Datos.AddItem 4.9
  '    Datos.AddItem 16.2
  '    Datos.AddItem 16.6
  '    Datos.AddItem 13.3
  '    Datos.AddItem 16.7
  '    Datos.AddItem 15.9
  '    Datos.AddItem 15.8
  '    Datos.AddItem 20.9
  '    Datos.AddItem 16.2
  '    Datos.AddItem 18.3
  '    Datos.AddItem 19

  'For k = 0 To Datos.ListCount - 1
  '   Datos.List(k) = Replace(Datos.List(k), ",", ".")
  'Next k

  ' Activación y Desactivación de Botones
  Command1.Enabled = False
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command13.Enabled = True
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

Private Sub mnuEjercicioBlanca_Click(Index As Integer)
  Command3_Click
  Open "Sturges.txt" For Output As #1
  Close #1

  Select Case Index
  Case Is = 0
    ' Ejercicio Blanca Clase
    Datos.AddItem 20
    Datos.AddItem 22
    Datos.AddItem 24
    Datos.AddItem 28
    Datos.AddItem 26
    Datos.AddItem 24
    Datos.AddItem 29
    Datos.AddItem 30
    Datos.AddItem 30
    Datos.AddItem 26
    Datos.AddItem 25
    Datos.AddItem 23
    Datos.AddItem 22
    Datos.AddItem 21
    Datos.AddItem 20
    Datos.AddItem 21
    Datos.AddItem 22
    Datos.AddItem 23
    Datos.AddItem 24
    Datos.AddItem 27
    Datos.AddItem 26
    Datos.AddItem 27
    Datos.AddItem 28
    Datos.AddItem 29
    Datos.AddItem 30
    Datos.AddItem 30
    Datos.AddItem 21
    Datos.AddItem 24
    Datos.AddItem 25
    Datos.AddItem 29

  Case Is = 1
    ' Practica Blanca 01
    Datos.AddItem 1
    Datos.AddItem 5
    Datos.AddItem 12
    Datos.AddItem 5
    Datos.AddItem 10
    Datos.AddItem 5
    Datos.AddItem 3
    Datos.AddItem 6
    Datos.AddItem 5
    Datos.AddItem 7
    Datos.AddItem 9
    Datos.AddItem 7
    Datos.AddItem 3
    Datos.AddItem 4
    Datos.AddItem 10
    Datos.AddItem 5
    Datos.AddItem 11
    Datos.AddItem 5
    Datos.AddItem 15
    Datos.AddItem 7
    Datos.AddItem 1
    Datos.AddItem 5
    Datos.AddItem 12
    Datos.AddItem 5
    Datos.AddItem 10
    Datos.AddItem 5
    Datos.AddItem 3
    Datos.AddItem 6
    Datos.AddItem 5
    Datos.AddItem 7
    Datos.AddItem 9
    Datos.AddItem 7
    Datos.AddItem 3
    Datos.AddItem 4
    Datos.AddItem 10
    Datos.AddItem 5
    Datos.AddItem 11
    Datos.AddItem 5
    Datos.AddItem 15
    Datos.AddItem 7
    Datos.AddItem 1
    Datos.AddItem 5
    Datos.AddItem 12
    Datos.AddItem 5
    Datos.AddItem 10
    Datos.AddItem 5
    Datos.AddItem 3
    Datos.AddItem 6
    Datos.AddItem 5
    Datos.AddItem 7
    Datos.AddItem 9
    Datos.AddItem 7
    Datos.AddItem 3
    Datos.AddItem 4
    Datos.AddItem 10
    Datos.AddItem 5
    Datos.AddItem 11
    Datos.AddItem 5
    Datos.AddItem 15
    Datos.AddItem 7

  Case Is = 2
    ' Practica Blanca 02
    Datos.AddItem 20
    Datos.AddItem 21
    Datos.AddItem 30
    Datos.AddItem 32
    Datos.AddItem 24
    Datos.AddItem 25
    Datos.AddItem 23
    Datos.AddItem 26
    Datos.AddItem 25
    Datos.AddItem 27
    Datos.AddItem 29
    Datos.AddItem 27
    Datos.AddItem 23
    Datos.AddItem 24
    Datos.AddItem 20
    Datos.AddItem 25
    Datos.AddItem 21
    Datos.AddItem 30
    Datos.AddItem 31
    Datos.AddItem 32
    Datos.AddItem 20
    Datos.AddItem 21
    Datos.AddItem 30
    Datos.AddItem 32
    Datos.AddItem 24
    Datos.AddItem 25
    Datos.AddItem 23
    Datos.AddItem 26
    Datos.AddItem 25
    Datos.AddItem 27
    Datos.AddItem 29
    Datos.AddItem 27
    Datos.AddItem 23
    Datos.AddItem 24
    Datos.AddItem 20
    Datos.AddItem 25
    Datos.AddItem 21
    Datos.AddItem 30
    Datos.AddItem 31
    Datos.AddItem 32
    Datos.AddItem 20
    Datos.AddItem 21
    Datos.AddItem 30
    Datos.AddItem 32
    Datos.AddItem 24
    Datos.AddItem 25
    Datos.AddItem 23
    Datos.AddItem 26
    Datos.AddItem 25
    Datos.AddItem 27

  Case Is = 3
    ' Practica Blanca 03
    Datos.AddItem 50
    Datos.AddItem 51
    Datos.AddItem 60
    Datos.AddItem 62
    Datos.AddItem 74
    Datos.AddItem 85
    Datos.AddItem 93
    Datos.AddItem 76
    Datos.AddItem 95
    Datos.AddItem 57
    Datos.AddItem 69
    Datos.AddItem 57
    Datos.AddItem 63
    Datos.AddItem 74
    Datos.AddItem 70
    Datos.AddItem 75
    Datos.AddItem 51
    Datos.AddItem 50
    Datos.AddItem 63
    Datos.AddItem 72
    Datos.AddItem 50
    Datos.AddItem 61
    Datos.AddItem 70
    Datos.AddItem 62
    Datos.AddItem 54
    Datos.AddItem 55
    Datos.AddItem 53
    Datos.AddItem 66
    Datos.AddItem 55
    Datos.AddItem 67
    Datos.AddItem 69
    Datos.AddItem 57
    Datos.AddItem 63
    Datos.AddItem 64
    Datos.AddItem 60
    Datos.AddItem 75
    Datos.AddItem 62
    Datos.AddItem 50
    Datos.AddItem 71
    Datos.AddItem 95

  Case Is = 5
    ' Practica Blanca 05
    Datos.AddItem 450
    Datos.AddItem 950
    Datos.AddItem 750
    Datos.AddItem 850
    Datos.AddItem 950
    Datos.AddItem 750
    Datos.AddItem 650
    Datos.AddItem 550
    Datos.AddItem 450
    Datos.AddItem 850
    Datos.AddItem 450
    Datos.AddItem 650
    Datos.AddItem 550
    Datos.AddItem 950
    Datos.AddItem 450
    Datos.AddItem 450
    Datos.AddItem 950
    Datos.AddItem 750
    Datos.AddItem 850
    Datos.AddItem 950
    Datos.AddItem 750
    Datos.AddItem 650
    Datos.AddItem 550
    Datos.AddItem 450
    Datos.AddItem 850
    Datos.AddItem 450
    Datos.AddItem 650
    Datos.AddItem 550
    Datos.AddItem 950
    Datos.AddItem 450
    Datos.AddItem 450
    Datos.AddItem 950
    Datos.AddItem 750
    Datos.AddItem 850
    Datos.AddItem 950
    Datos.AddItem 750
    Datos.AddItem 650
    Datos.AddItem 550
    Datos.AddItem 450
    Datos.AddItem 850
    Datos.AddItem 450
    Datos.AddItem 650
    Datos.AddItem 550
    Datos.AddItem 950
    Datos.AddItem 450
    Datos.AddItem 450
    Datos.AddItem 950
    Datos.AddItem 750
    Datos.AddItem 850
    Datos.AddItem 950
    Datos.AddItem 750
    Datos.AddItem 650
    Datos.AddItem 550
    Datos.AddItem 450
    Datos.AddItem 850
    Datos.AddItem 450
    Datos.AddItem 650
    Datos.AddItem 550
    Datos.AddItem 950
    Datos.AddItem 450

  Case Is = 6
    ' Practica Blanca 06
    Datos.AddItem 2
    Datos.AddItem 1
    Datos.AddItem 3
    Datos.AddItem 3
    Datos.AddItem 4
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 6
    Datos.AddItem 5
    Datos.AddItem 7
    Datos.AddItem 9
    Datos.AddItem 2
    Datos.AddItem 3
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 1
    Datos.AddItem 1
    Datos.AddItem 1
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 1
    Datos.AddItem 3
    Datos.AddItem 2
    Datos.AddItem 4
    Datos.AddItem 5
    Datos.AddItem 3
    Datos.AddItem 2
    Datos.AddItem 5
    Datos.AddItem 7
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 5
    Datos.AddItem 1
    Datos.AddItem 3
    Datos.AddItem 1
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 2
    Datos.AddItem 3
    Datos.AddItem 3
    Datos.AddItem 4
    Datos.AddItem 5
    Datos.AddItem 8
    Datos.AddItem 6
    Datos.AddItem 5
    Datos.AddItem 8

  Case Is = 7
    ' Practica Blanca 07
    Datos.AddItem 30
    Datos.AddItem 51
    Datos.AddItem 76
    Datos.AddItem 82
    Datos.AddItem 24
    Datos.AddItem 95
    Datos.AddItem 59
    Datos.AddItem 161
    Datos.AddItem 95
    Datos.AddItem 57
    Datos.AddItem 189
    Datos.AddItem 175
    Datos.AddItem 163
    Datos.AddItem 74
    Datos.AddItem 70
    Datos.AddItem 75
    Datos.AddItem 51
    Datos.AddItem 50
    Datos.AddItem 63
    Datos.AddItem 72
    Datos.AddItem 150
    Datos.AddItem 161
    Datos.AddItem 170
    Datos.AddItem 62
    Datos.AddItem 150
    Datos.AddItem 155
    Datos.AddItem 53
    Datos.AddItem 166
    Datos.AddItem 155
    Datos.AddItem 100

  Case Is = 8
    ' Practica Blanca 08
    Datos.AddItem 1
    Datos.AddItem 10
    Datos.AddItem 20
    Datos.AddItem 20
    Datos.AddItem 30
    Datos.AddItem 25
    Datos.AddItem 30
    Datos.AddItem 40
    Datos.AddItem 30
    Datos.AddItem 27
    Datos.AddItem 18
    Datos.AddItem 27
    Datos.AddItem 12
    Datos.AddItem 15
    Datos.AddItem 28
    Datos.AddItem 36
    Datos.AddItem 40
    Datos.AddItem 18
    Datos.AddItem 15
    Datos.AddItem 17
    Datos.AddItem 36
    Datos.AddItem 37
    Datos.AddItem 36
    Datos.AddItem 30
    Datos.AddItem 20
    Datos.AddItem 35
    Datos.AddItem 36
    Datos.AddItem 36
    Datos.AddItem 37
    Datos.AddItem 27
    Datos.AddItem 25
    Datos.AddItem 15
    Datos.AddItem 25
    Datos.AddItem 16
    Datos.AddItem 25
    Datos.AddItem 15
    Datos.AddItem 26
    Datos.AddItem 16
    Datos.AddItem 17
    Datos.AddItem 19

  Case Is = 9
    ' Practica Blanca 09
    Datos.AddItem 520
    Datos.AddItem 640
    Datos.AddItem 850
    Datos.AddItem 820
    Datos.AddItem 780
    Datos.AddItem 620
    Datos.AddItem 820
    Datos.AddItem 520
    Datos.AddItem 720
    Datos.AddItem 640
    Datos.AddItem 920
    Datos.AddItem 720
    Datos.AddItem 540
    Datos.AddItem 520
    Datos.AddItem 820
    Datos.AddItem 640
    Datos.AddItem 740
    Datos.AddItem 950
    Datos.AddItem 920
    Datos.AddItem 880
    Datos.AddItem 720
    Datos.AddItem 920
    Datos.AddItem 950
    Datos.AddItem 640
    Datos.AddItem 520
    Datos.AddItem 820
    Datos.AddItem 920
    Datos.AddItem 640
    Datos.AddItem 540
    Datos.AddItem 890
    Datos.AddItem 520
    Datos.AddItem 640
    Datos.AddItem 650
    Datos.AddItem 540
    Datos.AddItem 980
    Datos.AddItem 920
    Datos.AddItem 820
    Datos.AddItem 890
    Datos.AddItem 520
    Datos.AddItem 640
    Datos.AddItem 720
    Datos.AddItem 540
    Datos.AddItem 640
    Datos.AddItem 520
    Datos.AddItem 950
    Datos.AddItem 720
    Datos.AddItem 740
    Datos.AddItem 750
    Datos.AddItem 820
    Datos.AddItem 780

  Case Is = 10
    ' Desde Internet

    Datos.AddItem 1
    Datos.AddItem 2
    Datos.AddItem 4
    Datos.AddItem 4
    Datos.AddItem 5
    Datos.AddItem 6
    Datos.AddItem 7
    Datos.AddItem 8
    Datos.AddItem 9
    Datos.AddItem 11

    Datos.AddItem 1
    Datos.AddItem 2
    Datos.AddItem 4
    Datos.AddItem 4
    Datos.AddItem 5
    Datos.AddItem 6
    Datos.AddItem 7
    Datos.AddItem 8
    Datos.AddItem 10
    Datos.AddItem 12

    Datos.AddItem 1
    Datos.AddItem 2
    Datos.AddItem 4
    Datos.AddItem 4
    Datos.AddItem 6
    Datos.AddItem 6
    Datos.AddItem 7
    Datos.AddItem 8
    Datos.AddItem 10
    Datos.AddItem 12

    Datos.AddItem 1
    Datos.AddItem 3
    Datos.AddItem 4
    Datos.AddItem 4
    Datos.AddItem 6
    Datos.AddItem 6
    Datos.AddItem 7
    Datos.AddItem 8
    Datos.AddItem 10
    Datos.AddItem 12

    Datos.AddItem 2
    Datos.AddItem 3
    Datos.AddItem 4
    Datos.AddItem 5
    Datos.AddItem 6
    Datos.AddItem 7
    Datos.AddItem 8
    Datos.AddItem 9
    Datos.AddItem 11
    Datos.AddItem 12



  Case Is = 11
    ' Desde Internet
    Datos.AddItem 80.502
    Datos.AddItem 35.342
    Datos.AddItem 44.719
    Datos.AddItem 37.402
    Datos.AddItem 61.979
    Datos.AddItem 4.3068
    Datos.AddItem 41.539
    Datos.AddItem 62.215
    Datos.AddItem 51.269
    Datos.AddItem 82.919
    Datos.AddItem 34.182
    Datos.AddItem 37.654
    Datos.AddItem 51.179
    Datos.AddItem 74.582
    Datos.AddItem 58.708
    Datos.AddItem 48.035
    Datos.AddItem 67.124
    Datos.AddItem 41.83
    Datos.AddItem 61.03
    Datos.AddItem 58.267
    Datos.AddItem 74.239
    Datos.AddItem 60.727
    Datos.AddItem 56.155
    Datos.AddItem 86.07
    Datos.AddItem 90.565
    Datos.AddItem 53.751
    Datos.AddItem 76.58
    Datos.AddItem 68.629
    Datos.AddItem 48.24
    Datos.AddItem 57.884
    Datos.AddItem 55.257
    Datos.AddItem 84.656
    Datos.AddItem 48.662
    Datos.AddItem 10.504
    Datos.AddItem 60.951
    Datos.AddItem 38.42
    Datos.AddItem 79.426
    Datos.AddItem 67.662
    Datos.AddItem 53.324
    Datos.AddItem 49.011
    Datos.AddItem 29.48
    Datos.AddItem 41.128
    Datos.AddItem 30.252
    Datos.AddItem 33.412
    Datos.AddItem 47.012
    Datos.AddItem 71.36
    Datos.AddItem 78.635
    Datos.AddItem 41.715
    Datos.AddItem 72.635
    Datos.AddItem 41.463
    Datos.AddItem 48.996
    Datos.AddItem 48.172
    Datos.AddItem 55.643
    Datos.AddItem 55.912
    Datos.AddItem 46.681
    Datos.AddItem 66.519
    Datos.AddItem 59.168
    Datos.AddItem 66.313
    Datos.AddItem 35.884
    Datos.AddItem 28.625
    Datos.AddItem 84.588
    Datos.AddItem 40.709
    Datos.AddItem 50.238
    Datos.AddItem 61.39
    Datos.AddItem 85.72
    Datos.AddItem 45.313
    Datos.AddItem 46.724
    Datos.AddItem 61.752
    Datos.AddItem 63.692
    Datos.AddItem 70.003
    Datos.AddItem 65.996
    Datos.AddItem 55.989
    Datos.AddItem 49.677
    Datos.AddItem 46.502
    Datos.AddItem 67.467
    Datos.AddItem 64.398
    Datos.AddItem 44.411
    Datos.AddItem 41.886
    Datos.AddItem 34.754
    Datos.AddItem 59.888
    Datos.AddItem 59.449
    Datos.AddItem 67.632
    Datos.AddItem 89.116
    Datos.AddItem 69.483
    Datos.AddItem 48.698
    Datos.AddItem 65.854
    Datos.AddItem 75.85
    Datos.AddItem 36.949
    Datos.AddItem 75.548
    Datos.AddItem 69.01
    Datos.AddItem 61.477
    Datos.AddItem 65.585
    Datos.AddItem 52.452
    Datos.AddItem 50.432
    Datos.AddItem 37.748
    Datos.AddItem 51.831
    Datos.AddItem 73.808
    Datos.AddItem 61.065
    Datos.AddItem 35.807
    Datos.AddItem 57.277

  End Select

  ' Activación y Desactivación de Botones
  Command1.Enabled = False
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command13.Enabled = True
  Command15.Enabled = False
  Command16.Enabled = False

End Sub

Private Sub mnuExamen_Click()
  Command3_Click
  Open "Sturges.txt" For Output As #1
  Close #1

  ' Ejercicio Blanca Clase
  Datos.AddItem 25
  Datos.AddItem 2
  Datos.AddItem 5
  Datos.AddItem 15
  Datos.AddItem 22
  Datos.AddItem 30
  Datos.AddItem 18
  Datos.AddItem 10
  Datos.AddItem 15
  Datos.AddItem 18
  Datos.AddItem 28
  Datos.AddItem 15
  Datos.AddItem 30
  Datos.AddItem 5
  Datos.AddItem 2
  Datos.AddItem 10
  Datos.AddItem 30
  Datos.AddItem 25
  Datos.AddItem 2
  Datos.AddItem 14
  Datos.AddItem 16
  Datos.AddItem 18
  Datos.AddItem 25
  Datos.AddItem 22
  Datos.AddItem 30
  Datos.AddItem 2
  Datos.AddItem 23
  Datos.AddItem 22
  Datos.AddItem 23
  Datos.AddItem 30
  Datos.AddItem 16
  Datos.AddItem 17
  Datos.AddItem 20
  Datos.AddItem 15
  Datos.AddItem 25
  Datos.AddItem 30
  Datos.AddItem 25
  Datos.AddItem 22
  Datos.AddItem 2
  Datos.AddItem 5
  Datos.AddItem 18
  Datos.AddItem 27
  Datos.AddItem 22
  Datos.AddItem 18
  Datos.AddItem 17
  Datos.AddItem 28
  Datos.AddItem 2
  Datos.AddItem 4
  Datos.AddItem 16
  Datos.AddItem 32


  ' Activación y Desactivación de Botones
  Command1.Enabled = False
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command13.Enabled = True
  Command15.Enabled = False
  Command16.Enabled = False
End Sub

Private Sub mnuExamen02_Click()
  Command3_Click
  Open "Sturges.txt" For Output As #1
  Close #1

  ' Ejercicio Blanca Clase
  Datos.AddItem 23
  Datos.AddItem 2
  Datos.AddItem 5
  Datos.AddItem 15
  Datos.AddItem 22
  Datos.AddItem 30
  Datos.AddItem 18
  Datos.AddItem 10
  Datos.AddItem 15
  Datos.AddItem 18
  Datos.AddItem 18
  Datos.AddItem 15
  Datos.AddItem 30
  Datos.AddItem 5
  Datos.AddItem 2
  Datos.AddItem 10
  Datos.AddItem 30
  Datos.AddItem 25
  Datos.AddItem 2
  Datos.AddItem 4
  Datos.AddItem 16
  Datos.AddItem 18
  Datos.AddItem 25
  Datos.AddItem 22
  Datos.AddItem 30
  Datos.AddItem 2
  Datos.AddItem 23
  Datos.AddItem 22
  Datos.AddItem 23
  Datos.AddItem 30
  Datos.AddItem 16
  Datos.AddItem 17
  Datos.AddItem 20
  Datos.AddItem 15
  Datos.AddItem 25
  Datos.AddItem 30
  Datos.AddItem 25
  Datos.AddItem 22
  Datos.AddItem 2
  Datos.AddItem 5
  Datos.AddItem 18
  Datos.AddItem 17
  Datos.AddItem 22
  Datos.AddItem 18
  Datos.AddItem 17
  Datos.AddItem 22
  Datos.AddItem 2
  Datos.AddItem 4
  Datos.AddItem 16
  Datos.AddItem 30

  ' Activación y Desactivación de Botones
  Command1.Enabled = False
  Command2.Enabled = False
  Command4.Enabled = False
  Command5.Enabled = False
  Command6.Enabled = False
  Command7.Enabled = False
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  Command11.Enabled = False
  Command12.Enabled = False
  Command13.Enabled = True
  Command15.Enabled = False
  Command16.Enabled = False

End Sub

' RECIBE DATOS PARA LOS LIMITES INFERIORES
Private Sub Text1_Click()
  If Text1.Text = "" Or Not (IsNumeric(Text1.Text)) Then
    Text1.Text = ""
    Text1.SetFocus
  Else
    LimiteInferior.AddItem Replace(Text1.Text, ",", ".")
    Text1.Enabled = False
    Text2.Enabled = True
    Text2.SetFocus
  End If
End Sub

' RECIBE DATOS PARA LOS LIMITES SUPERIORES
Private Sub Text2_Click()
  If Text1.Text = "" Or Not (IsNumeric(Text1.Text)) Then
    Text1.Text = ""
    Text1.SetFocus
  Else
    If Text2.Text = "" Or Not (IsNumeric(Text2.Text)) Then
      Text2.Text = ""
      Text2.SetFocus
    Else
      LimiteSuperior.AddItem Replace(Text2.Text, ",", ".")
      Text2.Enabled = False
      Text3.Enabled = True
      Text3.SetFocus
    End If
  End If
End Sub

' RECIBE DATOS PARA LAS FRECUENCIAS ABSOLUTAS
Private Sub Text3_Click()
  If Text1.Text = "" Or Not (IsNumeric(Text1.Text)) Then
    Text1.Text = ""
    Text1.SetFocus
  Else
    If Text2.Text = "" Or Not (IsNumeric(Text2.Text)) Then
      Text2.Text = ""
      Text2.SetFocus
    Else
      If Text3.Text = "" Or Not (IsNumeric(Text3.Text)) Then
        Text3.Text = ""
        Text3.SetFocus
      Else
        Fi.AddItem Replace(Text3.Text, ",", ".")
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text2.Enabled = False
        Text3.Enabled = False
        Text1.Enabled = True
        Text1.SetFocus
        Command1.Enabled = True
      End If
    End If
  End If
End Sub

' RECIBE DATOS PARA LOS LIMITES INFERIORES
Private Sub Text4_Click()
  If Text4.Text = "" Or Not (IsNumeric(Text4.Text)) Then
    Text4.Text = ""
    Text4.SetFocus
  Else
    Datos.AddItem Replace(Text4.Text, ",", ".")
    Text4.Text = ""
    Text4.SetFocus

    ' Activación y Desactivación de Botones
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    Command9.Enabled = False
    Command10.Enabled = False
    Command11.Enabled = False
    Command12.Enabled = False
    Command13.Enabled = True
    Command16.Enabled = False
  End If
End Sub

' RECIBE DATOS PARA EL LIMITE INTERMEDIO INFERIOR
Private Sub Text5_Click()

End Sub

' RECIBE DATOS PARA EL LIMITE INTERMEDIO SUPERIOR
Private Sub Text6_Click()

End Sub

' MUESTRA EL PERCENTIL DESEADO
Private Sub Slider1_Change()
  IndicePercentil = Slider1.Value
  MiNPercentil.Text = IndicePercentil
  MiVPercentil.Text = Percentiles.List(IndicePercentil - 1)
End Sub

'MENÚ DE OPCIÓN PARA CAMBIER DE SKIN
Private Sub MNUCambiarSkin_Click(Index As Integer)
  Select Case Index
  Case Is = 0
    Mi_Skin = "\Skins\B-Studio.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "B-Studio.skn"
    Close #1

  Case Is = 1
    Mi_Skin = "\Skins\Comander.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Comander.skn"
    Close #1

  Case Is = 2
    Mi_Skin = "\Skins\Cool Breeze.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Cool Breeze.skn"
    Close #1

  Case Is = 3
    Mi_Skin = "\Skins\Copper.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Copper.skn"
    Close #1

  Case Is = 4
    Mi_Skin = "\Skins\Corona.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Corona.skn"
    Close #1

  Case Is = 5
    Mi_Skin = "\Skins\DogmaX.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "DogmaX.skn"
    Close #1

  Case Is = 6
    Mi_Skin = "\Skins\Droid.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Droid.skn"
    Close #1

  Case Is = 7
    Mi_Skin = "\Skins\Green.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Green.skn"
    Close #1

  Case Is = 8
    Mi_Skin = "\Skins\Gris.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Gris.skn"
    Close #1

  Case Is = 9
    Mi_Skin = "\Skins\KOZ.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "KOZ.skn"
    Close #1

  Case Is = 10
    Mi_Skin = "\Skins\LaST v1-2.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "LaST v1-2.skn"
    Close #1

  Case Is = 11
    Mi_Skin = "\Skins\LongHorn.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "LongHorn.skn"
    Close #1

  Case Is = 12
    Mi_Skin = "\Skins\Mac.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "green.skn"
    Close #1

  Case Is = 13
    Mi_Skin = "\Skins\Media.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Media.skn"
    Close #1

  Case Is = 14
    Mi_Skin = "\Skins\Messenger.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Messenger.skn"
    Close #1

  Case Is = 15
    Mi_Skin = "\Skins\Metallic.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Metallic.skn"
    Close #1

  Case Is = 16
    Mi_Skin = "\Skins\MMD.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "MMD.skn"
    Close #1

  Case Is = 17
    Mi_Skin = "\Skins\Neo.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Neo.skn"
    Close #1

  Case Is = 18
    Mi_Skin = "\Skins\Office.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Office.skn"
    Close #1

  Case Is = 19
    Mi_Skin = "\Skins\Office2007.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Office2007.skn"
    Close #1

  Case Is = 20
    Mi_Skin = "\Skins\Orange_Graf.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Orange_Graf.skn"
    Close #1

  Case Is = 21
    Mi_Skin = "\Skins\Paper.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Paper.skn"
    Close #1

  Case Is = 22
    Mi_Skin = "\Skins\SknR.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "SknR.skn"
    Close #1

  Case Is = 23
    Mi_Skin = "\Skins\SoftCrystal.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "SoftCrystal.skn"
    Close #1

  Case Is = 24
    Mi_Skin = "\Skins\St.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "St.skn"
    Close #1

  Case Is = 25
    Mi_Skin = "\Skins\TopSecret.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "TopSecret.skn"
    Close #1

  Case Is = 26
    Mi_Skin = "\Skins\Tp.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Tp.skn"
    Close #1

  Case Is = 27
    Mi_Skin = "\Skins\Web-II.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Web-II.skn"
    Close #1

  Case Is = 28
    Mi_Skin = "\Skins\Winamp 5.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Winamp 5.skn"
    Close #1

  Case Is = 29
    Mi_Skin = "\Skins\Zega.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Zega.skn"
    Close #1

  Case Is = 30
    Mi_Skin = "\Skins\Zhelezo.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Zhelezo.skn"
    Close #1

  Case Is = 31
    Mi_Skin = "\Skins\Zippo.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin MiEstadistica.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Zippo.skn"
    Close #1

  End Select
End Sub

