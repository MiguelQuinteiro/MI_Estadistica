VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Documento 
   BackColor       =   &H00FFFFFF&
   Caption         =   "REPORTES DEL SISTEMA"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13590
   LinkTopic       =   "Form2"
   ScaleHeight     =   10125
   ScaleWidth      =   13590
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "Documento.frx":0000
      Top             =   6240
   End
   Begin VB.ListBox lstTemporal 
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12375
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8916
      _Version        =   393217
      BackColor       =   14737632
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Documento.frx":0234
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Documento"
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
  Aplicar_skin Me

  Form_Resize

  ' Para mostrar la informacion del reporte
  'LIMPIA LIST TEMPORAL
  lstTemporal.Clear
  'LIMPIA RICHTEXTBOX
  rtfText.Text = ""
  ' ABRE REPORTE.TXT PARA LECTURA
  Open "Reporte.txt" For Input As #1
  Do Until EOF(1)
    Line Input #1, A$
    ' CARGA LAS REPORTE EN LIST TEMPORAL
    lstTemporal.AddItem A$
  Loop
  Close #1
  For k = 0 To lstTemporal.ListCount - 1
    lstTemporal.ListIndex = k
    rtfText.Text = rtfText.Text + lstTemporal & vbCr
  Next k
  rtfText.Text = rtfText.Text + " " & vbCr
  rtfText.Text = rtfText.Text + " " & vbCr

  ' Para mostrar la informacion de Sturges
  'LIMPIA LIST TEMPORAL
  lstTemporal.Clear
  ' ABRE STURGES.TXT PARA LECTURA
  Open "Sturges.txt" For Input As #1
  Do Until EOF(1)
    Line Input #1, A$
    ' CARGA LAS STURGES EN LIST TEMPORAL
    lstTemporal.AddItem A$
  Loop
  Close #1
  For k = 0 To lstTemporal.ListCount - 1
    lstTemporal.ListIndex = k
    rtfText.Text = rtfText.Text + lstTemporal & vbCr
  Next k
  rtfText.Text = rtfText.Text + " " & vbCr
  rtfText.Text = rtfText.Text + " " & vbCr

  ' Para mostrar la informacion de los calculos
  'LIMPIA LIST TEMPORAL
  lstTemporal.Clear
  ' ABRE CALCULOS.TXT PARA LECTURA
  Open "Calculos.txt" For Input As #1
  Do Until EOF(1)
    Line Input #1, A$
    ' CARGA LAS CALCULOS EN LIST TEMPORAL
    lstTemporal.AddItem A$
  Loop
  Close #1
  For k = 0 To lstTemporal.ListCount - 1
    lstTemporal.ListIndex = k
    rtfText.Text = rtfText.Text + lstTemporal & vbCr
  Next k
  rtfText.Text = rtfText.Text + " " & vbCr
  rtfText.Text = rtfText.Text + " " & vbCr

  ' Para mostrar la informacion de las Ecuaciones
  'LIMPIA LIST TEMPORAL
  lstTemporal.Clear
  ' ABRE ECUACIONES.TXT PARA LECTURA
  Open "Ecuaciones.txt" For Input As #1
  Do Until EOF(1)
    Line Input #1, A$
    ' CARGA LAS ECUACIONES EN LIST TEMPORAL
    lstTemporal.AddItem A$
  Loop
  Close #1
  For k = 0 To lstTemporal.ListCount - 1
    lstTemporal.ListIndex = k
    rtfText.Text = rtfText.Text + lstTemporal & vbCr
  Next k
  rtfText.Text = rtfText.Text + " " & vbCr
  rtfText.Text = rtfText.Text + " " & vbCr
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
  rtfText.RightMargin = rtfText.Width - 400
End Sub


