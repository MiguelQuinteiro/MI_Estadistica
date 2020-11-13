Attribute VB_Name = "Module1"
' *****************************************************************************
' * PROYECTO   : SISTEMA DE ESTADÍSTICA
' * FORMULARIO : Módulo General
' * AUTORES    : Miguel Quinteiro
' * FECHA      : 25 de Abril de 2008
' * ***************************************************************************

'DECLARA VARIABLES DE BASE DE DATOS Y TABLA
Global Mi_Skin As String


'PROCEDIMIENTOS GENERALES A TODO EL SISTEMA
Public Sub Aplicar_skin(ByVal Formulario As Form)
  MiEstadistica.Skin1.LoadSkin App.Path & Mi_Skin
  MiEstadistica.Skin1.ApplySkin Formulario.hWnd
End Sub


