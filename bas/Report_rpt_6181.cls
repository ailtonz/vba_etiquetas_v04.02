VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_6181"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub Cabe�alhoDoRelat�rio_Print(Cancel As Integer, PrintCount As Integer)
Call zeraVariaveis 'para imprimir em branco
End Sub

Private Sub Detalhe_Print(Cancel As Integer, PrintCount As Integer)
Dim Result As Long 'Variavel para chamar a fun��o codigo de barras
Dim numCopia As Integer, pula As Integer

On Error Resume Next
    'Chame a fun��o assim para criar os c�digos de Barra
    'Nome do controle caixa de texto para conter o desenho - n�o acoplada
    'e o relat�rio atual
'    Result = MD_Barcode39(CodBarra, Me)
    
    'c�digo para o gerador de etiquetas
    numCopia = Forms!admGerenciadorDeTabelas!txtCopias
    pula = Forms!admGerenciadorDeTabelas!txtBranco
    Call GeraEtiquetas(Me, numCopia, pula)

End Sub

Private Sub Report_Close()
On Error Resume Next
Forms!admGerenciadorDeTabelas.Visible = True
End Sub

Private Sub Report_Open(Cancel As Integer)
Call zeraVariaveis 'para imprimir em branco
DoCmd.Maximize
End Sub

