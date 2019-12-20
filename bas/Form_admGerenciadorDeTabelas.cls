VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_admGerenciadorDeTabelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub AtivaBotao()
'Ativa ou desativa botão de atualizar vínculos
If Me.lstTabelas.ItemsSelected.Count > 0 Then
    Me.cmdAtualizar.Enabled = True
Else
    Me.cmdVerificar.SetFocus
    Me.cmdAtualizar.Enabled = False
End If
End Sub

Sub PreencheLista()
Dim cnn As ADODB.Connection
Dim cat As Object
Dim tbl As Object

Dim strLista As String 'origem de linha
Dim strSource As String 'path do BD
Dim strStatus As String 'OK ou !

On Error GoTo ErrHandler

'Banco de dados atual
Set cnn = CurrentProject.Connection
Set cat = CreateObject("ADOX.Catalog")
cat.ActiveConnection = cnn

'Percorre todas as tabelas
For Each tbl In cat.Tables
    'Se for vinculada, inclui na listbox
    If tbl.Type = "LINK" Then
        'Armazena o data source
        strSource = _
        tbl.Properties("Jet OLEDB:Link Datasource")
        'Verifica se o link está OK
        On Error Resume Next
        tbl.Properties("Jet OLEDB:Link Datasource") = _
        strSource
        'Verifica se ocorreu erro
        If Err = 0 Then
            'Se não ocorreu erro, OK
            strStatus = "OK"
        Else
            'Se ocorreu erro, exclamação "!"
            strStatus = "!"
            Err.Clear
        End If
        'Três colunas: status, nome da tabela, endereço
        strLista = strLista & strStatus & ";" & tbl.Name _
        & ";" & strSource & ";"
    End If
Next tbl

'Origem da listbox
Me.lstTabelas.RowSource = strLista

Call AtivaBotao

ExitHere:
'Libera a memória
Set tbl = Nothing
Set cat = Nothing
Set cnn = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description
Resume ExitHere
End Sub

Private Sub cmdAtualizar_Click()
'Atualiza o vínculo das tabelas
Dim fd As Office.FileDialog
Dim strArq As String
Dim varItem As Variant
Dim strTabela As String

Dim cnn As ADODB.Connection
Dim cat As Object
Dim tbl As Object

On Error GoTo ErrHandler

'Diálogo de selecionar arquivo - Office
Set fd = Application.FileDialog(msoFileDialogFilePicker)

fd.Filters.Add "BDs do Access", "*.MDB;*.MDE"
fd.TITLE = "Localize a fonte de dados"
fd.AllowMultiSelect = False
If fd.Show = -1 Then
    strArq = fd.SelectedItems(1)
End If
'Se selecionou arquivo, atualiza os vínculos
If strArq <> "" Then

    'Banco de dados atual
    Set cnn = CurrentProject.Connection
    Set cat = CreateObject("ADOX.Catalog")
    cat.ActiveConnection = cnn
   
    'Percorre os itens da listbox
    For Each varItem In Me.lstTabelas.ItemsSelected
        strTabela = Me.lstTabelas.Column(1, varItem)
        On Error Resume Next
        'Define o novo vínculo
        Set tbl = cat.Tables(strTabela)
        tbl.Properties("Jet OLEDB:Link Datasource") = _
        strArq
        'Se houver erro, avisa
        If Not Err = 0 Then
            MsgBox "Erro ao vincular " & tbl.Name
            Err.Clear
        End If
    Next varItem
    'Atualiza a listbox
    Call PreencheLista
End If

ExitHere:
'Libera a memória
Set tbl = Nothing
Set cat = Nothing
Set cnn = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description
Resume ExitHere
End Sub

Private Sub cmdFiltrar_Click()
Dim intCurrentRow As Long
Dim strCategorias As String
Dim strVendedores As String
Dim strSQL As String


'---------------------
'SELECIONAR CATEGORIAS
'---------------------

For intCurrentRow = 0 To Me.lstCategorias.ListCount - 1
    DoEvents
    
    If Me.lstCategorias.Selected(intCurrentRow) Then
        
        strCategorias = strCategorias & "'" & Me.lstCategorias.ItemData(intCurrentRow) & "',"
        
        ''' DESMARCAR ITEM SELECIONADO
        Me.lstCategorias.Selected(intCurrentRow) = False
    End If
    
Next intCurrentRow

If Len(strCategorias) > 0 Then
    strCategorias = Left(strCategorias, Len(strCategorias) - 1) & ""
'    MsgBox strCategorias, vbInformation, "strCategorias"
'    Saida strCategorias, "strCategorias.log"
End If

'---------------------
'SELECIONAR VENDEDORES
'---------------------

For intCurrentRow = 0 To Me.lstVendedores.ListCount - 1
    DoEvents
    
    If Me.lstVendedores.Selected(intCurrentRow) Then
        
        strVendedores = strVendedores & "'" & Me.lstVendedores.ItemData(intCurrentRow) & "',"
        
        ''' DESMARCAR ITEM SELECIONADO
        Me.lstVendedores.Selected(intCurrentRow) = False
    End If
    
Next intCurrentRow

If Len(strVendedores) > 0 Then
    strVendedores = Left(strVendedores, Len(strVendedores) - 1) & ""
'    MsgBox strVendedores, vbInformation, "strVendedores"
'    Saida strVendedores, "strVendedores.log"
End If

If Len(strVendedores) > 0 And Len(strCategorias) > 0 Then
    strSQL = "SELECT Cadastros.codCadastro, Cadastros.Descricao, Cadastros.Endereco AS end_01, Cadastros.Bairro AS End_03, [Cidade] & ' - ' & [uf] AS end_04, Cadastros.CEP AS end_05, Cadastros.Vendedor, Cadastros.Categoria FROM Cadastros WHERE (((Cadastros.Descricao)<>NULL) AND ((Cadastros.Vendedor) In (" & strVendedores & ")) AND ((Cadastros.Categoria) In (" & strCategorias & "))) ORDER BY Cadastros.Descricao"
'    Saida strSQL, "strSQL_Categorias_Vendedores.log"
    
ElseIf Len(strCategorias) > 0 Then
    strSQL = "SELECT Cadastros.codCadastro, Cadastros.Descricao, Cadastros.Endereco AS end_01, Cadastros.Bairro AS End_03, [Cidade] & ' - ' & [uf] AS end_04, Cadastros.CEP AS end_05, Cadastros.Vendedor, Cadastros.Categoria FROM Cadastros WHERE ((Cadastros.Descricao)<>NULL) AND ((Cadastros.Categoria) In (" & strCategorias & ")) ORDER BY Cadastros.Descricao"
'    Saida strSQL, "strSQL_Categorias.log"

ElseIf Len(strVendedores) > 0 Then
    strSQL = "SELECT Cadastros.codCadastro, Cadastros.Descricao, Cadastros.Endereco AS end_01, Cadastros.Bairro AS End_03, [Cidade] & ' - ' & [uf] AS end_04, Cadastros.CEP AS end_05, Cadastros.Vendedor, Cadastros.Categoria FROM Cadastros WHERE ((Cadastros.Descricao)<>NULL) AND ((Cadastros.Vendedor) In (" & strVendedores & "))  ORDER BY Cadastros.Descricao"
'    Saida strSQL, "strSQL_Vendedores.log"
    
Else
    strSQL = "qryDados"

End If
'Saida strSQL, "strSQL.log"

Me.lstDados.RowSource = strSQL
Me.lstDados.Requery

End Sub

Private Sub cmdLimpar_Click()
'Limpa a seleção
Dim i As Integer

    For i = 0 To Me.lstTabelas.ListCount
        Me.lstTabelas.Selected(i) = False
    Next i
    Call AtivaBotao
    
End Sub

Private Sub cmdNenhum_Click()
Dim i As Integer
    
    For i = 0 To Me.lstDados.ListCount - 1
        Me.lstDados.Selected(i) = False
    Next i
End Sub

Private Sub cmdPesquisar_Click()
Dim strSQL As String

strSQL = Me.lstDados.RowSource
    
strSQL = "SELECT * FROM ( " & strSQL & " ) AS TMP WHERE TMP.Descricao Like '*" & Me.txtPesquisar & "*'"
    
'Saida strSQL, "strSQL.log"

Me.lstDados.RowSource = strSQL
Me.lstDados.Requery
    
End Sub

Private Sub cmdSelecionar_Click()
'Seleciona todos os itens da listbox
Dim i As Integer

    For i = 0 To Me.lstTabelas.ListCount
        Me.lstTabelas.Selected(i) = True
    Next i
    Call AtivaBotao

End Sub

Private Sub cmdTodos_Click()
Dim i As Integer
    
    For i = 0 To Me.lstDados.ListCount - 1
        Me.lstDados.Selected(i) = True
    Next i
End Sub

Private Sub cmdVerificar_Click()
    Call PreencheLista
End Sub

Private Sub cmdVisualizar_Click()
'Verifica se algum item foi selecionado
'Duas Formas
'If Me.lstDados.ListIndex = -1 Then Exit Sub

Dim strSQL As String

If Me.lstDados.ItemsSelected.Count = 0 Then
    strSQL = Me.lstDados.RowSource
    
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "DELETE FROM tmpCadastros"
    DoCmd.RunSQL "INSERT INTO tmpCadastros ( codCadastro, Descricao, end_01, End_03, end_04, end_05, Vendedor, Categoria ) SELECT * FROM ( " & strSQL & " ) as qryDados"
    
    Call SelecaoMultiplaR(Me.lstDados, "", "rpt_6181")
    Me.Visible = False
    
    DoCmd.SetWarnings True
Else
    Call SelecaoMultiplaR(Me.lstDados, "codCadastro", "rpt_6181_Selecao")
    Me.Visible = False
End If

End Sub

Private Sub Form_Open(Cancel As Integer)
    Call PreencheLista
End Sub

Private Sub lstTabelas_Click()
    Call AtivaBotao
End Sub

Private Sub pgeDADOS_Click()
    Me.lstCategorias.Requery
    Me.lstVendedores.Requery
    Me.lstDados.Requery
End Sub

