Attribute VB_Name = "Módulo11"

Sub abrircadastrocliente()
CadastroCliente.lbTIPOCADASTRO.Tag = 0
CadastroCliente.lbTIPOCADASTRO.Caption = "NOVO CADASTRO"
CadastroCliente.Show

End Sub

Sub AbrirEditarCadastro()
CadastroCliente.lbTIPOCADASTRO.Tag = Mid(Application.Caller, 10)
CadastroCliente.lbTIPOCADASTRO.Caption = "EDICAO DE CADASTRO"
PesquisadordeRegistro CadastroCliente.lbTIPOCADASTRO.Tag
CadastroCliente.Show

End Sub

Sub PesquisadordeRegistro(idcadastro As Integer)

Dim linha As Integer

linha = 5
Do While Planilha2.Cells(linha, 2) <> idcadastro
linha = linha + 1
Loop

With CadastroCliente
.TextNOME.Value = Planilha2.Cells(linha, 3)
.TextCELULAR.Value = Planilha2.Cells(linha, 4)
.TextDATA.Value = Planilha2.Cells(linha, 5)
.TextCPF.Value = Planilha2.Cells(linha, 6)
.TextCEP.Value = Planilha2.Cells(linha, 7)
.TextENDERECO.Value = Planilha2.Cells(linha, 8)
.TextNUMERO.Value = Planilha2.Cells(linha, 9)
.TextBAIRRO.Value = Planilha2.Cells(linha, 10)
.TextCIDADE.Value = Planilha2.Cells(linha, 11)
.TextESTADO.Value = Planilha2.Cells(linha, 12)
End With

End Sub

Sub ExcluirCadastro()

Dim linha As Integer
Dim idcadastro As Integer
Dim msgboxresposta As Integer

msgboxresposta = MsgBox("Tem certeza que deseja excluir este cadastro?", vbInformation + vbYesNo, "Exclusao de cadastro")

If msgboxresposta = vbNo Then Exit Sub

idcadastro = Mid(Application.Caller, 11)

linha = 5
 Do While Planilha2.Cells(linha, 2) <> idcadastro
 linha = linha + 1
 Loop
 
 Planilha2.Unprotect 123
 
 
 Rows(linha & ":" & linha).Delete shift:=xlUp
 Planilha2.Shapes("ICOeditar" & idcadastro).Delete
 Planilha2.Shapes("ICOexcluir" & idcadastro).Delete

Planilha2.Protect 123

MsgBox "Cadastro excluido com sucesso!", vbInformation, "Exclusao de cadstro"





End Sub

Sub inserirbotoesacao(linha As Integer, idcadastro As Integer)
'icone editar(lapis)'
Planilha2.Shapes("ICOeditar").Copy
Planilha2.Range("N" & linha).Activate
Planilha2.Paste

Selection.ShapeRange.Name = "ICOeditar" & idcadastro
Selection.ShapeRange.IncrementTop (0.555905512 * 4)
Selection.ShapeRange.IncrementLeft (0.555905512 * 6)
Selection.OnAction = "AbrirEditarCadastro"

'icone excluir(lixeira)'

Planilha2.Shapes("ICOexcluir").Copy
Planilha2.Range("O" & linha).Activate
Planilha2.Paste

Selection.ShapeRange.Name = "ICOexcluir" & idcadastro
Selection.ShapeRange.IncrementTop (0.555905512 * 4)
Selection.ShapeRange.IncrementLeft (0.555905512 * 6)
Selection.OnAction = "ExcluirCadastro"



End Sub
