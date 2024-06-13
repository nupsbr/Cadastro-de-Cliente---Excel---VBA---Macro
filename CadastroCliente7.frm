VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CadastroCliente 
   Caption         =   "Cadastro Cliente"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6435
   OleObjectBlob   =   "CadastroCliente7.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CadastroCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bntCANCELAR_Click()

Unload CadastroCliente
MsgBox "Cadastro nao foi salvo!", vbInformation, "Cadastro Cancelado"

End Sub
Private Sub bntSALVAR_Click()

If lbTIPOCADASTRO.Caption = "NOVO CADASTRO" Then
salvarnovocadastro
ElseIf lbTIPOCADASTRO.Caption = "EDICAO DE CADASTRO" Then
SalvarEdicaoCadastro
End If

End Sub

Private Sub SalvarEdicaoCadastro()
Dim campoembranco As Boolean
Dim linha As Integer

campoembranco = False

'inicar a verificacao se os campos obrigatorios estao preenchidos'

If TextNOME = Empty Then campoembranco = True
If TextCELULAR = Empty Then campoembranco = True

If campoembranco = True Then
MsgBox "Campos marcados com * sao obrigatorios", vbInformation, "Campos em branco"
Exit Sub
End If

linha = 5
Do While Planilha2.Cells(linha, 2) <> CDbl(lbTIPOCADASTRO.Tag)
linha = linha + 1
Loop
'inicar armazenamento de dados'

Planilha2.Unprotect 123

Planilha2.Cells(linha, 3) = TextNOME.Value
Planilha2.Cells(linha, 4) = TextCELULAR.Value
Planilha2.Cells(linha, 5) = TextDATA.Value
Planilha2.Cells(linha, 6) = TextCPF.Value
Planilha2.Cells(linha, 7) = TextCEP.Value
Planilha2.Cells(linha, 8) = TextENDERECO.Value
Planilha2.Cells(linha, 9) = TextNUMERO.Value
Planilha2.Cells(linha, 10) = TextBAIRRO.Value
Planilha2.Cells(linha, 11) = TextCIDADE.Value
Planilha2.Cells(linha, 12) = TextESTADO.Value

Planilha2.Protect 123

Unload CadastroCliente

MsgBox "Cadastro editado com sucesso!", vbInformation, "Cadastro editado"

End Sub


Private Sub salvarnovocadastro()

Dim campoembranco As Boolean
Dim linha As Integer

campoembranco = False

'inicar a verificacao se os campos obrigatorios estao preenchidos'

If TextNOME = Empty Then campoembranco = True
If TextCELULAR = Empty Then campoembranco = True

If campoembranco = True Then
MsgBox "Campos marcados com * sao obrigatorios", vbInformation, "Campos em branco"
Exit Sub
End If

'descobrir a ultima linha preenchida na tabela'

linha = 5
Do While Planilha2.Cells(linha, 3) <> Empty
linha = linha + 1
Loop

'inicia o armazenamento dos dados do cadastro do cliente'

Planilha2.Unprotect 123

Planilha2.Cells(linha, 2) = Planilha2.Range("N2").Value + 1
Planilha2.Cells(linha, 3) = TextNOME.Value
Planilha2.Cells(linha, 4) = TextCELULAR.Value
Planilha2.Cells(linha, 5) = TextDATA.Value
Planilha2.Cells(linha, 6) = TextCPF.Value
Planilha2.Cells(linha, 7) = TextCEP.Value
Planilha2.Cells(linha, 8) = TextENDERECO.Value
Planilha2.Cells(linha, 9) = TextNUMERO.Value
Planilha2.Cells(linha, 10) = TextBAIRRO.Value
Planilha2.Cells(linha, 11) = TextCIDADE.Value
Planilha2.Cells(linha, 12) = TextESTADO.Value
Planilha2.Cells(linha, 13) = Date
'fechar cadastro com sucesso'

'inserir botao editar e excluir registro'
inserirbotoesacao linha, Planilha2.Cells(linha, 2)

Unload CadastroCliente

Planilha2.Range("N2").Value = Planilha2.Range("N2").Value + 1

Planilha2.Protect 123

MsgBox "Cadastro salvo com sucesso!", vbInformation, "Cadastro Salvo"

End Sub










Private Sub TextCELULAR_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

TextCELULAR.MaxLength = 15

'somente numeros'
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

Select Case TextCELULAR.SelStart
'colocar caracters'
Case Is = 0
TextCELULAR.SelText = "("
Case Is = 3
TextCELULAR.SelText = ") "
Case Is = 9
TextCELULAR.SelText = "-"

End Select
End Sub



Private Sub TextCPF_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'somente numeros'
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
'colocar caracters'

TextCPF.MaxLength = 14
Select Case TextCPF.SelStart

Case Is = 3, 7
TextCPF.SelText = "."
Case Is = 11
TextCPF.SelText = "-"
End Select

End Sub

Private Sub TextDATA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

TextDATA.MaxLength = 10

'somente numeros'
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

Select Case TextDATA.SelStart
'colocar caracters'
Case Is = 2, 5
TextDATA.SelText = "/"
End Select


End Sub



Private Sub TextNUMERO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)


If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If

End Sub

