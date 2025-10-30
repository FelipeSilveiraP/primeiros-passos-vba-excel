Option Explicit

Private Sub Cmb_calcular_Click()
'-----------------variáveis --> MÉDIA PRECO NOVO P/ PROF "S" REFRIGERACAO E DO TIPO ALIMENTAR
Dim novopreco_cel As Double, soma_m As Double, media_novopreco As Double, cont_m As Integer
Dim refrig_cel As String, tipo_cel As String

'-----------------variável --> SOMA DOS IMPOSTOS
Dim cel_impostos As Double, soma_impostos As Double

'-----------------variavel --> SOMA QUANT. PROD VESTUARIO OU LIMPEZA
Dim quant_prod As Double

'-----------------variavel --> PROD COM > imposto
Dim MAIORIMPOSTO As Double, class_cel As String, classif As String

'-----------------variavel --> < Vlr PRECONOVO PROD. Vlr ADIC ENTRE R$4,5 --> R$ 7,5
Dim adc_cel As Double, menorVlrPrecoN As Double

'-----------------variavel --> < Vlr DESCONTO PROD. N ALIMENTO / REGRIGERACAO (N)
Dim desconto_cel As Double, menorDesconto As Double

'------------------VARIAVEL PARA FOR, LINHA, msg
Dim i As Integer, LINHA As Integer, msg As String

'------------------ comando linha ----------------------------
LINHA = Range("A300").End(xlUp).Row

soma_m = 0
quant_prod = 0
MAIORIMPOSTO = 0
menorVlrPrecoN = 30000
menorDesconto = 30000
classif = ""

For i = 2 To LINHA
    
    tipo_cel = Cells(i, 2).Value
    refrig_cel = Cells(i, 3).Value
    adc_cel = Cells(i, 4).Value
    cel_impostos = Cells(i, 5).Value
    
    desconto_cel = Cells(i, 7).Value
    novopreco_cel = Cells(i, 8).Value
    class_cel = Cells(i, 9).Value
       
    '---------------MÉDIA PRECO NOVO P/ PROF "S" REFRIGERACAO E DO TIPO ALIMENTAR
    If (refrig_cel = "S") And (tipo_cel = "A") Then
            soma_m = soma_m + novopreco_cel
            cont_m = cont_m + 1
    End If
    
    '---------------SOMA DE IMPOSTOS
    If (cel_impostos > 0) Then
        soma_impostos = soma_impostos + cel_impostos
    End If
    
    '---------------SOMA QUANT. PROD VESTUARIO OU LIMPEZA
    If (tipo_cel = "V") Or (tipo_cel = "L") Then
        quant_prod = quant_prod + 1
    End If
    
    '---------------PROD COM > IMPOSTO
    If (cel_impostos > MAIORIMPOSTO) Then
        MAIORIMPOSTO = cel_impostos
        classif = class_cel
    End If
    
    '--------------- < Vlr PRECONOVO PROD. Vlr ADIC ENTRE R$4,5 --> R$ 7,5
    If (adc_cel >= 4.5) And (adc_cel <= 7.5) Then
    
        If (novopreco_cel < menorVlrPrecoN) Then
            menorVlrPrecoN = novopreco_cel
        End If
        
    End If
    
    '--------------- < Vlr DESCONTO PROD. N ALIMENTO / REGRIGERACAO (N)
    If (tipo_cel <> "A") Or (refrig_cel = "N") Then
        If (desconto_cel < menorDesconto) Then
            menorDesconto = desconto_cel
        End If
    End If
        
Next i

media_novopreco = soma_m / cont_m

msg = "MÉDIA PRECO NOVO DOS PROD. REFRIGERADOS E DO TIPO ALIMENTOS = R$ " & media_novopreco & vbCrLf
msg = msg & vbCrLf & "Soma Todos Valores Impostos Pagos = R$ " & soma_impostos & vbCrLf
msg = msg & vbCrLf & "QUANT. PROD EM VESTUARIO OU LIMPEZA = " & quant_prod & " unids" & vbCrLf
msg = msg & vbCrLf & "CLASSIF. PROD COM MAIOR IMPOSTO = " & classif & " --> R$ " & MAIORIMPOSTO & vbCrLf
msg = msg & vbCrLf & "menor Vlr NOVO PRECO com ADICIONAL ENTRE R$ 4,50 --> R$ 7,50 = R$ " & menorVlrPrecoN & vbCrLf
msg = msg & vbCrLf & "menor Vlr DESCONTO PROD. NÃO ALIMENTÍCIOS E SEM REFRIGERAÇÃO = R$ " & menorDesconto & vbCrLf

Label5.Caption = msg
End Sub

Private Sub Cmb_cancel_Click()
Unload Me
End Sub


Private Sub Cmb_Ok_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'NOME: FELIPE SILVEIRA PESSOA
'TURMA: ADS (NOTURNO)

'declaracao de variáveis
Dim preco As Double, tipo_prod As String
Dim preco_custo As Double, preco_novo As Double

Dim refrig As String, adic As Double, classif As String

Dim imposto As Double, desc As Double, porcen As Double

Dim LINHA As Double

LINHA = Range("A200").End(xlUp).Row + 1

'leitura de dados

preco = Txb_p.Text

tipo_prod = 0
refrig = 0
adic = 0
preco_custo = 0
preco_novo = 0
imposto = 0
porcen = 0
desc = 0
classif = 0


'------------- REFRIGERAÇÃO ------------

'------ NÃO
If (Opt_n.Value) Then
    refrig = "N"
    
    If (Opt_a.Value) Then
    tipo_prod = "A"
        If (preco < 15) Then
            adic = 2
            
            ElseIf (preco >= 15) Then
            adic = 5
        End If
    End If
    
    If (Opt_l.Value) Then
    tipo_prod = "L"
        If (preco < 10) Then
            adic = 1.5
            
            ElseIf (preco >= 10) Then
            adic = 2.5
        End If
    End If
    
    If (Opt_v.Value) Then
    tipo_prod = "V"
        If (preco < 30) Then
            adic = 3
            
            ElseIf (preco >= 30) Then
            adic = 2.5
        End If
    End If
End If

'------ SIM
If (Opt_s.Value) Then
    refrig = "S"
    
        If (Opt_a.Value) Then
            tipo_prod = "A"
            adic = 8
        End If
        
            If (Opt_l.Value) Then
                tipo_prod = "L"
                adic = 0
            End If
            
                If (Opt_v.Value) Then
                    tipo_prod = "V"
                    adic = 0
                End If
End If

'---------- IMPOSTO -----------------

If (preco < 25) Then
    porcen = 5 / 100
    
        ElseIf (preco >= 25) Then
        porcen = 8 / 100
End If

imposto = porcen * preco


'---------- PREÇO CUSTO ----------
preco_custo = preco + imposto


'----------- DESCONTO ------------
If (tipo_prod <> "A") And (refrig = "N") Then
    desc = preco - (preco * 3 / 100)
 
    Else
    desc = 0
End If

'----------- NOVO PRECO ----------
preco_novo = preco_custo + adic - desc


If (preco_novo <= 50) Then
    classif = "BARATO"
    
        ElseIf (preco_novo > 50) And (preco_novo < 100) Then
        classif = "Normal"
        
            ElseIf (preco_novo >= 100) Then
            classif = "Caro"
End If

'saida de dados
Cells(LINHA, 1).Value = preco
Cells(LINHA, 2).Value = tipo_prod
Cells(LINHA, 3).Value = refrig
Cells(LINHA, 4).Value = adic
Cells(LINHA, 5).Value = imposto
Cells(LINHA, 6).Value = preco_custo
Cells(LINHA, 7).Value = desc
Cells(LINHA, 8).Value = preco_novo
Cells(LINHA, 9).Value = classif

'apaga caixas de texto
Txb_p.Value = ""

Opt_a.Value = False
Opt_l.Value = False
Opt_v.Value = False

Opt_s.Value = False
Opt_n.Value = False

End Sub

