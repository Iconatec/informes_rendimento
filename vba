Sub ExportarGuiaPDF()
Dim ws As Worksheet
Dim ws2 As Worksheet
Dim caminho As String
Dim OutlookApp As Object
Dim OutlookMail As Object
Set ws = ThisWorkbook.Sheets("Exportar")
Set ws2 = ThisWorkbook.Sheets("RESUMO")
Set OutlookApp = CreateObject("Outlook.Application")
Dim inicio As Double
inicio = Timer
    
ultimaLinha = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row

For i = 2 To ultimaLinha
If ws2.Cells(i, "M").Value = "GERAR" Then

    'Definição dos valores
    nome = ws2.Cells(i, "C").Value
    nomeArquivo = ws2.Cells(i, "B").Value & " - " & "Declaração IRRF" & " - " & nome
    cnpjInvestidor = ws2.Cells(i, "D").Value
    filial = ws2.Cells(i, "E").Value
    cnpjFilial = ws2.Cells(i, "F").Value
    saldo24 = Format(ws2.Cells(i, "G").Value, "#,##0.00")
    aporte = Format(ws2.Cells(i, "H").Value, "#,##0.00")
    capitalD = Format(ws2.Cells(i, "I").Value, "#,##0.00")
    saldo25 = Format(ws2.Cells(i, "J").Value, "#,##0.00")
    distribuicao = Format(ws2.Cells(i, "K").Value, "#,##0.00")
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>SE FOSSE MANDAR E-MAIL<<<<<<<<<<<<
    'destinatario = ws2.Cells(i, "J").Value
    
    'Define o caminho e nome do arquivo PDF
    caminho = "C:\Users\Eduardo Souza\Downloads\" & nomeArquivo & ".pdf"
     
    'Preencher os valores
    ws.Range("A8").Value = "Informamos para fins de Imposto de Renda Ano-Calendário 2025 as informações abaixo do SR. " & _
    nome & ", " & cnpjInvestidor & _
    " na Empresa do Grupo Elio Winter Incorporações"
    
    'Filial
    ws.Cells(10, "A").Value = filial
    ws.Cells(10, "A").Value = Replace(ws.Cells(10, "A").Value, vbLf, " ")
    'CNPJ
    ws.Cells(12, "A").Value = "CNPJ: " & cnpjFilial
    
    'Saldo de 2024
    ws.Cells(14, "A").Value = "Saldo do Capital Investido em 31/12/2024: " & "R$ " & saldo24
    'Aporte em 2024
    ws.Cells(16, "A").Value = "AFAC em 2025: " & "R$ " & aporte
    'Capital devolvido
    ws.Cells(18, "A").Value = "Devolução de AFAC em 2025: " & "R$ " & capitalD
    'Saldo investido
    ws.Cells(20, "A").Value = "Saldo do Capital Investido em 31/12/2025: " & "R$ " & saldo25
    'Distribuição
    ws.Cells(22, "A").Value = "Distribuição de Dividendos 2025: " & "R$ " & distribuicao
        
    'Exporta para PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminho, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
'Enviar o e-mail
'Set OutlookMail = OutlookApp.CreateItem(0)
'With OutlookMail
       ' .To = destinatario
       ' .Subject = "Informe de rendimento 2025"
        '.Body = "Prezado," & vbCrLf & vbCrLf & _
                "Segue em anexo o informe de rendimento." & vbCrLf & vbCrLf & "Por gentileza confirmar o recebimento" & vbCrLf & _
                "Atenciosamente,"
        '.Attachments.Add caminho
      '  .Send
'End With

ws2.Cells(i, "M").Value = "GERADO"
End If
                                                 
Next i

MsgBox "Tempo de execução: " & Format(Timer - inicio, "0.00") & " segundos"
End Sub



