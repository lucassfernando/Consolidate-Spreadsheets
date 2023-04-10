Sub extrairDados()

    Dim pasta As String
    Dim planilha As String
    Dim linha As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Definir a pasta onde as planilhas estão localizadas
    pasta = "PASTA"
    
    ' Definir a planilha onde os dados serão armazenados
    Set ws = ThisWorkbook.Sheets("Dados")
    linha = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    ' Loop pelas planilhas na pasta
    planilha = Dir(pasta & "*.xlsx")
    Do While planilha <> ""
        
        ' Abrir a planilha
        Set wb = Workbooks.Open(pasta & planilha)
        
        ' Copiar os dados da linha 2
        ws.Range("A" & linha).Value = wb.Sheets(1).Range("A2").Value
        ws.Range("B" & linha).Value = wb.Sheets(1).Range("B2").Value
        ws.Range("C" & linha).Value = wb.Sheets(1).Range("C2").Value
        ws.Range("D" & linha).Value = wb.Sheets(1).Range("D2").Value
        ws.Range("E" & linha).Value = wb.Sheets(1).Range("E2").Value
        ws.Range("F" & linha).Value = wb.Sheets(1).Range("F2").Value
        ws.Range("G" & linha).Value = wb.Sheets(1).Range("G2").Value
        
        ' Fechar a planilha sem salvar as alterações
        wb.Close False
        
        ' Selecionar a próxima planilha na pasta
        planilha = Dir
        
        ' Incrementar a linha
        linha = linha + 1
        
    Loop

End Sub
