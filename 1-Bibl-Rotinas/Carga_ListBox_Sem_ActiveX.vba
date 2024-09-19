Sub Carga_ListBox_Sem_ActiveX()
    
    Dim shApCel As Worksheet
    
    Dim pvtTable As pivotTable
    Dim pvtField As PivotField
    Dim pvtItem  As PivotItem
    
    Dim lstApCel As ListBox
    
    Dim intCont         As Integer
    Dim arrLstName_()   As String
    
    Dim intPosSeparador As Integer   ' Posição do caractere strCaractere na string
    Dim intTamStr       As Integer   ' Quantidade total de caracteres na string
    Dim intNovoTamStr   As Integer   ' Tamanho da String Tratada SEM o caractere [ mais a ESQUERDA
    
    Dim strCaractere    As String    ' Caractere que divide a String
    
    Set wb = ThisWorkbook
     
     ' Definir a planilha onde a ListBox está localizada
     'Set sh = wb.Sheets("pEmissaoTermos")
     
    strCaractere = "&"               ' Caractere que divide a String referente ao Nome do Funcionário
    ReDim arrLstName_(wb.Sheets.Count)
     
    For intCont = 1 To UBound(arrLstName_)
       arrLstName_(intCont) = Worksheets(intCont).Name
    Next intCont
     
     'fxSelectsheet (arrLstName_(3))
     
    Set sh = wb.Sheets(arrLstName_(1))                   ' Define a planilha para Emissão dos Termos
    Set lstApCel = sh.ListBoxes("lstApCelular")
    Set shApCel = wb.Sheets(arrLstName_(3))              ' Define a planilha e a tabela dinâmica
     
    For Each pvtTable In shApCel.PivotTables             ' Iterar sobre todas as Tabelas Dinâmicas
       '
       For Each pvtField In pvtTable.PivotFields    ' Iterar sobre todos os campos da Tabela Dinâmica
          '
          If Right(pvtField.Name, 17) = "[NomeFuncionario]" Then
             '
             Set pvtField = pvtTable.PivotFields(pvtField.Name)    ' Use o nome encontrado
                                 
                With lstApCel
                   .RemoveAllItems
                   For Each pvtItem In pvtField.PivotItems
                      '
                      intTamStr = Len(pvtItem.Name)                                              ' Quantidade total de caracteres na String
                      intPosSeparador = InStr(pvtItem.Name, strCaractere) + 1                    ' Posição do strCaractere na String
                      intNovoTamStr = Len(Right(pvtItem.Name, intTamStr - intPosSeparador))
                      
                      .AddItem Left(Right(pvtItem.Name, intNovoTamStr), (intNovoTamStr - 1))
                      '.AddItem strNomeFunc
                   Next pvtItem
                End With
                Exit Sub
          End If
       Next pvtField
    Next pvtTable
    
    MsgBox "Campo 'NomeFuncionario' não encontrado na Tabela Dinâmica."
    
 End Sub