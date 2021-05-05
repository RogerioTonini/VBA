' Auto....: Rogerio Tonini
' Data....: 03/05/2021
' Objetivo: Carregar dados de tabelas para qualquer objeto ComboBox

Sub CargaCombo(frm As UserForm, ByVal strObjeto As String, Optional wkbNomeArq As Workbook)

   Dim ctrControle   As Control
   
   Dim tbTable       As Excel.ListObject

   Dim lngQtdPlan    As Long     ' Contador para a quantidade de planilhas no arquivo
   Dim lngContLin    As Long     ' Contador de linha
   Dim lngContCol    As Long     ' Contador de coluna

   Dim arrTmp        As Variant
   
   Dim wksNomePlan   As Worksheet
     
   If wkbNomeArq Is Nothing Then Set wkbNomeArq = ThisWorkbook
   '
   ' Captura o CODENAME da planilha a ser utilizada
   '
   For lngQtdPlan = 1 To wkbNomeArq.Worksheets.Count
      If wkbNomeArq.Worksheets(lngQtdPlan).CodeName = "ws" & strObjeto Then
         Set wksNomePlan = wkbNomeArq.Worksheets(lngQtdPlan)
         Exit For
      End If
   Next lngQtdPlan
   Set tbTable = wksNomePlan.ListObjects("tb" & strObjeto)
   '
   ' Carrega a Tabela para a matriz temporária e posterior para a ComboBox
   '
   arrTmp = tbTable.Range
   arrTmp = Range("tb" & strObjeto).Value
   For Each ctrControle In frm.Controls
      With ctrControle
         If TypeOf ctrControle Is MSForms.ComboBox Then
            If .Name = "cmb" & strObjeto Then
               .Clear
               For lngContLin = LBound(arrTmp, 1) To UBound(arrTmp, 1)
                  For lngContCol = UBound(arrTmp, 2) To LBound(arrTmp, 2) Step -1
                     If lngContCol = 3 And arrTmp(lngContLin, 4) = "A" Then
                        .AddItem (arrTmp(lngContLin, lngContCol))
                        .List(0, 1) = arrTmp(lngContLin, 1)
                     End If
                  Next lngContCol
               Next lngContLin
            End If
         End If
      End With
   Next
   Erase arrTmp
   Set tbTable = Nothing
   Set wksNomePlan = Nothing
   Set wkbNomeArq = Nothing

End Sub
