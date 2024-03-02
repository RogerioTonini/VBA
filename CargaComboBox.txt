Sub CargaCombo(frm As UserForm, ByVal strObjeto As String)
'
' Autor.....: ROGERIO TONINI
' Data......: 03/05/2021
' Objetivo..: Popular objeto ComboBox
' Parametros: frm       - Formulário atual
'             strObjeto - Qual o nome do objeto esta sendo manipulado.
'                 Durante a execução é acrescentado um tipo diferente prefixo: "ws"  - CODENAME
'                                                                              "tb"  - Nome da tabela
'                                                                              "cmb" - Nome do ComboBox
'
   Dim ctrControle   As Control
   
   Dim tbTable       As Excel.ListObject

   Dim lngQtdPlan    As Long     ' Contador para a quantidade de planilhas no arquivo
   Dim lngContLin    As Long     ' Contador de linha
   Dim lngContCol    As Long     ' Contador de coluna

   Dim arrTmp        As Variant
   
   Dim wksNomePlan   As Worksheet
   Dim wkbNomeArq    As Workbook
     
   Set wkbNomeArq = ThisWorkbook
   '
   ' Captura o CODENAME da planilha
   '
   For lngQtdPlan = 1 To wkbNomeArq.Worksheets.Count
      If wkbNomeArq.Worksheets(lngQtdPlan).CodeName = "ws" & strObjeto Then
         Set wksNomePlan = wkbNomeArq.Worksheets(lngQtdPlan)
         Exit For
      End If
   Next lngQtdPlan
   Set tbTable = wksNomePlan.ListObjects("tb" & strObjeto)
   '
         ' Carrega a Tabela para a matriz temporária e posteriormente para o ComboBox
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
