' Lógica VBA
' Renam Fernando Ruthes
'
' Códigos adaptados da planilha SGE
' Módulos de controle de Eduardo Ongareli
'



Private Sub btn_dataFIM_Click()

    Me.Txt_FiltraATE = GetCalendário
    
   End Sub
   
   Private Sub btn_dataINI_Click()
   
   Me.Txt_FiltraDE = GetCalendário
   
   End Sub
   
   
   Private Sub btn_FiltroAtual_Click()
   
       Dim Linha  As Long
       Dim xMax   As Long
       Dim Regs   As Integer
       Dim DtINI  As Date
       Dim DtFIM  As Date
       Dim xData  As Date
       Dim xFalta As Currency
       Dim xPagos As Currency
       Dim Lv
       xData = Date
       
          
       Plan01.Activate
       xMax = Range("A" & Rows.Count).End(xlUp).Row + 1
       Linha = 2
       DtINI = Me.Txt_FiltraDE
       DtFIM = Me.Txt_FiltraATE
           
       Me.ListView1.ListItems.Clear
       
       Do Until Plan01.Cells(Linha, 1) = ""
           
            If Cells(Linha, 2) >= DtINI And Cells(Linha, 2) <= DtFIM Then
               
                 Set Lv = ListView1.ListItems.Add(Text:=Plan01.Cells(Linha, 1)) 'Código
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 2) ' Data
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 3) ' Produto
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 4) ' Unidade
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 5) ' Quantidade
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 6) ' Valor Unitário
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 7) ' Valor Total
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 8) ' Fornecedor
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 9) ' Endereço
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 10) 'Cidade
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 11) 'Estado
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 12) 'CEP
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 13) 'Telefone
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 14) 'Contato
                   
                   Regs = Regs + 1
           
           End If
           
           Linha = Linha + 1
           
       Loop
       
       Set Lv = Nothing
      
       Me.lbl_registros = Regs & " Registros localizados"
       
       
       
   End Sub
   
   Private Sub Btn_todos_Click()
   
   Me.txt_busca = ""
   
   
   End Sub
   
   Private Sub CommandButton1_Click()
   
   End Sub
   
   Private Sub txt_busca_Change()
   
   
       Dim Abas  As Worksheet
       Dim Linha As Integer
       Dim xPesq As String
       Dim xCel  As String
       Dim Regs  As Integer
       Dim xRegs As Integer
       Dim xPlan As String
       Dim Lv
       
       Plan01.Activate
       xPlan = Plan01.Name
       xRegs = Range("A" & Rows.Count).End(xlUp).Row
       
       xPesq = txt_busca.Text
       Linha = 2
       
           
       Set Abas = ThisWorkbook.Worksheets(xPlan)
       With Abas
       
           Me.ListView1.ListItems.Clear
           While .Cells(Linha, 1) <> Empty
          
               
               For coluna = 3 To 14 'faz a pesquisa entre as colunas 3 e 14
               xCel = .Cells(Linha, coluna)
               
               If InStr(1, UCase(xCel), UCase(xPesq), 1) Then
                                      
                   Set Lv = ListView1.ListItems.Add(Text:=Plan01.Cells(Linha, 1)) 'Código
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 2) ' Data
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 3) ' Produto
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 4) ' Unidade
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 5) ' Quantidade
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 6) ' Valor Unitário
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 7) ' Valor Total
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 8) ' Fornecedor
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 9) ' Endereço
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 10) 'Cidade
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 11) 'Estado
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 12) 'CEP
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 13) 'Telefone
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 14) 'Contato
                        GoTo proxima_linha
                        
                  
               End If
               
               Next coluna
               
   proxima_linha:
               
             Linha = Linha + 1
                
             Wend
             
           
       End With
       Set Lv = Nothing
      
       Me.lbl_registros = ListView1.ListItems.Count & " Registros localizados"
   
   
   End Sub
   
   Private Sub UserForm_Initialize()
   
    Me.ListView1.ListItems.Clear
       With ListView1
           .Gridlines = True
           .View = lvwReport
           .FullRowSelect = True
           .ColumnHeaders.Add(, , "Código", Width:=35, Alignment:=0).Tag = "number"
           .ColumnHeaders.Add(, , "Data Cad.", Width:=55, Alignment:=0).Tag = "date"
           .ColumnHeaders.Add(, , "Produto", Width:=120, Alignment:=0).Tag = ""
           .ColumnHeaders.Add(, , "Unidade", Width:=55, Alignment:=2).Tag = ""
           .ColumnHeaders.Add(, , "Quantidade", Width:=55, Alignment:=2).Tag = "number"
           .ColumnHeaders.Add(, , "Valor Unit.", Width:=50, Alignment:=2).Tag = "number"
           .ColumnHeaders.Add(, , "Valor Total", Width:=60, Alignment:=0).Tag = "number"
           .ColumnHeaders.Add(, , "Fornecedor", Width:=80, Alignment:=0).Tag = ""
           .ColumnHeaders.Add(, , "Endereço", Width:=100, Alignment:=0).Tag = ""
           .ColumnHeaders.Add(, , "Cidade", Width:=80, Alignment:=0).Tag = ""
           .ColumnHeaders.Add(, , "Estado", Width:=20, Alignment:=0).Tag = ""
           .ColumnHeaders.Add(, , "CEP", Width:=60, Alignment:=2).Tag = ""
           .ColumnHeaders.Add(, , "Telefone", Width:=70, Alignment:=2).Tag = ""
           .ColumnHeaders.Add(, , "Contato", Width:=80, Alignment:=0).Tag = ""
       End With
   
       Dim Linha As Long
       Dim xCont As Long
       Dim x
       Dim Lv
       
       Plan01.Activate
       xCont = Plan01.Range("A" & Rows.Count).End(xlUp).Row
       Me.ListView1.ListItems.Clear
       
       For Linha = 2 To xCont
           Set Lv = ListView1.ListItems.Add(Text:=Plan01.Cells(Linha, 1)) 'Código
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 2) ' Data
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 3) ' Produto
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 4) ' Unidade
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 5) ' Quantidade
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 6) ' Valor Unitário
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 7) ' Valor Total
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 8) ' Fornecedor
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 9) ' Endereço
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 10) 'Cidade
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 11) 'Estado
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 12) 'CEP
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 13) 'Telefone
                        Lv.ListSubItems.Add Text:=Plan01.Cells(Linha, 14) 'Contato
       Next Linha
       
       Set Lv = Nothing
       lbl_registros = xCont - 1 & " Registros listados"
   
   
   
   End Sub
   
   
   Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
      
       On Error Resume Next
          
       '    Começa ordenar o listview pela coluna clicada
       Dim vbHourglass
       
       With ListView1
       
           ' Mostrar o cursor ampulheta enquanto faz o filtro
           
           Dim lngCursor As Long
           lngCursor = .MousePointer
           .MousePointer = vbHourglass
           
           'A rotina impede que o controle ListView faça atualização na tela
           'Isto é para esconder as mudanças que estão sendo feitas aos listitems
           'E também para acelerar o código
           
           'Verifique o tipo de dados da coluna de ser classificada,
           'para nomeá-la em conformidade
           
           Dim l As Long
           Dim strFormat As String
           Dim strData() As String
           
           Dim lngIndex As Long
           lngIndex = ColumnHeader.Index - 1
       
           '***************************************************************************
           ' Ordenar por data.
           
           Select Case UCase$(ColumnHeader.Tag)
           Case "DATE"
           
               
               
               strFormat = "YYYYMMDDHhNnSs"
           
               'O Loop através dos valores desta coluna organizam
               'As datas de modo que eles possam ser classificados em ordem alfabética,
           
               With .ListItems
                   If (lngIndex > 0) Then
                       For l = 1 To .Count
                           With .Item(l).ListSubItems(lngIndex)
                               .Tag = .Text & Chr$(0) & .Tag
                               If IsDate(.Text) Then
                                   .Text = Format(CDate(.Text), _
                                                       strFormat)
                               Else
                                   .Text = ""
                               End If
                           End With
                       Next l
                   Else
                       For l = 1 To .Count
                           With .Item(l)
                               .Tag = .Text & Chr$(0) & .Tag
                               If IsDate(.Text) Then
                                   .Text = Format(CDate(.Text), _
                                                       strFormat)
                               Else
                                   .Text = ""
                               End If
                           End With
                       Next l
                   End If
               End With
               
               ' Ordenar a lista em ordem alfabética por esta coluna
               
               .SortOrder = (.SortOrder + 1) Mod 2
               .SortKey = ColumnHeader.Index - 1
               .Sorted = True
               
               ' Restaura os valores anteriores das "células" nesta
               ' Coluna da lista das tags, e também restaura
               ' as tags para os valores originais
               
               With .ListItems
                   If (lngIndex > 0) Then
                       For l = 1 To .Count
                           With .Item(l).ListSubItems(lngIndex)
                               strData = Split(.Tag, Chr$(0))
                               .Text = strData(0)
                               .Tag = strData(1)
                           End With
                       Next l
                   Else
                       For l = 1 To .Count
                           With .Item(l)
                               strData = Split(.Tag, Chr$(0))
                               .Text = strData(0)
                               .Tag = strData(1)
                           End With
                       Next l
                   End If
               End With
               
           '***************************************************************************
           'Ordenar Numericamente
           
           Case "NUMBER"
           
          
               strFormat = String(30, "0") & "." & String(30, "0")
           
               ' Loop através dos valores desta coluna. Ordena os valores de modo que eles
               ' Podem ser classificados em ordem
           
               With .ListItems
                   If (lngIndex > 0) Then
                       For l = 1 To .Count
                           With .Item(l).ListSubItems(lngIndex)
                               .Tag = .Text & Chr$(0) & .Tag
                               If IsNumeric(.Text) Then
                                   If CDbl(.Text) >= 0 Then
                                       .Text = Format(CDbl(.Text), _
                                           strFormat)
                                   Else
                                       .Text = "&" & InvNumber( _
                                           Format(0 - CDbl(.Text), _
                                           strFormat))
                                   End If
                               Else
                                   .Text = ""
                               End If
                           End With
                       Next l
                   Else
                       For l = 1 To .Count
                           With .Item(l)
                               .Tag = .Text & Chr$(0) & .Tag
                               If IsNumeric(.Text) Then
                                   If CDbl(.Text) >= 0 Then
                                       .Text = Format(CDbl(.Text), _
                                           strFormat)
                                   Else
                                       .Text = "&" & InvNumber( _
                                           Format(0 - CDbl(.Text), _
                                           strFormat))
                                   End If
                               Else
                                   .Text = ""
                               End If
                           End With
                       Next l
                   End If
               End With
               
               ' Ordenar a lista em ordem alfabética por esta coluna
               
               .SortOrder = (.SortOrder + 1) Mod 2
               .SortKey = ColumnHeader.Index - 1
               .Sorted = True
               
                         
               With .ListItems
                   If (lngIndex > 0) Then
                       For l = 1 To .Count
                           With .Item(l).ListSubItems(lngIndex)
                               strData = Split(.Tag, Chr$(0))
                               .Text = strData(0)
                               .Tag = strData(1)
                           End With
                       Next l
                   Else
                       For l = 1 To .Count
                           With .Item(l)
                               strData = Split(.Tag, Chr$(0))
                               .Text = strData(0)
                               .Tag = strData(1)
                           End With
                       Next l
                   End If
               End With
           
           Case Else   ' Assume ordenação como string
               
               
           
               .SortOrder = (.SortOrder + 1) Mod 2
               .SortKey = ColumnHeader.Index - 1
               .Sorted = True
               
           End Select
      
           .MousePointer = lngCursor
       
       End With
       
   
   
   
   End Sub
   
   
   
   
   '*****************************************************************************
   'InvNumber
   'Função usada para permitir que os números negativos possam ser classificados
   '-----------------------------------------------------------------------------
   
   Private Function InvNumber(ByVal Number As String) As String
       Static i As Integer
       For i = 1 To Len(Number)
           Select Case Mid$(Number, i, 1)
           Case "-": Mid$(Number, i, 1) = " "
           Case "0": Mid$(Number, i, 1) = "9"
           Case "1": Mid$(Number, i, 1) = "8"
           Case "2": Mid$(Number, i, 1) = "7"
           Case "3": Mid$(Number, i, 1) = "6"
           Case "4": Mid$(Number, i, 1) = "5"
           Case "5": Mid$(Number, i, 1) = "4"
           Case "6": Mid$(Number, i, 1) = "3"
           Case "7": Mid$(Number, i, 1) = "2"
           Case "8": Mid$(Number, i, 1) = "1"
           Case "9": Mid$(Number, i, 1) = "0"
           End Select
       Next
       InvNumber = Number
   End Function
   