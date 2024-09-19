Private Sub CommandButton1_Click()

  Dim Item As Integer
  Dim soma As Double
    
  For Item = 0 To ListBox1.ListCount - 1 'para item = 0 até a quantidade de itens da minha listbox -1 pois começa em zero
                                         'irá fazer o teste para ver se o item está selecionado
                                            
        If ListBox1.Selected(Item) = True Then
        soma = soma + ListBox1.List(Item, 4)
        End If
        
   Next

    Me.TextBox2 = FormatNumber(soma, 2) 'FormatNumber para ficar no formato 1.000,00

End Sub

Private Sub CommandButton2_Click()

    Me.TextBox2 = ""

    For Item = 0 To ListBox1.ListCount - 1 'para item = 0 até a quantidade de itens da minha listbox -1 pois começa em zero
                                        'irá fazer o teste para ver se o item está selecionado
            
            If ListBox1.Selected(Item) = True Then
                ListBox1.Selected(Item) = False    'desmarca os que estiverem marcados
            End If
    Next Item

End Sub

Private Sub CommandButton3_Click()

    'faz o teste para ver se há mais de um item selecionado
    Dim Item As Integer
    Dim contador As Integer
    contador = 0
        
    For Item = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(Item) = True Then
            contador = contador + 1
            End If
    Next Item

    If contador > 1 Then
        MsgBox "Selecione somente um item para editar", vbExclamation
        Exit Sub
    ElseIf contador = 0 Then
        MsgBox "Selecione um item para editar", vbExclamation
    Else
        UserForm2.Show 'depois de fechado  vai continuar as ações abaixo

        Call UserForm_Initialize
        MsgBox "Dados Atualizados"
    End If

End Sub

Private Sub btn_transfere_tudo_Click()

   Sheets("Relatorio").Select

    'testa se a list esta vazia
    If ListBox1.ListCount = 0 Then
        MsgBox ("Não há itens a serem impressos..."), vbInformation, ("Erro")
    Else
        'limpa dados antes de lançar os novos dados
        If Range("A4").Select = "" Then
            'não faz nanda
        Else
            'apaga o intervalo
            Range("A4").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            Range("A4").Select
        End If

        'lança os dados na planilha
    
        Dim Item As Double
        Dim linha As Integer
        Dim valor_unitario As Double
        Dim valor_total As Double

        linha = 4
    
        For Item = 0 To ListBox1.ListCount - 1
                                        
            valor_unitario = ListBox1.List(Item, 3)
            valor_total = ListBox1.List(Item, 4)
                                        
            Sheets("Relatorio").Cells(linha, 1) = ListBox1.List(Item, 0)
            Sheets("Relatorio").Cells(linha, 2) = ListBox1.List(Item, 1)
            Sheets("Relatorio").Cells(linha, 3) = ListBox1.List(Item, 2)
            Sheets("Relatorio").Cells(linha, 4) = valor_unitario
            Sheets("Relatorio").Cells(linha, 5) = valor_total
            
            linha = linha + 1
        Next Item
    End If

End Sub


Private Sub btn_transfere_selecao_Click()

    Sheets("Relatorio").Select

    'testa se a list esta vazia
    If ListBox1.ListCount = 0 Then
        MsgBox ("Não há itens a serem impressos..."), vbInformation, ("Erro")
    Else
      'limpa dados antes de lançar os novos dados
        If Range("A4").Select = "" Then
            'não faz nanda
        Else
            'apaga o intervalo
            Range("A4").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            Range("A4").Select
        End If

        'lança os dados na planilha
    
        Dim Item As Double
        Dim linha As Integer
        Dim valor_unitario As Double
        Dim valor_total As Double
        
        linha = 4
            
        For Item = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(Item) = True Then
        
                valor_unitario = ListBox1.List(Item, 3)
                valor_total = ListBox1.List(Item, 4)
                                            
                Sheets("Relatorio").Cells(linha, 1) = ListBox1.List(Item, 0)
                Sheets("Relatorio").Cells(linha, 2) = ListBox1.List(Item, 1)
                Sheets("Relatorio").Cells(linha, 3) = ListBox1.List(Item, 2)
                Sheets("Relatorio").Cells(linha, 4) = valor_unitario
                Sheets("Relatorio").Cells(linha, 5) = valor_total
                        
                linha = linha + 1
            End If
        Next Item
    End If
End Sub

Private Sub CommandButton4_Click()

    For Item = 0 To ListBox1.ListCount - 1 'para item = 0 até a quantidade de itens da minha listbox -1 pois começa em zero
                                           'irá fazer o teste para ver se o item está selecionado
        
        If ListBox1.Selected(Item) = False Then
           ListBox1.Selected(Item) = True    'seleciona os que estiverem desmarcados
        End If
    Next Item


End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub TextBox3_Change()

    valor_pesquisa = TextBox3.Text

    Dim guia As Worksheet
    Dim linha As Integer
    Dim coluna As Integer
    Dim linhalistbox As Integer
    Dim valor_celula As String
    Dim soma As Double
    
    Set guia = ThisWorkbook.Worksheets("Produtos")
    
    linhalistbox = 0
    linha = 2 'linha de inicio dos dados
    coluna = 2 'coluna referente busca de produtos
    
    ListBox1.Clear
    
    guia.Select
     
    With guia
        While .Cells(linha, coluna).Value <> Empty 'enquanto for diferente de vazio faça
            
            valor_celula = .Cells(linha, coluna).Value 'recebe o valor da célula para fazer o teste
            
            'Condição para satisfazer a busca tem que ser igual ao valor da texbox3
            If UCase(Left(valor_celula, Len(valor_pesquisa))) = UCase(valor_pesquisa) Then
            
                'adiciona itens a listbox
                With ListBox1
                .AddItem
                .List(linhalistbox, 0) = guia.Cells(linha, 1)
                .List(linhalistbox, 1) = guia.Cells(linha, 2)
                .List(linhalistbox, 2) = guia.Cells(linha, 3)
                .List(linhalistbox, 3) = FormatNumber(guia.Cells(linha, 4), 2)
                .List(linhalistbox, 4) = FormatNumber(ListBox1.List(linhalistbox, 2) * ListBox1.List(linhalistbox, 3), 2)
                'Na coluna 5 faço a multiplicação do valor que está na coluna 3 pelo valor da coluna 4
                
                soma = soma + ListBox1.List(linhalistbox, 4)
                'Acumulador declarado como double para acumular a soma dos valores totais
                
                End With
                linhalistbox = linhalistbox + 1
            End If
            
            linha = linha + 1
        Wend
    End With
    Me.TextBox1 = FormatNumber(soma, 2)

End Sub

Private Sub UserForm_Initialize()

    '''carrega dados

    Dim soma As Double

    ListBox1.Clear
    linha = 2
    linhalistbox = 0

    Sheets("Produtos").Select

    Do Until Sheets("Produtos").Cells(linha, 1) = ""

        With ListBox1
            .AddItem
            .List(linhalistbox, 0) = Sheets("Produtos").Cells(linha, 1)
            .List(linhalistbox, 1) = Sheets("Produtos").Cells(linha, 2)
            .List(linhalistbox, 2) = Sheets("Produtos").Cells(linha, 3)
            .List(linhalistbox, 3) = FormatNumber(Sheets("Produtos").Cells(linha, 4), 2)
            .List(linhalistbox, 4) = FormatNumber(ListBox1.List(linhalistbox, 2) * ListBox1.List(linhalistbox, 3), 2)
            'Na coluna 5 faço a multiplicação do valor que está na coluna 3 pelo valor da coluna 4
            
            soma = soma + ListBox1.List(linhalistbox, 4)
            'Acumulador declarado como double para acumular a soma dos valores totais
        End With
        linhalistbox = linhalistbox + 1
        linha = linha + 1
    Loop
    Me.TextBox1 = FormatNumber(soma, 2) 'FormatNumber para ficar no formato 1.000,00

End Sub