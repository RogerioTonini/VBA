
Private Sub CommandButton1_Click()

    Dim linha As Integer
    Dim id As Integer
    Dim quantidade As Double
    Dim valor As Double

    On Error Resume Next 'para não gerar erro caso o registro não exista

    id = Me.TextBox1
    quantidade = Me.TextBox3
    valor = Me.TextBox4

    'Localiza um registro pelo método find
    linha = Sheets("Produtos").Range("A:A").Find(id).Row

    Sheets("Produtos").Cells(linha, 2) = Me.TextBox2  'coluna produto 2ª coluna
    Sheets("Produtos").Cells(linha, 3) = quantidade 'coluna quantidade 3ª coluna
    Sheets("Produtos").Cells(linha, 4) = valor 'coluna valor unitario 4ª coluna

    Unload Me
End Sub

Private Sub UserForm_Initialize()
 
    Dim Item As Integer
    
    For Item = 0 To UserForm1.ListBox1.ListCount - 1
        If UserForm1.ListBox1.Selected(Item) = True Then
            TextBox1 = UserForm1.ListBox1.List(Item, 0)
            TextBox2 = UserForm1.ListBox1.List(Item, 1)
            TextBox3 = UserForm1.ListBox1.List(Item, 2)
            TextBox4 = UserForm1.ListBox1.List(Item, 3)
        End If
    Next

End Sub
