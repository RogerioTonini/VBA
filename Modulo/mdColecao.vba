Option Explicit

Sub ColeçãoDeObjetos()
    Dim tblEmpregados As ListObject
    Dim oEmpregado As clsEmpregado
    Dim cEmpregados As Collection
    Dim iListRow As ListRow
    Dim iItem As Long

    'Definir tabela de Empregados
    Set tblEmpregados = ThisWorkbook.Worksheets("Coleções").ListObjects("Empregados")
    
    'Inicializar Coleção?
    Set cEmpregados = New Collection
    
    'Criar e popular objetos
    For Each iListRow In tblEmpregados.ListRows
        Set oEmpregado = New clsEmpregado
        oEmpregado.Nome = iListRow.Range(1)
        oEmpregado.Endereço = iListRow.Range(2)
        oEmpregado.Salário = iListRow.Range(3)
        
        'Adiciona objeto à coleção:
        cEmpregados.Add oEmpregado
    Next iListRow
    
    'Mostrar resultados
    For Each oEmpregado In cEmpregados
        Debug.Print oEmpregado.Nome, oEmpregado.Endereço, oEmpregado.Salário
    Next oEmpregado
    
    'Ou, com laço do tipo For...Next
    For iItem = 1 To cEmpregados.Count
        Set oEmpregado = cEmpregados(iItem)
        Debug.Print oEmpregado.Nome, oEmpregado.Endereço, oEmpregado.Salário
    Next iItem
        
End Sub
