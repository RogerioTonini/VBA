' Autor: danieltakeshi - Hobbyst Developer, interested in learning and knowledge
' link.: https://pt.stackoverflow.com/questions/381623/passar-array-como-par%C3%A2metro-em-vba-excel
' Data.: 06/05/2019

' Via Função

Public Type pessoa
    Nome As String
    Idade As Long
End Type

Public Function SomaIdade(ByRef arrPessoa() As pessoa)
    Dim soma As Long
    For i = 0 To UBound(arrPessoa)
        soma = soma + arrPessoa(i).Idade
    Next i
    SomaIdade = soma
End Function

Sub main()
    Dim p() As pessoa
    Dim soma_idade As Long
    ReDim p(0 To 2)
    p(0).Idade = 10
    p(1).Idade = 20
    p(2).Idade = 30
    soma_idade = SomaIdade(p())
    Debug.Print soma_idade
End Sub

' Via Sub

Public Type pessoa
    Nome As String
    Idade As Long
End Type

Public Sub SomaIdade(ByRef arrPessoa() As pessoa)
    Dim soma As Long
    For i = 0 To UBound(arrPessoa)
        soma = soma + arrPessoa(i).Idade
    Next i
    Debug.Print (soma)
End Sub

Sub main()
    Dim p() As pessoa
    ReDim p(0 To 2)
    p(0).Idade = 10
    p(1).Idade = 20
    p(2).Idade = 30
    SomaIdade p()
End Sub
