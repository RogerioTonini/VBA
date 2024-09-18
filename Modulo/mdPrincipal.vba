Option Explicit

Sub GerarPropriedade()
    Dim sAtributo As String
    Dim sInstrução As String
    Dim sParâmetro As String
    Dim sPropriedade As String
    Dim sSaída As String
    Dim sTipoDeDados As String
    Dim vInstruções() As String
    
    sInstrução = InputBox(Prompt:="Digite a linha de declaração do atributo:", _
                          Default:="Private aNome As String")
    If sInstrução = "" Then Exit Sub
    
    vInstruções = Split(sInstrução)
    sAtributo = vInstruções(1)
    sPropriedade = Mid(sAtributo, 2)
    sParâmetro = "p" & sPropriedade
    sTipoDeDados = vInstruções(3)
    
    sSaída = ""
    sSaída = sSaída & "Property Get " & sPropriedade & "() As " & sTipoDeDados & vbNewLine
    sSaída = sSaída & vbTab & sPropriedade & " = " & sAtributo & vbNewLine
    sSaída = sSaída & "End Property" & vbNewLine & vbNewLine
    sSaída = sSaída & "Property Let " & sPropriedade & "(" & sParâmetro & " As " & sTipoDeDados & ")" & vbNewLine
    sSaída = sSaída & vbTab & sAtributo & " = " & sParâmetro & vbNewLine
    sSaída = sSaída & "End Property" & vbNewLine
    Debug.Print sSaída
End Sub