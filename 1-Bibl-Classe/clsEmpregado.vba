Option Explicit

Private aEndereço As String
Private aDólares As Double

'Eventos
Private Sub Class_Initialize()
    Me.Endereço = "Rua das Flores, 105"
End Sub
Private Sub Class_Terminate()
    MsgBox "Um objeto cuja propriedade Endereço é " & Me.Endereço & " foi destruído.", vbInformation
End Sub

'Propriedades
Property Get Endereço() As String
    Endereço = aEndereço
End Property
Property Let Endereço(pEndereço As String)
    aEndereço = pEndereço
End Property

Property Get Dólares() As Double
    Dólares = aDólares
End Property
Property Let Dólares(pDólares As Double)
    aDólares = pDólares
End Property
