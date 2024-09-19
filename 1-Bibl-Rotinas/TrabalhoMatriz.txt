 Popular planilha com informações de vetores

Sub FillWithArray()
	Dim i 						 as Double
	Dim j 						 as Double
	Dim arr(1 To 1000, 1 To 256) as Double
	
	begin = Timer
	Application.EnableEvents = False
	ActiveSheet.Cells.Clear
	
	For i = 1 To 1000
		For j = 1 To 256
			arr(i, j) = i * j
		Next j
	Next i
	Range("A1:IV1000"). Value = arr

	Application.EnableEvents = True
	Finish = Timer
	Debug.Print Format(Finish - Begin, "0.0000000000")
End Sub

---------------------

- Capturar conteudo de uma tabela para uma Matriz

Sub Teste()

	Dim wb   As Excel.Workbook
	Dim ws   As Excel.Worksheet
	Dim tb   As Excel.ListObject
	Dim arr  As Variant

	Set wb = ThisWorkbook
	Set ws = wb.Sheets("Hiperlinks")
	Set tb = ws.ListObjects("Tabela")

	arr = tb.Range
	
	' Carrega os valores da Planilha/Tabela para o vetores
	arr = Range("[Intervalo Nomeado]").Value
	
	' Apaga a matriz da memória RAM
	Erase arr

	Set wb = Nothing
	Set ws = Nothing
	Set tb = Nothing

End Sub


Ao criar uma tabela, podemos criar um intervalo dinâmico através do recurso de criar intervalos nomeados
e utilizando DESLOC + CONT.VALORES =(DESLOC(pCrtCredito!$I$3;0;0;CONT.VALORES(pCrtCredito!$I:$I)-1;1))

Sub Teste()

	Dim arr  		As Variant
	Dim lngLinIni 	As Long
	Dim lngLinFim 	As Long
	Dim lngColIni 	As Long
	Dim lngColFim 	As Long
	Dim lngContLin	As Long
	Dim lngContCol	As Long
	
	' (1) Carrega os valores da Planilha/Tabela para o vetores
	arr = Range("[Intervalo Nomeado]").Value

	lngLinIni = LBound(arr, 1)	' Captura a primeira linha da matriz
	lngLinFim = UBound(arr, 1)	' Captura a última linha da matriz
	lngColIni = LBound(arr, 2)	' Captura a primeira posição da direita da matriz
	lngColFim = UBound(arr, 2)	' Captura a última posição da direita da matriz
	
	' Realizar operações diversas na matriz
	' Neste caso ao invés de utilizar as variáveis para utilização do For...Next,
	' será usado diretamente as funções UBound / LBound
	For lngContLin = LBound(arr, 1) To UBound(arr, 1)
	
		For lngContCol = LBound(arr, 2) To UBound(arr, 2)
			arr(lngContLin, lngContCol) = arr(lngContLin, lngContCol) * 3
		Next lngContCol
		
	Next lngContLin
	
	' Descarregar novamente na Planilha / Tabela
	Range("A1", Cells(UBound(arr, 1), UBound(arr, 2))).Value = arr
	
	' Descarregar novamente na Planilha / Tabela - (Modo Raiz) 
	Range(Cells(LBound(arr, 1), LBound(arr, 2)), Cells(UBound(arr, 1), UBound(arr, 2))).Value = arr
	
	' Caso seja necessário o redimensionamento da matriz, poderá ser feito durante o processo, porém,
	' só poderá ser feito sem a clausula PRESERVE, ou seja, apagando todo o conteúdo da matriz e 
	' inicializando ela novamente.
	
	Redim arr(1 To 50, 1 To 10)
	
	' ** Para realizar este procedimento deve-se utilizar o comando Application.Transpose
	' Este comando transpõe as colunas para linhas e as linhas para coluna'
	arr = Application.Transpose
	
	' Redimensiona a qtde de linhas da matriz, pois, através da transposição (troca de lugar entre linhas e colunas)
	' para poder aumentar a qtde de linhas
	Redim Preserve arr(1 To 10, 1 To UBound(arr, 2) * 2)
	
	' Novamente faz a transpoisição para regularizar a posição de linha e coluna.
	arr = Application.Transpose
	
	' Apaga a matriz da memória RAM
	Erase arr

	Set wb = Nothing
	Set ws = Nothing
	Set tb = Nothing

End Sub
