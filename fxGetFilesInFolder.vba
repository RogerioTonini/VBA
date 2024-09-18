Function fxGetFilesInFolder() As Boolean
    
    Dim fso     As Object  ' Cria um objeto FileSystemObject
    Dim file    As Object  ' Cria um objeto File
    Dim folder  As Object  ' Cria um objeto Folder
    Dim intCont As Integer ' Contador
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(strCamArquivos)
    
    For intCont = 1 To intQtdeArquivos
       For Each file In folder.Files
          If arrListFiles(intCont, 1) = file.Name Then
             arrListFiles(intCont, 2) = True
             Exit For
          End If
       Next file
    Next intCont
         
    For intCont = 1 To intQtdeArquivos
       If arrListFiles(intCont, 2) = False Then
          fxGetFilesInFolder = False
          Exit For
       Else
          fxGetFilesInFolder = True
       End If
    Next intCont
    
 End Function