Sub ClearComboBox()

    Dim xOle As OLEObject
    Dim xDrop As DropDown
    
    For Each xOle In ActiveSheet.OLEObjects
        If TypeName(xOle.Object) = "ComboBox" Then
            xOle.ListFillRange = ""
        End If
    Next
    
    For Each xDrop In ActiveSheet.DropDowns
        xDrop.ListFillRange = ""
    Next
    Application.ScreenUpdating = True

End Sub