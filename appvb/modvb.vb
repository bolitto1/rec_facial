Module modvb
    Public Function isnumber(num As String) As Boolean
        Dim Res1 As Boolean
        If IsNumeric(num) Then
            Res1 = True
        Else
            Res1 = False

        End If
        Return Res1

    End Function
End Module
