Attribute VB_Name = "modRegisteredFunctions"
Option Explicit

Public Sub ListRegisteredFunctionNames()
    Dim registered As Variant
    Dim results() As Variant
    Dim regId As Variant, functionName As Variant
    Dim r As Long, c As Long, outputRow As Long
    Dim firstCol As Long, count As Long
    Dim expr As String, errorText As String
    Dim ws As Worksheet

    registered = Application.RegisteredFunctions

    If IsNull(registered) Then
        MsgBox "No registered functions found.", vbInformation
        Exit Sub
    End If

    If Not IsArray(registered) Then
        MsgBox "RegisteredFunctions did not return an array.", vbExclamation
        Exit Sub
    End If

    count = UBound(registered, 1) - LBound(registered, 1) + 1
    firstCol = LBound(registered, 2)
    ReDim results(1 To count + 1, 1 To 6)

    results(1, 1) = "DLL / XLL"
    results(1, 2) = "Export / stub"
    results(1, 3) = "Type string"
    results(1, 4) = "Registration ID"
    results(1, 5) = "Worksheet name"
    results(1, 6) = "Status"

    outputRow = 1

    For r = LBound(registered, 1) To UBound(registered, 1)
        outputRow = outputRow + 1

        For c = 0 To 2
            results(outputRow, c + 1) = registered(r, firstCol + c)
        Next c

        ' Omit type_text: query an existing registration.
        expr = "REGISTER.ID(" & _
               XlmArgument(registered(r, firstCol)) & "," & _
               XlmArgument(registered(r, firstCol + 1)) & ")"

        If TryExcel4(expr, regId, errorText) Then
            If IsNumeric(regId) Then
                results(outputRow, 4) = regId

                ' Numeric definition; omitted document; include hidden names.
                ' Str$ uses "." as the decimal separator.
                expr = "GET.DEF(" & Trim$(Str$(CDbl(regId))) & ",,3)"

                If TryExcel4(expr, functionName, errorText) Then
                    results(outputRow, 5) = functionName
                    results(outputRow, 6) = "Resolved (first matching name)"
                Else
                    results(outputRow, 6) = "GET.DEF: " & errorText
                End If
            Else
                results(outputRow, 6) = "REGISTER.ID returned a non-numeric value"
            End If
        Else
            results(outputRow, 6) = "REGISTER.ID: " & errorText
        End If
    Next r

    Set ws = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    ws.Name = "Registered functions"

    With ws
        ' Preserve names and type strings as literal text.
        .Columns("A:C").NumberFormat = "@"
        .Columns("E:F").NumberFormat = "@"
        .Range("A1").Resize(count + 1, 6).Value2 = results
        .Rows(1).Font.Bold = True
        .Range("A1").Resize(count + 1, 6).AutoFilter
        .Columns("A:F").AutoFit
        .Columns("A").ColumnWidth = 65
    End With
End Sub

Private Function XlmArgument(ByVal value As Variant) As String
    If VarType(value) = vbString Then
        XlmArgument = """" & Replace$(CStr(value), """", """""") & """"
    Else
        ' Preserve numeric export ordinals as numbers.
        XlmArgument = Trim$(Str$(CDbl(value)))
    End If
End Function

Private Function TryExcel4(ByVal expression As String, _
                           ByRef result As Variant, _
                           ByRef errorText As String) As Boolean
    On Error GoTo Failed

    result = Empty
    errorText = vbNullString
    result = Application.ExecuteExcel4Macro(expression)

    If IsError(result) Then
        errorText = CStr(result)
    Else
        TryExcel4 = True
    End If
    Exit Function

Failed:
    errorText = CStr(Err.Number) & ": " & Err.Description
End Function

