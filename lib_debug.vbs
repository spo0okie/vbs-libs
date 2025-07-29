Option Explicit
Function gettype(s)

    Select Case VarType(s)
    Case 0
        gettype = "vbEnpty"
    Case 1
        gettype = "vbNull"
    Case 2
        gettype = "vbInteger"
    Case 3
        gettype = "vbLong"
    Case 4
        gettype = "vbSingle"
    Case 5
        gettype = "vbDouble"
    Case 6
        gettype = "vbCurrency"
    Case 7
        gettype = "vbDate"
    Case 8
        gettype = "vbString"
    Case 9
        gettype = "vbObject"
    Case 10
        gettype = "vbError"
    Case 11
        gettype = "vbBoolean"
    Case 12
        gettype = "vbVariant"
    Case 13
        gettype = "vbDataObject"
    Case 17
        gettype = "vbByte"
    Case 8192
        gettype = "vbArray"
    Case 8204
        gettype = "vbArray"
    Case 8209
        gettype = "vbBinary"
    End Select

End Function

Sub var_dump(expression)
    var_dump_helper expression,0
End Sub

Sub var_dump_helper(expression,tab)

    If VarType(tab) <> 2 Then tab = 0

    Dim strTab : strTab = String(tab,vbTab)

    If IsObject(expression) Then
        msg_ "Dictionary Object(" & expression.count & ")" & vbCrLf
    ElseIf IsArray(expression) Then
        msg_ "Array(" & (uBound(expression)+1) & ")" & vbCrLf
    End If

	msg_ strTab & "(" & vbCrLf

    Dim a,i
    i = 0
    If IsObject(expression) Then
        For Each a In expression
            msg_ strTab
            If IsArray(a) or IsObject(a) Then
                msg_ vbTab & "[] => "
                call var_dump_helper(a,tab + 1)
            ElseIf isArray(expression(a)) or isObject( expression(a) ) Then
                msg_ vbTab & "[" & a & "] => "
                call var_dump_helper(expression(a),tab + 1)

            Else
               msg_ vbTab & "[" & a & "]" & " => " & _
                              gettype(expression(a)) & "(" & expression(a) & ")" & vbCrLf
            End If
        Next
    ElseIf IsArray(expression) Then
        For Each a In expression
            msg_ strTab
            If IsArray(a) or IsObject(a) Then
                msg_ vbTab & "[" & i & "] => "
                call var_dump_helper(a,tab + 1)
            Else
                msg_ vbTab & "[" & i & "] => " & _
                               gettype(a) & "(" & a & ")" & vbCrLf
            End If

            i =  i+1
        Next
    Else
        msg_ strTab & gettype(expression) & "(" & expression & ")" & vbCrLf
    End If

    msg_ strTab & ")" & vbCrLf

End Sub
