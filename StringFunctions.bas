Attribute VB_Name = "Module1"
Function OnSumbolTrim(p_string As String, p_one_symbol As String)
 If (Len(p_one_symbol) > 1) Then
    MsgBox ("Только один символ")
    OnSumbolTrim = p_string
    Exit Function
 End If
 
 Do While Left(p_string, 1) = p_one_symbol
    p_string = Right(p_string, Len(p_string) - 1)
 Loop
 Do While Right(p_string, 1) = p_one_symbol
    p_string = Left(p_string, Len(p_string) - 1)
 Loop
 OnSumbolTrim = p_string
End Function

Function MultiSumbolTrim(p_string As String, p_one_symbol As String)
 
 Do While Left(p_string, Len(p_one_symbol)) = p_one_symbol
    p_string = Right(p_string, Len(p_string) - Len(p_one_symbol))
 Loop
 Do While Right(p_string, Len(p_one_symbol)) = p_one_symbol
    p_string = Left(p_string, Len(p_string) - Len(p_one_symbol))
 Loop
 MultiSumbolTrim = p_string
End Function

Function OneSumbolCleaner(p_string As String, p_one_symbol As String)
Dim v_two_symbols As String
 If (Len(p_one_symbol) > 1) Then
    MsgBox ("Только один символ")
    OneSumbolCleaner = p_string
    Exit Function
 End If
 v_two_symbols = p_one_symbol + p_one_symbol
 Do While InStr(1, p_string, v_two_symbols)
 p_string = Replace(p_string, v_two_symbols, p_one_symbol)
 Loop
 OneSumbolCleaner = p_string
End Function

Sub sb()
    ok = OneSumbolCleaner("+'+++'+f++++++++gf+++df++g+++++++*'+'", "+")
    MsgBox (ok)
End Sub
