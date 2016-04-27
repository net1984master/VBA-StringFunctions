Option Explicit
Public Function OneSymbolTrim(p_string As String, ByVal p_one_symbol As String)
 If (Len(p_one_symbol) > 1) Then
    MsgBox ("Только один символ")
    OneSymbolTrim = p_string
    Exit Function
 End If
 
 Do While Left(p_string, 1) = p_one_symbol
    p_string = Right(p_string, Len(p_string) - 1)
 Loop
 Do While Right(p_string, 1) = p_one_symbol
    p_string = Left(p_string, Len(p_string) - 1)
 Loop
 OneSymbolTrim = p_string
End Function

Public Function OneStringTrim(p_string As String, p_one_string As String)
 
 Do While Left(p_string, Len(p_one_string)) = p_one_string
    p_string = Right(p_string, Len(p_string) - Len(p_one_string))
 Loop
 Do While Right(p_string, Len(p_one_string)) = p_one_string
    p_string = Left(p_string, Len(p_string) - Len(p_one_string))
 Loop
 OneStringTrim = p_string
End Function

Public Function OneSumbolCleaner(p_string As String, p_one_symbol As String)
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


Private Function MultiTrim0(p_string As String, symbols() As String)
Dim i As Integer
Dim j As Integer
For i = LBound(symbols) To UBound(symbols)
    p_string = OneStringTrim(p_string, symbols(i))
    If i > 0 Then
    For j = 0 To i
        If Left(p_string, Len(symbols(j))) = symbols(j) Or Right(p_string, Len(symbols(j))) = symbols(j) Then
            p_string = MultiTrim0(p_string, symbols())
        End If
    Next
    End If
Next i

MultiTrim0 = p_string
End Function
Public Function MultiTrim(p_string As String, ParamArray symbols() As Variant)
Dim i As Integer
Dim arr() As String
ReDim arr(UBound(symbols))
For i = LBound(symbols) To UBound(symbols)
    arr(i) = symbols(i)
Next i

MultiTrim = MultiTrim0(p_string, arr)
End Function





Public Sub sb()
Dim ok As String
    ok = MultiTrim("+=+//+//+==+=+=/+=+o!o!ooo++=рк/++oo+1=+/=+", "oo", "рк", "+", "=", "/", 1)
    MsgBox (ok)
End Sub
