<%
'=======================================================================
' 文字列を反復する
'=======================================================================
'【引数】
'  input        = string    繰り返す文字列。
'  multiplier   = int       input を繰り返す回数。multiplier は 0 以上でなければなりません。 multiplier が 0 に設定された場合、この関数は空文字を返します。
'【戻り値】
'  繰り返した文字列を返します。
'【処理】
'  ・input  を multiplier  回を繰り返した文字列を返します。
'=======================================================================
Function str_repeat(input, multiplier)
    If multiplier < 0 Then Exit Function
    str_repeat = String(multiplier,input)
End Function
%>
