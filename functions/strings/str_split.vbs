<%
'=======================================================================
' 文字列を配列に変換する
'=======================================================================
'【引数】
'  string       = string 入力文字列。
'  split_length = string 分割した部分の最大長。
'【戻り値】
'  オプションのパラメータ split_length  が指定されている場合、 返される配列の各要素は、split_length  の長さとなります。それ以外の場合、1 文字ずつ分割された配列となります。
'  split_length が 1 より小さい場合に FALSE を返します。
'  split_length が string の長さより大きい場合、文字列全体が 最初の(そして唯一の)要素となる配列を返します。 
'【処理】
'  ・文字列を配列に変換します。
'=======================================================================
Function str_split(string, split_length)

    str_split = false
    If len(string) = 0 Then Exit Function
    If len(split_length) = 0 Then split_length = 1
    If split_length < 1 Then Exit Function

    Dim counter,i,pointer
    counter = len(string)
    counter = counter / split_length + 0.9999
    counter = int(counter) -1

    ReDim tmp_ar(counter)

    For i = 0 to counter
        pointer = i * split_length + 1
        tmp_ar(i) = Mid(string,pointer,split_length)
    Next

    str_split = tmp_ar

End Function
%>
