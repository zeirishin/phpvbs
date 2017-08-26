<%
'=======================================================================
' 文字列に rot13 変換を行う
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'【戻り値】
'  指定した文字列を ROT13 変換した結果を返します。
'【処理】
'  ・ROT13 は、各文字をアルファベット順に 13 文字シフトさせ、 アルファベット以外の文字はそのままとするエンコードを行います。
'  エンコードとデコードは同じ関数で行われます。
'  引数にエンコードされた文字列を指定した場合には、元の文字列が返されます。
'=======================================================================
Function str_rot13(str)

    Dim str_rotated : str_rotated = ""
    Dim i,j,k

    For i = 1 to Len(str)
        j = Mid(str, i, 1)
        k = Asc(j)
        if k >= 97 and k =< 109 then
            k = k + 13 ' a ... m
        elseif k >= 110 and k =< 122 then
            k = k - 13 ' n ... z
        elseif k >= 65 and k =< 77 then
            k = k + 13 ' A ... M
        elseif k >= 78 and k =< 90 then
            k = k - 13 ' N ... Z
        end if

        str_rotated = str_rotated & Chr(k)
    Next

    str_rot13 = str_rotated

End Function
%>
