<%
'=======================================================================
' 文字列の中から任意の文字を探す
'=======================================================================
'【引数】
'  haystack     =   string  char_list  を探す文字列。
'  char_list    =   string  このパラメータは大文字小文字を区別します。
'【戻り値】
'  見つかった文字から始まる文字列、あるいは見つからなかった場合に FALSE を返します。
'【処理】
'  ・ strpbrk() は、文字列 haystack  から char_list  を探します。
'=======================================================================
Function strpbrk( haystack, char_list )

    haystack  = Cstr( haystack )
    char_list = Cstr( char_list )

    Dim intlen : intlen = len( haystack )
    Dim i,char
    For i = 1 to intlen
        char = Mid(haystack,i,1)
        If [!==](strpos(char_list,char,""),false) Then
            strpbrk = Mid(haystack,i)
            Exit Function
        End If
    Next

    strpbrk = false

End Function
%>
