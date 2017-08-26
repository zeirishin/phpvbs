<%
'=======================================================================
' 文字列中に文字が最後に現れる場所を取得する
'=======================================================================
'【引数】
'  haystack = string    検索を行う文字列。
'  needle   = string    needle  がひとつ以上の文字を含んでいる場合は、 最初のもののみが使われます。この動作は、 strstr()  とは異なります。
'【戻り値】
'  この関数は、部分文字列を返します。 needle  が見つからない場合は FALSE を返します。
'【処理】
'  ・ この関数は、文字列 haystack  の中で needle  が最後に現れた位置から、 haystack  の終わりまでを返します。
'=======================================================================
Function strrchr( haystack, needle )

    haystack = Cstr( haystack )
    needle   = Cstr( needle )
    If len(needle) > 1 Then needle = Left(needle,1)

    strrchr = false

    Dim i
    i = strrpos(haystack, needle,"")

    If i > 0 Then
        strrchr = Mid(haystack,i)
    End If

End Function
%>
