<%
'=======================================================================
' 特定の文字を変換する
'=======================================================================
'【引数】
'  str      = string    変換する文字列。
'  from     = string    strTo  に変換される文字列。
'  strTo    = int       from  を置換する文字列。
'【戻り値】
'  この関数は str  を走査し、 from  に含まれる文字が見つかると、そのすべてを strTo  の中で対応する文字に置き換え、 その結果を返します。
'【処理】
'  ・ この関数は str  を走査し、 from  に含まれる文字が見つかると、そのすべてを to  の中で対応する文字に置き換え、 その結果を返します。
'  ・ from と to の長さが異なる場合、長い方の余分な文字は無視されます。 
'=======================================================================
Function strtr(ByVal str, from, strTo)

    If isObject(from) Then
        Dim key
        For Each key In from
            str = Replace(str,key,from(key))
        Next

    Else

        Dim len1 : len1 = len(from)
        Dim len2 : len2 = len(strTo)

        If len1 > len2 Then
            from = Left(from,len2)
        ElseIf len2 > len1 Then
            strTo = Left(strTo,len1)
        End If

        str = Replace(str,from,strTo)
    End if

    strtr = str
End Function
%>
