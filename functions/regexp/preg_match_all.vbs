<%
'=======================================================================
'繰り返し正規表現検索を行う
'=======================================================================
'【引数】
'  pattern  = string    検索するパターンを表す文字列。
'  subject  = string    入力文字列。
'  matches  = array     matches  を指定した場合、検索結果が代入されます。matches(0) にはパターン全体にマッチしたテキストが代入され、 matches(1)には 1 番目ののキャプチャ用サブパターンにマッチした 文字列が代入され、といったようになります。
'  flags    = int       戻り値の形式を指定
'  offset   = int       通常、検索は対象文字列の先頭から開始されます。 オプションのパラメータ offset  を使用して 検索の開始位置を指定することも可能です。
'【戻り値】
'  パターンがマッチした総数を返します（ゼロとなる可能性もあります）。 
'  または、エラーが発生した場合に FALSE を返します。
'【処理】
'  ・ subject  を検索し、 pattern  に指定した正規表現にマッチした すべての文字列を、flags  で指定した 順番で、matches  に代入します。
'  ・ 正規表現にマッチすると、そのマッチした文字列の後から 検索が続行されます。 
'=======================================================================
Function preg_match_all(pattern, ByVal subject, ByRef matches, flags, offset)

    Dim regEx,matchall,matchone
    Dim cnt,counter : counter = 0
    Dim helper

    preg_match_all = false
    If vartype(matches) <> 0 Then Exit Function
    If len(flags) = 0 Then flags = PREG_PATTERN_ORDER

    set helper = new RegExp_Helper
    helper.parseOption(pattern)

    Set regEx = new RegExp
    With regEx
        .IgnoreCase = helper.withIgnoreCase
        .Global     = True
        .pattern    = helper.withPattern
        .MultiLine  = helper.withMultiLine
    End With
    set helper = Nothing

    If len(offset) > 0 Then
        offset = int(offset)
        subject = Mid(subject,offset)
    End If

    Set matchall = regEx.execute(subject)
    Set regEx = Nothing
    If matchall.count = 0 Then exit Function

    If flags = PREG_PATTERN_ORDER Then

        ReDim matches(matchall(0).SubMatches.count)

        For cnt = 0 to uBound(matches)
            toReDim matches(cnt),(matchall.count-1)
        Next

        counter = 0
        For Each matchone In matchall
            matches(0)(counter) = matchone.value
            For cnt = 1 to matchone.SubMatches.count
                matches(cnt)(counter) = matchone.SubMatches(cnt-1)
            Next
            counter = counter + 1
        Next

    Elseif flags = PREG_SET_ORDER Then

        ReDim matches(matchall.count-1)

        counter = 0
        For Each matchone In matchall
            toReDim matches(counter),(matchone.SubMatches.count)
            matches(counter)(0) = matchone.value
            For cnt = 1 to matchone.SubMatches.count
                matches(counter)(cnt) = matchone.SubMatches(cnt-1)
            Next
            counter = counter + 1
        Next

    ElseIf PREG_OFFSET_CAPTURE Then

        ReDim matches(matchall(0).SubMatches.count)

        For cnt = 0 to uBound(matches)
            toReDim matches(cnt),(matchall.count-1)
        Next

        counter = 0
        For Each matchone In matchall
            toReDim matches(0)(counter),1
            matches(0)(counter)(0) = matchone.value
            matches(0)(counter)(1) = matchone.FirstIndex
            For cnt = 1 to matchone.SubMatches.count
                toReDim matches(cnt)(counter),1
                matches(cnt)(counter)(0) = matchone.SubMatches(cnt-1)
                matches(cnt)(counter)(1) = InStr( matchone.value, matchone.SubMatches(cnt-1) ) -1
            Next
            counter = counter + 1
        Next

    End If

    preg_match_all = matchall.Count

End Function
%>
