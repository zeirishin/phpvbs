<%
'=======================================================================
'正規表現によるマッチングを行う
'=======================================================================
'【引数】
'  pattern  = string    検索するパターンを表す文字列。
'  subject  = string    入力文字列。
'  matches  = array     matches  を指定した場合、検索結果が代入されます。matches(0) にはパターン全体にマッチしたテキストが代入され、 matches(1)には 1 番目ののキャプチャ用サブパターンにマッチした 文字列が代入され、といったようになります。
'  flags    = int       PREG_OFFSET_CAPTURE   このフラグを設定した場合、各マッチに対応する文字列のオフセットも返されます。 これにより、返り値は配列となり、配列の要素 0 はマッチした文字列、 要素 1は対象文字列中におけるマッチした文字列のオフセット値 となることに注意してください。
'  offset   = int       通常、検索は対象文字列の先頭から開始されます。 オプションのパラメータ offset  を使用して 検索の開始位置を指定することも可能です。
'【戻り値】
'  preg_match() は、pattern  がマッチした回数を返します。
'  つまり、0 回（マッチせず）または 1 回となります。
'  これは、最初にマッチした時点でpreg_match()  は検索を止めるためです。
'【処理】
'  ・pattern  で指定した正規表現により subject  を検索します。
'=======================================================================
Const PREG_PATTERN_ORDER  = 1
Const PREG_SET_ORDER      = 2
Const PREG_OFFSET_CAPTURE = 256
Function preg_match(pattern, ByVal subject, ByRef matches, flags, offset)

    Dim matchAll,matchone
    Dim cnt,helper

    preg_match = false

    set helper = new RegExp_Helper
    helper.parseOption(pattern)

    Set matchAll = new RegExp
    With matchAll
        .IgnoreCase = helper.withIgnoreCase
        .Global     = false
        .pattern    = helper.withPattern
        .MultiLine  = helper.withMultiLine
    End With

    set helper = Nothing

    If not is_empty(offset) Then
        offset = int(offset)
        subject = Mid(subject,offset)
    End If

    offset = intval( offset )

    If vartype(matches) <> 8 Then
        Set matchone = matchAll.execute(subject)
        Set matchAll = Nothing
        If matchone.count = 0 Then exit Function

        If flags = PREG_OFFSET_CAPTURE Then

            ReDim matches(1)
            matches(0) = matchone(0).value
            matches(1) = offset + matchone(0).FirstIndex
        Else
            ReDim matches(0)
            matches(0) = matchone(0).value
        End If

        preg_match = true
    Else
        preg_match = matchAll.Test(subject)
        Set matchAll = Nothing
    End If

End Function
%>
