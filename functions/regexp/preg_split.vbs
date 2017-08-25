<%
'=======================================================================
'正規表現で文字列を分割する
'=======================================================================
'【引数】
'  pattern      = string 検索するパターンを表す文字列。
'  subject      = string 入力文字列。
'  limit        = int    これを指定した場合、最大 limit  個の部分文字列が返されます。
'  flags        = int    flags  は、フラグを組み合わせたものとする （ビット和演算子｜で組み合わせる）ことが可能です。
'【戻り値】
'  pattern  にマッチした境界で分割した subject  の部分文字列の配列を返します。
'【処理】
'  ・指定した文字列を、正規表現で分割します。
'=======================================================================
Const PREG_SPLIT_NO_EMPTY       = 1
Const PREG_SPLIT_DELIM_CAPTURE  = 2
Const PREG_SPLIT_OFFSET_CAPTURE = 4
Function preg_split(pattern, subject,limit,flags)

    If is_empty(limit) Then limit = 0

    Dim key,matches,tmp_sp,tmp_str
    Dim cnt,counter,strMid,pointer : pointer = 1 : counter = 0
    Dim strRegExp,intPoint,strPoint

    cnt = preg_match_all(pattern,subject, matches, PREG_OFFSET_CAPTURE, "")
    If cnt > 0 Then
        For key = 0 to uBound(matches(0))

            counter = counter + 1
            If limit > 0 Then
                If counter >= limit Then Exit For
            End If

            intPoint  = matches(0)(key)(1)
            strPoint  = matches(0)(key)(0)
            strRegExp = Mid(subject, pointer,intPoint-pointer+1)

            Select Case flags
            Case PREG_SPLIT_NO_EMPTY
                if len(strRegExp) > 0 Then [] tmp_sp , strRegExp
            Case PREG_SPLIT_DELIM_CAPTURE
                [] tmp_sp , strRegExp
                [] tmp_sp , matches(1)(key)(0)
            Case PREG_SPLIT_OFFSET_CAPTURE
                [] tmp_sp , array(strRegExp,pointer-1)
            Case Else
                [] tmp_sp , strRegExp
            End Select
            pointer = intPoint + 1 + len(strPoint)
        Next

        strRegExp = Mid(subject, pointer)
        Select Case flags
        Case PREG_SPLIT_NO_EMPTY
            if len(strRegExp) > 0 Then [] tmp_sp , strRegExp
        Case PREG_SPLIT_DELIM_CAPTURE
            [] tmp_sp , strRegExp
        Case PREG_SPLIT_OFFSET_CAPTURE
            [] tmp_sp , array(strRegExp,pointer-1)
        Case Else
            [] tmp_sp , strRegExp
        End Select
    End If

    preg_split = tmp_sp

End Function
%>
