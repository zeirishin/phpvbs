'=======================================================================
'正規表現検索および置換を行う
'=======================================================================
'【引数】
'  pattern      = mixed 検索を行うパターン。文字列もしくは配列とすることができます。
'  callback     = mixed このコールバック関数は、検索対象文字列でマッチした要素の配列が指定されて コールされます。このコールバック関数は、置換後の文字列を返す必要があります。
'  subject      = mixed 検索・置換対象となる文字列もしくは文字列の配列
'  limit        = int   subject  文字列において、各パターンによる 置換を行う最大回数。デフォルトは -1 (制限無し)。
'  cnt          = int   この引数が指定されると、置換回数が渡されます。
'【戻り値】
'  subject  引数が配列の場合は配列を、 その他の場合は文字列を返します。
'  パターンがマッチした場合、〔置換が行われた〕新しい subject  を返します。
'  マッチしなかった場合、subject  をそのまま返します。
'【処理】
'  ・subject  に関して pattern  を用いて検索を行い、 callback  に置換します。
'=======================================================================
Function preg_replace_callback(pattern,callback,ByVal subject,limit,ByRef cnt)

    Dim key,counter
    cnt = 0
    If len(limit) = 0 Then limit = 0

    If isArray(subject) Then
        For key = 0 to uBound(subject)
            subject(key) = preg_replace( pattern, callback, subject(key),limit,cnt)
        Next
    ElseIf isObject(subject) Then
        For Each key In subject
            subject(key) = preg_replace( pattern, callback, subject(key),limit,cnt)
        Next
    Else

        If isArray(pattern) Then
                For key = 0 to uBound(pattern)
                    subject = preg_replace( pattern(key), callback, _
                                   subject,limit,cnt)
                Next

        Else

            Dim matchAll,strCallback
            If is_empty(limit) Then
                If preg_match_all(pattern, subject, matchAll,PREG_PATTERN_ORDER,"") <> false Then
                    For Each key In matchAll(0)
                        execute("strCallback = " & callback & "(key)")
                        subject = Replace(subject,key,strCallback)
                    Next
                End If

            Else
                If preg_match_all(pattern, subject, matchAll,PREG_PATTERN_ORDER,"") <> false Then
                    For Each counter In matchAll(0)
                        cnt = cnt + 1
                        If cnt > limit Then Exit For
                        execute("strCallback = " & callback & "(counter)")
                        subject = Replace(subject,counter,strCallback)
                    Next
                End If
            End If

        End If
    End If

    preg_replace_callback = subject

End Function
