'=======================================================================
'正規表現検索および置換を行う
'=======================================================================
'【引数】
'  pattern      = mixed 検索を行うパターン。文字列もしくは配列とすることができます。
'  replacement  = mixed 置換を行う文字列もしくは文字列の配列。
'  subject      = mixed 検索・置換対象となる文字列もしくは文字列の配列
'  limit        = int   subject  文字列において、各パターンによる 置換を行う最大回数。デフォルトは -1 (制限無し)。
'  cnt          = int   この引数が指定されると、置換回数が渡されます。
'【戻り値】
'  subject  引数が配列の場合は配列を、 その他の場合は文字列を返します。
'  パターンがマッチした場合、〔置換が行われた〕新しい subject  を返します。
'  マッチしなかった場合、subject  をそのまま返します。
'【処理】
'  ・subject  に関して pattern  を用いて検索を行い、 replacement  に置換します。
'=======================================================================
Function preg_replace(pattern,replacement,ByVal subject,limit,ByRef cnt)

    Dim key
    cnt = 0

    If isArray(subject) Then
        For key = 0 to uBound(subject)
            subject(key) = preg_replace( pattern, replacement, subject(key),limit,cnt)
        Next
    ElseIf isObject(subject) Then
        For Each key In subject
            subject(key) = preg_replace( pattern, replacement, subject(key),limit,cnt)
        Next
    Else

        If isArray(pattern) Then
            If not isArray(replacement) Then
                For key = 0 to uBound(pattern)
                    subject = preg_replace( pattern(key), replacement, _
                                   subject,limit,cnt)
                Next
            ElseIf isArray(replacement) Then

                If uBound(pattern) <> uBound(replacement) Then
                    ReDim Preserve replacement(uBound(pattern))
                End If
                For key = 0 to uBound(pattern)
                    subject = preg_replace( pattern(key), replacement(key), _
                                   subject,limit,cnt)
                Next
            End If

        Else

            Dim strRetValue, RegEx,helper

            set helper = new RegExp_Helper
            helper.parseOption(pattern)

            Set RegEx = New RegExp
            With RegEx
                .IgnoreCase = helper.withIgnoreCase
                .Global     = [?](limit,false,true)
                .pattern    = helper.withPattern
                .MultiLine  = helper.withMultiLine
            End With
            set helper = Nothing

            If RegEx.Global Then
                If  len(subject) > 0 Then _
                subject = RegEx.Replace(subject, replacement)
            Else
                For key = 1 to limit
                    subject = RegEx.Replace(subject, replacement)
                    cnt = cnt + 1
                Next
            End If

            set RegEx = Nothing

        End If
    End If

    preg_replace = subject

End Function
