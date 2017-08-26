<%
'=======================================================================
'検索文字列に一致したすべての文字列を置換する
'=======================================================================
'【引数】
'  search    = mixed 検索文字列
'  replace   = mixed 置換文字列
'  subject   = mixed 置換対象文字列
'【戻り値】
'  置換後の文字列あるいは配列を返します。
'【処理】
'  ・subject  の中の search  を全て replace  に置換します。
'=======================================================================
Function str_replace(ByVal search, ByVal strReplace, ByVal subject)

    Dim tmp
    Dim J

    If IsObject(search) or IsObject(strReplace) or IsObject(subject) Then Exit Function

    If IsArray(search) and Not IsArray(strReplace) Then
        tmp = strReplace
        ReDim strReplace(UBound(search))
        strReplace(0) = tmp
    ElseIf Not IsArray(search) and IsArray(strReplace) Then
        tmp = search
        ReDim search(UBound(strReplace))
        search(0) = tmp
    End If

    If IsArray(search) and IsArray(strReplace) Then

        If UBound(search) <> UBound(strReplace) Then

            If UBound(search) > UBound(strReplace) Then

                ReDim strReplace(UBound(search))

            ElseIf UBound(search) < UBound(strReplace) Then

                ReDim search(UBound(strReplace))

            End If

        End If

    End If

    If IsArray(subject) Then
        For J = 0 To UBound(subject)
            subject(J) = str_replace(search, strReplace, subject(J))
        Next

    Else

        If IsArray(search) Then
            For J = 0 To UBound(search)
                subject = Replace(subject,search(J),strReplace(J),1,len(subject),vbBinaryCompare)
            Next
        Else
            subject = Replace(subject,search,strReplace,1,len(subject),vbBinaryCompare)
        End If

    End If

    str_replace = subject

End Function
%>
