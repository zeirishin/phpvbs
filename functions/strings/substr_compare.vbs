<%
'=======================================================================
' 指定した位置から指定した長さの 2 つの文字列について、バイナリ対応で比較する
'=======================================================================
'【引数】
'  main_str             = string    最初の文字列。
'  str                  = string    次の文字列。
'  offset               = int       比較を開始する位置。 負の値を指定した場合は、文字列の最後から数えます。
'  length               = int       比較する長さ。
'  case_insensitivity   = bool      case_insensitivity  が TRUE の場合、 大文字小文字を区別せずに比較します。
'【戻り値】
'  main_str  の offset  以降が str  より小さい場合に負の数、 str  より大きい場合に正の数、 等しい場合に 0 を返します。
'【処理】
'  ・ substr_compare() は、main_str  の offset  文字目以降の最大 length  文字を、str  と比較します。
'=======================================================================
Function substr_compare(main_str,str,offset,length, case_insensitivity)

    If len(offset) > 0 Then
        If offset > 0 Then
            main_str = Mid(main_str,offset)
        Else
            main_str = Mid(main_str,len(main_str) + offset + 1)
        End If
    End If

    If len(length) > 0 Then
        main_str = Left(main_str,length)
        str = Left(str,length)
    End If
    var_dump main_str
    var_dump str
    If case_insensitivity = true Then
        substr_compare = strcasecmp(main_str,str)
    Else
        substr_compare = strcmp(main_str,str)
    End If

End Function
%>
