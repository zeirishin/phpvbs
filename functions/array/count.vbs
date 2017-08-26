<%
'=======================================================================
'変数に含まれる要素、 あるいはオブジェクトに含まれるプロパティの数を数える
'=======================================================================
'【引数】
'  var    = mixed   配列。
'  mode   = int     オプションのmode  引数が COUNT_RECURSIVE  (または 1) にセットされた場合、count()  は再帰的に配列をカウントします。
'【戻り値】
'   var に含まれる要素の数を返します。 他のものには、1つの要素しかありませんので、通常 var  は配列です。
'   もし var が配列もしくは Countable インターフェースを実装したオブジェクトではない場合、 1 が返されます。
'   ひとつ例外があり、var が NULL の場合、 0 が返されます。 
'【処理】
'  ・変数に含まれる要素、 あるいはオブジェクトに含まれるプロパティの数を数えます。
'=======================================================================
Const COUNT_RECURSIVE = 1
Function count(var,mode)

    If not isArray(var) and not isObject(var) Then
        If isNull(var) Then
            count = 0
        Else
            count = 1
        End If
        Exit Function
    End If

    If mode <> COUNT_RECURSIVE Then

        If isArray(var) Then
            count = uBound(var) + 1
        ElseIf isObject(var) Then
            count = var.Count
        End If
        Exit Function

    Else

        Dim key,output : output = 0
        If isArray(var) Then
            For key = 0 to uBound(var)
                If isArray(var(key)) or isObject(var(key)) Then output = output + 1
                output = output + count(var(key),mode)
            Next
        ElseIf isObject(var) Then
            For Each key In var
                If isArray(var(key)) or isObject(var(key)) Then output = output + 1
                output = output + count(var(key),mode)
            Next
        End If

        count = output
    End If

End Function
%>
