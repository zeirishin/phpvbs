<%
'=======================================================================
'変数名とその値から配列を作成する
'=======================================================================
'【引数】
'  varname    = mixed   変数名の配列とすることができます。
'【戻り値】
'  追加された全ての変数を値とする出力配列を返します。
'【処理】
'  ・数名とその値から配列を作成します。
'  ・各引数について、compact() は現在のシンボルテーブルにおいてその名前を有する変数を探し、 変数名がキー、変数の値がそのキーに関する値となるように追加します。
'=======================================================================
Function compact(varname)

    If Not isArray(varname) Then Exit Function

    Dim output : set output = Server.CreateObject("Scripting.Dictionary")
    Dim var,code

    For Each var In varname
        code = "If output.Exists(var) Then" & vbCrLf & _
                "   output.Item(var) = " & var & vbCrLf & _
                "Else" & vbCrLf & _
                "   output.Add var, " & var & vbCrLf & _
                "End If" & vbCrLf
        execute (code)

    Next

    set compact = output

End Function
%>
