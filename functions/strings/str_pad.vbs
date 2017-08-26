<%
'=======================================================================
' 文字列を固定長の他の文字列で埋める
'=======================================================================
'【引数】
'  input        = string    入力文字列。
'  pad_length   = int       pad_length  の値が負、 または入力文字列の長さよりも短い場合、埋める操作は行われません。
'  pad_string   = string    必要とされる埋める文字数が pad_string  の長さで均等に分割できない場合、pad_string  は切り捨てられます。 
'  pad_type     = int       オプションの引数 pad_type  には、 STR_PAD_RIGHT, STR_PAD_LEFT, STR_PAD_BOTH  を指定可能です。 pad_type が指定されない場合、 STR_PAD_RIGHT  を仮定します。
'【戻り値】
'  フォーマット文字列 format  に基づき生成された文字列を返します。
'【処理】
'  ・この関数は文字列 input  の左、右または両側を指定した長さで埋めます。オプションの引数 pad_string  が指定されていない場合は、 input  は空白で埋められ、それ以外の場合は、 pad_string  からの文字で制限まで埋められます。
'=======================================================================
Const STR_PAD_LEFT  = 0
Const STR_PAD_RIGHT = 1
Const STR_PAD_BOTH  = 2
Function str_pad(byVal input, pad_length, pad_string, pad_type)

    Dim half : half = ""
    Dim pad_to_go

    If pad_type <> STR_PAD_LEFT and pad_type <> STR_PAD_RIGHT and pad_type <> STR_PAD_BOTH Then
        pad_type = STR_PAD_RIGHT
    End If

    If len(pad_string) = 0 Then pad_string = " "

    pad_to_go = pad_length - len( input )
    If pad_to_go > 0 Then
        If pad_type = STR_PAD_LEFT Then
            input = str_pad_helper(pad_string, pad_to_go) & input
        ElseIf pad_type = STR_PAD_RIGHT Then
            input = input & str_pad_helper(pad_string, pad_to_go)
        ElseIf pad_type = STR_PAD_BOTH Then
            half = str_pad_helper(pad_string,intval(pad_to_go/2))
            input = half & input & half
            input = Left(input,pad_length)
        End If
    End if

    str_pad = input

End Function

'***************************
Function str_pad_helper(s, intlen)

    Dim collect : collect = ""
    Dim i

    Do Until len( collect ) > intlen
        collect = collect & s
    Loop

    collect = Left(collect,intlen)

    str_pad_helper = collect

End Function
%>
