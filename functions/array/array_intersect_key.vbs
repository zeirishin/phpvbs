<%
'=======================================================================
'キーを基準にして配列の共通項を計算する
'=======================================================================
'【引数】
'  mAry1    = array  値を調べるもととなる配列。
'  mAry2    = array  値を比較する対象となる配列。
'【戻り値】
'  mAry1 の値のうち、 すべての引数に存在するキーのものを含む連想配列を返します。
'【処理】
'  ・他の全ての引数に存在する mAry1 の値を全て有する配列を返します。
'=======================================================================
Function array_intersect_key(mAry1,mAry2)

    Dim result : set result = Server.CreateObject("Scripting.Dictionary")
    Dim result_keys,key
    ReDim arg_keys(1)

    arg_keys(0) = array_keys(mAry1,"",false)
    arg_keys(1) = array_keys(mAry2,"",false)
    set result_keys = array_intersect(arg_keys(0),arg_keys(1))

    For Each key In result_keys
        result.Add result_keys(key) ,mAry1(result_keys(key))
    Next
    set array_intersect_key = result

End Function
%>
