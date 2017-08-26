'=======================================================================
'配列をシャッフルする
'=======================================================================
'【引数】
'  arr        = Array   配列。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・この関数は、配列をシャッフル (要素の順番をランダムに) します。
'=======================================================================
Function shuffle(ByRef arr)

    shuffle = false
    If not isArray(arr) Then Exit Function

    Dim key,j,x,i : i = count(arr,0)

    Randomize

    For key = 0 to uBound(arr)
        i = i -1
        j = Round(Rnd * i)
        [=] x , arr(i)
        [=] arr(i) , arr(j)
        [=] arr(j) , x
    Next

    shuffle = true

End Function
