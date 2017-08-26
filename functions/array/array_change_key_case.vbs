'=======================================================================
'配列のすべてのキーを変更する
'=======================================================================
'【引数】
'  mObj     = objec  処理を行う連想配列。
'  flag     = int    CASE_UPPER あるいは CASE_LOWER (デフォルト)。
'【戻り値】
'  すべてのキーを小文字あるいは大文字にした配列を返します。 input  が配列でない場合は false を返します。
'【処理】
'  ・mAry  のすべてのキーを小文字あるいは大文字にした配列を返します。 数値添字はそのままとなります。
'=======================================================================
Const CASE_UPPER = 1
Const CASE_LOWER = 0
Function array_change_key_case(ByRef mObj, flag)

    if flag <> CASE_UPPER and flag <> CASE_LOWER then Exit Function
    if Not isObject(mObj) Then Exit Function

    Dim cnt,i

    arykey = mObj.Keys
    aryval = mObj.Items
    cnt    = mObj.Count -1
    mObj.RemoveAll

    Select Case flag
    Case CASE_UPPER

        For i = 0 to cnt
            mObj.Add Ucase(arykey(i)), aryval(i)
        Next

    Case CASE_LOWER

        For i = 0 to cnt
            mObj.Add Lcase(arykey(i)), aryval(i)
        Next

    End Select

End Function
