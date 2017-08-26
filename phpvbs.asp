<%
'/**
' *
' * php.asp :  Rapid Development Framework
' * Copyright 2005-2015, zeirishin
' *
' * @filesource
' * @copyright		Copyright 2005-2017, zeirishin
' * @link			http://phpvbs.verygoodtown.com/
' * @package		php.asp
' */

'array_change_key_case()で使用
Const CASE_UPPER = 1
Const CASE_LOWER = 0

'sort用
Const SORT_REGULAR = 0
Const SORT_NUMERIC = 1
Const SORT_STRING  = 2

'coount用
Const COUNT_RECURSIVE = 1

'=======================================================================
'一つの要素を配列の最後に追加する
'=======================================================================
'【引数】
'  mAry     = mixed  配列
'  mVal     = mixed  追加する要素
'【戻り値】
'  値を返しません。
'【処理】
'  ・渡された変数を mAry  の最後に加えます。
'=======================================================================
Sub [](ByRef mAry, ByVal mVal)

    If IsArray(mAry) Then
        Dim counter : counter = UBound(mAry) + 1
        ReDim Preserve mAry(counter)
        mAry(counter) = mVal
    Else
        mAry = Array(mVal)
    End If

End Sub

'=======================================================================
'配列をディクショナリに変換する
'=======================================================================
'【引数】
'  arr  = array  配列
'【戻り値】
'  ディクショナリオブジェクト。
'【処理】
'  ・渡された配列を ディクショナリオブジェクトに変換します。
'=======================================================================
Function array2Dic(ByVal myAry)

    Dim i,tmpObj
    set tmpObj = Server.CreateObject("Scripting.Dictionary")
    For i = 0 to uBound(myAry)
        tmpObj.add i, myAry(i)
    Next
    set array2Dic = tmpObj

End Function

'=======================================================================
'配列を作成する
'=======================================================================
'【引数】
'  mAry     = mixed  配列
'  mVal     = mixed  追加する要素の数
'【戻り値】
'  値を返しません。
'【処理】
'  ・mAryを配列にします。
'=======================================================================
Sub toReDim(ByRef mAry, ByVal mVal)

    If isArray(mAry) Then
        ReDim Preserve mAry(mVal)
    Else
        ReDim mAry(mVal)
    End If

End Sub

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

'=======================================================================
'配列を分割する
'=======================================================================
'【引数】
'  mAry     = Array         処理を行う配列。
'  size     = int           各部分のサイズ。
'  preserve_keys = bool     TRUE の場合はキーをそのまま保持します。 デフォルトは FALSE で、各部分のキーをあらためて数字で振りなおします。
'【戻り値】
'  数値添字の多次元配列を返します。添え字はゼロから始まり、 各次元の要素数が size  となります。
'【処理】
'  ・配列を、size  個ずつの要素に分割します。 
'  ・最後の部分の要素数は size  より小さくなることもあります。
'=======================================================================
Function array_chunk(mAry,size)

    If not isNumeric(size) Then Exit Function
    If size < 1 then Exit Function

    Dim x,i,c : x = 0 : c = -1
    Dim l : l = uBound(mAry)
    Dim n : n = int(l / size)
    ReDim tmpAry(n)

    For i = 0 to l
        x = i Mod size

        If x >= 1 Then
            If isObject(mAry(i)) Then
                set tmpAry(c)(x) = mAry(i)
            Else
                tmpAry(c)(x) = mAry(i)
            End If
        Else
            c = c +1
            If n <> c Then
                toReDim tmpAry(c),size -1
            Else
                toReDim tmpAry(c),l -i
            End If

            If isObject(mAry(i)) Then
                set tmpAry(c)(0) = mAry(i)
            Else
                tmpAry(c)(0) = mAry(i)
            End If

        End If
    Next

    array_chunk = tmpAry

End Function

'=======================================================================
'一方の配列をキーとして、もう一方の配列を値として、ひとつの配列を生成する
'=======================================================================
'【引数】
'  keys     = array  キーとして使用する配列。
'  values   = array  値として使用する配列。
'【戻り値】
'  作成した配列を返します。
'  互いの配列の要素の数が合致しない場合に FALSE を返します。
'【処理】
'  ・keys  配列の値をキーとして、
'  ・また values  配列の値を対応する値として生成した 配列 を作成します。
'=======================================================================
Function array_combine(keys,values)

    Dim obj
    set obj = Server.CreateObject("Scripting.Dictionary")

    If uBound(keys) <> uBound(values) Then
        set array_combine = obj
        Exit Function
    End If

    Dim i
    For i = 0 to uBound(keys)
        If obj.Exists( keys(i) ) Then
            obj.Item( keys(i) ) = values(i)
        Else
            obj.Add keys(i) , values(i)
        End If
    Next

    set array_combine = obj

End Function

'=======================================================================
'配列の値の数を数える
'=======================================================================
'【引数】
'  mAry     = array  値を数える配列。
'【戻り値】
'   mAry のキーとその登場回数を組み合わせた連想配列を作成します。
'【処理】
'  ・配列 mAry の値をキーとし、mAry におけるその値の出現回数を値とした配列を返します。
'【エラー / 例外】
'  ・string あるいは integer 以外の型の要素が登場すると致命的なエラーが発生します。
'=======================================================================
Function array_count_values(mAry)

    Dim obj
    set obj = Server.CreateObject("Scripting.Dictionary")
    if Not isArray(mAry) and Not isObject(mAry) Then
        Set array_count_values = obj
        Exit Function
    End If

    Dim j,k
    Dim intCounter


    For Each j In mAry

        intCounter = 0

        For Each k In mAry
            If j = k Then intCounter = intCounter + 1
        Next

        If Not obj.Exists(j) Then obj.Add j, intCounter

    Next

    Set array_count_values = obj

End Function

'=======================================================================
'追加された添字の確認を含めて配列の差を計算する
'=======================================================================
'【引数】
'  mAry1    = array  比較元の配列。
'  mAry2    = array  比較する対象となる配列。
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・array_diff() とは異なり、 配列のキーを用いて比較を行います。
'=======================================================================
Function array_diff_assoc(ByVal mAry1,ByVal mAry2)

    Dim retAry
    set retAry = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = retAry : Exit Function
    End If

    Dim j,k
    For Each j in mAry1

        retAry.Add j, mAry1(j)

        For Each k In mAry2
            if j = k and mAry1(j) = mAry2(k) Then retAry.Remove k
        Next
    Next

    set array_diff_assoc = retAry

End Function

'=======================================================================
'キーを基準にして配列の差を計算する
'=======================================================================
'【引数】
'  mAry1    = array  比較元の配列。
'  mAry2    = array  比較する対象となる配列。
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・この関数は array_diff() に似ていますが、 値ではなくキーを用いて比較するという点が異なります。
'=======================================================================
Function array_diff_key(ByVal mAry1,ByVal mAry2)

    Dim arr_dif
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) then
        set array_diff_uassoc = arr_dif
        Exit Function
    End If

    Dim key
    For Each key In mAry1
        arr_dif.Add key, mAry1(key)
    Next

    If isObject(mAry2) Then
        For Each key In mAry2
            If arr_dif.Exists( key ) Then arr_dif.Remove key
        Next
    End If

    set array_diff_key = arr_dif

End Function

'=======================================================================
'ユーザが指定したコールバック関数を利用し、 追加された添字の確認を含めて配列の差を計算する
'=======================================================================
'【引数】
'  mAry1            = array     比較元の配列。
'  mAry2            = array     比較する対象となる配列。
'  key_compare_func = callback  使用するコールバック関数。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・ユーザが指定したコールバック関数を用いて添字を比較します。
'=======================================================================
Function array_diff_uassoc(ByVal mAry1,ByVal mAry2,key_compare_func)

    Dim arr_dif
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = arr_dif : Exit Function
    End If

    Dim j,k,callback_ret
    For Each j in mAry1

        arr_dif.Add j, mAry1(j)

        For Each k In mAry2
            If mAry1(j) = mAry2(k) Then
                execute("callback_ret = " & key_compare_func & "(j,k)")
                If callback_ret = 0 Then
                    If arr_dif.Exists(j) Then arr_dif.Remove j
                ElseIf callback_ret < 0 Then
                    arr_dif.Remove j
                    If arr_dif.Exists(k) Then
                        arr_dif.Item( k ) = mAry2(k)
                    Else
                        arr_dif.Add k ,mAry2(k)
                    End If
                End If
            End If
        Next
    Next

    set array_diff_uassoc = arr_dif

End Function

'=======================================================================
'配列の差を計算する
'=======================================================================
'【引数】
'  mAry1    = array  比較元の配列。
'  mAry2    = array  比較する対象となる配列。
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・二つの要素は、(string) elem1 = (string) elem2  の場合のみ等しいと見直されます。
'  ・言い換えると、文字列表現が同じ場合となります。 
'=======================================================================
Function array_diff(ByVal mAry1,ByVal mAry2)

    Dim arr_dif,key_c,key,found
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then
        set mAry1 = array2Dic(mAry1)
    End If

    If isArray(mAry2) Then
        set mAry2 = array2Dic(mAry2)
    End If


    For Each key In mAry1

        found = false
        For Each key_c In mAry2
            If mAry1(key) = mAry2(key_c) Then
                found = true
                Exit For
            End If
        Next

        If Not found Then
            arr_dif.add key, mAry1(key)
        End If
    Next


    set array_diff = arr_dif

End Function

'=======================================================================
'キーを指定して、配列を値で埋める
'=======================================================================
'【引数】
'  keys     = array     キーとして使用する値の配列。
'  val      = string    文字列か、あるいは値の配列。
'【戻り値】
'  値を埋めた配列を返します。
'【処理】
'  ・パラメータ val  で指定した値で配列を埋めます。 
'  ・キーとして、配列 keys  で指定した値を使用します。
'=======================================================================
Function array_fill_keys(keys, val)

    Dim ary_fill,i
    set ary_fill = Server.CreateObject("Scripting.Dictionary")
    set array_fill_keys = ary_fill
    if Not isArray(keys) then Exit Function
    If isArray(val) Then
        If uBound(val) > uBound(keys) then Exit Function
    End If

    If IsArray(val) Then
        For i = 0 to uBound(keys)
            If Not ary_fill.Exists( keys(i) ) Then ary_fill.Add keys(i), val(i)
        Next
    Else
        For i = 0 to uBound(keys)
            If Not ary_fill.Exists( keys(i) ) Then ary_fill.Add keys(i), val
        Next
    End If

    set array_fill_keys = ary_fill

End Function

'=======================================================================
'配列を指定した値で埋める
'=======================================================================
'【引数】
'  start_index  = int       返される配列の最初のインデックス。
'  num          = int       挿入する要素数。
'  val          = string    要素に使用する値。
'【戻り値】
'  値を埋めた配列を返します。
'【処理】
'  ・パラメータ value  を値とする num  個のエントリからなる配列を埋めます。 
'  ・この際、キーは、start_index  パラメータから開始します。
'=======================================================================
Function array_fill(start_index, num, val)

    If Not isNumeric(num) or num < 1 then Exit Function

    Dim intCounter,ary()
    Dim i

    intCounter = start_index + num -1
    ReDim ary(intCounter)

    For i = start_index to intCounter
        ary(i) = val
    Next

    array_fill = ary

End Function

'=======================================================================
'配列の要素をフィルタリングする
'=======================================================================
'【引数】
'  mAry         = aarray    処理する配列。
'  callback     = callback  使用するコールバック関数。コールバック関数が与えられなかった場合、 input のエントリの中で FALSE と等しいもの (boolean への変換 を参照ください) がすべて削除されます

'【戻り値】
'  フィルタリングされた結果の配列を返します。
'【処理】
'  ・mAry のエントリの中で FALSE と等しいもの がすべて削除されます。
'=======================================================================
Function array_filter(ByRef mAry,callback)

    If isArray(mAry) Then

        Dim intCounter,i,strType,callback_ret
        intCounter = uBound(mAry)

        For i = 0 to intCounter
            callback_ret = true
            If Len( callback ) > 0 Then _
                execute("callback_ret = " & callback & "(mAry(i))")

            If callback_ret = true and ( mAry(i) = empty or isNull(mAry(i)) ) Then
                mAry = array_remove(mAry,i)
                call array_filter(mAry,callback)
                Exit For
            End If
        Next

    ElseIf isObject(mAry) Then
        Dim j
        For Each j IN mAry
            callback_ret = true
            If Len( callback ) > 0 Then _
                execute("callback_ret = " & callback & "(mAry(i))")

            If callback_ret = true and ( mAry(j) = empty or isNull(mAry(j)) ) Then _
                mAry.Remove j
        Next

    End If

    array_filter = true

End Function

'=======================================================================
'配列のキーと値を反転する
'=======================================================================
'【引数】
'  trans    = array  反転を行うキー/値の組。
'【戻り値】
'  成功した場合に反転した配列、失敗した場合に 空のオブジェクト を返します。
'【処理】
'  ・配列を反転して返します。
'  ・すなわち、trans  のキーが値となり、 trans  の値がキーとなります。
'=======================================================================
Function array_flip(trans)

    Dim aryObj : set aryObj = Server.CreateObject("Scripting.Dictionary")

    If Not isArray(trans) and Not isObject(trans) Then
        set array_flip = aryObj
        Exit Function
    End If


    If isArray(trans) Then

        Dim i
        For i = 0 to uBound(trans)
            aryObj( trans(i) ) = i
        Next

    Elseif isObject(trans) Then

        Dim j
        For Each j In trans
            aryObj( trans(j) ) = j
        Next

    End If

    set array_flip = aryObj
End Function

'=======================================================================
'追加された添字の確認も含めて配列の共通項を確認する
'=======================================================================
'【引数】
'  mAry1    = array  値を調べるもととなる配列。
'  mAry2    = array  値を比較する対象となる配列。
'【戻り値】
'  mAry1 の値のうち、すべての引数に存在するものを含む連想配列を返します。
'【処理】
'  ・全ての引数に現れる mAry1 の全ての値を含む配列を返します。 
'  ・array_intersect() と異なり、 キーが比較に使用されることに注意してください。
'=======================================================================
Function array_intersect_assoc(mAry1,mAry2)

    Dim intersect : set intersect = Server.CreateObject("Scripting.Dictionary")
    Dim key,counter

    If isArray(mAry2) Then
        counter = uBound(mAry2)
    Else
        counter = null
    End If

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)
            intersect.Add key, mAry1(key)
            If counter >= key or isNull(counter) Then
                If mAry2(key) <> mAry1(key) Then
                    intersect.Remove key
                End If
            End If
        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1
            intersect.Add key, mAry1(key)
            If isNull(counter) or (isNumeric(key) and counter >= key) Then
                If mAry2(key) <> mAry1(key) Then
                    intersect.Remove key
                End If
            Else
               intersect.Remove key
            End If
        Next
    End If

    set array_intersect_assoc = intersect

End Function

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

'=======================================================================
'追加された添字の確認も含め、コールバック関数を用いて 配列の共通項を確認する
'=======================================================================
'【引数】
'  mAry1            = array     比較元となる最初の配列。
'  mAry2            = array     キーを比較する対象となる最初の配列。
'  key_compare_func = callback  比較に使用する、ユーザ定義のコールバック関数。
'【戻り値】
'  mAry1 の値のうち、 すべての引数に存在するもののみを返します。
'【処理】
'  ・全ての引数に現れる mAry1 の全ての値を含む配列を返します。 array_intersect() と異なり、 キーが比較に使用されることに注意してください。
'  ・比較は、ユーザが指定したコールバック関数を利用して行われます。 
'  ・この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'=======================================================================
Function array_intersect_uassoc(mAry1,mAry2,key_compare_func)

    Dim result : set result = Server.CreateObject("Scripting.Dictionary")
    Dim key,k,found,compare

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)
            found = false

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            End If

            If found = true Then
                result.Add k, mAry1(key)
            End if
        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1
            found = false

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 and mAry1(key) = mAry2(k) Then
                        found = true
                        Exit For
                    End If
                Next
            End If

            If found = true Then
                result.Add k, mAry1(key)
            End if
        Next
    End If

    set array_intersect_uassoc = result

End Function

'=======================================================================
'キーを基準にし、コールバック関数を用いて 配列の共通項を計算する
'=======================================================================
'【引数】
'  mAry1    = array  比較元となる最初の配列。
'  mAry2    = array  キーを比較する対象となる最初の配列。
'  key_compare_func = callback  比較に使用する、ユーザ定義のコールバック関数。
'【戻り値】
'  mAry1 の値のうち、 すべての引数に存在するキーのものを含む連想配列を返します。
'【処理】
'  ・他の全ての引数に存在する mAry1 の値を全て有する配列を返します。
'=======================================================================
Function array_intersect_ukey(mAry1,mAry2,key_compare_func)

    Dim result : set result = Server.CreateObject("Scripting.Dictionary")
    Dim key,k,compare

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            End If

        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1

            If isArray(mAry2) Then
                For k = 0 to uBound(mAry2)
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            ElseIf isObject(mAry2) Then
                For Each k In mAry2
                    execute("compare = " & key_compare_func & "(key,k)")
                    If compare = 0 Then
                        result.Add key, mARy1(key)
                        Exit For
                    End If
                Next
            End If

        Next
    End If

    set array_intersect_ukey = result

End Function

'=======================================================================
'配列の共通項を計算する
'=======================================================================
'【引数】
'  mAry1    = array  値を調べるもととなる配列。
'  mAry2    = array  値を比較する対象となる配列。
'【戻り値】
'  mAry1 の値のうち、すべての引数に存在する値のものを含む配列を返します。
'【処理】
'  ・他の全ての引数に存在する mAry1 の値を全て有する配列を返します。
'=======================================================================
Function array_intersect(mAry1,mAry2)

    Dim key
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then
        For key = 0 to uBound(mAry1)
            If len(array_search(mAry1(key),mAry2,false)) > 0 Then
                output.Add key, mAry1(key)
            End If
        Next
    ElseIf isObject(mAry1) Then
        For Each key In mAry1
            If len(array_search(mAry1(key),mAry2,false)) > 0 Then
                output.Add key, mAry1(key)
            End If
        Next
    End If

    set array_intersect = output

End Function

'=======================================================================
'指定したキーまたは添字が配列にあるかどうかを調べる
'=======================================================================
'【引数】
'  key      = mixed  配列
'  sAry     = array  キーが存在するかどうかを調べたい配列。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・指定した key  が配列に設定されている場合、 array_key_exists() は TRUE を返します。 
'  ・key  は配列添字として使用できる全ての値を使用可能です。
'=======================================================================
Function array_key_exists(key, sAry)

    array_key_exists = false
    If isObject(sAry) Then
        if sAry.Exists( key ) then array_key_exists = true
    ElseIf isArray(sAry) and isNumeric(key) Then
        If (uBound(sAry)+1) > key Then
            If Not isNull(sAry(key)) Then array_key_exists = true
        End If
    End If

End Function

'=======================================================================
'配列のキーをすべて返す
'=======================================================================
'【引数】
'  mAry         = array  返すキーを含む配列。
'  search_value = mixed  指定した場合は、これらの値を含むキーのみを返します。
'  strict       = mixed  検索時に型比較を行います。
'【戻り値】
'  mAry のすべてのキーを配列で返します。
'【処理】
'  ・配列 mAry から全てのキー (数値および文字列) を返します。
'  ・オプション search_value  が指定された場合、 指定した値に関するキーのみが返されます。
'  ・指定されない場合は、mAry から全てのキーが返されます。
'  ・strict  パラメータを使って、 比較に型も比較することができます。
'=======================================================================
Function array_keys(mAry,search_value,strict)

    Dim tmp_arr
    Dim key
    Dim include
    Dim addArr
    Dim cnt : cnt = 0

    addArr = true
    If [==](search_value,empty) Then
        addArr = false
        ReDim tmp_arr( count(mAry,0)-1 )
    End If

    If isObject( mAry ) Then

        For Each key In mAry
            include = true
            If [!=](search_value,empty) Then
                If strict = true Then
                    If [!=](mAry(key) , search_value) or (varType(mAry(key)) <> varType(search_value)) Then
                        include = false
                    End If
                ElseIf [!=](mAry(key) , search_value) Then
                    include = false
                End If
            End If

            If include = true Then
                If addArr Then
                    [] tmp_arr, key
                Else
                    tmp_arr(cnt) = key
                    cnt = cnt + 1
                End If
            End If
        Next

    ElseIf isArray(mAry) Then

        For cnt = 0 to uBound(mAry)

            include = true
            If [!=](search_value,empty) Then

                If strict = true Then
                    If [!=](mAry(cnt) , search_value) or (varType(mAry(cnt)) <> varType(search_value)) Then
                        include = false
                    End If
                ElseIf [!=](mAry(cnt) , search_value) Then
                    include = false
                End If
            End If

            If include = true Then
                If addArr Then
                    [] tmp_arr, cnt
                Else
                    tmp_arr(cnt) = cnt
                End If
            End if
        Next
    End If

    array_keys = tmp_arr

End Function

'=======================================================================
'指定した配列の要素にコールバック関数を適用する
'=======================================================================
'【引数】
'  callback = callback  配列の各要素に適用するコールバック関数。
'  arr      = array     コールバック関数を適用する配列。
'【戻り値】
'  arr の各要素に callback  関数を適用した後、 その全ての要素を含む配列を返します。
'【処理】
'  ・arr の各要素に callback  関数を適用します。
'=======================================================================
Function array_map(callback, arr)

    Dim key
    Dim tmp_ar

    If isArray( arr ) Then

        If Len( callback ) = 0 Then
            array_map = arr
            Exit Function
        End If

        ReDim tmp_ar( uBound(arr) )
        For key = 0 to uBound( arr )
            If isObject( arr(key) ) Then
                execute("set tmp_ar(key) = " & callback & "(arr(key))")
            Else
                execute("tmp_ar(key) = " & callback & "(arr(key))")
            End If
        Next

        array_map = tmp_ar

    ElseIf isObject( arr ) Then

        If Len( callback ) = 0 Then
            set array_map = arr
            Exit Function
        End If

        Dim return_val

        set tmp_ar = Server.CreateObject("Scripting.Dictionary")
        For Each key In arr
            return_val = ""
            If isObject( arr.Item(key) ) Then
                execute("set return_val = " & callback & "(arr.Item(key))")
            Else
                execute("return_val = " & callback & "(arr.Item(key))")
            End If
            tmp_ar.Add key, return_val
        Next

        set array_map = tmp_ar

    End If

End Function

'=======================================================================
'二つ以上の配列を再帰的にマージする
'=======================================================================
'【引数】
'  mAry1    = array  マージするもとの配列。
'  mAry2    = array  再帰的にマージしていく配列。
'【戻り値】
'  すべての引数の配列をマージした結果の配列を返します。
'【処理】
'  ・ 一つ以上の配列の要素をマージし、 前の配列の最後にもう一つの配列の値を追加します。 
'  ・ マージした後の配列を返します。
'  ・ 入力配列が同じ文字列のキーを有している場合、 これらのキーの値は配列に一つのマージされます。
'  ・ これは再帰的に行われます。 
'  ・ このため、値の一つが配列自体を指している場合、 この関数は別の配列の対応するエントリもマージします。 
'  ・ しかし、配列が同じ数値キーを有している場合、 後の値は元の値を上書せず、追加されます。 
'=======================================================================
Function array_merge_recursive(mAry1,mAry2)

    Dim j
    Dim retAry : set retAry = Server.CreateObject("Scripting.Dictionary")

    If isObject( mAry1 ) Then
        For Each j In mAry1
            if Not retAry.Exists(j) then retAry.Add j, mAry1(j)
        Next

    ElseIf isArray( mAry1 ) Then
        For j = 0 to uBound( mAry1 )
            retAry.Add j, mAry1(j)
        Next

    End If

    If isObject( mAry2 ) Then
        For Each j In mAry2
            If isObject( mAry2(j) ) Then

                set retAry(j) = array_merge_recursive(retAry(j),mAry2(j))

            Elseif retAry.Exists(j) then
                retAry.Item(j) = array(retAry.Item(j) , mAry2(j))
            Else
                retAry.Add j, mAry2(j)
            End If
        Next

    ElseIf isArray( mAry2 ) Then
        For j = 0 to uBound( mAry2 )
            if retAry.Exists(j) then
                retAry.Item(j) = array(retAry.Item(j) , mAry2(j))
            Else
                retAry.Add j, mAry2(j)
            End If
        Next
    End If

    set array_merge_recursive = retAry

End Function

'=======================================================================
'ひとつまたは複数の配列をマージする
'=======================================================================
'【引数】
'  mAry1    = array  最初の配列。
'  mAry2    = array  再帰的にマージしていく配列。
'【戻り値】
'  結果の配列を返します。
'【処理】
'  ・前の配列の後ろに配列を追加することにより、 ひとつまたは複数の配列の要素をマージし、得られた配列を返します。
'  ・入力配列が同じキー文字列を有していた場合、そのキーに関する後に指定された値が、 前の値を上書きします。
'  ・しかし、配列が同じ添字番号を有していても 値は追記されるため、このようなことは起きません。
'  ・配列が一つだけ指定され、その配列が数字で添字指定されていた場合、 キーの添字が連続となるように振り直されます。 
'=======================================================================
Function array_merge(mAry1,mAry2)

    Dim j,k
    Dim ret,retAry

    If isArray(mAry1) AND isArray(mAry2) Then

        If is_empty(mAry1) Then
            array_merge = mAry2
            Exit Function
        End If

        If is_empty(mAry2) Then
            array_merge = mAry1
            Exit Function
        End If

        Dim cnt : cnt = 0
        Dim uBoundCnt : uBoundCnt = count(mAry1,0) + count(mAry2,0)
        ReDim retAry(uBoundCnt-1)

        For Each j In mAry1
            If isObject(j) Then
                set retAry(cnt) = j
            Else
                retAry(cnt) = j
            End If
            cnt = cnt + 1
        Next

        For Each j In mAry2
            If isObject(j) Then
                set retAry(cnt) = j
            Else
                retAry(cnt) = j
            End If
            cnt = cnt + 1
        Next

        array_merge = retAry

    Else
        If Not isObject(retAry) Then
            set retAry = Server.CreateObject("Scripting.Dictionary")
        End If
    
        If isObject( mAry1 ) Then
            For Each j In mAry1
                if Not retAry.Exists(j) then retAry.Add j, mAry1(j)
            Next
    
        ElseIf isArray( mAry1 ) Then
            For j = 0 to uBound( mAry1 )
                retAry.Add j, mAry1(j)
            Next
    
        End If
    
        If isObject( mAry2 ) Then
            For Each j In mAry2
                If isObject( mAry2(j) ) Then

                    set retAry(j) = Server.CreateObject("Scripting.Dictionary")
                    set ret = array_merge(retAry(j),mAry2(j))

                    For Each k in ret
                        if retAry(j).Exists(k) then
                            retAry(j).Item(k) = ret(k)
                        Else
                            retAry(j).Add k, ret(k)
                        End If
                    Next
                Elseif retAry.Exists(j) then
                    retAry.Item(j) = mAry2(j)
                Else
                    retAry.Add j, mAry2(j)
                End If
            Next
    
        ElseIf isArray( mAry2 ) Then
            For j = 0 to uBound( mAry2 )
                if retAry.Exists(j) then
                    retAry.Item(j) = mAry2(j)
                Else
                    retAry.Add j, mAry2(j)
                End If
            Next
        End If
    
        set array_merge = retAry
    End If

End Function

'=======================================================================
'複数の多次元の配列をソートする
'=======================================================================
'【引数】
'  arr  = array  ソートしたい配列。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  array_multisort() は、多次元の配列をその次元の一つでソートする際に使用可能です。
'=======================================================================
Function array_multisort(ByRef arr)

    array_multisort = false
    If not isArray(arr) Then Exit Function

    Dim key
    For key = 0 to uBound(arr)
        array_multisort arr(key)
    Next

    sort arr,0

    array_multisort = true

End Function

'=======================================================================
'指定長、指定した値で配列を埋める
'=======================================================================
'【引数】
'  mAry         = array  値を埋めるもととなる配列。
'  pad_size     = int    新しい配列のサイズ。
'  pad_value    = mixed  mAry が pad_size より小さいときに、 埋めるために使用する値。
'【戻り値】
'  pad_size  で指定した長さになるように値 pad_value  で埋めて mAry のコピーを返します。
'  pad_size  が正の場合、配列の右側が埋められます。 
'  負の場合、配列の左側が埋められます。 
'  pad_size  の絶対値が mAry の長さ以下の場合、埋める処理は行われません。
'【処理】
'  pad_size  で指定した長さになるように値 pad_value  で埋めて mAry のコピーを返します。
'=======================================================================
Function array_pad(ByVal mAry, pad_size, pad_value)

    If Not isArray( mAry ) Then Exit Function
    If Not isNumeric( pad_size ) Then Exit Function

    Dim pad,aryCounter,newLength,i,intCounter

    If pad_size < 0 Then
        newLength = pad_size * -1
    Else
        newLength = pad_size
    End If
    newLength = newLength -1

    aryCounter = uBound(mAry)
    If newLength > aryCounter Then

        ReDim pad(newLength)
        intCounter = 0
        For i = 0 to newLength
            If pad_size < 0 Then
                If newLength - aryCounter > i Then
                    pad(i) = pad_value
                Else
                    pad(i) = mAry(intCounter)
                    intCounter = intCounter + 1
                End If
            Else
                If i > aryCounter Then
                    pad(i) = pad_value
                Else
                    pad(i) = mAry(i)
                End If
            End If
        Next
    Else
        pad = mAry
    End If

    array_pad = pad

End Function

'=======================================================================
'配列の末尾から要素を取り除く
'=======================================================================
'【引数】
'  mAry = array  値を取り出す配列。
'【戻り値】
'  配列 mAry の最後の値を取り出して返します。
'  mAry が空 (または、配列でない) の場合、 NULL が返されます。
'【処理】
'  ・array  の最後の値を取り出して返します。
'  ・配列 array  は、要素一つ分短くなります。
'=======================================================================
Function array_pop(ByRef mAry)

    If Not isArray(mAry) Then
        array_pop = null
        Exit Function
    End If

    Dim intCounter
    intCounter = uBound( mAry )
    array_pop = mAry( intCounter )
    ReDim Preserve mAry(intCounter - 1)

End Function

'=======================================================================
'配列の値の積を計算する
'=======================================================================
'【引数】
'  mAry     = array  配列
'【戻り値】
'  積を、integer あるいは float で返します。
'【処理】
'  ・配列の各要素の積を計算します。
'=======================================================================
Function array_product(mAry)

    If Not isArray( mAry ) Then Exit Function

    Dim j,product
    product = 1

    For Each j In mAry
        If isNumeric(j) Then product = product * j
    Next

    array_product = product

End Function

'=======================================================================
'一つ以上の要素を配列の最後に追加する
'=======================================================================
'【引数】
'  mAry     = array  配列
'  mVal     = mixed  追加する要素
'【戻り値】
'  処理後の配列の中の要素の数を返します。
'【処理】
'  ・渡された変数を mAry の最後に加えます。
'  ・配列 mAry の長さは渡された変数の数だけ増加します。
'=======================================================================
Function array_push(ByRef mAry, ByVal mVal)

    Dim intCounter
    Dim intElementCount

    If IsArray(mAry) Then
        If IsArray(mVal) Then

            intElementCount = UBound(mAry)
            ReDim Preserve mAry(intElementCount + UBound(mVal) + 1)

            For intCounter = 0 to UBound(mVal)
                mAry(intElementCount + intCounter + 1) = mVal(intCounter)
            Next

        Else
            ReDim Preserve mAry(UBound(mAry) + 1)
            mAry(UBound(mAry)) = mVal
        End If
    Else

        If IsArray(mVal) Then
            mAry = mVal
        Else
            mAry = Array(mVal)
        End If
    End If

    array_push = UBound(mAry)

End Function

'=======================================================================
'配列から一つ以上の要素をランダムに取得する
'=======================================================================
'【引数】
'  mAry     = array  入力の配列。
'  num_req  = int    取得するエントリの数を指定します。 指定されない場合は、デフォルトの 1 になります。
'【戻り値】
'  エントリを一つだけ取得する場合、 array_rand() はランダムなエントリのキーを返します。
'  その他の場合は、ランダムなエントリのキーの配列を返します。 
'  これにより、ランダムなキーを取得し、 配列から値を取得することが可能になります。
'【処理】
'  ・配列から一つ以上のランダムなエントリを取得しようとする場合に有用です。
'=======================================================================
Function array_rand(mAry, ByVal num_req)

    If Not isArray( mAry ) Then Exit Function
    If Not isNumeric( num_req ) Then num_req = 1

    Dim rand,i,intCounter,aryCounter,indexes

    intCounter = uBound(mAry)
    aryCounter = num_req -1

    If intCounter < aryCounter Then Exit Function

    Randomize

    ReDim indexes( aryCounter )
    For i = 0 to aryCounter
        Do While true
            rand = Round( Rnd * uBound(mAry) )
            If Not in_array(rand, indexes,true) Then
                indexes(i) = rand
                Exit Do
            End If
        Loop
    Next

    If num_req = 1 Then
        array_rand = indexes(0)
    Else
        array_rand = indexes
    End If

End Function

'=======================================================================
'コールバック関数を用いて配列を普通の値に変更することにより、配列を再帰的に減らす
'=======================================================================
'【引数】
'  mAry     = array     入力の配列。
'  callback = callback  コールバック関数。
'  initial  = int       オプションの intial が利用可能な場合、処理の最初で使用されたり、 配列が空の場合の最終結果として使用されます。
'【戻り値】
'   結果の値を返します。
'   配列が空で initial が渡されなかった場合は、 array_reduce() は NULL を返します。 
'【処理】
'  ・配列 mAry の各要素に callback 関数を繰り返し適用し、 配列を一つの値に減らします。
'=======================================================================
Function array_reduce(ByVal mAry, callback, ByVal initial)

    array_reduce = null
    If len( initial ) > 0 Then array_reduce = initial
    If not isArray( mAry ) and not isObject( mAry ) Then Exit Function

    Dim acc : acc = initial
    Dim key

    If isObject( mAry ) Then
        For Each key In mAry
            execute("acc = " & callback & "(acc, mAry(key))")
        Next

    ElseIf isArray( mAry ) Then

        Dim lon : lon = uBound( mAry )
        For key = 0 to lon
            execute("acc = " & callback & "(acc, mAry(key))")
        Next
    End If

    array_reduce = acc

End Function

'=======================================================================
'配列の指定した要素を一つ削除する
'=======================================================================
'【引数】
'  mAry     = array  対象となる配列
'  num      = int    削除する要素番号
'【戻り値】
'  処理後の配列を返します。
'【処理】
'  ・mAry のnum番目の要素を削除します。
'  ・配列 mAry の長さは一つ減少します。
'=======================================================================
Function array_remove(mAry,num)

    if Not isArray(mAry) Then Exit Function
    If Not isNumeric(num) Then Exit Function

    Dim strCount
    strCount = uBound(mAry)
    If strCount+1 < num Then
        array_remove = mAry
        Exit Function
    End If

    If (strCount+1) = num Then
        ReDim Preserve mAry(strCount - 1)
        array_remove = mAry
        Exit Function
    End If

    If num = 0 Then
        call array_shift(mAry)
        array_remove = mAry
        Exit Function
    End If

    Dim tmpAry,retAry
    tmpAry = array_chunk(mAry,num)
    call array_shift( tmpAry(1) )

    call array_push(tmpAry(0),tmpAry(1))
    retAry = tmpAry(0)

    if uBound(tmpAry) > 1 Then

        Dim intCounter
        For intCounter = 2 to uBound(tmpAry)
            call array_push(retAry,tmpAry(intCounter))
        Next

    end if

    array_remove = retAry

End Function

'=======================================================================
'要素を逆順にした配列を返す
'=======================================================================
'【引数】
'  ary              = Array 入力の配列。
'【戻り値】
'  ・逆転させた配列を返します。
'【処理】
'  ・配列を受け取って、要素の順番を逆にした新しい配列を返します。
'=======================================================================
Function array_reverse(mAry)

    Dim arr_len,i

    If isArray(mAry) Then

        Dim tmp_ar()
        Dim newkey

        arr_len = uBound(mAry)
        ReDim tmp_ar(arr_len)

        For i = 0 to arr_len
            newkey = arr_len -i
            tmp_ar(i) = mAry(newkey)
        Next

        array_reverse = tmp_ar

    ElseIf isObject(mAry) Then

        Dim tmpObj,j,cnt

        cnt = 0
        set tmpObj = Server.CreateObject("Scripting.Dictionary")
        arr_len = mAry.Count-1

        ReDim index_values(arr_len),index_keys(arr_len)

        For Each j In mAry
            index_values(cnt) = mAry(j)
            index_keys(cnt)   = j
            cnt = cnt + 1
        Next

        For i = cnt-1 To 0 Step -1

            If Not tmpObj.Exists(Cstr(index_keys(i))) Then
                tmpObj.add Cstr(index_keys(i)),index_values(i)
            End if
        Next

        set array_reverse = tmpObj
    End If
End Function

'=======================================================================
'指定した値を配列で検索し、見つかった場合に対応するキーを返す
'=======================================================================
'【引数】
'  needle   = mixed 探したい値。
'  haystack = array 配列。
'  strict   = mixed TRUE が指定された場合、array_search() は haystack  の中で needle  の型に一致するかどうかも確認します。
'【戻り値】
'  ・needle  が見つかった場合に配列のキー、 それ以外の場合に FALSE を返します。
'【処理】
'  ・haystack  において needle  を検索します。
'=======================================================================
Function array_search(needle,haystack,strict)

    array_search = false

    If IsArray(needle) or isObject(needle) Then Exit Function

    If VarType(strict) <> 11 Then strict = false

    Dim key
    If isObject( haystack ) Then
        For Each key In haystack
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                array_search = array_search(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    array_search = key
            End If
            If array_search <> false Then Exit For
        Next
    ElseIf isArray( haystack ) Then
        For key = 0 to uBound( haystack )
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                array_search = array_search(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    array_search = key
            End If
            If array_search <> false Then Exit For
        Next
    End If

End Function

'=======================================================================
'配列の先頭から要素を一つ取り出す
'=======================================================================
'【引数】
'  ary     = Array 配列
'【戻り値】
'  ・ary  の最初の値を取り出して返します。
'  ・array  が空の場合 (または配列でない場合)、 NULL が返されます。
'【処理】
'  ・配列 ary  は、要素一つ分だけ短くなり、全ての要素は前にずれます。 
'  ・数値添字の配列のキーはゼロから順に新たに振りなおされますが、 リテラルのキーはそのままになります。
'=======================================================================
Function array_shift(ByRef ary)

    If Not isArray(ary) and Not isObject(ary) then
        array_shift = null
        Exit Function
    End If

    Dim i,key : i = 0

    If isArray(ary) Then
        array_shift = ary(0)

        For i = 0 to uBound(ary)-1
            ary(i) = ary(i+1)
        Next
        Redim Preserve ary(UBound(ary) - 1)

    ElseIf isObject(ary) Then
        For Each key In ary
            array_shift = ary(key)
            ary.Remove(key)
            Exit For
        Next
    End if

End Function

'=======================================================================
'配列の一部を展開する
'=======================================================================
'【引数】
'  mAry     = Array 入力の配列。
'  offset   = int   offset  が負の値ではない場合、要素位置の計算は、 配列 array  の offset から始められます。 offset  が負の場合、要素位置の計算は array  の最後から行われます。
'  level    = int  level が指定され、正の場合、 連続する複数の要素が返されます。level が指定され、負の場合、配列の末尾から連続する複数の要素が返されます。 省略された場合、offset  から配列の最後までの全ての要素が返されます。
'【戻り値】
'  ・切り取った部分を返します。
'【処理】
'  ・mAry から引数 offset  および level で指定された連続する要素を返します。
'=======================================================================
Function array_slice(mAry,offset,level)

    array_slice = false

    If Not isArray(mAry) Then Exit Function
    If Not isNumeric(offset) Then Exit Function
    If Not isNumeric(level) Then level = uBound(mAry)

    Dim s,e,arynum
    arynum = uBound(mAry)

    If offset >= 0 Then _
        s = offset _
    Else _
        s = arynum + offset + 1

    If level >= 0 Then _
        e = s + level _
    Else _
        e = arynum + level

    If e > arynum Then e = arynum

    Dim i,counter
    counter = 0
    ReDim tmp_ar(e-s)
    For i = s to e
        tmp_ar(counter) = mAry(i)
        counter = counter +1
    Next

    array_slice = tmp_ar

End Function

'=======================================================================
'配列の一部を削除し、他の要素で置換する
'=======================================================================
'【引数】
'  mAry         = Array 入力の配列。
'  offset       = int   offset  が正の場合、削除される部分は 配列 input  の最初から指定オフセットの ぶんだけ進んだ位置からとなります。 offset  が負の場合、削除される部分は、 input  の末尾から数えた位置からとなります。
'  level        = int   length  が省略された場合、 offset  から配列の最後までが全て削除されます。 length  が指定され、正の場合、複数の要素が削除されます。 length  が指定され、負の場合、 削除される部分は配列の末尾から複数の要素となります。 ヒント: replacement  も指定した場合に offset  から配列の最後まで全てを削除するには、 length  を求めるために count($input)  を使用してください。
'  replacement  = int    配列 replacement  が指定された場合、 削除された要素は、この配列の要素で置換されます。offset および length で何も削除しないと指定した場合、配列 replacement の要素は offset で指定された位置に挿入されます。 置換される配列のキーは保存されないことに注意してください。もし replacement に一つしか要素がない場合、 要素そのものが配列でない限り、array() で括る必要はありません
'【戻り値】
'  ・抽出された要素を含む配列を返します。
'【処理】
'  ・配列 input  から offset  および length  で指定された要素を削除し、配列 replacement  でそれを置換します。
'=======================================================================
Function array_splice(ByRef mARy,offset,level,replacement)
End Function

'=======================================================================
'配列の中の値の合計を計算する
'=======================================================================
'【引数】
'  mAry         = Array 入力の配列。
'【戻り値】
'  ・値の合計を整数または float として返します。
'【処理】
'  ・配列の中の値の合計を整数または float として返します。
'=======================================================================
Function array_sum(mAry)

    array_sum = 0
    If Not isArray(mAry) and Not isObject(mAry) Then Exit Function

    Dim key
    If isObject(mAry) Then
        For Each key in mAry
            array_sum = array_sum + mAry(key)
        Next
    Else
        For Each key in mAry
            array_sum = array_sum + key
        Next
    End If

End Function

'=======================================================================
'データの比較にコールバック関数を用い、 追加された添字の確認を含めて配列の差を計算する
'=======================================================================
'【引数】
'  mAry1     = Array 最初の配列。
'  mAry2     = Array 2 番目の配列。
'  mAry1     = callback 比較用のコールバック関数。ユーザが指定したコールバック関数を用いてデータの比較を行います。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。

'【戻り値】
'  ・他の引数のいずれにも存在しない mAry1 の値の全てを有する配列を返します。
'【処理】
'  ・データの比較にコールバック関数を用い、配列の差を計算します。 
'  ・この関数は array_diff_assoc() と異なり、 データの比較に内部関数を利用します。
'=======================================================================
Function array_udiff_assoc(ByVal mAry1,ByVal mAry2, data_compare_func)

    Dim arr_udif
    set arr_udif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) then
        set array_diff_uassoc = arr_udif
        Exit Function
    End If

    Dim key,found
    For Each key In mAry1
        arr_udif.Add key, mAry1(key)
    Next

    If Not isObject(mAry2) Then Exit Function

    For Each key In mAry2
        If arr_udif.Exists( key ) Then
            execute("found = " & data_compare_func & "(arr_udif(key), mAry2(key))")
            If found = 0 Then
                If arr_udif.Exists( key ) Then arr_udif.Remove key
            ElseIf found < 0 Then
                If isObject(mAry2(key)) Then
                    set arr_udif.Item( key ) = mAry2(key)
                Else
                    arr_udif.Item( key ) = mAry2(key)
                End If
            End if
        End If
    Next

    set array_udiff_assoc = arr_udif

End Function

'=======================================================================
'ユーザが指定したコールバック関数を利用し、 追加された添字の確認を含めて配列の差を計算する
'=======================================================================
'【引数】
'  mAry1                = array     比較元の配列。
'  mAry2                = array     比較する対象となる配列。
'  date_compare_func    = callback  使用するコールバック関数。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'  key_compare_func     = callback  キー（添字）の比較は、コールバック関数
'【戻り値】
'  mAry1  の要素のうち、 その他の配列のいずれにも含まれないものだけを残した配列を返します。
'【処理】
'  ・mAry1  を mAry2 と比較し、その差を返します。
'  ・ユーザが指定したコールバック関数を用いて添字を比較します。
'=======================================================================
Function array_udiff_uassoc(ByVal mAry1,ByVal mAry2, data_compare_func,key_compare_func)

    Dim arr_dif
    set arr_dif = Server.CreateObject("Scripting.Dictionary")

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = arr_dif : Exit Function
    End If

    Dim j,k,key_result,data_result,found
    For Each j in mAry1

        found = false
        For Each k In mAry2
            execute("key_result  = " & key_compare_func & "(j,k)")
            execute("data_result = " & data_compare_func & "(mAry1(j),mAry2(k))")

            If key_result = 0 and data_result = 0 Then
                found = true
                Exit For
            End If
        Next

        If Not found Then
             arr_dif.Add j , mAry1(j)
        End If
    Next

    set array_udiff_uassoc = arr_dif

End Function

'=======================================================================
'データの比較にコールバック関数を用い、配列の差を計算する
'=======================================================================
'【引数】
'  mAry1     = Array 最初の配列。
'  mAry2     = Array 2 番目の配列。
'  mAry1     = callback 比較用のコールバック関数。ユーザが指定したコールバック関数を用いてデータの比較を行います。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。

'【戻り値】
'  ・他の引数のいずれにも存在しない mAry1 の値の全てを有する配列を返します。
'【処理】
'  ・データの比較にコールバック関数を用い、配列の差を計算します。 
'  ・この関数は array_diff() と異なり、 データの比較に内部関数を利用します。
'=======================================================================
Function array_udiff(mAry1,mAry2,data_compare_func)

    Dim arr_udif,key_c,key,found

    If Not isObject(mAry1) or Not isObject(mAry2) Then
        set array_diff_assoc = retAry : Exit Function
    End If

    If isArray(mAry1) and isArray(mAry2) Then

        For Each key In mAry1

            found = 0
            For Each key_c In mAry2
                execute("found = " & data_compare_func & "(key, key_c)")
                If found <> 0 Then
                    Exit For
                End If
            Next

            If found > 0 Then
                [] arr_udif, mAry1(key)
            ElseIf found < 0 Then
                [] arr_udif, mAry2(key_c)
            End If
        Next

        array_udiff = arr_udif

    ElseIf isObject(mAry1) and isObject(mAry2) Then

        set arr_udif = Server.CreateObject("Scripting.Dictionary")

        For Each key In mAry1

            found = 0
            For Each key_c In mAry2
                execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                If found <> 0 Then
                    Exit For
                End If
            Next

            If found > 0 Then
                If arr_udif.Exists(key) Then
                    arr_udif.Item(key) = mAry1(key)
                Else
                    arr_udif.Add key, mAry1(key)
                End If
            ElseIf found < 0 Then
                If arr_udif.Exists(key_c) Then
                    arr_udif.Item(key_c) = mAry2(key_c)
                Else
                    arr_udif.Add key_c, mAry2(key_c)
                End If
            End If

        Next

        set array_udiff = arr_udif

    End If

End Function

'=======================================================================
'データの比較にコールバック関数を用い、配列の共通項を計算する
'=======================================================================
'【引数】
'  mAry1                = Array     最初の配列。
'  mAry2                = Array     2 番目の配列。
'  data_compare_func    = callback  比較用のコールバック関数。比較は、ユーザが指定したコールバック関数を利用して行われます。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'【戻り値】
'  ・他の全ての引数に存在する mAry1 の値を全て有する配列を返します。
'【処理】
'  ・データの比較にコールバック関数を用い、配列の共通項を計算します。
'=======================================================================
Function array_uintersect_assoc(mAry1,mAry2,data_compare_func)
'Callbackの例
'function rmul(v, w)
'    rmul = 0
'    If isObject(v) or isArray(v) Then
'        rmul = 1
'    Elseif isObject(w) or isArray(w) Then
'        rmul = 1
'    End If
'    If rmul = 1 then Exit FUnction
'    If v = w Then
'        rmul = 0
'    Else
'        rmul = 1
'    End If
'End Function

    Dim key,key_c
    Dim found
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then

        For key = 0 to uBound(mAry1)

            If array_key_exists(key, mAry2) Then
                execute("found = " & data_compare_func & "(mAry1(key), mAry2(key))")
                If found = 0 Then
                    If output.Exists(key) Then
                        output(key) = mAry1(key)
                    Else
                        output.Add key, mAry1(key)
                    End If
                End If
            End If
        Next

    ElseIf isObject(mAry1) Then

        For Each key In mAry1
            If array_key_exists(key, mAry2) Then
                execute("found = " & data_compare_func & "(mAry1(key), mAry2(key))")
                If found = 0 Then

                    If output.Exists(key) Then
                        output(key) = mAry1(key)
                    Else
                        output.Add key, mAry1(key)
                    End If
                End If
            End If
        Next

    End If

    set array_uintersect_assoc = output

End Function

'=======================================================================
'データと添字の比較にコールバック関数を用い、 追加された添字の確認も含めて配列の共通項を計算する
'=======================================================================
'【引数】
'  mAry1                = Array     最初の配列。
'  mAry2                = Array     2 番目の配列。
'  data_compare_func    = callback  比較用のコールバック関数。比較は、ユーザが指定したコールバック関数を利用して行われます。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'  key_compare_func     = callback  キーの比較用のコールバック関数。
'【戻り値】
'  ・他の全ての引数に現れる mAry1 の値を含む配列を返します。
'【処理】
'  ・データの比較にコールバック関数を用い、配列の共通項を計算します。
'  ・キーが比較に使用されることに注意してください。 
'  ・データと添字は、それぞれ個別のコールバック関数を用いて比較されます。
'=======================================================================
Function array_uintersect_uassoc(mAry1,mAry2,data_compare_func,key_compare_func)

'Callbackの例
'function rmul(v, w)
'    rmul = 0
'    If isObject(v) or isArray(v) Then
'        rmul = 1
'    Elseif isObject(w) or isArray(w) Then
'        rmul = 1
'    End If
'    If rmul = 1 then Exit FUnction
'    If v = w Then
'        rmul = 0
'    Else
'        rmul = 1
'    End If
'End Function

    Dim key,key_c
    Dim found,key_found
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then

        For key = 0 to uBound(mAry1)

            If isArray(mAry2) Then
                For key_c = 0 to uBound(mAry2)
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next

            ElseIf isObject(mAry2) Then

                For Each key_c In mAry2
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next
            End If
        Next

    ElseIf isObject(mAry1) Then

        For Each key In mAry1
            If isArray(mAry2) Then
                For key_c = 0 to uBound(mAry2)
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next

            ElseIf isObject(mAry2) Then

                For Each key_c In mAry2
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    execute("key_found = " & key_compare_func & "(key, key_c)")
                    If found = 0 and key_found = 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next
            End If
        Next

    End If

    set array_uintersect_uassoc = output

End Function

'=======================================================================
'データの比較にコールバック関数を用い、配列の共通項を計算する
'=======================================================================
'【引数】
'  mAry1                = Array     入力の配列。
'  mAry2                = Array     2 番目の配列。
'  data_compare_func    = callback  比較用のコールバック関数。比較は、ユーザが指定したコールバック関数を利用して行われます。 この関数は、1 つめの引数が 2 つめより小さい / 等しい / 大きい 場合にそれぞれ 負の数 / ゼロ / 正の数 を返す必要があります。
'【戻り値】
'  ・他の全ての引数に存在する mAry1 の値を全て有する配列を返します。
'【処理】
'  ・データの比較にコールバック関数を用い、配列の共通項を計算します。
'=======================================================================
Function array_uintersect(mAry1,mAry2,data_compare_func)

'Callbackの例
'function rmul(v, w)
'    rmul = 1
'    If isObject(v) or isArray(v) Then
'        rmul = 1
'    Elseif isObject(w) or isArray(w) Then
'        rmul = 1
'    End If
'    If rmul = 1 then Exit FUnction
'    If v = w Then
'        rmul = 1
'    Else
'        rmul = 0
'    End If
'End Function

    Dim key,key_c
    Dim found
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(mAry1) Then

        For key = 0 to uBound(mAry1)

            If isArray(mAry2) Then
                For key_c = 0 to uBound(mAry2)
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    If found <> 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next

            ElseIf isObject(mAry2) Then

                For Each key_c In mAry2
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    If found <> 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next
            End If
        Next

    ElseIf isObject(mAry1) Then

        For Each key In mAry1
            If isArray(mAry2) Then
                For key_c = 0 to uBound(mAry2)
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    If found <> 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next

            ElseIf isObject(mAry2) Then

                For Each key_c In mAry2
                    execute("found = " & data_compare_func & "(mAry1(key), mAry2(key_c))")
                    If found <> 0 Then
                        If output.Exists(key) Then
                            output(key) = mAry1(key)
                        Else
                            output.Add key, mAry1(key)
                        End If
                    End If
                Next
            End If
        Next

    End If

    set array_uintersect = output

End Function

'=======================================================================
'配列から重複した値を削除する
'=======================================================================
'【引数】
'  mAry     = Array 入力の配列。
'【戻り値】
'  ・処理済の配列を返します。
'【処理】
'  ・値に重複のない新規配列を返します。
'=======================================================================
function array_unique(arr)

    Dim key,key_c
    Dim found
    Dim output : set output = Server.CreateObject("Scripting.Dictionary")

    If isArray(arr) Then
        For key = 0 to uBound(arr)
            found = array_search(arr(key),output,false)
            If found = false and varType(found) = 11 Then
                output.Add key, arr(key)
            End If
        Next
    ElseIf isObject(arr) Then
        For Each key In arr
            found = array_search(arr(key),output,false)
            If found = false and varType(found) = 11 Then
                    output.Add key, arr(key)
            End If
        next
    End If

    set array_unique = output
end function

'=======================================================================
'一つ以上の要素を配列の最初に加える
'=======================================================================
'【引数】
'  mAry     = Array 配列
'  mVal     = mixed 追加する要素
'【戻り値】
'  ・処理後の mAry  の要素の数を返します。
'【処理】
'  ・リストの要素は全体として加えられるため、 加えられた要素の順番は変わらないことに注意してください。 
'  ・配列の数値添字はすべて新たにゼロから振りなおされます。 
'  ・リテのキーについては変更されません。
'=======================================================================
Function array_unshift(ByRef mAry, ByVal mVal)

    Dim intCounter
    Dim intElementCount

    If IsArray(mAry) Then
        If IsArray(mVal) Then

            ret = array_push(mVal,mAry)
            mAry = mVal

        Else

            ReDim Preserve mAry(UBound(mAry) + 1)

            For intCounter = UBound(mAry) to 1 Step -1
                mAry(intCounter) = mAry(intCounter -1)
            Next

            mAry(0) = mVal

        End If
    Else
        If IsArray(mVal) Then
            mAry = mVal
        Else
            mAry = Array(mVal)
        End If
    End If

    array_unshift = UBound(mAry)

End Function

'=======================================================================
'配列の全ての値を返す
'=======================================================================
'【引数】
'  mAry     = array 配列。
'【戻り値】
'  数値添字の値の配列を返します。
'【処理】
'  ・配列から全ての値を取り出し、数値添字をつけた配列を返します。
'=======================================================================
Function array_values(mAry)

    Dim tmp_ar
    Dim key,counter : counter= 0

    If isObject(mAry) Then

        ReDim tmp_ar(mAry.Count -1)

        For Each key In mAry
            If isObject(mAry(key)) Then
                set tmp_ar(counter) = mAry(key)
            Else
                tmp_ar(counter) = mAry(key)
            End if
            counter = counter + 1
        Next

    ElseIf isArray(mAry) Then
        tmp_ar = mAry
    End If

    array_values = tmp_ar

End Function

'=======================================================================
'配列の全ての要素に、ユーザー関数を再帰的に適用する
'=======================================================================
'【引数】
'  arr      = array     入力の配列。
'  callback = callback  引数を二つとります。 array  パラメータの値が最初の引数、 キー/添字は二番目の引数となります。funcname  により配列の値そのものを変更する必要がある場合、 funcname  の最初の引数は 参照  として渡す必要があります。この場合、配列の要素に加えた変更は、 配列自体に対して行われます。 
'  userdata = array     userdata  パラメータが指定された場合、 コールバック関数 funcname  への三番目の引数として渡されます。
'【戻り値】
'  成功した場合に TRUE を返します。
'【処理】
'  ・arr の各要素に callback  関数を適用します。
'  ・この関数は配列の要素内を再帰的にたどっていきます。
'=======================================================================
Function array_walk_recursive(ByRef arr, callback, userdata)

    Dim key
    Dim thisCall

    If Len( callback ) = 0 Then Exit Function

    If isArray( arr ) Then

        For key = 0 to uBound( arr )
            If isArray(arr(key)) or isObject(arr(key)) Then
                thisCall = "array_walk_recursive"
            Else
                thisCall = callback
            End If

            execute("call " & callback & "(arr(key),key,userdata)")
        Next

    ElseIf isObject( arr ) Then

        For Each key In arr
            If isArray(arr(key)) or isObject(arr(key)) Then
                thisCall = "array_walk_recursive"
            Else
                thisCall = callback
            End If

            execute("call " & callback & "(arr(key),key,userdata)")
        Next

    End If

    array_walk_recursive = true

End Function

'=======================================================================
'配列の全ての要素にユーザ関数を適用する
'=======================================================================
'【引数】
'  arr      = array     入力の配列。
'  callback = callback  引数を二つとります。 array  パラメータの値が最初の引数、 キー/添字は二番目の引数となります。funcname  により配列の値そのものを変更する必要がある場合、 funcname  の最初の引数は 参照  として渡す必要があります。この場合、配列の要素に加えた変更は、 配列自体に対して行われます。 
'  userdata = array     userdata  パラメータが指定された場合、 コールバック関数 funcname  への三番目の引数として渡されます。
'【戻り値】
'  成功した場合に TRUE を返します。
'【処理】
'  ・arr の各要素に callback  関数を適用します。
'=======================================================================
Function array_walk(ByRef arr, callback, userdata)

    Dim key

    If Len( callback ) = 0 Then Exit Function

    If isArray( arr ) Then

        For key = 0 to uBound( arr )
            execute("call " & callback & "(arr(key),key,userdata)")
        Next

    ElseIf isObject( arr ) Then

        Dim return_val

        For Each key In arr
            execute("call " & callback & "(arr.Item(key),key,userdata)")
        Next

    End If

    array_walk = true

End Function

'=======================================================================
'連想キーと要素との関係を維持しつつ配列を逆順にソートする
'=======================================================================
'【引数】
'  ary        = Array   入力の配列。
'  sort_flags = int     オプションのパラメータ sort_flags  によりソートの動作を修正可能です。詳細については、 sort() を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・この関数は、 連想配列において各配列のキーと要素との関係を維持しつつ配列をソートします。
'  ・この関数は、 主に実際の要素の並び方が重要である連想配列をソートするために使われます。
'=======================================================================
Function arsort(ByRef arr, sort_flags)

    arsort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")
    Dim found
    Dim cnt

    keys = array_values(arr)
    rsort keys,sort_flags

    For Each key In keys
        found = array_keys(arr,key,true)
        If isArray(found) Then
            For cnt = 0 to uBound(found)
                If Not new_arr.Exists(found(cnt)) Then
                    new_arr.Add found(cnt), arr(found(cnt))
                End If
            Next
        Else
            new_arr.Add found, arr(found)
        End If
    Next

    set arr = new_arr

    arsort = true

End Function

'=======================================================================
'連想キーと要素との関係を維持しつつ配列をソートする
'=======================================================================
'【引数】
'  ary        = Array   入力の配列。
'  sort_flags = int     オプションのパラメータ sort_flags  によりソートの動作を修正可能です。詳細については、 sort() を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・この関数は、 連想配列において各配列のキーと要素との関係を維持しつつ配列をソートします。
'  ・この関数は、 主に実際の要素の並び方が重要である連想配列をソートするために使われます。
'=======================================================================
Function asort(ByRef arr, sort_flags)

    asort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")
    Dim found
    Dim cnt

    keys = array_values(arr)
    sort keys,sort_flags

    For Each key In keys
        found = array_keys(arr,key,true)
        If isArray(found) Then
            For cnt = 0 to uBound(found)
                If Not new_arr.Exists(found(cnt)) Then
                    new_arr.Add found(cnt), arr(found(cnt))
                End If
            Next
        Else
            new_arr.Add found, arr(found)
        End If
    Next

    set arr = new_arr

    asort = true

End Function

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


'=======================================================================
'配列に値があるかチェックする
'=======================================================================
'【引数】
'  needle     = mixed 探す値。
'  haystack   = Array  配列。
'  strict     = bool   三番目のパラメータ strict が TRUE に設定された場合、 haystack の中の needle の型も確認します。
'【戻り値】
'  配列で needle  が見つかった場合に TRUE、それ以外の場合は、FALSE を返します。
'【処理】
'  ・haystack配列内にneedleが含まれるかチェック
'=======================================================================
Function in_array(needle, haystack,strict)

    in_array = False

    If Not IsArray(needle) Then
        If Len( needle ) = 0 Then Exit Function
    End If

    If VarType(strict) <> 11 Then strict = false

    Dim key

    If isArray(needle) Then
        For Each key In needle
            in_array = in_array(key,haystack,strict)
            If in_array = true Then Exit For
        Next
        Exit Function
    End If

    If isObject( haystack ) Then
        For Each key In haystack
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                in_array = in_array(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    in_array = true
            End If
            If in_array = true Then Exit For
        Next
    ElseIf isArray( haystack ) Then
        For key = 0 to uBound( haystack )
            If isArray(haystack( key )) or isObject(haystack( key )) Then
                in_array = in_array(needle,haystack( key ),strict)
            ElseIf ( strict and vartype(haystack( key )) = vartype(needle) and haystack( key ) = needle ) or _
               ( Not strict and haystack( key ) = needle ) Then
                    in_array = true
            End If
            If in_array = true Then Exit For
        Next
    End If

End Function


'=======================================================================
'配列をキーで逆順にソートする
'=======================================================================
'【引数】
'  ary        = Array   入力の配列。
'  sort_flags = int     オプションのパラメータ sort_flags  によりソートの動作を修正可能です。詳細については、 sort() を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・配列をキーにより逆順にソートします。
'  ・キーとデータとの関係は維持されます。
'  ・この関数は、主として連想配列において有用です。
'=======================================================================
Function krsort(ByRef arr, sort_flags)

    krsort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")

    keys = array_keys(arr,"",false)
    rsort keys,sort_flags

    For Each key In keys
        new_arr.Add key, arr(key)
    Next

    set arr = new_arr

    krsort = true

End Function

'=======================================================================
'配列をキーでソートする
'=======================================================================
'【引数】
'  ary        = Array   入力の配列。
'  sort_flags = int     オプションのパラメータ sort_flags  によりソートの動作を修正可能です。詳細については、 sort() を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・キーとデータの関係を維持しつつ、配列をキーでソートします。 
'  ・この関数は、主として連想配列において有用です。
'=======================================================================
Function ksort(ByRef arr, sort_flags)

    ksort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")

    keys = array_keys(arr,"",false)
    sort keys,sort_flags

    For Each key In keys
        new_arr.Add key, arr(key)
    Next

    set arr = new_arr

    ksort = true

End Function

Function natcasesort()
End Function

Function natsort()
End Function

'=======================================================================
'ある範囲の整数を有する配列を作成する
'=======================================================================
'【引数】
'  low  = mixed 下限値。
'  high = mixed 上限値。
'  step = mixed step  が指定されている場合、それは 要素毎の増加数となります。step  は正の数でなければなりません。デフォルトは 1 です。
'【戻り値】
'  low  から high  までの整数の配列を返します。 low > high の場合、順番は high から low となります。
'【処理】
'  ・ある範囲の整数を有する配列を作成します。
'=======================================================================
Function range(low,high,step)

    Dim matrix
    Dim inival, endval, plus
    Dim walker : If len(step) > 0 Then walker = step Else walker = 1
    Dim chars : chars = false

    If isNumeric(low) and isNumeric(high) Then
        inival = low
        endval = high
    ElseIf Not isNumeric(low) and Not isNumeric(high) Then
        chars  = true
        inival = Asc(low)
        endval = Asc(high)
    Else
        inival = [?](isNumeric(low),low,0)
        endval = [?](isNumeric(high),high,0)
    End If

    plus = true
    If inival > endval Then plus = false

    If plus Then
        Do While inival <= endval
            [] matrix, [?](chars,Chr(inival),inival)
            inival = inival + walker
        Loop
    Else
        Do While inival >= endval
            [] matrix, [?](chars,Chr(inival),inival)
            inival = inival - walker
        Loop
    End If

    range = matrix

End Function

'=======================================================================
'配列を逆順にソートする
'=======================================================================
'【引数】
'  ary        = Array   入力の配列。
'  sort_flags = int     オプションのパラメータ sort_flags  によりソートの動作を修正可能です。詳細については、 sort() を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・この関数は、配列を逆順に(高位から低位に)ソートします。
'=======================================================================
Function rsort(ByRef arr, sort_flags)

    rsort = false
    If Not IsArray(arr) Then  Exit Function

    Dim i,j
    Dim temp

    For i = 1 to (uBound(arr))

        [=] temp, arr(i)

        j = i -1
        Do While rsort_helper(temp,arr(j),sort_flags)

            [=] arr(j+1) , arr(j)
            j = j -1
            If j < 0 Then Exit Do
        Loop

        [=] arr(j + 1), temp
    Next

    rsort = true

End Function

'******************************************
Function rsort_helper(temp , arr, sort_flags)

    rsort_helper = true
    If isArray(temp) or isObject(temp) Then Exit Function

    rsort_helper = false
    If isArray(arr) or isObject(arr) Then Exit Function

    If varType(sort_flags) <> 2 Then sort_flags = 0

    If sort_flags = SORT_REGULAR Then
        rsort_helper = (temp > arr)
    ElseIf sort_flags = SORT_NUMERIC Then
        rsort_helper = (intval(temp) > intval(arr))
    ElseIf sort_flags = SORT_STRING Then
        rsort_helper = (Cstr(temp) > Cstr(arr))
    End If

End Function
'******************************************

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

'=======================================================================
'count() のエイリアス
'=======================================================================
'【引数】
'  var    = mixed   配列。
'  mode   = int     オプションのmode  引数が COUNT_RECURSIVE  (または 1) にセットされた場合、count()  は再帰的に配列をカウントします。
'【処理】
'  ・この関数は次の関数のエイリアスです。 count().
'=======================================================================
Function sizeof(var,mode)
    sizeof = count(var,mode)
End Function

'=======================================================================
'配列をソートする
'=======================================================================
'【引数】
'  ary        = Array   ソートする配列
'  sort_flags = int     ソートの動作を修正するために使用することが可能です。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・配列をソートします。
'　・この関数が正常に終了すると、 各要素は低位から高位へ並べ替えられます。
'　・http://www.thinkit.co.jp/article/62/3/
'=======================================================================
Function sort(ByRef arr, sort_flags)

    sort = false
    If Not IsArray(arr) Then  Exit Function

    Dim i,j
    Dim temp

    For i = 1 to (uBound(arr))

        [=] temp, arr(i)

        j = i -1
        Do While sort_helper(temp,arr(j),sort_flags)

            [=] arr(j+1) , arr(j)
            j = j -1
            If j < 0 Then Exit Do
        Loop

        [=] arr(j + 1), temp
    Next

    sort = true

End Function

'******************************************
Function sort_helper(temp , arr, sort_flags)

    sort_helper = false
    If isArray(temp) or isObject(temp) Then Exit Function

    sort_helper = true
    If isArray(arr) or isObject(arr) Then Exit Function

    If varType(sort_flags) <> 2 Then sort_flags = 0

    If sort_flags = SORT_REGULAR Then
        sort_helper = (temp < arr)
    ElseIf sort_flags = SORT_NUMERIC Then
        sort_helper = (intval(temp) < intval(arr))
    ElseIf sort_flags = SORT_STRING Then
        sort_helper = (Cstr(temp) < Cstr(arr))
    End If

End Function
'******************************************

'=======================================================================
'ユーザ定義の比較関数で配列をソートし、連想インデックスを保持する
'=======================================================================
'【引数】
'  ary          = Array   入力の配列。
'  cmp_function = int     ユーザ定義の比較関数の例については、 usort() および uksort()  を参照ください。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・ この関数は、配列インデックスが関連する配列要素との関係を保持するような配列をソートします。
'  ・ 主に実際の配列の順序に意味がある連想配列をソートするためにこの関数は使用されます。 
'=======================================================================
Function uasort(ByRef arr, cmp_function)

    uasort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")
    Dim found
    Dim cnt

    keys = array_values(arr)
    usort keys,cmp_function

    For Each key In keys
        found = array_keys(arr,key,true)
        If isArray(found) Then
            For cnt = 0 to uBound(found)
                If Not new_arr.Exists(found(cnt)) Then
                    new_arr.Add found(cnt), arr(found(cnt))
                End If
            Next
        Else
            new_arr.Add found, arr(found)
        End If
    Next

    set arr = new_arr

    uasort = true

End Function

'=======================================================================
'ユーザ定義の比較関数を用いて、キーで配列をソートする
'=======================================================================
'【引数】
'  ary          = Array   入力の配列。
'  cmp_function = int     比較用のコールバック関数。関数 cmp_function は、 array のキーペアによって満たされる 2 つのパラメータを受け取ります。 この比較関数が返す値は、最初の引数が二番目より小さい場合は負の数、 等しい場合はゼロ、そして大きい場合は正の数でなければなりません。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・uksort() は、 ユーザ定義の比較関数を用いて配列のキーをソートします。
'  ・ソートしたい配列を複雑な基準でソートする必要がある場合には、 この関数を使う必要があります。
'=======================================================================
Function uksort(ByRef arr, cmp_function)

    uksort = false
    If Not IsObject(arr) Then  Exit Function

    Dim key,keys
    Dim new_arr : set new_arr = Server.CreateObject("Scripting.Dictionary")

    keys = array_keys(arr,"",false)
    usort keys,cmp_function

    For Each key In keys
        new_arr.Add key, arr(key)
    Next

    set arr = new_arr

    uksort = true

End Function


'=======================================================================
'ユーザー定義の比較関数を使用して、配列を値でソートする
'=======================================================================
'【引数】
'  ary          = Array   入力の配列。
'  cmp_function = int     比較関数は、最初の引数が 2 番目の引数より小さいか、等しいか、大きい場合に、 それぞれゼロ未満、ゼロに等しい、ゼロより大きい整数を返す 必要があります。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・この関数は、ユーザー定義の比較関数により配列をその値でソートします。 
'  ・ソートしたい配列を複雑な基準でソートする必要がある場合、 この関数を使用するべきです。
'=======================================================================
Function usort(ByRef arr, cmp_function)

'ユーザー定義関数の例
'きちんとtrue falseを返さないと動かない
'Function cmp(a,b)
'
'    If [==](a,b) Then
'        cmp = false
'    Else
'        If (a < b) Then
'            cmp = false
'        Else
'            cmp = true
'        End If
'    End If
'End Function



    usort = false
    If Not IsArray(arr) Then  Exit Function

    Dim i,j
    Dim temp

    For i = 1 to (uBound(arr))

        [=] temp, arr(i)
        j = i -1
        Do While usort_helper(temp,arr(j),cmp_function)
            [=] arr(j+1) , arr(j)
            j = j -1
            If j < 0 Then Exit Do
        Loop

        [=] arr(j + 1), temp
    Next

    usort = true


End Function

'*****************************************************
Function usort_helper(temp,arr,cmp_function)

    Dim output
    execute ("output = " & cmp_function & "(temp,arr)")
    usort_helper = output

End Function
'******************************************


'=======================================================================
'英数字かどうかを調べる
'=======================================================================
'【引数】
'  text     = string  調べる文字列。
'【戻り値】
'  text  のすべての文字が英字または数字だった場合に TRUE 、そうでない場合に FALSE を返します。
'【処理】
'  ・与えられた文字列 text  のすべての文字が英字または 数字であるかどうかを調べます。
'=======================================================================
Function ctype_alnum(text)

    ctype_alnum = false
    If len( text ) = 0 Then Exit Function


End Function


Function ctype_alpha()
End Function


Function ctype_cntrl()
End Function


Function ctype_digit()
End Function


Function ctype_graph()
End Function


Function ctype_lower()
End Function


Function ctype_print()
End Function


Function ctype_punct()
End Function


Function ctype_space()
End Function


Function ctype_upper()
End Function


Function ctype_xdigit()
End Function
%>
