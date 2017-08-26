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
' グレグリオ歴の日付/時刻の妥当性を確認します
'=======================================================================
'【引数】
'  mm   = int   月は 1 から 12 の間となります。
'  dd   = int   日は、指定された month  の日数の範囲内になります。year  がうるう年の場合は、それも考慮されます。
'  yyyy = int   年は 1 から 32767 の間となります。
'【戻り値】
'  指定した日付が有効な場合に TRUE、そうでない場合に FALSE を返します。
'【処理】
'  ・引数で指定された日付の妥当性をチェックします。 各パラメータが適切に指定されている場合に、妥当であると判断されます。
'=======================================================================
Function checkdate(mm,dd,yyyy)
    Dim myDate : myDate = DateSerial(yyyy,mm,dd)
    checkdate = eval(Month(myDate) = mm)
End Function

'=======================================================================
' ローカルな日付・時間をフォーマットする
'=======================================================================
'【引数】
'  format   = string   フォーマット文字列
'【戻り値】
'  指定したフォーマット文字列に基づき文字列をフォーマットして返します。
'【処理】
'  ・設定にもとづいてフォーマットします。
'=======================================================================
Function strftime(format)
    Dim bobj : set bobj = Server.CreateObject("Basp21")
    strftime = bobj.strftime(format)
    set bobj = Nothing
End Function


'file
Const FILE_IGNORE_NEW_LINES = 2
Const FILE_SKIP_EMPTY_LINES = 4

'pathinfo
Const PATHINFO_DIRNAME = 1
Const PATHINFO_BASENAME = 2
Const PATHINFO_EXTENSION = 4
Const PATHINFO_FILENAME = 3

Class File_System

    Private fso
    Private ts

    '
    'Initialize Class
    ' 
    '@access private
    '
    Private Sub Class_Initialize()
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
    End Sub

    '
    'Terminate Class
    ' 
    '@access private
    '
    Private Sub Class_Terminate()
        Set fso = Nothing
    End Sub

    '=======================================================================
    ' パス中のファイル名の部分を返す
    '=======================================================================
    '【引数】
    '  path      = string   パス。
    '  suffix    = string   ファイル名が、 suffix  で終了する場合、 この部分もカットされます。
    '【戻り値】
    '  指定した path  のベース名を返します。
    '【処理】
    '  ・この関数は、ファイルへのパスを有する文字列を引数とし、 ファイルのベース名を返します。
    '=======================================================================
    Function basename(path, suffix)

        Dim b
        b = preg_replace("/^.*[¥/¥¥]/g","",path,null,null)

        If len(suffix) > 0 Then
            If Right(b,len(suffix)) = suffix Then
                b = Left(b,len(b) - len(suffix))
            End If
        End If

        basename = b

    End Function

    '=======================================================================
    ' ファイルをコピーする
    '=======================================================================
    '【引数】
    '  source  = string   コピー元ファイルへのパス。
    '  dest    = string   コピー先のパス。
    '【戻り値】
    '  成功した場合に TRUE を、失敗した場合に FALSE を返します。
    '【処理】
    '  ・ ファイル source  を dest  にコピーします。
    '  ・ ファイルを移動したいならは、rename() 関数を使用してください。 
    '=======================================================================
    Public Function copy(source,dest)
        fso.CopyFile source,dest
    End Function

    '=======================================================================
    ' オープンされたファイルポインタをクローズする
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  
    '【処理】
    '  ファイルをクローズします。
    '=======================================================================
    Public function fclose
        ts.close
        Set ts = Nothing
    end function

    '=======================================================================
    ' ファイルポインタがファイル終端に達しているかどうか調べる
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  ファイルポインタが EOF に達しているかまたはエラー (ソケットタイムアウトを含みます) の場合に TRUE 、 その他の場合に FALSE を返します。
    '【処理】
    '  ファイルポインタがファイル終端に達しているかどうかを調べます。
    '=======================================================================
    Public function feof
        feof = ts.AtEndofStream
    end function

    '=======================================================================
    ' ファイルポインタから1文字取り出す
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  ファイルポインタから 1 文字読み出し、 その文字からなる文字列を返します。EOF の場合に FALSE を返します。
    '【処理】
    '  指定したファイルポインタから 1 文字読み出します。
    '=======================================================================
    Public function fgetc
        If ts.AtEndofStream Then
            fgetc = false
        Else
            fgetc = ts.Read(1)
        End If
    end function

    '=======================================================================
    ' ファイルポインタから行を取得し、CSVフィールドを処理する
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  読み込んだフィールドの内容を含む数値添字配列を返します。
    '【処理】
    '  fgets() に動作は似ていますが、 fgetcsv() は行を CSV  フォーマットのフィールドとして読込み処理を行い、 読み込んだフィールドを含む配列を返すという違いがあります。
    '=======================================================================
    Public Function fgetcsv(delim)

        Dim tmp,d
        If len(delim) > 0 Then d = delim Else d = ","
        tmp = ts.ReadLine
        fgetcsv = fgetcsv_helper(tmp,d)

    End Function

    '************************************
    Public Function fgetcsv_helper(str,d)

        Dim matchAll,key
        Dim data,field,record : field = 0 : record = 0
        ReDim data(0)

        If preg_match_all(_
        "/" & d & "|" & vbCr & "?" & vbLf & "|[^" & d & """" & vbCrLf & "][^" & d & "" & vbCrLf & "]*|""(?:[^""]|"""")*""/",_
        str, matchAll,PREG_PATTERN_ORDER,"") Then
            For Each key In matchAll(0)
                Select Case key
                Case d
                    field = field + 1
                Case vbCrLf
                    [] data , ""
                    record = record +1
                Case Else
                    If left(key,1) = """" Then
                        key = Replace(substr(key,2,-1),"""""","""")
                    End if
                    [] data(record), key
                End Select
            Next
        End If

        fgetcsv_helper = data

    End Function

    '=======================================================================
    ' ファイルポインタから 1 行取得する
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  ファイルポインタから 1 行取得する
    '【処理】
    '  ファイルポインタから 1 行取得します。
    '=======================================================================
    Public function fgets
        fgets = ts.ReadLine
    end function

    '=======================================================================
    ' ファイルポインタから 1 行取り出し、HTML タグを取り除く
    '=======================================================================
    '【引数】
    '  
    '【戻り値】
    '  HTML や PHP コードを取り除いた文字列を返します。
    '【処理】
    '  fgets() と同じですが、 fgetss() は読み込んだテキストから HTML および PHP のタグを取り除こうとすることが異なります。
    '=======================================================================
    Public function fgetss
        fgets = strip_tags(ts.ReadLine)
    end function

    '=======================================================================
    ' ファイルまたはディレクトリが存在するかどうか調べる
    '=======================================================================
    '【引数】
    '  path      = string   ファイルあるいはディレクトリへのパス。
    '【戻り値】
    '  ファイルあるいはディレクトリが存在するかどうかを調べます。
    '【処理】
    '  ファイルあるいはディレクトリが存在するかどうかを調べます。
    '=======================================================================
    Public Function file_exists(ByVal filename)

        file_exists = false
        filename = fileMapPath(filename)
        If fso.FileExists(filename) Then file_exists = true
        If fso.FolderExists(filename) Then file_exists = true

    End Function

    '=======================================================================
    'ファイルの内容を全て文字列に読み込む
    '=======================================================================
    '【引数】
    '  filename  = string データを読み込みたいファイルの名前。
    '【戻り値】
    '  読み込んだデータを返します。失敗した場合は FALSE を返します。
    '【処理】
    '  ファイルの内容を文字列に読み込む
    '=======================================================================
    Public function file_get_contents(filename)

        Dim ts
        Dim contents

        if left(filename,7) <> "http://" and file_exists( filename ) then
            Set TS = fso.OpenTextFile( fileMapPath(filename),1)

            '空のファイルの場合、エラーになってしまう
            If TS.AtEndOfStream <> True Then
               contents = TS.ReadAll
            End If

            file_get_contents = contents
            Exit Function
        end if

        if left(filename,7) <> "http://" then
            file_get_contents = false
            Exit Function
        end if

        Dim objWinHttp
        'Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
        'Set objWinHttp = Server.CreateObject("MSXML2.XMLHTTP")
        Set objWinHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
'        on error resume next

        objWinHttp.Open "GET", filename, false
        objWinHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        objWinHttp.Send
        'Response.Write objWinHttp.Status & " " & objWinHttp.StatusText

        file_get_contents = objWinHttp.ResponseText

        Set objWinHttp = Nothing

        '有効にしたエラー処理を無効にする
'        on error goto 0

    end function

    '=======================================================================
    ' ファイルの最終アクセス時刻を取得する
    '=======================================================================
    '【引数】
    '  filename = string   ファイルへのパス。
    '【戻り値】
    '  ファイルの最終アクセス時刻を返し、エラーの場合は FALSE を返します。
    '【処理】
    '  指定したファイルの最終アクセス時刻を取得します。
    '=======================================================================
    Public Function fileatime(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        fileatime = f.DateLastAccessed

    End Function

    '=======================================================================
    ' ファイルの更新時刻を取得する
    '=======================================================================
    '【引数】
    '  filename = string   ファイルへのパス。
    '【戻り値】
    '  ファイルの最終更新時刻を返し、エラーの場合は FALSE  を返します。
    '【処理】
    '  この関数は、ファイルのブロックデータが書き込まれた時間を返します。 これは、ファイルの内容が変更された際の時間です。
    '=======================================================================
    Public Function filemtime(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        filemtime = f.DateLastModified

    End Function

    '*************************************
    Private Function fileMapPath(filename)

        Dim tmp
        tmp = Left(filename,3)
        tmp = Lcase(tmp)
        If tmp <> "d:¥" and tmp <> "c:¥" and left(filename,7) <> "http://" then
                fileMapPath = Server.MapPath(filename)
        Else
            fileMapPath = filename
        End If

    End Function

    '=======================================================================
    ' ファイルのサイズを取得する
    '=======================================================================
    '【引数】
    '  filename = string   ファイルへのパス。
    '【戻り値】
    '  ファイルのサイズを返し、エラーの場合は FALSE を返します (また E_WARNING レベルのエラーを発生させます) 。
    '【処理】
    '  指定したファイルのサイズを取得します。
    '=======================================================================
    Public Function filesize(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        filesize = f.Size

    End Function

    '=======================================================================
    ' ファイルタイプを取得する
    '=======================================================================
    '【引数】
    '  filename = string   ファイルへのパス。
    '【戻り値】
    '  ファイルのタイプを返します。
    '【処理】
    '  指定したファイルのタイプを返します。
    '=======================================================================
    Public Function filetype(filename)

        Dim f
        filename = fileMapPath(filename)
        set f = fso.GetFile(filename)
        filetype = f.Type

    End Function

    '=======================================================================
    ' ファイルまたは URL をオープンする
    '=======================================================================
    '【引数】
    '  filename  =  string データを読み込みたいファイルの名前。
    '  mode      =  string ストリームに要するアクセス形式を指定します
    '【戻り値】
    '  成功した場合にファイルポインタリソースを返します。
    '【処理】
    '  fopen() は、filename  で指定されたリソースをストリームに結び付けます。
    '=======================================================================
    Public function fopen(filename, mode)

        Dim filePath
        filePath = fileMapPath(filename)

        If left(filePath,len("http://")) = "http://" Then
            fopen = file_get_contents(filePath)
            Exit Function
        End If

        Select Case mode
        Case "r"
            '読み込みのみでオープンします。
            Set ts = fso.OpenTextFile(filePath,1,false)
        Case "w"
            '書き込みでオープンします。
            Set ts = fso.OpenTextFile(filePath,2,true)
        Case "a"
            '追記でオープンします。
            Set ts = fso.OpenTextFile(filePath,8,true)
        Case "x"
            '書き込みでオープンします。ファイルが存在した場合はfalseを返します。
            If is_file(filePath) Then
                fopen = false
            Else
                Set ts = fso.OpenTextFile(filePath,2,true)
            End If
        Case Else
            'empty
            ts = false
        End Select

    end function

    '=======================================================================
    ' 行を CSV 形式にフォーマットし、ファイルポインタに書き込む
    '=======================================================================
    '【引数】
    '  fields  =  string    値の配列。
    '  delimiter  =  string オプションの delimiter  はフィールド区切り文字 (一文字だけ) を指定します。デフォルトはカンマ (,) です。
    '  enclosure  =  string オプションの enclosure  はフィールドを囲む文字 (一文字だけ) を指定します。デフォルトは二重引用符 (") です。
    '【戻り値】
    '  書き込んだ文字列の長さを返します。失敗した場合は FALSE を返します。
    '【処理】
    '  fputcsv() は、行（fields  配列として渡されたもの）を CSV としてフォーマットし、それを ファイルに書き込みます (いちばん最後に改行を追加します)。
    '=======================================================================
    Public function fputcsv(fields,delimiter,enclosure)

        fputcsv = false
        If len(delimiter) = 0 Then delimiter = ","
        If len(enclosure) = 0 Then enclosure = """"

        Dim key,replaced
        For key = 0 to uBound(fields)
            replaced = false
            If inStr(fields(key),delimiter) or inStr(fields(key),enclosure) or inStr(fields(key),vbCrLf) Then
                fields(key) = Replace(fields(key),enclosure,enclosure & enclosure)
                fields(key) = enclosure & fields(key) & enclosure
            End If
        Next

        Dim str : str = join(fields,delimiter)
        ts.WriteLine str
        fputcsv = len(str)
    end function

    '=======================================================================
    ' fwrite() のエイリアス
    '=======================================================================
    '【引数】
    '  str  =  string 書き込む文字列。
    '【説明】
    '  この関数は次の関数のエイリアスです。 fwrite().
    '=======================================================================
    Public function fputs(str)
        fputs = fwrite(str)
    end function

    '=======================================================================
    ' バイナリセーフなファイル書き込み処理
    '=======================================================================
    '【引数】
    '  str     =  string   書き込む文字列。
    '【戻り値】
    '  
    '【処理】
    '  string の内容を ファイル・ストリームに書き込みます。
    '=======================================================================
    Public function fwrite(str)
        ts.WriteLine str
    end function

    '=======================================================================
    ' ファイルがディレクトリかどうかを調べる
    '=======================================================================
    '【引数】
    '  filename = string    ファイルへのパス。filename  が相対パスの場合は、現在の作業ディレクトリからの相対パスとして処理します。
    '【戻り値】
    '  ファイルがが存在して、かつそれがディレクトリであれば TRUE、それ以外の場合は FALSE を返します。
    '【処理】
    '  指定したファイルがディレクトリかどうかを調べます。
    '=======================================================================
    Public Function is_dir(filename)

        Dim fn
        is_dir = false
        fn = fileMapPath(filename)

        If fso.FolderExists(fn) Then is_dir = true

    End Function

    '=======================================================================
    ' 通常ファイルかどうかを調べる
    '=======================================================================
    '【引数】
    '  filename = string    ファイルへのパス。
    '【戻り値】
    '  ファイルが存在し、かつそれが通常のファイルである場合に TRUE、 それ以外の場合に FALSE を返します。
    '【処理】
    '  指定したファイルが通常のファイルかどうかを調べます。
    '=======================================================================
    Public Function is_file(ByVal filename)

        is_file = false
        filename = fileMapPath(filename)
        If fso.FileExists(filename) Then is_file = true

    End Function

    '=======================================================================
    ' ディレクトリを作る
    '=======================================================================
    '【引数】
    '  pathname = string    ディレクトリのパス。
    '【戻り値】
    '  成功した場合に TRUE を、失敗した場合に FALSE を返します。
    '【処理】
    '  指定したディレクトリを作成します。
    '=======================================================================
    Public Function mkdir(ByVal pathname)

        mkdir = false
        pathname = fileMapPath(pathname)
        If not fso.FolderExists(pathname) Then
            mkdir = fso.CreateFolder(pathname)
        End If

    End Function

    '=======================================================================
    ' ファイルパスに関する情報を返す
    '=======================================================================
    '【引数】
    '  path     = string    調べたいパス。
    '  options  = string    どの要素を返すのかをオプションのパラメータ options  で指定します。これは PATHINFO_DIRNAME、 PATHINFO_BASENAME、 PATHINFO_EXTENSION および PATHINFO_FILENAME の組み合わせとなります。 デフォルトではすべての要素を返します。
    '【戻り値】
    '   以下の要素を含む連想配列を返します。 dirname (ディレクトリ名)、basename (ファイル名) そして、もし存在すれば extension (拡張子)。
    '   options を使用すると、 すべての要素を選択しない限りこの関数の返り値は文字列となります。 
    '【処理】
    '  pathinfo() は、path  に関する情報を有する連想配列を返します。
    '=======================================================================
    Public Function pathinfo(path,options)

        Dim obj : set obj = Server.CreateObject("Scripting.Dictionary")
        Dim tmp


        obj("dirname") = dirname(path)
        obj("basename") = basename(path,"")
        obj("extension") = obj("basename")

        If inStr(obj("basename"),".") Then
            tmp = Split(obj("basename"),".")
            obj("extension") = tmp( uBound(tmp) )
        End if

        obj("filename") = Replace(obj("basename"),"." & obj("extension"),"")

        If len(options) > 0 Then

            If options = PATHINFO_DIRNAME Then
                pathinfo = obj("dirname")
            ElseIf options = PATHINFO_BASENAME Then
                pathinfo = obj("basename")
            ElseIf options = PATHINFO_EXTENSION Then
                pathinfo = obj("extension")
            ElseIf options = PATHINFO_FILENAME Then
                pathinfo = obj("filename")
            End if
            Exit Function
        End If

        set pathinfo = obj
    End Function

    '=======================================================================
    ' ディレクトリを削除する
    '=======================================================================
    '【引数】
    '  dirname     = string    ディレクトリへのパス。
    '【戻り値】
    '   成功した場合に TRUE を、失敗した場合に FALSE を返します。
    '【処理】
    '  dirname で指定されたディレクトリを 削除しようと試みます。
    '  ディレクトリは空でなくてはならず、また 適切なパーミッションが設定されていなければなりません。
    '=======================================================================
    Public Function rmdir(ByVal dirname)

        dirname = fileMapPath(dirname)
        fso.DeleteFolder dirname
        rmdir = true

    End Function
    
    '=======================================================================
    ' ファイルを削除する
    '=======================================================================
    '【引数】
    '  filename     = string    ファイルへのパス。
    '【戻り値】
    '   成功した場合に TRUE を、失敗した場合に FALSE を返します。
    '【処理】
    '  filename  を削除します。 Unix C 言語の関数 unlink() と動作は同じです。
    '=======================================================================
    Public Function unlink(ByVal filename)

        filename = fileMapPath(filename)
        fso.DeleteFile filename
        unlink = true

    End Function

End Class

'=======================================================================
' パス中のファイル名の部分を返す
'=======================================================================
'【引数】
'  path      = string   パス。
'  suffix    = string   ファイル名が、 suffix  で終了する場合、 この部分もカットされます。
'【戻り値】
'  指定した path  のベース名を返します。
'【処理】
'  ・この関数は、ファイルへのパスを有する文字列を引数とし、 ファイルのベース名を返します。
'=======================================================================
Function basename(path, suffix)

    Dim b
    b = preg_replace("/^.*[¥/¥¥]/","",path,"","")

    If len(suffix) > 0 Then
        If Right(b,len(suffix)) = suffix Then
            b = Left(b,len(b) - len(suffix))
        End If
    End If

    basename = b

End Function

'=======================================================================
' ファイルをコピーする
'=======================================================================
'【引数】
'  source  = string   コピー元ファイルへのパス。
'  dest    = string   コピー先のパス。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・ ファイル source  を dest  にコピーします。
'  ・ ファイルを移動したいならは、rename() 関数を使用してください。 
'=======================================================================
Function copy(source, dest)

    copy = false
    Dim fileObj : set fileObj = new File_System
    copy = fileObj.copy(source,dest)
    set fileObj = nothing

End Function

'=======================================================================
' パス中のディレクトリ名の部分を返す
'=======================================================================
'【引数】
'  path  = string   パス。
'【戻り値】
'  ディレクトリの名前を返します。 path  にスラッシュが無い場合は、 カレントディレクトリを示すドット ('.') を返します。それ以外の場合は、スラッシュ以降の /component 部分を取り除いた path  を返します。
'【処理】
'  ・ この関数は、ファイルへのパス名を有する文字列を引数とし、 ディレクトリの名前を返します。
'=======================================================================
Function dirname(path)

    Dim d
    d = preg_replace("/¥¥/","/",path,"","")

    If inStr(d,"/") > 0 Then
        d = preg_replace("//[^/]*/?$/","",d,"","")
    Else
        d = "."
    End If
    dirname = d

End Function

'=======================================================================
' ファイルまたはディレクトリが存在するかどうか調べる
'=======================================================================
'【引数】
'  path      = string   ファイルあるいはディレクトリへのパス。
'【戻り値】
'  ファイルあるいはディレクトリが存在するかどうかを調べます。
'【処理】
'  ファイルあるいはディレクトリが存在するかどうかを調べます。
'=======================================================================
Function file_exists(ByVal filename)

    Dim fileObj : set fileObj = new File_System
    file_exists = fileObj.file_exists(filename)
    set fileObj = nothing

End Function

'=======================================================================
' ファイルの内容を全て文字列に読み込む
'=======================================================================
'【引数】
'  filename  = string データを読み込みたいファイルの名前。
'【戻り値】
'  読み込んだデータを返します。失敗した場合は FALSE を返します。
'【処理】
'  ファイルの内容を文字列に読み込む
'=======================================================================
Function file_get_contents(filename)

    file_get_contents = false
    Dim fileObj : set fileObj = new File_System
    file_get_contents = fileObj.file_get_contents(filename)
    set fileObj = nothing

End Function

'=======================================================================
' ファイル全体を読み込んで配列に格納する
'=======================================================================
'【引数】
'  filename  = string ファイルへのパス。
'  flags     = int    オプションのパラメータ flags  は、以下の定数のうちのひとつ、あるいは複数の組み合わせとなります。
'【戻り値】
'  ファイルを配列に入れて返します。 配列の各要素はファイルの各行に対応します。
'  改行記号はついたままとなります。 失敗すると file() は FALSE を返します。
'【処理】
'  ファイル全体を配列に読み込みます。
'=======================================================================
Function file(filename,flags)

    Dim req,tmp,key
    req = file_get_contents(filename)

    If flags = FILE_SKIP_EMPTY_LINES Then
        var_dump req
        tmp = preg_replace("/^" & vbCrLf & "/is","",req,"","")
        tmp = Split(tmp,vbCrLf)
    Else
        tmp = Split(req,vbCrLf)

        If flags = FILE_IGNORE_NEW_LINES Then
            For key = 0 to uBound(tmp)
                tmp(key) = tmp(key) & vbCrLf
            Next
        End If
    End If

    file = tmp

End Function

'=======================================================================
' ファイルパスに関する情報を返す
'=======================================================================
'【引数】
'  path     = string    調べたいパス。
'  options  = string    どの要素を返すのかをオプションのパラメータ options  で指定します。これは PATHINFO_DIRNAME、 PATHINFO_BASENAME、 PATHINFO_EXTENSION および PATHINFO_FILENAME の組み合わせとなります。 デフォルトではすべての要素を返します。
'【戻り値】
'   以下の要素を含む連想配列を返します。 dirname (ディレクトリ名)、basename (ファイル名) そして、もし存在すれば extension (拡張子)。
'   options を使用すると、 すべての要素を選択しない限りこの関数の返り値は文字列となります。 
'【処理】
'  pathinfo() は、path  に関する情報を有する連想配列を返します。
'=======================================================================
Function pathinfo(ByVal path,ByVal options)

    Dim fileObj : set fileObj = new File_System
    set pathinfo = fileObj.pathinfo(path,options)
    set fileObj = nothing

End Function


'=======================================================================
'パラメータの配列を指定してユーザ関数をコールする
'=======================================================================
'【引数】
'  callback     = mixed  コールする関数。このパラメータに array($classname, $methodname) を指定することにより、 クラスメソッドも静的にコールすることができます。
'  param_arr    = array  関数に渡すパラメータを指定する配列。
'【戻り値】
'  関数の結果、あるいはエラー時に FALSE を返します。
'【処理】
'  ・param_arr  にパラメータを指定して、 function  で指定したユーザ定義関数をコールします。
'=======================================================================
Function call_user_func_array(callback,param_arr)

    Dim thisFunc,thisParam,retval
    If isArray(callback) Then
        thisFunc  = callback(0) & "." & callback(1)
    Else
        thisFunc = callback
    End If

    If isArray(param_arr) Then
        Dim key
        For Each key In parameter
            If len( thisParam ) > 0 Then
                thisParam = thisParam & "," & key
            Else
                thisParam = key
            End IF
        Next
    Else
        thisParam = param_arr
    End If
    execute("retval = " & thisFunc & "(" & thisParam & ")")
    call_user_func_array = retval

End Function

'=======================================================================
'最初の引数で指定したユーザ関数をコールする
'=======================================================================
'【引数】
'  callback     = mixed  コールする関数。このパラメータに array(classname, methodname) を指定することにより、 クラスメソッドも静的にコールすることができます。
'  parameter    = mixed  この関数に渡す、ゼロ個以上のパラメータ。
'【戻り値】
'  関数の結果、あるいはエラー時に FALSE を返します。
'【処理】
'  ・パラメータ callback で指定した ユーザ定義のコールバック関数をコールします。
'=======================================================================
Function call_user_func(callback,parameter)

    Dim thisFunc,retval
    If isArray(callback) Then
        thisFunc  = callback(0) & "." & callback(1)
    Else
        thisFunc = callback
    End If

    execute("retval = " & thisFunc & "(parameter)")
    call_user_func = retval
End Function



'PHPの言語構造を関数として実装

'=======================================================================
'a に b を代入
'=======================================================================
'【引数】
'  a    = mixed  値
'  b    = mixed  代入する値
'【戻り値】
'  値を返しません。
'【処理】
'  ・左辺に右辺を代入します。
'=======================================================================
Function [=](ByRef a, ByVal b)

    If isObject(b) Then
        set a = b
    Else
        a = b
    End if

End Function

'=======================================================================
'a が b に等しい時に TRUE。
'=======================================================================
'【引数】
'  a    = mixed  値
'  b    = mixed  比較する値
'【戻り値】
'  aとbが等しい場合にTRUE を、等しくない場合に FALSE を返します。
'【処理】
'  ・左辺と右辺を比較します。型は厳密にチェックしません。
'=======================================================================
Function [==](ByVal a, ByVal b)

    [==] = false

    Dim tmp_a,tmp_b
    Dim key

    If (isArray(a) or isArray(b)) or (isObject(a) or isObject(b)) Then

        If isObject(a) and isObject(b) Then
            If a.count <> b.count Then Exit Function

            tmp_a = a.keys : tmp_b = b.keys
            If Not [==](tmp_a,tmp_b) Then Exit Function

            tmp_a = a.Items : tmp_b = b.Items
            If Not [==](tmp_a,tmp_b) Then Exit Function
            [==] = true
        End If

        If isArray(a) and isArray(b) Then
            If uBound(a) <> uBound(b) Then Exit Function

            For key = 0 to uBound(a)
                If Not [==](a(key),b(key) ) Then Exit Function
            Next

            [==] = true
        End If

    Else
        If isNull(a) Then a = ""
        If isNull(b) Then b = ""

        [==] = (Cstr(a) = Cstr(b))
    End If

End Function

'=======================================================================
'a が b に等しい時に TRUE。
'=======================================================================
'【引数】
'  a    = mixed  値
'  b    = mixed  比較する値
'【戻り値】
'  aとbが等しい場合にTRUE を、等しくない場合に FALSE を返します。
'【処理】
'  ・左辺と右辺を比較します。型は厳密にチェックします。
'=======================================================================
Function [===](a, b)

    [===] = false

    Dim tmp_a,tmp_b
    Dim key
    If (isArray(a) or isArray(b)) or (isObject(a) or isObject(b)) Then

        If isObject(a) and isObject(b) Then
            If a.count <> b.count Then Exit Function

            tmp_a = a.keys : tmp_b = b.keys
            If Not [===](tmp_a,tmp_b) Then Exit Function

            tmp_a = a.Items : tmp_b = b.Items
            If Not [===](tmp_a,tmp_b) Then Exit Function
            [===] = true
        End If

        If isArray(a) and isArray(b) Then
            If uBound(a) <> uBound(b) Then Exit Function

            For key = 0 to uBound(a)
                If Not [===](a(key),b(key) ) Then Exit Function
            Next

            [===] = true
        End If

    Else
        [===] = eval(a = b and vartype(a) = vartype(b))
    End If

End Function

'=======================================================================
'三項演算子
'=======================================================================
'【引数】
'  i    = mixed  式1
'  j    = mixed  式2
'  k    = mixed  式3
'【戻り値】
'  式1 が TRUE の場合に 式2 を、 式1 が FALSE の場合に 式3 を値とします。
'【処理】
'  ・式1 が TRUE の場合に 式2 を、 式1 が FALSE の場合に 式3 を値とします。
'=======================================================================
Function [?](i,j,k)
	If Not is_empty(i) Then [?] = j Else [?] = k
End Function

'=======================================================================
'a が b に等しくない時に TRUE。
'=======================================================================
'【引数】
'  a    = mixed  値
'  b    = mixed  比較する値
'【戻り値】
'  aとbが等しくない場合にTRUE を、等しくない場合に FALSE を返します。
'【処理】
'  ・左辺と右辺を比較します。型は厳密にチェックしません。
'=======================================================================
Function [!=](a, b)

    If [==](a,b) Then
        [!=] = false
    Else
        [!=] = true
    End If

End Function

'=======================================================================
'a が b に等しくない時に TRUE。
'=======================================================================
'【引数】
'  a    = mixed  値
'  b    = mixed  比較する値
'【戻り値】
'  aとbが等しくない場合にTRUE を、等しくない場合に FALSE を返します。
'【処理】
'  ・左辺と右辺を比較します。型は厳密にチェックします。
'=======================================================================
Function [!==](a, b)

    If [===](a,b) Then
        [!==] = false
    Else
        [!==] = true
    End If

End Function

'=======================================================================
'a が b より多い TRUE。
'=======================================================================
'【引数】
'  a    = mixed  値
'  b    = mixed  比較する値
'【戻り値】
'  aがbより多い場合にTRUE を、多くない場合に FALSE を返します。
'【処理】
'  ・左辺と右辺を比較します。
'=======================================================================
Function [>](a, b)

    [>] = false

    Dim tmp_a,tmp_b
    Dim key

    If (isArray(a) or isArray(b)) or (isObject(a) or isObject(b)) Then

        If isObject(a) and isObject(b) Then
            If a.count < b.count Then Exit Function

            tmp_a = a.keys : tmp_b = b.keys
            If Not [>](uBound(tmp_a),uBound(tmp_b)) Then Exit Function
            [>] = true
        End If

        If isObject(a) and not isObject(b) Then
            [>] = true
        End If

        If isArray(a) and isArray(b) Then
            If uBound(a) < uBound(b) Then Exit Function

            For key = 0 to uBound(a)
                If Not [>](a(key),b(key) ) Then Exit Function
            Next

            [>] = true
        End If

        If isArray(a) and not isArray(b) Then
            [>] = true
        End If
    Else
        [>] = eval(a > b)
    End If

End Function


'rand
Const RAND_MAX = 32768

'=======================================================================
' 乱数を生成する
'=======================================================================
'【引数】
'  min = int   返す値の最小値 (デフォルトは 0)。
'  max = int   返す値の最大値 (デフォルトは RAND_MAX)。
'【戻り値】
'  min  (あるいは 0) から max  (あるいは RAND_MAX、それぞれ端点を含む) までの間の疑似乱数値を返します。
'【処理】
'  ・オプションの引数 min ,max  を省略してコールした場合、rand() は 0 と RAND_MAX の間の擬似乱数(整数)を返します。
'  ・例えば、5 から 15 まで（両端を含む）の乱数を得たい場合、 rand(5,15) とします。
'=======================================================================
Function rand(min,max)

    If len(min) = 0 Then min = 0
    If len(max) = 0 Then max = RAND_MAX

    Randomize
    rand = intval( Rnd * (max - min + 1)) + min

End Function

'=======================================================================
' 値が数値でないかどうかを判定する
'=======================================================================
'【引数】
'  val = float   調べる値。
'【戻り値】
'  val  が '非数値 (not a number)' の場合に TRUE、そうでない場合に FALSE を返します。
'【処理】
'  ・val  が '非数値 (not a number)' であるかどうかを調べます。たとえば acos(1.01) の結果などがこれにあたります。
'=======================================================================
Function is_nan(val)
    is_nan = not isNumeric(val)
End Function

'=======================================================================
' 指数表現
'=======================================================================
'【引数】
'  base = number    使用する基数。
'  exp  = number    指数。
'【戻り値】
'  base  の exp  乗を 返します。可能な場合、この関数は、vbDouble 型の値を 返します。累乗が計算できない場合は FALSE を返します。
'【処理】
'  ・base  の exp  乗を返します。
'=======================================================================
Function pow(base,exp)
    pow = base ^ exp
End Function

'=======================================================================
' 平方根
'=======================================================================
'【引数】
'  arg = float  処理する引数。
'【戻り値】
'  arg  の平方根を返します。 負の数を指定した場合は、null を返します。
'【処理】
'  ・arg  の平方根を返します。
'=======================================================================
Function sqrt(arg)

    If arg < 0 Then
        sqrt = null
    Else
        sqrt = sqr(arg)
    End If

End Function


'=======================================================================
' 文字エンコーディングを変換する
'=======================================================================
'【引数】
'  str          = string    変換する文字列。
'  to_encoding  = string    str  の変換後の文字エンコーディング。
'  from_encoding= string    変換前の文字エンコーディング名を指定します。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・文字列 str の文字エンコーディングを、 オプションで指定した from_encoding  から to_encoding  に変換します。
'=======================================================================
Function mb_convert_encoding(str,to_encoding,from_encoding)

    Dim bobj : set bobj = Server.CreateObject("basp21")
    mb_convert_encoding = bobj.Kconv(str,_
                          mb_convert_encoding_helper(to_encoding),_
                          mb_convert_encoding_helper(from_encoding))
End Function

'*******************************************
Function mb_convert_encoding_helper(encoding)

    Dim tmp
    Select Case lcase(encoding)
    Case "shift_jis","sjis"
        tmp = 1
    Case "euc","euc-jp"
        tmp = 2
    Case "jis"
        tmp = 3
    Case "ucs2"
        tmp = 4
    Case "utf-8","utf8"
        tmp = 5
    Case "auto"
        tmp = 0
    End Select

End Function

'BASP21を使用しない場合
'http://www.ac.cyberhome.ne.jp/‾mattn/AcrobatASP/1.html
'http://www.ac.cyberhome.ne.jp/‾mattn/AcrobatASP/StrConv.inc
'=======================================================================
' カナを("全角かな"、"半角かな"等に)変換する
'=======================================================================
'【引数】
'  str      = string  文字列
'  option   = mixed   変換オプション
'【戻り値】
'  変換された文字列を返します。
'【処理】
'  ・文字列 str  に関して「半角」-「全角」変換を行います。
'=======================================================================
Function mb_convert_kana( str , option_name )

    Dim outstr
    Dim J
    Dim tmp
    Dim bobj

    If Len( str ) = 0 Then Exit Function

    Set bobj = Server.CreateObject("basp21")

    If Len( option_name ) = 0 Then
        option_name = "K"
    End If

    outstr = str

    For J = 1 To Len( option_name )

        tmp = mid(option_name,J,1)

        If tmp = "r" Then
            '全角文字 (2 バイト) を半角文字 (1 バイト) に変換
            outstr = bobj.StrConv(outstr,8)

        ElseIf tmp = "R" Then
            '半角文字 (1 バイト) を全角文字 (2 バイト) に変換
            outstr = bobj.StrConv(outstr,4)

        ElseIf tmp = "c" Then
            'カタカナをひらがなに変換
            outstr = bobj.StrConv(outstr,32)

        ElseIf tmp = "C" Then
            'ひらがなをカタカナに変換
            outstr = bobj.StrConv(outstr,16)

        ElseIf tmp = "K" or tmp = "V" THen
            '文字列の中の半角カナを全角カナに変換します。濁音にも対応。
            outstr = bobj.HAN2ZEN(outstr)

        End If
    Next

    mb_convert_kana = outstr

End Function

'=======================================================================
' メールを送信する
'=======================================================================
'【引数】
'  to       = string    文字列
'  subject  = string    件名
'  message  = string    本文
'  file     = string    添付ファイル
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・email を送信します。
'=======================================================================
function mb_send_mail( strto, subject, message, from, strFile )

    Dim basp
    Dim msg

    if inStr(subject,vbCrLf) then
        Response.Write "題名に改行を含めることはできません。"
        Response.End
    end if

	Set basp = Server.CreateObject("basp21")
	msg = basp.SendMail(SMTP_SERVER, strto, from, subject, message, strFile)

    mb_send_mail = msg

end function


'=======================================================================
'メッセージを出力し、現在のスクリプトを終了する
'=======================================================================
'【引数】
'  val   = 文字列
'【戻り値】
'  
'【処理】
'  ・文字列出力後、スクリプト停止
'=======================================================================
Sub die(val)

    If Len(val) > 0 Then
        Response.Write val
    End If

    Response.End

End Sub

'=======================================================================
'文字列を ASP コードとして評価する
'=======================================================================
'【引数】
'  code_str = string   文字列
'【戻り値】
'  NULL を返します。
'【処理】
'  ・code_str  で与えられた文字列を PHP コードとして評価します。 
'  ・中でも、データベースのテキストフィールドにコードを保存し、 後で実行するためには便利です。
'=======================================================================
sub eval(code_str)
    execute code_str
end sub


'=======================================================================
' (IPv4) インターネットプロトコルドット表記のアドレスを、適当なアドレスを有する文字列に変換する
'=======================================================================
'【引数】
'  ip_address = string   標準形式のアドレス。
'【戻り値】
'  IPv4 アドレス、あるいは ip_address  が不正な形式の場合に FALSE を返します。
'【処理】
'  ・関数 ip2long() は、インターネット標準形式 (ドット表記の文字列) 表現から IPv4 インターネットネットアドレスを生成します。
'=======================================================================
Function ip2long( ip_address )

    ip2long = false

    If preg_match("/^¥d{1,3}¥.¥d{1,3}¥.¥d{1,3}¥.¥d{1,3}$/",ip_address,"","","") Then
        Dim parts
        parts = Split(ip_address,".")
        ip2long = parts(0) * pow(256,3) + _
                  parts(1) * pow(256,2) + _
                  parts(2) * pow(256,1) + _
                  parts(3) * pow(256,0)
    End If

End Function

'=======================================================================
' (IPv4) インターネットアドレスをインターネット標準ドット表記に変換する
'=======================================================================
'【引数】
'  proper_address = string   正しい形式のアドレス。
'【戻り値】
'  インターネットの IP アドレスを表す文字列を返します。
'【処理】
'  ・関数long2ip() は、適切なアドレス表現からドット表記 (例:aaa.bbb.ccc.ddd)のインターネットアドレスを生成します。
'=======================================================================
Function long2ip( proper_address )

    long2ip = false

    If isNumeric(proper_address) Then
        If proper_address >= 0 and proper_address <= 4294967295 Then

            long2ip = ( proper_address / pow(256,3) ) & "." & _
                      ( (proper_address Mod pow(256,3)) / pow(256,2) ) & "." & _
                      ( ((proper_address Mod pow(256,3)) / pow(256,2)) / pow(256,1) ) & "." & _
                      ( (((proper_address Mod pow(256,3)) / pow(256,2)) / pow(256,1)) / pow(256,0) )

        End If
    End If
End Function

Sub aspinfo
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "DTD/xhtml1-transitional.dtd">
<html><head>
<style type="text/css"><!--
body {background-color: #ffffff; color: #000000;}
body, td, th, h1, h2 {font-family: sans-serif;}
pre {margin: 0px; font-family: monospace;}
a:link {color: #000099; text-decoration: none; background-color: #ffffff;}
a:hover {text-decoration: underline;}
table {border-collapse: collapse;}
.center {text-align: center;}
.center table { margin-left: auto; margin-right: auto; text-align: left;}
.center th { text-align: center !important; }
td, th { border: 1px solid #000000; font-size: 75%; vertical-align: baseline;}
h1 {font-size: 150%;}
h2 {font-size: 125%;}
.p {text-align: left;}
.e {background-color: #ccccff; font-weight: bold; color: #000000;}
.h {background-color: #9999cc; font-weight: bold; color: #000000;}
.v {background-color: #cccccc; color: #000000;}
i {color: #666666; background-color: #cccccc;}
img {float: right; border: 0px;}
hr {width: 600px; background-color: #cccccc; border: 0px; height: 1px; color: #000000;}
//--></style>
<title>aspinfo()</title></head>
<body><div class="center">
<table border="0" cellpadding="3" width="600">
<tr class="h"><td>
<h1 class="p">ASP</h1>
</td></tr>
</table><br />
<h2>Request</h2>
<table border="0" cellpadding="3" width="600">
<tr class="h"><th colspan="2">ServerVariables</th></tr>
<%
Dim key
For Each key In Request.ServerVariables
    Response.Write "<tr>" & vbCrlf
    Response.Write "<td class=""e"">" & key & "</td>" & vbCrLf
    Response.Write "<td class=""v"">" & Request.ServerVariables(key) & "</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
Next
%>
<tr class="h"><th colspan="2">ClientCertificate</th></tr>
<%
For Each key In Request.ClientCertificate
    Response.Write "<tr>" & vbCrlf
    Response.Write "<td class=""e"">" & key & "</td>" & vbCrLf
    Response.Write "<td class=""v"">" & Request.ClientCertificate(key) & "</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
Next
%>

</table><br />
<h2>Application</h2>
<table border="0" cellpadding="3" width="600">
<tr class="h"><th colspan="2">Contents</th></tr>
<%
For Each key In Application.Contents
    Response.Write "<tr>" & vbCrlf
    Response.Write "<td class=""e"">" & key & "</td>" & vbCrLf
    Response.Write "<td class=""v"">" & Application.Contents(key) & "</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
Next
%>
</table><br />
</div></body></html>
<%
End Sub


'preg_grep
Const PREG_GREP_INVERT    = 1

'preg_match
Const PREG_PATTERN_ORDER  = 1
Const PREG_SET_ORDER      = 2
Const PREG_OFFSET_CAPTURE = 256

'preg_split
Const PREG_SPLIT_NO_EMPTY       = 1
Const PREG_SPLIT_DELIM_CAPTURE  = 2
Const PREG_SPLIT_OFFSET_CAPTURE = 4

'=======================================================================
'パターンにマッチする配列の要素を返す
'=======================================================================
'【引数】
'  pattern  = string    検索するパターンを表す文字列。
'  input    = array     入力の配列。
'  flags    = array     PREG_GREP_INVERT  を設定すると、この関数は 与えた pattern  にマッチ しない  要素を返します。
'【戻り値】
'  input  配列のキーを使用した配列を返します。
'【処理】
'  ・ input  配列の要素のうち、 指定した pattern  にマッチするものを要素とする配列を返します。 
'=======================================================================
Function preg_grep(pattern, input, flags)

    Dim obj
    set obj = Server.CreateObject("Scripting.Dictionary")

    If not isArray(input) and not isObject(input) Then
        set preg_grep =  obj
        Exit Function
    End If

    Dim key
    If isArray(input) Then
        For key = 0 to uBound(input)
            If flags = PREG_GREP_INVERT Then
                If Not preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            Else
                If preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            End If
        Next
    ElseIf isObject(input) Then
        For Each key In input
            If flags = PREG_GREP_INVERT Then
                If Not preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            Else
                If preg_match(pattern, input(key),"","","") Then
                    obj.Add key, input(key)
                End If
            End If
        Next
    End If

    set preg_grep =  obj

End Function

'=======================================================================
'繰り返し正規表現検索を行う
'=======================================================================
'【引数】
'  pattern  = string    検索するパターンを表す文字列。
'  subject  = string    入力文字列。
'  matches  = array     matches  を指定した場合、検索結果が代入されます。matches(0) にはパターン全体にマッチしたテキストが代入され、 matches(1)には 1 番目ののキャプチャ用サブパターンにマッチした 文字列が代入され、といったようになります。
'  flags    = int       戻り値の形式を指定
'  offset   = int       通常、検索は対象文字列の先頭から開始されます。 オプションのパラメータ offset  を使用して 検索の開始位置を指定することも可能です。
'【戻り値】
'  パターンがマッチした総数を返します（ゼロとなる可能性もあります）。 
'  または、エラーが発生した場合に FALSE を返します。
'【処理】
'  ・ subject  を検索し、 pattern  に指定した正規表現にマッチした すべての文字列を、flags  で指定した 順番で、matches  に代入します。
'  ・ 正規表現にマッチすると、そのマッチした文字列の後から 検索が続行されます。 
'=======================================================================
Function preg_match_all(pattern, ByVal subject, ByRef matches, flags, offset)

    Dim regEx,matchall,matchone
    Dim cnt,counter : counter = 0
    Dim helper

    preg_match_all = false
    If vartype(matches) <> 0 Then Exit Function
    If len(flags) = 0 Then flags = PREG_PATTERN_ORDER

    set helper = new RegExp_Helper
    helper.parseOption(pattern)

    Set regEx = new RegExp
    With regEx
        .IgnoreCase = helper.withIgnoreCase
        .Global     = True
        .pattern    = helper.withPattern
        .MultiLine  = helper.withMultiLine
    End With
    set helper = Nothing

    If len(offset) > 0 Then
        offset = int(offset)
        subject = Mid(subject,offset)
    End If

    Set matchall = regEx.execute(subject)
    Set regEx = Nothing
    If matchall.count = 0 Then exit Function

    If flags = PREG_PATTERN_ORDER Then

        ReDim matches(matchall(0).SubMatches.count)

        For cnt = 0 to uBound(matches)
            toReDim matches(cnt),(matchall.count-1)
        Next

        counter = 0
        For Each matchone In matchall
            matches(0)(counter) = matchone.value
            For cnt = 1 to matchone.SubMatches.count
                matches(cnt)(counter) = matchone.SubMatches(cnt-1)
            Next
            counter = counter + 1
        Next

    Elseif flags = PREG_SET_ORDER Then

        ReDim matches(matchall.count-1)

        counter = 0
        For Each matchone In matchall
            toReDim matches(counter),(matchone.SubMatches.count)
            matches(counter)(0) = matchone.value
            For cnt = 1 to matchone.SubMatches.count
                matches(counter)(cnt) = matchone.SubMatches(cnt-1)
            Next
            counter = counter + 1
        Next

    ElseIf PREG_OFFSET_CAPTURE Then

        ReDim matches(matchall(0).SubMatches.count)

        For cnt = 0 to uBound(matches)
            toReDim matches(cnt),(matchall.count-1)
        Next

        counter = 0
        For Each matchone In matchall
            toReDim matches(0)(counter),1
            matches(0)(counter)(0) = matchone.value
            matches(0)(counter)(1) = matchone.FirstIndex
            For cnt = 1 to matchone.SubMatches.count
                toReDim matches(cnt)(counter),1
                matches(cnt)(counter)(0) = matchone.SubMatches(cnt-1)
                matches(cnt)(counter)(1) = InStr( matchone.value, matchone.SubMatches(cnt-1) ) -1
            Next
            counter = counter + 1
        Next

    End If

    preg_match_all = matchall.Count

End Function

'=======================================================================
'正規表現によるマッチングを行う
'=======================================================================
'【引数】
'  pattern  = string    検索するパターンを表す文字列。
'  subject  = string    入力文字列。
'  matches  = array     matches  を指定した場合、検索結果が代入されます。matches(0) にはパターン全体にマッチしたテキストが代入され、 matches(1)には 1 番目ののキャプチャ用サブパターンにマッチした 文字列が代入され、といったようになります。
'  flags    = int       PREG_OFFSET_CAPTURE   このフラグを設定した場合、各マッチに対応する文字列のオフセットも返されます。 これにより、返り値は配列となり、配列の要素 0 はマッチした文字列、 要素 1は対象文字列中におけるマッチした文字列のオフセット値 となることに注意してください。
'  offset   = int       通常、検索は対象文字列の先頭から開始されます。 オプションのパラメータ offset  を使用して 検索の開始位置を指定することも可能です。
'【戻り値】
'  preg_match() は、pattern  がマッチした回数を返します。
'  つまり、0 回（マッチせず）または 1 回となります。
'  これは、最初にマッチした時点でpreg_match()  は検索を止めるためです。
'【処理】
'  ・pattern  で指定した正規表現により subject  を検索します。
'=======================================================================
Function preg_match(pattern, ByVal subject, ByRef matches, flags, offset)

    Dim matchAll,matchone
    Dim cnt,helper

    preg_match = false

    set helper = new RegExp_Helper
    helper.parseOption(pattern)

    Set matchAll = new RegExp
    With matchAll
        .IgnoreCase = helper.withIgnoreCase
        .Global     = false
        .pattern    = helper.withPattern
        .MultiLine  = helper.withMultiLine
    End With

    set helper = Nothing

    If not is_empty(offset) Then
        offset = int(offset)
        subject = Mid(subject,offset)
    End If

    offset = intval( offset )

    If vartype(matches) <> 8 Then
        Set matchone = matchAll.execute(subject)
        Set matchAll = Nothing
        If matchone.count = 0 Then exit Function

        If flags = PREG_OFFSET_CAPTURE Then

            ReDim matches(1)
            matches(0) = matchone(0).value
            matches(1) = offset + matchone(0).FirstIndex
        Else
            ReDim matches(0)
            matches(0) = matchone(0).value
        End If

        preg_match = true
    Else
        preg_match = matchAll.Test(subject)
        Set matchAll = Nothing
    End If

End Function

'=======================================================================
'正規表現文字をクオートする
'=======================================================================
'【引数】
'  str          = string    入力文字列。
'  delimiter    = string    オプションの delimiter  を指定すると、 ここで指定した文字もエスケープされます。これは、PCRE 関数が使用する デリミタをエスケープする場合に便利です。'/' がデリミタとしては 最も一般的に使用されています。
'【戻り値】
'  クォートされた文字列を返します。
'【処理】
'  ・ preg_quote() は、str  を引数とし、正規表現構文の特殊文字の前にバックスラッシュを挿入します。
'  ・ この関数は、実行時に生成される文字列をパターンとしてマッチングを行う必要があり、 その文字列には正規表現の特殊文字が含まれているかも知れない場合に有用です。
'  ・ 正規表現の特殊文字は、次のものです。 . ¥ + * ? [ ^ ] $ ( ) { } = ! < > | : 
'=======================================================================
Function preg_quote(ByVal str,delimiter)

    Dim pattern : pattern = array("¥",".","+","*","?","[","^","]","$","(",")","{","}","=","!","<",">","|",":")
    If len(delimiter) > 0 Then [] pattern , delimiter

    Dim key
    For key = 0 to uBound(pattern)
        str = Replace(str,pattern(key),"¥" & pattern(key))
    Next
    preg_quote = str

End Function

'=======================================================================
'正規表現検索および置換を行う
'=======================================================================
'【引数】
'  pattern      = mixed 検索を行うパターン。文字列もしくは配列とすることができます。
'  callback     = mixed このコールバック関数は、検索対象文字列でマッチした要素の配列が指定されて コールされます。このコールバック関数は、置換後の文字列を返す必要があります。
'  subject      = mixed 検索・置換対象となる文字列もしくは文字列の配列
'  limit        = int   subject  文字列において、各パターンによる 置換を行う最大回数。デフォルトは -1 (制限無し)。
'  cnt          = int   この引数が指定されると、置換回数が渡されます。
'【戻り値】
'  subject  引数が配列の場合は配列を、 その他の場合は文字列を返します。
'  パターンがマッチした場合、〔置換が行われた〕新しい subject  を返します。
'  マッチしなかった場合、subject  をそのまま返します。
'【処理】
'  ・subject  に関して pattern  を用いて検索を行い、 callback  に置換します。
'=======================================================================
Function preg_replace_callback(pattern,callback,ByVal subject,limit,ByRef cnt)

    Dim key,counter
    cnt = 0
    If len(limit) = 0 Then limit = 0

    If isArray(subject) Then
        For key = 0 to uBound(subject)
            subject(key) = preg_replace( pattern, callback, subject(key),limit,cnt)
        Next
    ElseIf isObject(subject) Then
        For Each key In subject
            subject(key) = preg_replace( pattern, callback, subject(key),limit,cnt)
        Next
    Else

        If isArray(pattern) Then
                For key = 0 to uBound(pattern)
                    subject = preg_replace( pattern(key), callback, _
                                   subject,limit,cnt)
                Next

        Else

            Dim matchAll,strCallback
            If is_empty(limit) Then
                If preg_match_all(pattern, subject, matchAll,PREG_PATTERN_ORDER,"") <> false Then
                    For Each key In matchAll(0)
                        execute("strCallback = " & callback & "(key)")
                        subject = Replace(subject,key,strCallback)
                    Next
                End If

            Else
                If preg_match_all(pattern, subject, matchAll,PREG_PATTERN_ORDER,"") <> false Then
                    For Each counter In matchAll(0)
                        cnt = cnt + 1
                        If cnt > limit Then Exit For
                        execute("strCallback = " & callback & "(counter)")
                        subject = Replace(subject,counter,strCallback)
                    Next
                End If
            End If

        End If
    End If

    preg_replace_callback = subject

End Function

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

'=======================================================================
'正規表現で文字列を分割する
'=======================================================================
'【引数】
'  pattern      = string 検索するパターンを表す文字列。
'  subject      = string 入力文字列。
'  limit        = int    これを指定した場合、最大 limit  個の部分文字列が返されます。
'  flags        = int    flags  は、フラグを組み合わせたものとする （ビット和演算子｜で組み合わせる）ことが可能です。
'【戻り値】
'  pattern  にマッチした境界で分割した subject  の部分文字列の配列を返します。
'【処理】
'  ・指定した文字列を、正規表現で分割します。
'=======================================================================
Function preg_split(pattern, subject,limit,flags)

    If is_empty(limit) Then limit = 0

    Dim key,matches,tmp_sp,tmp_str
    Dim cnt,counter,strMid,pointer : pointer = 1 : counter = 0
    Dim strRegExp,intPoint,strPoint

    cnt = preg_match_all(pattern,subject, matches, PREG_OFFSET_CAPTURE, "")
    If cnt > 0 Then
        For key = 0 to uBound(matches(0))

            counter = counter + 1
            If limit > 0 Then
                If counter >= limit Then Exit For
            End If

            intPoint  = matches(0)(key)(1)
            strPoint  = matches(0)(key)(0)
            strRegExp = Mid(subject, pointer,intPoint-pointer+1)

            Select Case flags
            Case PREG_SPLIT_NO_EMPTY
                if len(strRegExp) > 0 Then [] tmp_sp , strRegExp
            Case PREG_SPLIT_DELIM_CAPTURE
                [] tmp_sp , strRegExp
                [] tmp_sp , matches(1)(key)(0)
            Case PREG_SPLIT_OFFSET_CAPTURE
                [] tmp_sp , array(strRegExp,pointer-1)
            Case Else
                [] tmp_sp , strRegExp
            End Select
            pointer = intPoint + 1 + len(strPoint)
        Next

        strRegExp = Mid(subject, pointer)
        Select Case flags
        Case PREG_SPLIT_NO_EMPTY
            if len(strRegExp) > 0 Then [] tmp_sp , strRegExp
        Case PREG_SPLIT_DELIM_CAPTURE
            [] tmp_sp , strRegExp
        Case PREG_SPLIT_OFFSET_CAPTURE
            [] tmp_sp , array(strRegExp,pointer-1)
        Case Else
            [] tmp_sp , strRegExp
        End Select
    End If

    preg_split = tmp_sp

End Function

'*******************
Class Regexp_Helper

    Private strPattern
    Private boolIgnoreCase
    Private boolMultiLine

    'IgnoreCase = 大文字小文字を区別しないよう設定します。
    'Global     = 文字列全体を検索するよう設定します。
    'pattern    = 正規表現パターンを設定します。
    'MultiLine  = 文字列を複数行として扱わない。

    Private Sub Class_Initialize()
        'empty
    End Sub

    Private Sub Class_Terminate()
        'empty
    End Sub

    Private Property Let withPattern(str)
        strPattern = str
    End Property

    Public Property Get withPattern
        withPattern = strPattern
    End Property

    Private Property Let withIgnoreCase(bool)
        boolIgnoreCase = bool
    End Property

    Public Property Get withIgnoreCase
        withIgnoreCase = boolIgnoreCase
    End Property

    Private Property Let withMultiLine(bool)
        boolMultiLine = bool
    End Property

    Public Property Get withMultiLine
        withMultiLine =boolMultiLine
    End Property

    Public Function parseOption(str)

        If left(str,1) <> "/" Then Exit Function

        Dim tmp,options
        tmp = Split(str,"/")
        withPattern = tmp(1)

        If uBound( tmp ) > 2 Then
            Dim key
            For key = 2 to uBound( tmp ) -1
                withPattern = withPattern & "/" & tmp(key)
            Next
        End If

        withMultiLine = false
        withIgnoreCase = false

        options = tmp( uBound(tmp) )
        If inStr(options,"s") > 0 Then withMultiLine = true
        If inStr(options,"i") > 0 Then withIgnoreCase = true

    End Function

End Class


'htmlspecialchars
Const ENT_NOQUOTES = 0
Const ENT_COMPAT   = 2
Const ENT_QUOTES   = 3

'str_pad
Const STR_PAD_LEFT  = 0
Const STR_PAD_RIGHT = 1
Const STR_PAD_BOTH  = 2

'=======================================================================
' バイナリデータを文字列に変換
'=======================================================================
'【引数】
'  str  = string    文字列に変換したいバイナリデータ
'【戻り値】
'  文字列を返します。
'【処理】
'  ・日本語対応
'=======================================================================
Function Bin2Str(byteData)

    Dim i,u,s,intChar : i = 1

    Bin2Str =""
    Do While i <= LenB(byteData)
        u = Hex(AscB(MidB(byteData, i, 1)))
        If ((CInt("&H" & u) >= &H81) And (CInt("&H" & u) <= &H9F)) _
            Or ((CInt("&H" & u) >= &HE0) And (CInt("&H" & u) <= &HFC)) Then 'Code Page 932
            l = Hex(AscB(MidB(byteData, i + 1, 1)))
            intChar = CInt("&H" & u & l)
            s = Chr(intChar)
            i = i + 2
        Else
            intChar = CInt("&H" & u)
            s = Chr(intChar)
            i = i + 1
        End If
        Bin2Str = Bin2Str & s
    Loop
End Function

'=======================================================================
' 文字列をバイナリ変換
'=======================================================================
'【引数】
'  str  = string    バイナリ変換したい文字列。
'【戻り値】
'  バイナリ変換された文字列を返します。
'【処理】
'  ・日本語対応
'=======================================================================
Function Str2Bin(strData)

    Dim strChar,strHex
    For i = 1 To Len(strData)
        strChar = Mid(strData, i, 1)
        strHex = CStr(Hex(Asc(strChar)))
        Select Case Len(strHex)
            Case 1 '1Byte
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 1, 1)))
            Case 2 '1Byte
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 1, 2)))
            Case 4 '2Byte
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 1, 2)))
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 3, 2)))
        End Select
    Next
End Function

'=======================================================================
' 文字列をスラッシュでクォートする
'=======================================================================
'【引数】
'  str  = string    エスケープしたい文字列。
'【戻り値】
'  エスケープされた文字列を返します。
'【処理】
'  ・データベースへの問い合わせなどに際してクォートされるべき文字の前に バックスラッシュを挿入した文字列を返します。 
'=======================================================================
Function addslashes(ByVal str)

    If isNull(str) Then
        str = ""
    End If

    str = Replace(str,"¥","¥¥")
    str = Replace(str,"""","¥""")
    str = Replace(str,"'","¥'")

    addslashes = str

End Function

'=======================================================================
' rtrim() のエイリアス
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'【処理】
'  ・この関数は次の関数のエイリアスです。 rtrim().
'=======================================================================
Function chop(str)
    chop = RTrim(str)
End Function

'=======================================================================
' 1 つ以上の文字列を出力する
'=======================================================================
'【引数】
'  str  = string    出力したいパラメータ。
'【戻り値】
'  値を返しません。
'【処理】
'  ・すべてのパラメータを出力します。
'=======================================================================
Sub echo(str)

    If isObject(str) then
        Response.Write "Object"
    ElseIf IsArray(str) then
        Response.Write "Array"
    Else
        Response.Write str
    End if

End Sub

'=======================================================================
' 文字列を文字列により分割する
'=======================================================================
'【引数】
'  delimiter    = string    区切り文字列。
'  string       = string    入力文字列。
'  limit        = string    limit  が指定された場合、返される配列には 最大 limit  の要素が含まれ、その最後の要素には string  の残りの部分が全て含まれます。
'【戻り値】
'  空の文字列 ("") が delimiter  として使用された場合、 explode() は FALSE  を返します。
'  delimiter  に引数 string  に含まれていない値が含まれている場合、 explode() は、引数 string  を含む配列を返します。
'【処理】
'  ・文字列の配列を返します。この配列の各要素は、 string  を文字列 delimiter  で区切った部分文字列となります。
'=======================================================================
Function explode(delimiter,string,limit)

    explode = false
    If len(delimiter) = 0 Then Exit Function
    If len(limit) = 0 Then limit = 0

    If limit > 0 Then
        explode = Split(string,delimiter,limit)
    Else
        explode = Split(string,delimiter)
    End If

End Function

'=======================================================================
'特殊な HTML エンティティを文字に戻す
'=======================================================================
'【引数】
'  str         = string デコードする文字列。
'  quote_style = int    クォートの形式。以下の定数のいずれかです。
'【戻り値】
'  デコードされた文字列を返します。
'【処理】
'  ・特殊な HTML エンティティを文字に戻します。
'=======================================================================
function htmlspecialchars_decode(str,quote_style)

    Dim I
    Dim sText

    if empty_(quote_style) then quote_style = ENT_COMPAT
    sText = str

    if quote_style <> ENT_NOQUOTES then
        sText = Replace(sText, "&quot;", Chr(34))
    end if

    sText = Replace(sText, "&lt;"  , Chr(60))
    sText = Replace(sText, "&gt;"  , Chr(62))
    sText = Replace(sText, "&amp;" , Chr(38))
    sText = Replace(sText, "&nbsp;", Chr(32))

    For I = 1 to 255
        if I = 39 then
            if quote_style <> ENT_COMPAT then
                sText = Replace(sText, "&#" & I & ";", Chr(I))
            end if
        else
            sText = Replace(sText, "&#" & I & ";", Chr(I))
        end if
    Next

    htmlspecialchars_decode = sText

end function

'=======================================================================
' 特殊文字を HTML エンティティに変換する
'=======================================================================
'【引数】
'  str  = string    変換される文字列。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・文字の中には HTML において特殊な意味を持つものがあり、 それらの本来の値を表示したければ HTML の表現形式に変換してやらなければなりません。
'  ・この関数は、これらの変換を行った結果の文字列を返します。 
'=======================================================================
Function htmlSpecialChars(ByVal str)

    if len( str ) > 0 then
        str = Server.HTMLEncode(str)
        str = Replace(str,"'","&#039;")
    end if
    htmlspecialchars = str

End Function

'=======================================================================
' 配列要素を文字列により連結する
'=======================================================================
'【引数】
'  glue     = string    デフォルトは空文字 ('') です。 これは implode() の好ましい使用法ではありません。 下位互換性のため、常に 2 つのパラメータを使用することが推奨されています。
'  pieces   = array     連結したい文字列の配列。
'【戻り値】
'  すべての配列要素の順序を変えずに、各要素間に glue  文字列をはさんで 1 つの文字列にして返します。
'【処理】
'  ・配列の要素を glue  文字列で連結します。
'=======================================================================
Function implode(glue,pieces)
    implode = join(pieces,glue)
End Function

'=======================================================================
' 文字列の最初の文字を小文字にする
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・str  の最初の文字がアルファベットであれば、 それを小文字にします。
'=======================================================================
Function lcfirst(byVal str)

    Dim tmp
    tmp = left(str,1)
    tmp = Lcase(tmp)
    lcfirst = tmp & Mid(str,2)

End Function

'=======================================================================
' 二つの文字列のレーベンシュタイン距離を計算する
'=======================================================================
'【引数】
'  str1 = string    レーベンシュタイン距離を計算する文字列のひとつ。
'  str2 = string    レーベンシュタイン距離を計算する文字列のひとつ。
'【戻り値】
'  この関数は、引数で指定した二つの文字列のレーベンシュタイン距離を返します。
'  引数文字列の一つが 255 文字の制限より長い場合に -1 を返します。
'【処理】
'  ・レーベンシュタイン距離は、str1  を str2  に変換するために置換、挿入、削除 しなければならない最小の文字数として定義されます。
'  ・アルゴリズムの複雑さは、 O(m*n) です。
'  ・ここで、n および m はそれぞれ str1  および str2  の長さです (O(max(n,m)**3) となる similar_text() よりは良いですが、 まだかなりの計算量です)。
'  ・上記の最も簡単な形式では、この関数はパラメータとして引数を二つだけとり、 str1  から str2  に変換する際に必要な 挿入、置換、削除演算の数のみを計算します。
'=======================================================================
Function levenshtein( str1, str2 )

    Dim s,l,t,i,j,m,n,u
    Dim a,tmp

    s = str_split(str1,1)
    u = str_split(str2,1)

    If isArray(s) Then l = count(s,"") Else l = 0
    If isArray(u) Then t = count(u,"") Else t = 0

    If is_empty(l) or is_empty(t) Then
        If [>](l , t) Then levenshtein = l
        If [>](t , l) Then levenshtein = t
        If isEmpty(levenshtein) Then levenshtein = 0
        Exit Function
    End If

    ReDim a(l)
    For i = 0 to l
        toReDim a(i),t
    Next

    For i = l to 0 Step -1
       a(i)(0) = i
    Next

    For i = t to 0 Step -1
       a(0)(i) = i
    Next

    i = 0
    m = l

    Do While(i < m)

        j = 0
        n = t

        Do While(j < n)
            tmp = a(i)(j + 1) + 1
            If tmp > a(i+1)(j) + 1 Then tmp = a(i+1)(j) + 1
            If tmp > a(i)(j) + intval([!=](s(i) ,u(j))) Then tmp = a(i)(j) + intval([!=](s(i) ,u(j)))
            a(i+1)(j+1) = tmp

            j = j + 1
        Loop

        i = i +1
    Loop

    levenshtein = a(l)(t)

End Function

'=======================================================================
' 指定したファイルのMD5ハッシュ値を計算する
'=======================================================================
'【引数】
'  filename = string    ファイル名
'【戻り値】
'  成功時は文字列、そうでなければ FALSE
'【処理】
'  ・filename パラメータで指定したファイルの MD5ハッシュを計算し、そのハッシュを返します。
'=======================================================================
Function md5_file(filename)
    md5_file = md5( file_get_contents(filename) )
End Function

'=======================================================================
'文字列のmd5ハッシュ値を計算する
'=======================================================================
'【引数】
'  str      = string  文字列
'【戻り値】
'  32 文字の 16 進数からなるハッシュを返します。
'【処理】
'  ・str の MD5 ハッシュ値を計算し、 そのハッシュを返します。
'=======================================================================
Function md5(str)

    Dim bobj
    Set bobj = Server.CreateObject("basp21")
    md5 = bobj.MD5(str)

End Function

'=======================================================================
' 改行文字の前に HTML の改行タグを挿入する
'=======================================================================
'【引数】
'  str      = string  入力文字列。
'【戻り値】
'  変更後の文字列を返します。
'【処理】
'  ・string  に含まれるすべての改行文字の前に '<br />' を挿入して返します。
'=======================================================================
Function nl2br(str)
    nl2br = preg_replace("/([^>])" & vbCrLf & "/","$1<br />", str,"","")
End Function

'=======================================================================
' 数字を千位毎にグループ化してフォーマットする
'=======================================================================
'【引数】
'  number           = float     フォーマットする数値。
'  decimals         = int       小数点以下の桁数。
'  dec_point        = string    小数点を表す区切り文字。
'  thousands_sep    = string    千位毎の区切り文字。thousands_sep は最初の文字だけが使用されます。 例えば、数字の 1000 に対する thousands_sep として bar を使用した場合、number_format() は 1b000 を返します。
'【戻り値】
'  変更後の文字列を返します。
'【処理】
'  ・パラメータが 1 つだけ渡された場合、 number  は千位毎にカンマ (",") が追加され、 小数なしでフォーマットされます。
'  ・パラメータが 2 つ渡された場合、number は decimals 桁の小数の前にドット (".") 、 千位毎にカンマ (",") が追加されてフォーマットされます。
'  ・パラメータが 4 つ全て渡された場合、number はドット (".") の代わりに dec_point が decimals 桁の小数の前に、千位毎にカンマ (",") の代わりに thousands_sep が追加されてフォーマットされます。 
'=======================================================================
Function number_format( number, decimals, dec_point, thousands_sep )

    Dim n,c,d,t,i,s
    n = number
    c = [?]( isNumeric(decimals),decimals,2 )
    c = abs( c )

    d = [?]( len(dec_point) = 0,",",dec_point)
    t = [?]( len(thousands_sep) = 0, ".", left(thousands_sep,1) )

    n = FormatNumber (n, c,true,false,true)
    n = Replace(n,",",d)
    n = Replace(n,".",t)

    number_format = n

End Function

'=======================================================================
' 文字の ASCII 値を返す
'=======================================================================
'【引数】
'  str  = string    文字
'【戻り値】
'  ASCII 値を返します。
'【処理】
'  ・string  の先頭文字の ASCII 値を返します。
'=======================================================================
Function ord(str)
    ord = asc( left(str,1) )
End Function

'=======================================================================
' 文字列を処理し、変数に代入する
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'  arr  = array     2 番目の引数 arr  が指定された場合、 変数は、代わりに配列の要素としてこの変数に保存されます。
'【戻り値】
'  作成した配列を返します。
'【処理】
'  ・URL 経由で渡されるクエリ文字列と同様に str  を処理し、現在のスコープに変数をセットします。
'=======================================================================
Function parse_str(str,ByRef arr)

    Dim glue1 : glue1 = "="
    Dim glue2 : glue2 = "&"
    Dim array2,array3 : set array3 = Server.CreateObject("Scripting.Dictionary")
    Dim x,tmp,counter,tmp_ar

    array2 = Split(str,glue2)
    If uBound( array2 ) > 0 Then
        For x = 0 to uBound( array2 )
            tmp = Split( array2(x), glue1 )
            If uBound( tmp ) > 0 Then

                tmp(0) = urldecode( tmp(0) )
                tmp(1) = Replace( urldecode(tmp(1)), "+", " ")

                If array3.Exists( tmp(0) ) Then
                    If isArray( array3.Item( tmp(0) ) ) Then
                        tmp_ar = array_values( array3.Item( tmp(0) ) )
                        [] tmp_ar, tmp(1)
                        array3.Item( tmp(0) ) = tmp_ar
                    Else
                        array3.Item( tmp(0) ) = array(array3.Item( tmp(0) ), tmp(1))
                    End If
               Else
                    array3.Add urldecode(tmp(0)), Replace( urldecode(tmp(1)), "+", " ")
                End If
            End If
        Next
    End If

    If vartype(arr) = 0 Then
        [=] arr, array3
    Else
        [=] parse_str, array3
    End If

End Function

'=======================================================================
' 文字列を出力する
'=======================================================================
'【引数】
'  str  = string    入力データ。
'【戻り値】
'  常に 1 を返します。
'【処理】
'  strを出力します。
'=======================================================================
Function print(str)
    echo str
    print = 1
End Function

'=======================================================================
' フォーマット済みの文字列を出力する
'=======================================================================
'【引数】
'  format   = string フォーマット文字列
'  args     = mixed  数値や文字列
'【戻り値】
'  フォーマット文字列 format  に基づき生成された文字列を出力します。
'【処理】
'  ・format  にしたがって、出力を生成します。
'=======================================================================
Function printf( format , args)
    Response.Write sprintf(format,args)
End Function

'=======================================================================
' メタ文字をクォートする
'=======================================================================
'【引数】
'  format   = string 入力文字列。
'【戻り値】
'  メタ文字をクォートした文字列を返します。
'【処理】
'  ・ 文字列 str  について、. ¥ + * ? [ ^ ] ( $ ) の前にバックスラッシュ文字 (¥) でクォートして返します。
'=======================================================================
Function quotemeta(byVal str)

    Dim pattern : pattern = array("¥",".","+","*","?","[","^","]","$","(",")")

    Dim key
    For key = 0 to uBound(pattern)
        str = Replace(str, pattern(key),"¥" & pattern(key))
    Next
    quotemeta = str

End Function

'=======================================================================
' 文字列の soundex キーを計算する
'=======================================================================
'【引数】
'  format   = string 入力文字列。
'【戻り値】
'  メタ文字をクォートした文字列を返します。
'【処理】
'  ・ str  の soundex キーを計算します。
'  ・ soundex キーには、似たような発音の単語に関して同じ soundex キーが生成されるという特性があります。 このため、発音は知っているが、スペルがわからない場合に、 データベースを検索することを容易にすることができます。
'  ・ soundex 関数は、ある文字から始まる 4 文字の文字列を返します。
'  ・ この soundex 関数についての説明は、Donald Knuth の "The Art Of Computer Programming, vol. 3: Sorting And Searching", Addison-Wesley (1973), pp. 391-392 にあります。 
'=======================================================================
Function soundex(str)

    Dim i,j, l, r, p, m, s
    p = [?](Not isNumeric(p) , 4 , [?](p > 10 , 10 , [?](p < 4 , 4 , p) ) )

    set m = Server.CreateObject("Scripting.Dictionary")
    m.Add "BFPV", 1
    m.Add "CGJKQSXZ", 2
    m.add "DT", 3
    m.add "L", 4
    m.add "MN", 5
    m.add "R", 6

    s = Ucase( str )
    s = preg_replace("/[^A-Z]/","",s,"","")
    s = str_split(s,1)
    r = array( array_shift(s) )

    For i = 0 to uBound(s)
        For Each j In m
            if inStr(j,s(i)) and r( uBound(r) ) <> m.Item(j) Then
                array_push r,m(j)
                Exit For
            End If
        Next
    Next

    If uBound(r) + 1 > p Then
        r = array_slice(r,0,p-1)
    End If

    Dim newArray()
    ReDim newArray(p - (uBound(r)+1))

    soundex = join(r,"") & join( newArray, "0" )

End Function

'=======================================================================
' フォーマットされた文字列を返す
'=======================================================================
'【引数】
'  format   = string フォーマット文字列
'  args     = mixed  数値や文字列
'【戻り値】
'  フォーマット文字列 format  に基づき生成された文字列を返します。
'【処理】
'  ・フォーマット文字列 format  に基づき生成された文字列を返します。
'=======================================================================
Function sprintf(format , args)

    If is_empty(args) Then args = ""
    Dim bobj : set bobj = Server.CreateObject("basp21")
    sprintf = bobj.sprintf(format,args)

End Function

'=======================================================================
' 大文字小文字を区別しない str_replace()
'=======================================================================
'【引数】
'  search    = mixed 検索文字列
'  strReplace= mixed 置換文字列
'  subject   = mixed subject  が配列の場合は、そのすべての要素に 対して検索と置換が行われ、返される結果も配列となります。
'  cnt       = mixed needles  の中で、マッチして置換を行った数を count  に返します。このパラメータは参照渡しとします。
'【戻り値】
'  置換した文字列あるいは配列を返します。
'【処理】
'  ・この関数は、subject  の中に現れるすべての search (大文字小文字を区別しない)を replace  に置き換えた文字列あるいは配列を 返します。
'=======================================================================
Function str_ireplace(search, strReplace, subject, ByRef cnt)

    If is_string(search) and isArray(strReplace) Then Exit Function

    If Not isArray(search) Then search = array(search)
    search = array_values(search)

    Dim replace_string,i
    If not isArray(strReplace) Then
        replace_string = strReplace

        strReplace = array()
        For i = 0 to uBound(search)
            [] strReplace, replace_string
        Next
    End if

    strReplace = array_values(strReplace)

    Dim length_replace,length_search
    length_replace = count(strReplace,"")
    length_search  = count(search,"")
    if length_replace < length_search Then
        For i = length_replace to length_search
            strReplace(i) = ""
        Next
    End If

    Dim was_array : was_array = false
    If isArray(subject) Then
        was_array = true
        subject = array( subject )
    End If

    For i = 0 to uBound( search )
        search(i) = "/" & preg_quote(search(i),"") & "/"
    Next

    For i = 0 to uBound( strReplace )
        strReplace(i) = str_replace( array(chr(92),"$"),array(chr(92) & chr(92), "¥$"),strReplace(i) )
    Next

    Dim result
    result = preg_replace(search,strReplace,subject,"",cnt)

    If was_array = true Then
        str_ireplace = result(0)
    Else
        str_ireplace = result
    End If

End Function

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

'=======================================================================
' 文字列を反復する
'=======================================================================
'【引数】
'  input        = string    繰り返す文字列。
'  multiplier   = int       input を繰り返す回数。multiplier は 0 以上でなければなりません。 multiplier が 0 に設定された場合、この関数は空文字を返します。
'【戻り値】
'  繰り返した文字列を返します。
'【処理】
'  ・input  を multiplier  回を繰り返した文字列を返します。
'=======================================================================
Function str_repeat(input, multiplier)
    If multiplier < 0 Then Exit Function
    str_repeat = String(multiplier,input)
End Function

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

'=======================================================================
' 文字列に rot13 変換を行う
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'【戻り値】
'  指定した文字列を ROT13 変換した結果を返します。
'【処理】
'  ・ROT13 は、各文字をアルファベット順に 13 文字シフトさせ、 アルファベット以外の文字はそのままとするエンコードを行います。
'  エンコードとデコードは同じ関数で行われます。
'  引数にエンコードされた文字列を指定した場合には、元の文字列が返されます。
'=======================================================================
Function str_rot13(str)

    Dim str_rotated : str_rotated = ""
    Dim i,j,k

    For i = 1 to Len(str)
        j = Mid(str, i, 1)
        k = Asc(j)
        if k >= 97 and k =< 109 then
            k = k + 13 ' a ... m
        elseif k >= 110 and k =< 122 then
            k = k - 13 ' n ... z
        elseif k >= 65 and k =< 77 then
            k = k + 13 ' A ... M
        elseif k >= 78 and k =< 90 then
            k = k - 13 ' N ... Z
        end if

        str_rotated = str_rotated & Chr(k)
    Next

    str_rot13 = str_rotated

End Function

'=======================================================================
' 文字列をランダムにシャッフルする
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'【戻り値】
'  シャッフルされた文字列を返します。
'【処理】
'  ・str_shuffle() は文字列をシャッフルします。
'  ・考えられるすべての順列のうちのひとつを作成します。
'=======================================================================
Function str_shuffle(str)

    If len( str ) = 0 Then Exit Function

    Dim tmp
    tmp = str_split(str,1)
    shuffle tmp
    str_shuffle = join(tmp,"")

End Function

'=======================================================================
' 文字列を配列に変換する
'=======================================================================
'【引数】
'  string       = string 入力文字列。
'  split_length = string 分割した部分の最大長。
'【戻り値】
'  オプションのパラメータ split_length  が指定されている場合、 返される配列の各要素は、split_length  の長さとなります。それ以外の場合、1 文字ずつ分割された配列となります。
'  split_length が 1 より小さい場合に FALSE を返します。
'  split_length が string の長さより大きい場合、文字列全体が 最初の(そして唯一の)要素となる配列を返します。 
'【処理】
'  ・文字列を配列に変換します。
'=======================================================================
Function str_split(string, split_length)

    str_split = false
    If len(string) = 0 Then Exit Function
    If len(split_length) = 0 Then split_length = 1
    If split_length < 1 Then Exit Function

    Dim counter,i,pointer
    counter = len(string)
    counter = counter / split_length + 0.9999
    counter = int(counter) -1

    ReDim tmp_ar(counter)

    For i = 0 to counter
        pointer = i * split_length + 1
        tmp_ar(i) = Mid(string,pointer,split_length)
    Next

    str_split = tmp_ar

End Function

'=======================================================================
' 大文字小文字を区別しないバイナリセーフな文字列比較を行う
'=======================================================================
'【引数】
'  str1 =   string  最初の文字列。
'  str2 =   string  次の文字列。
'【戻り値】
'  str1  が str2  より小さい場合は負、str1  が str2  より大きい場合は正、等しい場合は 0 を返します。
'【処理】
'  ・大文字小文字を区別しないバイナリセーフな文字列比較を行います。
'=======================================================================
Function strcasecmp(str1, str2)
    strcasecmp = StrComp(str1,str2,vbTextCompare)
End Function

'=======================================================================
' strstr() のエイリアス
'=======================================================================
'【引数】
'  haystack     = string    入力文字列。
'  needle       = string    needle  が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  before_needle= string    TRUE にすると (デフォルトは FALSE です)、strstr()  の返り値は、haystack  の中で最初に needle  があらわれる箇所より前の部分となります。
'【処理】
'  ・この関数は次の関数のエイリアスです。 strstr().
'=======================================================================
Function strchr( haystack, needle, before_needle )
    strchr = strstr( haystack, needle, before_needle )
End Function

'=======================================================================
' バイナリセーフな文字列比較
'=======================================================================
'【引数】
'  str1 =   string  最初の文字列。
'  str2 =   string  次の文字列。
'【戻り値】
'  str1  が str2  よりも小さければ < 0 を、str1 が str2 よりも大きければ > 0 を、 等しければ 0 を返します。
'【処理】
'  ・この比較は大文字小文字を区別することに注意してください。
'=======================================================================
Function strcmp(str1, str2)
    strcmp = StrComp(str1,str2,vbBinaryCompare)
End Function

'=======================================================================
' 文字列から HTMLタグを取り除く
'=======================================================================
'【引数】
'  str              = string 入力文字列。
'  allowable_tags   = string オプションの2番目の引数により、取り除かないタグを指定できます。
'【戻り値】
'  タグを除去した文字列を返します。
'【処理】
'  ・指定した文字列 ( str ) から全ての HTMLタグを取り除きます。
'=======================================================================
Function strip_tags( str )

    Dim objRegExp
    Dim plane

    plane = Trim( str & "" )

    If Len( plane ) > 0 Then

        Set objRegExp = New RegExp
        objRegExp.IgnoreCase = True
        objRegExp.Global = True
        objRegExp.Pattern= "</?[^>]+>"
        plane = objRegExp.Replace(str, "")
        Set objRegExp = Nothing

    End If

    strip_tags = plane

End Function

'=======================================================================
' 大文字小文字を区別せずに文字列が最初に現れる位置を探す
'=======================================================================
'【引数】
'  haystack = string    検索を行う文字列。
'  needle   = string    needle は、 ひとつまたは複数の文字であることに注意しましょう。needle が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  offset   = string    オプションのパラメータ offset  により、検索を開始する haystack  の位置を指定することができます。 この場合でも返される位置は、 haystack  の先頭からの位置のままとなります。
'【戻り値】
'  needle  がみつからない場合、 stripos() は boolean FALSE  を返します。
'【処理】
'  ・ 文字列 haystack  の中で needle  が最初に現れる位置を数字で返します。
'  ・ strpos() と異なり、stripos() は大文字小文字を区別しません。 
'=======================================================================
Function stripos( haystack, needle, offset)

    Dim i
    stripos = false

    If len(offset) = 0 Then
        offset = 1
    End If

    i = inStr(offset,haystack,needle,vbTextCompare)

    If i > 0 Then
        stripos = i
    End If

End Function

'=======================================================================
' addslashes() でクォートされた文字列のクォート部分を取り除く
'=======================================================================
'【引数】
'  str  = string    元に戻したい文字列。
'【戻り値】
'  バックスラッシュが取り除かれた文字列を返します(¥'  が ' になるなど)。
'  2 つ並んだバックスラッシュ (¥¥) は 1 つのバックスラッシュ (¥) になります。
'【処理】
'  ・クォートされた文字列を元に戻します。
'=======================================================================
Function stripslashes(ByVal str)
    str = preg_replace("/¥¥(.)/","$1",str,"","")
    stripslashes = str
End Function

'=======================================================================
' 大文字小文字を区別しない strstr()
'=======================================================================
'【引数】
'  haystack     = string    検索を行う文字列。
'  needle       = string    needle は、 ひとつまたは複数の文字であることに注意しましょう。needle が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  before_needle= string    TRUE にすると (デフォルトは FALSE です)、stristr()  の返り値は、haystack  の中で最初に needle  があらわれる箇所より前の部分となります。
'【戻り値】
'  マッチした部分文字列を返します。needle  が見つからない場合は FALSE を返します。
'【処理】
'  ・haystack  において needle  が最初に見つかった位置から最後までを返します。
'  ・needle  および haystack  は大文字小文字を区別せずに評価されます。
'=======================================================================
Function stristr( haystack, needle, before_needle )

    Dim pos
    If varType(before_needle) <> 11 Then before_needle = false

    pos = Instr(1,haystack,needle,vbTextCompare)

    If pos <= 0 Then
        stristr = false
    Else
        If before_needle Then
            stristr = Mid(haystack,1,pos-1)
        Else
            stristr = Mid(haystack,pos)
        End If
    End If

End Function

'=======================================================================
' 文字列の長さを得る
'=======================================================================
'【引数】
'  str  = string    長さを調べる文字列。
'【戻り値】
'  成功した場合に str の長さ、 str  が空の文字列だった場合に 0 を返します。
'【処理】
'  ・与えられた str の長さを返します。
'=======================================================================
Function strlen(str)
    strlen = len(str)
End Function

'=======================================================================
' "自然順"アルゴリズムにより大文字小文字を区別しない文字列比較を行う
'=======================================================================
'【引数】
'  str1 =   string  最初の文字列。
'  str2 =   string  次の文字列。
'【戻り値】
'  他の文字列比較関数と同様に、この関数は、 str1  が str2  より小さいに場合に < 0、str1  が str2  より大きい場合に > 0 、等しい場合に 0 を返します。
'【処理】
'  ・この関数は、人間が行うような手法でアルファベットまたは数字の 文字列の順序を比較するアルゴリズムを実装します。この手法は、"自然順" と言われます。
'  ・この関数の動作は、 strnatcmp() に似ていますが、 比較が大文字小文字を区別しない違いがあります。
'=======================================================================
Function strnatcasecmp( str1, str2 )

    Dim array1,array2
    array1 = strnatcmp_helper(str1)
    array2 = strnatcmp_helper(str2)

    Dim intlen,text,result,r
    intlen = uBound(array1)
    text   = true

    result = -1
    r      = 0

    if intlen > uBound(array2) Then
        intlen = uBound(array2)
        result = 1
    End If

    Dim i
    strnatcasecmp = false
    For i = 0 to intlen
        If not isNumeric( array1(i) ) Then
            If Not isNumeric( array2(i) ) Then
                text = true

                r = strcasecmp(array1(i),array2(i))
                If r <> 0 Then
                    strnatcasecmp = r
                End If

            ElseIf text Then
                strnatcasecmp = 1
            Else
                strnatcasecmp = 1
            End If

        ElseIf not isNumeric( array2(i) ) Then
            If text Then
                strnatcasecmp = -1
            Else
                strnatcasecmp = 1
            End If
        Else
            If text Then
                r = array1(i) - array2(i)
                If r <> 0 Then
                    strnatcasecmp = r
                End If
            Else
                r = strcasecmp(array1(i),array2(i))
                If r <> 0 Then
                    strnatcasecmp = r
                End If
            End If

            text = false
        End If

        if [!==](strnatcasecmp,false) Then Exit Function
    Next

    strnatcasecmp = result

End Function

'=======================================================================
' "自然順"アルゴリズムにより文字列比較を行う
'=======================================================================
'【引数】
'  str1 =   string  最初の文字列。
'  str2 =   string  次の文字列。
'【戻り値】
'  他の文字列比較関数と同様に、この関数は、 str1  が str2  より小さいに場合に < 0、str1  が str2  より大きい場合に > 0 、等しい場合に 0 を返します。
'【処理】
'  ・この関数は、人間が行うような手法でアルファベットまたは数字の 文字列の順序を比較するアルゴリズムを実装します。この手法は、"自然順" と言われます。
'  ・この比較は、大文字小文字を区別することに注意してください。
'=======================================================================
Function strnatcmp( str1, str2 )

    Dim array1,array2
    array1 = strnatcmp_helper(str1)
    array2 = strnatcmp_helper(str2)

    Dim intlen,text,result,r
    intlen = uBound(array1)
    text   = true

    result = -1
    r      = 0

    if intlen > uBound(array2) Then
        intlen = uBound(array2)
        result = 1
    End If

    Dim i
    strnatcmp = false
    For i = 0 to intlen
        If not isNumeric( array1(i) ) Then
            If Not isNumeric( array2(i) ) Then
                text = true

                r = strcmp(array1(i),array2(i))
                If r <> 0 Then
                    strnatcmp = r
                End If

            ElseIf text Then
                strnatcmp = 1
            Else
                strnatcmp = 1
            End If

        ElseIf not isNumeric( array2(i) ) Then
            If text Then
                strnatcmp = -1
            Else
                strnatcmp = 1
            End If
        Else
            If text Then
                r = array1(i) - array2(i)
                If r <> 0 Then
                    strnatcmp = r
                End If
            Else
                r = strcmp(Cstr( array1(i) ),Cstr( array2(i) ) )
                If r <> 0 Then
                    strnatcmp = r
                End If
            End If

            text = false
        End If

        if [!==](strnatcmp,false) Then Exit Function
    Next

    strnatcmp = result

End Function

'*****************************
Function strnatcmp_helper(str)

    Dim result
    Dim buffer : buffer = ""
    Dim strChr : strChr = ""
    Dim text   : text   = true

    Dim i
    For i = 1 to len(str)
        strChr = Mid(str,i,1)

        If preg_match("/[0-9]/i",strChr,"","","") Then
            If text Then
                If len( buffer ) > 0 Then
                    [] result, buffer
                    buffer = ""
                End If

                text = false
            End If
            buffer  = buffer & strChr
        ElseIf text = false and strChr = "." and i < (len( str )-1) Then
            If preg_match("/[0-9]/",Mid(str,i+1,1),"","","") Then
                [] result, buffer
                buffer = ""
            End If
        Else
            If text = false Then
                If len( buffer ) > 0 Then
                    [] result , intval(buffer)
                    buffer = ""
                End If
                text = true
            End If
            buffer = buffer & strChr
        End If
    Next

    If len( buffer ) > 0 Then
        if text Then
            [] result, buffer
        Else
            [] result, intval(buffer)
        End If
    End If

    strnatcmp_helper = result
End Function

'=======================================================================
' バイナリセーフで大文字小文字を区別しない文字列比較を、最初の n 文字について行う
'=======================================================================
'【引数】
'  str1     =   string  最初の文字列。
'  str2     =   string  次の文字列。
'  intlen   =   string  比較する文字列の長さ。
'【戻り値】
' str1  が str2  より短い場合に < 0 を返し、str1  が str2  より大きい場合に > 0、等しい場合に 0 を返します。
'【処理】
'  ・この関数は、strcasecmp() に似ていますが、 各文字列から比較する文字数(の上限)(len ) を指定できるという違いがあります。
'  ・どちらかの文字列が len より短い場合、その文字列の長さが比較時に使用されます。
'=======================================================================
Function strncasecmp(ByVal str1,ByVal str2,intlen)

    If len(str1) > intlen Then str1 = Left(str1,intlen)
    If len(str2) > intlen Then str2 = Left(str2,intlen)

    strncasecmp = StrComp(str1,str2,vbTextCompare)
End Function

'=======================================================================
' 最初の n 文字についてバイナリセーフな文字列比較を行う
'=======================================================================
'【引数】
'  str1     =   string  最初の文字列。
'  str2     =   string  次の文字列。
'  intlen   =   string  比較する文字数。
'【戻り値】
' str1  が str2  より短い場合に < 0 を返し、str1  が str2  より大きい場合に > 0、等しい場合に 0 を返します。
'【処理】
'  ・ この関数は strcmp() に似ていますが、 各文字列から(最大)文字数(len ) を比較に使用するところが異なります。
'  ・ 比較は大文字小文字を区別することに注意してください。 
'=======================================================================
Function strncmp(ByVal str1,ByVal str2,intlen)

    If len(str1) > intlen Then str1 = Left(str1,intlen)
    If len(str2) > intlen Then str2 = Left(str2,intlen)

    strncmp = StrComp(str1,str2,vbBinaryCompare)
End Function

'=======================================================================
' 文字列の中から任意の文字を探す
'=======================================================================
'【引数】
'  haystack     =   string  char_list  を探す文字列。
'  char_list    =   string  このパラメータは大文字小文字を区別します。
'【戻り値】
'  見つかった文字から始まる文字列、あるいは見つからなかった場合に FALSE を返します。
'【処理】
'  ・ strpbrk() は、文字列 haystack  から char_list  を探します。
'=======================================================================
Function strpbrk( haystack, char_list )

    haystack  = Cstr( haystack )
    char_list = Cstr( char_list )

    Dim intlen : intlen = len( haystack )
    Dim i,char
    For i = 1 to intlen
        char = Mid(haystack,i,1)
        If [!==](strpos(char_list,char,""),false) Then
            strpbrk = Mid(haystack,i)
            Exit Function
        End If
    Next

    strpbrk = false

End Function

'=======================================================================
' 文字列が最初に現れる場所を見つける
'=======================================================================
'【引数】
'  haystack = string    検索を行う文字列。
'  needle   = string   needle  が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  offset   = string    オプションのパラメータ offset  により、検索を開始する haystack  の位置を指定することができます。 この場合でも返される位置は、 haystack  の先頭からの位置のままとなります。
'【戻り値】
'  位置を表す整数値を返します。 needle  が見つからない場合、 strpos() は boolean FALSE を返します。
'【処理】
'  ・ 文字列 haystack  の中で、 needle  が最初に現れた位置を数字で返します。
'  ・ PHP 5 以前の strrpos() とは異なり、この関数は needle  パラメータとして文字列全体をとり、 その文字列全体が検索対象となります。
'=======================================================================
Function strpos( haystack, needle, offset)

    Dim i
    strpos = false

    If len(offset) = 0 Then
        offset = 1
    End If

    i = inStr(offset,haystack,needle,vbBinaryCompare)

    If i > 0 Then
        strpos = i
    End If

End Function

'=======================================================================
' 文字列中に文字が最後に現れる場所を取得する
'=======================================================================
'【引数】
'  haystack = string    検索を行う文字列。
'  needle   = string    needle  がひとつ以上の文字を含んでいる場合は、 最初のもののみが使われます。この動作は、 strstr()  とは異なります。
'【戻り値】
'  この関数は、部分文字列を返します。 needle  が見つからない場合は FALSE を返します。
'【処理】
'  ・ この関数は、文字列 haystack  の中で needle  が最後に現れた位置から、 haystack  の終わりまでを返します。
'=======================================================================
Function strrchr( haystack, needle )

    haystack = Cstr( haystack )
    needle   = Cstr( needle )
    If len(needle) > 1 Then needle = Left(needle,1)

    strrchr = false

    Dim i
    i = strrpos(haystack, needle,"")

    If i > 0 Then
        strrchr = Mid(haystack,i)
    End If

End Function

'=======================================================================
' 文字列を逆順にする
'=======================================================================
'【引数】
'  str  = string    逆順にしたい文字列。
'【戻り値】
'  逆順にした文字列を返します。
'【処理】
'  ・ str  を逆順にして返します。
'=======================================================================
Function strrev(str)
    strrev = StrReverse(str)
End Function

'=======================================================================
' 文字列中で、特定の(大文字小文字を区別しない)文字列が最後に現れた位置を探す
'=======================================================================
'【引数】
'  haystack = string    検索を行う文字列。
'  needle   = string   needle  が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  offset   = string    オプションのパラメータ offset  により、検索を開始する haystack  の位置を指定することができます。 この場合でも返される位置は、 haystack  の先頭からの位置のままとなります。
'【戻り値】
'  needle  が最後に現れた位置を返します。
'  needle  が見つからない場合、FALSE を返します。
'【処理】
'  ・ 文字列の中で、 大文字小文字を区別しないある文字列が最後に現れた位置を返します。
'  ・ strrpos() と異なり、strripos()  は大文字小文字を区別しません。
'=======================================================================
Function strripos( haystack, needle, offset)

    Dim i
    strripos = false

    If len(offset) = 0 Then
        offset = len( haystack)
    End If

    If len(needle) > 1 Then needle = Left(needle,1)

    i = InStrRev(haystack,needle,offset,vbTextCompare)

    If i > 0 Then
        strripos = i
    End If

End Function

'=======================================================================
' 文字列中に、ある文字が最後に現れる場所を探す
'=======================================================================
'【引数】
'  haystack = string    検索を行う文字列。
'  needle   = string   needle  が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  offset   = string    オプションのパラメータ offset  により、検索を開始する haystack  の位置を指定することができます。 この場合でも返される位置は、 haystack  の先頭からの位置のままとなります。
'【戻り値】
'  needle  が最後に現れた位置を返します。
'  needle  が見つからない場合、FALSE を返します。
'【処理】
'  ・ 文字列 haystack  の中で、 needle  が最後に現れた位置を数字で返します。
'  ・ needle  に文字列が指定された場合、その文字列の最初の文字だけが使われます。
'=======================================================================
Function strrpos( haystack, needle, offset)

    Dim i
    strrpos = false

    If len(offset) = 0 Then
        offset = len( haystack)
    End If

    If len(needle) > 1 Then needle = Left(needle,1)

    i = InStrRev(haystack,needle,offset,vbBinaryCompare)

    If i > 0 Then
        strrpos = i
    End If

End Function

'=======================================================================
' 文字列が最初に現れる位置を見つける
'=======================================================================
'【引数】
'  haystack     = string    入力文字列。
'  needle       = string    needle  が文字列でない場合は、 それを整数に変換し、その番号に対応する文字として扱います。
'  before_needle= string    TRUE にすると (デフォルトは FALSE です)、strstr()  の返り値は、haystack  の中で最初に needle  があらわれる箇所より前の部分となります。
'【戻り値】
'  部分文字列を返します。 needle  が見つからない場合は FALSE を返します。
'【処理】
'  ・haystack  の中で needle  が最初に現れる場所から文字列の終わりまでを返します。
'=======================================================================
Function strstr( haystack, needle, before_needle )

    Dim pos
    If varType(before_needle) <> 11 Then before_needle = false

    pos = Instr(1,haystack,needle,vbBinaryCompare)

    If pos <= 0 Then
        strstr = false
    Else
        If before_needle Then
            strstr = Mid(haystack,1,pos-1)
        Else
            strstr = Mid(haystack,pos)
        End If
    End If

End Function

'=======================================================================
' 文字列を小文字にする
'=======================================================================
'【引数】
'  str     = string    入力文字列。
'【戻り値】
'  小文字に変換した文字列を返します。
'【処理】
'  ・str  のアルファベット部分をすべて小文字にして返します｡
'=======================================================================
Function strtolower(str)
    strtolower = Lcase(str)
End Function

'=======================================================================
' 文字列を大文字にする
'=======================================================================
'【引数】
'  str     = string    入力文字列。
'【戻り値】
'  大文字にした文字列を返します。
'【処理】
'  ・str  のアルファベット部分をすべて大文字にして返します｡
'=======================================================================
Function strtoupper(str)
    strtoupper = Ucase(str)
End Function

'=======================================================================
' 特定の文字を変換する
'=======================================================================
'【引数】
'  str      = string    変換する文字列。
'  from     = string    strTo  に変換される文字列。
'  strTo    = int       from  を置換する文字列。
'【戻り値】
'  この関数は str  を走査し、 from  に含まれる文字が見つかると、そのすべてを strTo  の中で対応する文字に置き換え、 その結果を返します。
'【処理】
'  ・ この関数は str  を走査し、 from  に含まれる文字が見つかると、そのすべてを to  の中で対応する文字に置き換え、 その結果を返します。
'  ・ from と to の長さが異なる場合、長い方の余分な文字は無視されます。 
'=======================================================================
Function strtr(ByVal str, from, strTo)

    If isObject(from) Then
        Dim key
        For Each key In from
            str = Replace(str,key,from(key))
        Next

    Else

        Dim len1 : len1 = len(from)
        Dim len2 : len2 = len(strTo)

        If len1 > len2 Then
            from = Left(from,len2)
        ElseIf len2 > len1 Then
            strTo = Left(strTo,len1)
        End If

        str = Replace(str,from,strTo)
    End if

    strtr = str
End Function

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

'=======================================================================
' 副文字列の出現回数を数える
'=======================================================================
'【引数】
'  haystack     = string    検索対象の文字列
'  needle       = string    検索する副文字列
'  offset       = int       開始位置のオフセット
'  length       = int       指定したオフセット以降に副文字列で検索する最大長。
'【戻り値】
'  この関数は 整数 を返します。
'【処理】
'  ・substr_count() は、文字列 haystack  の中での副文字列 needle  の出現回数を返します。
'  ・needle  は英大小文字を区別することに注意してください。
'=======================================================================
Function substr_count( haystack, needle, offset, length )

    Dim pos,cnt : cnt = 0

    If not isNumeric(offset) Then offset = 1
    If not isNumeric(length) Then length = 0

    Do While inStr(offset+1,haystack,needle,vbBinaryCompare) > 0
        offset = inStr(offset+1,haystack,needle,vbBinaryCompare)
        If length > 0 and offset + len(needle) > length Then
            Exit Do
        Else
            cnt = cnt + 1
        End If
    Loop

    substr_count = cnt

End Function

'=======================================================================
' 文字列の一部を置換する
'=======================================================================
'【引数】
'  str          = string    入力文字列。
'  replacement  = string    置換する文字列。
'  start        = string    start が正の場合、置換は string で start 番目の文字から始まります。start が負の場合、置換は string の終端から start 番目の文字から始まります。
'  length       = string    正の値を指定した場合、 string 　の置換される部分の長さを表します。 負の場合、置換を停止する位置が string  の終端から何文字目であるかを表します。このパラメータが省略された場合、 デフォルト値は strlen(string )、すなわち、 string  の終端まで置換することになります。 当然、もし length  がゼロだったら、 この関数は string  の最初から start  の位置に replacement  を挿入するということになります。
'【戻り値】
'  結果の文字列を返します。もし、string  が配列の場合、配列が返されます。
'【処理】
'  ・substr_replace()は、文字列 string の start  および (オプションの) length  パラメータで区切られた部分を replacement  で指定した文字列に置換します。
'=======================================================================
Function substr_replace(str, replacement, start, length)

    Dim key

    If isArray(str) Then
        For key = 0 to uBound(str)
            substr_replace(key) = substr_replace(str(key), replacement, start, length)
        Next
        Exit Function
    ElseIf isObject(str) Then
        For Each key In str
            substr_replace(key) = substr_replace(str(key), replacement, start, length)
        Next
        Exit Function
    End If

    Dim result

    If start < 0 Then
        start = len(str) + start
    End If

    If start <> 0 Then
        result = Left(str,start)
    End If

    result = result & replacement

    If len(length) > 0 Then
        If length > 0 Then
            result = result & Mid(str,start + length)
        Else
            result = result & Right(str,abs(length))
        End If

    End If

    substr_replace = result
End Function

'=======================================================================
' 文字列の一部分を返す
'=======================================================================
'【引数】
'  str          = string    入力文字列。
'  start        = string     start  が正の場合、返される文字列は、 string  の 0 から数えて start 番目から始まる文字列となります。 例えば、文字列'abcdef'において位置 0にある文字は、'a'であり、 位置2には'c'があります。start が負の場合、返される文字列は、 string の後ろから数えて start 番目から始まる文字列となります。 
'  intLength    = string    入力文字列。
'【戻り値】
'  文字列の一部を返します。
'【処理】
'  ・文字列 str  の、start  で指定された位置から length  バイト分の文字列を返します。
'=======================================================================
Function substr(ByVal str,ByVal start,ByVal intLength)

	intStart = start
    If len(intStart) = 0 Then intStart = 0
    if intStart = 0 And Len(intLength) < 1 Then
    	substr = str
    	Exit Function
    End If
    
    If len(intLength) = 0 Then intLength = abs(start)
    
    If intStart < 0 Then
        intStart = len(str)+1 + intStart
    Else
    	intStart = intStart + 1
    End If

    Dim tmp
    If intStart <> 0 Then
	    tmp = Mid(str,intStart)
	Else 
		tmp = str
	End If

    Dim intLen
    intLen = len(tmp)


    If len(intLength) > 0 Then

        If intLen >= abs(intLength) Then
            If intLength > 0 Then

        		tmp = Left(tmp,intLength)

            Else
                tmp = Left(tmp,len(tmp) + intLength)
            End If
        Else
            tmp = False
        End If
    End If

    substr = tmp

End Function

'=======================================================================
' 文字列の最初の文字を大文字にする
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・str  の最初の文字がアルファベットであれば、 それを大文字にします。
'=======================================================================
Function ucfirst(byVal str)

    Dim tmp
    tmp = left(str,1)
    tmp = Ucase(tmp)
    ucfirst = tmp & Mid(str,2)

End Function

'=======================================================================
' 文字列の各単語の最初の文字を大文字にする
'=======================================================================
'【引数】
'  str  = string    入力文字列。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・ 文字がアルファベットの場合、str  の各単語の最初の文字を大文字にしたものを返します。
'  ・ 単語の定義は、空白文字 (スペース、フォームフィード、改行、キャリッジリターン、 水平タブ、垂直タブ) の直後にあるあらゆる文字からなる文字列です。 
'=======================================================================
Function ucwords(str)
    ucwords = preg_replace_callback("/^(.)|¥s(.)/","Ucase",str,"","")
End Function

'=======================================================================
' 文字列分割文字を使用して指定した文字数数に文字列を分割する
'=======================================================================
'【引数】
'  str      = string    入力文字列。
'  width    = int       カラムの幅。デフォルトは 75。
'  break    = string    オプションのパラメータ break  を用いて行を分割します。 デフォルトは 'vbCrLf' です。
'  cut      = bool      cut  を TRUE に設定すると、 文字列は常に指定した幅でラップされます。このため、 指定した幅よりも長い単語がある場合には、分割されます (2 番目の例を参照ください)。
'【戻り値】
'  変換後の文字列を返します。
'【処理】
'  ・ 指定した文字数で、指定した文字を用いて文字列を分割します。
'=======================================================================
Function wordwrap( str, int_width, str_break, cut )

    If len(int_width) = 0 Then int_width = 75
    If len(str_break) = 0 Then str_break = vbCrLf

    Dim m : m = int_width
    Dim b : b = str_break
    Dim c : c = cut

    Dim i,j, l, s, r
    Dim matches

    If m < 1 Then
        wordwrap = str
        Exit Function
    End If

    r = split(str,vbCrLf)
    l = uBound(r)
    i = -1

    Do While i < l
        i = i +1

        s = r(i)
        r(i) = ""

        Do While len(s) > m
            j = [==](c, 2)
            If is_empty(j) Then
                If preg_match("/¥S*(¥s)?$/",Left(s,m+1),matches,"","") Then
                    If len( trim(matches(0)) ) = 0 Then
                        j = m
                    Else
                        j = len( Left(s,m+1) ) - len(matches(0))
                    End If
                End If

                If is_empty(j) Then
                    j = [?]([==](c, true),m,false)
                End If

                If is_empty(j) Then
                    call preg_match("/^¥S*/",Mid(s,m),matches,"","")
                    j = len( Left(s,m) ) + len(matches(0))
                End If
            End If

            r(i) = r(i) & Left(s, j)
            s = Mid(s,j+1)
            r(i) = r(i) & [?](len(s), b , "")
        Loop

        r(i) = r(i) & s

    Loop

    wordwrap = join(r,vbCrLf)

End Function

'=======================================================================
'  MIME base64 方式によりエンコードされたデータをデコードする
'=======================================================================
'【引数】
'  data = mixed  デコードされるデータ。
'【戻り値】
'  もとのデータを返します。
'  失敗した場合は FALSE を返します。 返り値はバイナリになることもあります。
'【処理】
'  ・base64 でエンコードされた data  をデコードします。
'=======================================================================
Function base64_decode(data)

    Dim obj
    set obj=server.createobject("basp21")
    base64_decode = obj.Base64(data,1)
    set obj = nothing

    'BASP21を使用しない場合
'    Dim ST, DM, EL
'    Dim bin
' 
'    Set DM = CreateObject("Microsoft.XMLDOM")
'    Set EL = DM.createElement("tmp")
'    EL.DataType = "bin.base64"
'    EL.Text = Base64Text
'    bin = EL.NodeTypedValue
' 
'    Set ST = CreateObject("ADODB.Stream")
'    ST.Open
'    ST.Charset = "Shift-JIS"
'    ST.Type = adTypeBinary
'    ST.Write bin
'    ST.Position = 0
'    ST.Type = adTypeText
'    base64_decode = ST.ReadText
'    ST.Close

End Function

'=======================================================================
'  MIME base64 方式でデータをエンコードする
'=======================================================================
'【引数】
'  data = mixed  エンコードするデータ。
'【戻り値】
'  エンコードされたデータを文字列で返します。
'【処理】
'  ・指定した data  を base64 でエンコードします。
'=======================================================================
Function base64_encode(data)

    Dim obj
    set obj=server.createobject("basp21")
    base64_encode = obj.Base64(data,0)
    set obj = nothing

    'basp21を使用しない場合
'    Dim ST, DM, EL, bin
'  
'    Set ST = CreateObject("ADODB.Stream")
'    ST.Type = adTypeText
'    ST.Charset = "Shift-JIS"
'    ST.Open
'    ST.WriteText PlainText
'    ST.Position = 0
'    ST.Type = adTypeBinary
'    bin = ST.Read
'    ST.Close
' 
'    Set DM = CreateObject("Microsoft.XMLDOM")
'    Set EL = DM.CreateElement("tmp")
'    EL.DataType = "bin.base64"
'    EL.NodeTypedValue = bin
'    base64_encode = EL.Text

End Function

'=======================================================================
'  URL エンコードされたクエリ文字列を生成する
'=======================================================================
'【引数】
'  formdata       = array   配列もしくはオブジェクト。
'  numeric_prefix = string  配列の要素に対する数値インデックスの前にこれが追加されます。
'  arg_separator  = string  区分のためのセパレータとして使用されます。
'【戻り値】
'  URL エンコードされた文字列を返します。
'【処理】
'  ・与えられた連想配列 (もしくは添字配列) から URL エンコードされたクエリ文字列を生成します。
'=======================================================================
Function http_build_query(formdata , numeric_prefix , arg_separator )

    If Not isArray(formdata) and Not isObject(formdata) Then Exit Function

    Dim i,key
    Dim url
    Dim separator

    separator = "&"
    If len(arg_separator) > 0 then
        separator = arg_separator
    end if

    If isArray(formdata) Then
        For key = 0 to uBound(formdata)
            If isArray(formdata(key)) or isObject(formdata(key)) Then
                url = url & separator & http_build_query(formdata(key) , _
                                            numeric_prefix , arg_separator )
            else
                url = url & separator & _
                    key & "=" & Server.URLEncode(formdata(key))
            end if
        Next
    ElseIf isObject(formdata) Then

        For Each i In formdata
            if isArray(i) or isObject(i) then
                url = url & separator & http_build_query(i , numeric_prefix , arg_separator )

            elseif isArray(formdata(i)) or isObject(formdata(i)) Then

                If isArray( formdata(i) ) Then
                    For Each key In formdata(i)
                        If isObject(key) or isArray(key) Then
                            url = url & separator & http_build_query(key , numeric_prefix , arg_separator )
                        Else
                            url = url & separator & _
                                i & "=" & Server.URLEncode(key)
                        End If
                    Next
                Else
                    url = url & separator & http_build_query(formdata(i) , numeric_prefix , arg_separator )
                End If
            else
                if isArray( formdata ) and len(numeric_prefix) > 0 then
                        url = url & separator & _
                            numeric_prefix & i & "=" & Server.URLEncode(formdata(i))
                else
                    url = url & separator & _
                        i & "=" & Server.URLEncode(formdata(i))
                end if
            end if
        Next
    End If

    If len( url ) > 0 Then url = Mid(url,2)

    http_build_query = url

End Function

'=======================================================================
'  URLエンコードされた文字列をデコードする
'=======================================================================
'【引数】
'  sText   = string デコードする文字列。
'【戻り値】
'  デコードした文字列を返します。
'【処理】
'  ・与えられた文字列中のあらゆるエンコード文字 %## をデコードします。
'=======================================================================
function urldecode(sText)

    Dim obj
    Dim strDecode
    Dim strOutput

    set obj=server.createobject("basp21")
    strDecode = obj.Base64(sText,5)
    strOutput = Server.HTMLEncode(strDecode)
    set obj=Nothing

    URLDecode = strOutput

    'BASP21が使用できない場合
    '-----------------------
'    On Error Resume Next
'    sTmp=""
'    iCount = 1
'    lSrcLen=Len(Source)
'    Do Until iCount > lSrcLen
'        sChr = Mid(Source,iCount,1)
'        iCount = iCount+1
'        If sChr="+" Then
'            sChr = " "
'        ElseIf sChr="%" Then
'            sHex = Mid(Source,iCount,2)
'            iCount = iCount + 2
'            iAsc = CByte("&H" & sHex)
'            If (&H00 <= iAsc And iAsc <= &H80) Or _
'               (&HA0 <= iAsc And iAsc <= &HDF) Then
'                '1バイト文字
'                sChr=Chr(iAsc)
'            ElseIf (&H81 <= iAsc And iAsc <= &H9F) Or _
'               (&HE0 <= iAsc And iAsc <= &HFF) Then
'                '2バイト文字
'                sChr = Mid(Source,iCount,1)
'                iCount = iCount + 1
'                If sChr="%" Then
'                    sHex2 = Mid(Source,iCount,2)
'                    iCount = iCount + 2
'                Else
'                    sHex2 = Hex(Asc(sChr))
'                    If Len(sHex2) = 1 Then
'                        sHex2 = "0" & sHex2
'                    End If
'                End If
'                sChr=Chr(CInt("&H" & sHex & sHex2))
'            End If
'        End If
'        sTmp=sTmp & sChr
'    Loop
'    urldecode = sTmp
End function

'=======================================================================
'  文字列を URL エンコードする
'=======================================================================
'【引数】
'  str   = string エンコードする文字列。
'【戻り値】
'  -_. を除くすべての非英数文字が % 記号 (%)に続く二桁の数字で置き換えられ、 空白は + 記号(+)にエンコードされます。 同様の方法で、WWW のフォームからポストされたデータはエンコードされ、 application/x-www-form-urlencoded  メディア型も同様です。歴史的な理由により、この関数は RFC1738 エンコード( rawurlencode() を参照してください) とは異なり、 空白を + 記号にエンコードします。
'【処理】
'  ・この関数は、URL の問い合わせ部分に使用する文字列のエンコードや 次のページへ変数を渡す際に便利です。
'=======================================================================
Function urlencode(str)
    urlencode = Server.URLEncode(str)
End Function


'=======================================================================
'変数の float 値を取得する
'=======================================================================
'【引数】
'  str  = mixed  あらゆるスカラ型を指定できます。配列あるいはオブジェクトに floatval() を使用することはできません。
'【戻り値】
'  指定した変数の float 値を返します。
'【処理】
'  ・変数 str の float 値を返します。
'=======================================================================
Function floatval(str)

    floatval = false
    If isArray(str) or isObject(str) Then Exit Function
    If not isNumeric(str) Then Exit Function
    floatval = CDbl(str)

End Function

'=======================================================================
'変数の型を取得する
'=======================================================================
'【引数】
'  str  = mixed  型を調べたい変数。
'【戻り値】
'  型の文字列を返します。
'【処理】
'  ・変数 str の型を返します。
'=======================================================================
Function gettype(s)

    Select Case VarType(s)
    Case 0
        gettype = "vbEnpty"
    Case 1
        gettype = "vbNull"
    Case 2
        gettype = "vbInteger"
    Case 3
        gettype = "vbLong"
    Case 4
        gettype = "vbSingle"
    Case 5
        gettype = "vbDouble"
    Case 6
        gettype = "vbCurrency"
    Case 7
        gettype = "vbDate"
    Case 8
        gettype = "vbString"
    Case 9
        gettype = "vbObject"
    Case 10
        gettype = "vbError"
    Case 11
        gettype = "vbBoolean"
    Case 12
        gettype = "vbVariant"
    Case 13
        gettype = "vbDataObject"
    Case 17
        gettype = "vbByte"
    Case 8192
        gettype = "vbArray"
    Case 8204
        gettype = "vbArray"
    Case 8209
        gettype = "vbBinary"
    End Select

End Function

'=======================================================================
'変数の整数としての値を取得する
'=======================================================================
'【引数】
'  var = mixed 文字列
'【戻り値】
'  整数
'【処理】
'  ・var  の integer としての値を返します。
'=======================================================================
Function intval(str)

    intval = 1
    If IsObject(str) or IsArray(str) Then Exit Function
    If str = true Then Exit Function

    intval = 0
    If is_empty(str) or Not isNumeric(str) Then Exit Function

    str = int(str)
    If str > 32767 Then
        intval = 32767
    Else
        intval = Cint(str)
    End If

End Function

'=======================================================================
'変数が配列かどうかを検査する
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str が 配列型 の場合 TRUE、 そうでない場合 FALSE を返します。
'【処理】
'  ・与えられた変数が配列かどうかを検査します。
'=======================================================================
Function is_array(str)
    is_array = isArray(str)
End Function

'=======================================================================
'変数がネイティブバイナリ文字列かどうかを調べる
'=======================================================================
'【引数】
'  str   = mixed 調べる変数。
'【戻り値】
'  str がネイティブバイナリ文字列である場合に TRUE、それ以外の場合に FALSE を返します。
'【処理】
'  ・指定した変数が、ネイティブのバイナリ文字列かどうかを調べます。
'=======================================================================
Function is_binary(str)
    is_binary = (varType(str) = 8209)
End Function

'=======================================================================
'変数が boolean であるかを調べる
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str が boolean の場合 TRUE、 そうでない場合 FALSE を返します。
'【処理】
'  ・指定した変数が boolean であるかどうかを調べます。
'=======================================================================
Function is_bool(str)
    is_bool = (varType(str) = 11)
End Function

'=======================================================================
'is_float() のエイリアス
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【処理】
'  ・この関数は次の関数のエイリアスです。 is_float().
'=======================================================================
Function is_double(str)
    is_double = is_float(str)
End Function

'=======================================================================
'変数が空であるかどうかを検査する
'=======================================================================
'【引数】
'  s   = mixed チェックする変数
'【戻り値】
'  var が空でないか、0でない値であれば True を返します。
'【処理】
'  ・変数が空であるかどうかを検査する
'=======================================================================
Function is_empty(s)

    is_empty = false

    If isArray(s) Then
        If uBound(s) < 0 Then
            is_empty = true
            Exit Function
        Else
            Exit Function
        End If
    End If

    If isObject(s) Then
        If s.Count < 1 Then
            is_empty = true
            Exit Function
        Else
            Exit Function
        End If
    End If

    If isEmpty(s) or isNull(s) Then is_empty = true
    If s = empty Then is_empty = true

End Function

'=======================================================================
'変数の型が float かどうか調べる
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str が float型 の場合 TRUE、 そうでない場合 FALSE を返します。
'【処理】
'  ・与えられた変数の型が float かどうかを調べます。
'=======================================================================
Function is_float(str)
    is_float = (varType(str) = 5)
End Function

'=======================================================================
'変数が整数型かどうかを検査する
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str 整数型 の場合 TRUE、 そうでない場合は FALSE を返します。
'【処理】
'  ・与えられた変数の型が整数型かどうかを検査します。
'=======================================================================
Function is_int(str)

    is_int = false
    if Not isNumeric(str) Then Exit Function
    if str < 0 Then Exit Function
    is_int = (varType(str) = 2 or varType(str) = 3)

End Function

'=======================================================================
'is_int() のエイリアス
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【処理】
'  ・この関数は次の関数のエイリアスです。 is_int().
'=======================================================================
Function is_integer(str)
    is_integer = is_int(str)
End Function

'=======================================================================
'is_int() のエイリアス
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【処理】
'  ・この関数は次の関数のエイリアスです。 is_int().
'=======================================================================
Function is_long(str)
    is_long = is_int(str)
End Function

'=======================================================================
'変数が NULL かどうか調べる
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str が null の場合 TRUE、 そうでない場合は FALSE を返します。
'【処理】
'  ・指定した変数が NULL かどうかを調べます。
'=======================================================================
Function is_null(str)
    is_null = isNull(str)
End Function

'=======================================================================
'変数が数字または数値形式の文字列であるかを調べる
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str が数値または数値形式の文字列である場合 TRUE、 そうでない場合は FALSE を返します。
'【処理】
'  ・指定した変数が数値であるかどうかを調べます。
'  ・数値形式の文字列は以下の要素から なります。
'  ・（オプションの）符号、任意の数の数字、（オプションの）小数部、 そして（オプションの）指数部。つまり、+0123.45e6  は数値として有効な値です。
'  ・16 進表記（0xFF）は 認められません。
'=======================================================================
Function is_numeric(str)
    is_numeric = isNumeric(str)
End Function

'=======================================================================
'変数がオブジェクトかどうかを検査する
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str がobject 型の場合 TRUE、 そうでない場合は FALSE を返します。
'【処理】
'  ・与えられた変数がオブジェクトかどうかを調べます。
'=======================================================================
Function is_object(str)
    is_object = isObject(str)
End Function

'=======================================================================
'is_float() のエイリアス
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【処理】
'  ・この関数は次の関数のエイリアスです。 is_float().
'=======================================================================
Function is_real(str)
    is_real = is_float(str)
End Function

'=======================================================================
'変数がスカラかどうかを調べる
'=======================================================================
'【引数】
'  str   = mixed 評価する変数
'【戻り値】
'  str がスカラの場合 TRUE、 そうでない場合は FALSE を返します。
'【処理】
'  ・ 指定した変数がスカラかどうかを調べます。
'  ・ スカラ変数には integer、float、string あるいは boolean が含まれます。
'  ・ array、object および resource はスカラではありません。 
'=======================================================================
Function is_scalar(str)

    is_scalar = false
    If isArray(str) or isObject(str) Then Exit Function
    if isNull(str) Then Exit Function
    is_scalar = true

End Function

'=======================================================================
'変数の型が文字列かどうかを調べる
'=======================================================================
'【引数】
'  str   = mixed 評価する変数。
'【戻り値】
'  str の型が string である場合に TRUE、それ以外の場合に FALSE を返します。
'【処理】
'  ・指定した変数の型が文字列かどうかを調べます。
'=======================================================================
Function is_string(str)

    is_string = false
    If varType(str) = 8 Then is_string = true

End Function

'=======================================================================
'変数がセットされているか確認
'=======================================================================
'【引数】
'  val   = mixed 変数
'【戻り値】
'  val が存在すればTRUE、 そうでなければFALSEを返します。
'【処理】
'  ・変数がセットされているかどうかを調べます。
'=======================================================================
Function isset(val)

    isset = false
    If Not IsNull(val) Then isset = true

End Function

'=======================================================================
'指定した変数に関する情報を解りやすく出力する
'=======================================================================
'【引数】
'  expression   = mixed 表示したい式。
'  ret          = bool  print_r() はデフォルトでは結果を直接表示してしまいますが この引数が TRUE の場合には結果を戻します。
'【戻り値】
'  値が出力されます。
'【処理】
'  ・変数の値に関する情報を解り易い形式で表示します。
'=======================================================================
Function print_r(expression,ret)
    print_r = print_r_helper(expression,ret,0)
End Function

'*************************
Function print_r_helper(expression,ret,tab)

    If VarType(tab) <> 2 Then tab = 0
    If VarType(ret) <> 11 Then ret = false

    Dim strPrint

    If IsObject(expression) Then
        strPrint = strPrint & "Dictionary Object" & vbCrLf
    ElseIf IsArray(expression) Then
        strPrint = strPrint & "Array" & vbCrLf
    End If

    strPrint = strPrint & String(tab,vbTab) & "(" & vbCrLf

    Dim a,i
    i = 0
    If IsObject(expression) Then
        For Each a In expression
            strPrint = strPrint & String(tab,vbTab)
            If IsArray(a) or IsObject(a) Then
                strPrint = strPrint & vbTab & "[] => " & _
                           print_r_helper(a,true,tab + 1)
            ElseIf isArray(expression(a)) or isObject( expression(a) ) Then
                strPrint = strPrint & vbTab & "[" & a & "] => " & _
                           print_r_helper(expression(a),true,tab + 1)

            Else
               strPrint = strPrint & vbTab & ("[" & a & "]" & " => " & _
                          expression(a)) & vbCrLf
            End If
        Next
    ElseIf IsArray(expression) Then
        For Each a In expression
            strPrint = strPrint & String(tab,vbTab)
            If IsArray(a) or IsObject(a) Then
                strPrint = strPrint & vbTab & "[" & i & "] => " & _
                           print_r_helper(a,true,tab + 1)
            Else
                strPrint = strPrint & vbTab & ("[" & i & "] => " & a) & vbCrLf
            End If

            i =  i+1
        Next
    Else
        strPrint = strPrint & String(tab,vbTab) & expression & vbCrLf
    End If

    strPrint = strPrint & String(tab,vbTab) & ")" & vbCrLf

    If Not ret Then
        Response.Write strPrint
    Else
        print_r_helper = strPrint
    End If

End Function

'=======================================================================
'値の保存可能な表現を生成する
'=======================================================================
'【引数】
'  val   = mixed    シリアル化する値。
'【戻り値】
'  val  の保存可能なバイトストリーム表現を含む文字列を返します。 
'【処理】
'  ・ 値の保存可能な表現を生成します。
'  ・ 型や構造を失わずに ASP の値を保存または渡す際に有用です。
'  ・ シリアル化された文字列を ASP の値に戻すには、 unserialize() を使用してください。 
'=======================================================================
Function serialize(ByVal val)

    Dim strstrType
    strType = getType(val)

    Dim str
    Dim cnt : cnt = 0
    Dim strs
    Dim key

    Select Case strType

        Case "vbEnpty","vbNull"
            str = "N"
        Case "vbBoolean"
            str = "b:" & [?](val,1,0)
        Case "vbInteger","vbLong","vbSingle","vbDouble","vbCurrency"
            str = [?]([==](int(val),val),"i","d") & ":" & val
        Case "vbDate","vbString","vbVariant"
            str = "s:" & len(val) & ":""" & val & """"
        Case "vbArray"
            str = "a"

            For key = 0 to uBound(val)
                strs = strs & serialize(key) & _
                        serialize(val(key))
                cnt = cnt + 1
            Next
            str = str & ":" & cnt & ":{" & strs & "}"
            str = str & ";"

        Case "vbObject"
            str = "O"

            For Each key In val
                strs = strs & serialize(key) & _
                        serialize(val(key))
                cnt = cnt + 1
            Next
            str = str & ":" & cnt & ":{" & strs & "}"

        Case Else
            'empty
    End Select

    If strType <> "vbArray" AND strType <> "vbObject" Then
        str = str & ";"
    End If

    serialize = str
End Function

'=======================================================================
'変数の型をセットする
'=======================================================================
'【引数】
'  val   = mixed    破棄する変数。
'  type  = string   type  の値は以下の命令のいずれかです。
'【戻り値】
'  成功した場合に TRUE を、失敗した場合に FALSE を返します。
'【処理】
'  ・変数 str の型を type  にセットします。
'=======================================================================
Function settype(ByRef str,strType)

    settype = true

    Select Case strType
    Case "bool"
        str = CBool(str)
    Case "boolean"
        str = CBool(str)
    Case "byte"
        str = CByte(str)
    Case "currency"
        str = CCur(str)
    Case "date"
        str = CDate(str)
    Case "float"
        str = CDbl(str)
    Case "double"
        str = CDbl(str)
    Case "int"
        str = Cint(str)
    Case "integer"
        str = Cint(str)
    Case "long"
        str = CLng(str)
    Case "single"
        str = CSng(str)
    Case "string"
        str = Cstr(str)
    Case "array"
        If not isArray(str) Then
            str = array(str)
        End If
    Case "null"
        str = null
    Case Else
        settype = false
    End Select

End Function

'=======================================================================
'変数の文字列としての値を得ます
'=======================================================================
'【引数】
'  str   = mixed 文字列に変換したい変数。str は、全てのスカラー値にできます。 strval() に配列あるいはオブジェクトは使用できません。
'【戻り値】
'  str文字列値を返します。
'【処理】
'  ・strのstring としての値を返します。
'=======================================================================
Function strval(ByVal str)

    strval = false
    If isArray(str) or isObject(str) Then Exit Function
    strval = Cstr(str)

End Function

'=======================================================================
'指定した変数の割当を解除する
'=======================================================================
'【引数】
'  val   = mixed 破棄する変数。
'【戻り値】
'  値を返しません。
'【処理】
'  ・指定した変数を破棄します。
'=======================================================================
Function unset(ByRef val)

    If isObject(val) Then
        set val = Nothing
    Else
        val = null
    End If

End Function

'=======================================================================
'変数に関する情報をダンプする
'=======================================================================
'【引数】
'  val   = mixed 破棄する変数。
'【戻り値】
'  値を返しません。
'【処理】
'  ・この関数は、指定した式に関してその型や値を含む構造化された情報を 返します。
'  ・配列の場合、その構造を表示するために各値について再帰的に 探索されます。
'=======================================================================
Sub var_dump(expression)
    var_dump_helper expression,0
End Sub

'***************************
Sub var_dump_helper(expression,tab)

    If VarType(tab) <> 2 Then tab = 0

    Dim strTab : strTab = String(tab,vbTab)

    If IsObject(expression) Then
        Response.Write "Dictionary Object(" & expression.count & ")" & vbCrLf
    ElseIf IsArray(expression) Then
        Response.Write "Array(" & (uBound(expression)+1) & ")" & vbCrLf
    End If

    Response.Write strTab & "(" & vbCrLf

    Dim a,i
    i = 0
    If IsObject(expression) Then
        For Each a In expression
            Response.Write strTab
            If IsArray(a) or IsObject(a) Then
                Response.Write vbTab & "[] => "
                call var_dump_helper(a,tab + 1)
            ElseIf isArray(expression(a)) or isObject( expression(a) ) Then
                Response.Write vbTab & "[""" & a & """] => "
                call var_dump_helper(expression(a),tab + 1)

            Else
               Response.Write vbTab & "[""" & a & """]" & " => " & _
                              gettype(expression(a)) & "(" & expression(a) & ")" & vbCrLf
            End If
        Next
    ElseIf IsArray(expression) Then
        For Each a In expression
            Response.Write strTab
            If IsArray(a) or IsObject(a) Then
                Response.Write vbTab & "[" & i & "] => "
                call var_dump_helper(a,tab + 1)
            Else
                Response.Write vbTab & "[" & i & "] => " & _
                               gettype(a) & "(" & a & ")" & vbCrLf
            End If

            i =  i+1
        Next
    Else
        Response.Write strTab & gettype(expression) & "(" & expression & ")" & vbCrLf
    End If

    Response.Write strTab & ")" & vbCrLf

End Sub


'=======================================================================
'XMLファイルをパースし、配列に代入する
'=======================================================================
'【引数】
'  filename     = String XML ファイルへのパス。
'  encode       = String XML ファイルの文字コード。
'【戻り値】
'  ドキュメント内のデータを返します。
'  エラー時には エラーメッセージを返します。
'【処理】
'  ・指定したファイルの中の整形式 XML 配列に変換します。
'=======================================================================
Function simplexml_load_file(filename,encode)

    Dim objDoc,result
    Dim ret,strXml,rtResult,xPE

    Set ret = Server.CreateObject("Scripting.Dictionary")

    If encode <> "Shift_JIS" and encode <> "sjis" Then _
        Set objDoc = Server.CreateObject("MSXML2.DOMDocument") _
    Else _
        Set objDoc = Server.CreateObject("MSXML.DOMDocument")

    objDoc.async = true

    If inStr(filename,"http://") = 1 Then

        Dim file
        set file = new File_System
        strXml = file.file_get_contents(filename)
        rtResult = objDoc.LoadXML(strXml)
    Else
        rtResult = objDoc.Load(filename)
    End If

    Set xPE = objDoc.parseerror
    If xPE.errorcode <> 0 then
        ret("error") = xPE
        set simplexml_load_file = ret
        Exit Function
    End If

    If rtResult = True Then
        call simplexml_parse(objDoc.childNodes, ret)
    Else
        ret("error") = "XMLを取得できません。"
    End If

    Set simplexml_load_file = ret
    Set objDoc = Nothing

End Function

Function simplexml_parse(objNode,ByRef ret)

    Dim obj,tmp_ob,tmp_ar()
    Dim intCounter,objData,att
    Dim counter,j

    If Not isObject(ret) Then Set ret = Server.CreateObject("Scripting.Dictionary")
    Set counter = Server.CreateObject("Scripting.Dictionary")

    ReDim tmp_ar(objNode.length-1)
    For j = 0 to (objNode.length-1)
        tmp_ar(j) = objNode(j).nodeName
    Next

    Set tmp_ob = array_count_values(tmp_ar)

    For Each obj In tmp_ob
        counter.Add obj, 0
    Next

    For Each obj In objNode

        objData = obj.nodeName

        If obj.nodeTypeString = "element" Then
            If obj.attributes.length > 0 Then
                Set ret(objData & "_attr") = Server.CreateObject("Scripting.Dictionary")

                For Each att IN obj.attributes
                    ret(objData & "_attr").Add _
                        preg_replace("/(.+)="".*""/","$1",att.xml), att.value
                Next
            End If
        End If

        If obj.hasChildNodes Then

            If obj.childNodes.length = 1 and (obj.childNodes(0).nodeName = "#text" or _
               obj.childNodes(0).nodeName = "#cdata-section") Then
                If Not ret.Exists( objData ) Then ret.Add objData , obj.text
            Else

                If Not isObject(ret(objData)) Then _
                    Set ret(objData) = Server.CreateObject("Scripting.Dictionary")
                If tmp_ob(objData) = 1 Then
                    call simplexml_parse(obj.childNodes, ret(objData))
                Else

                    If Not isObject(ret(objData)(counter(objData))) Then _
                        Set ret(objData)(counter(obj.nodeName)) = _
                            Server.CreateObject("Scripting.Dictionary")

                    call simplexml_parse(obj.childNodes, ret(objData)(counter(objData)))
                    counter(objData) = counter(objData) +1
                End If
           End If

        Else
            If Not ret.Exists( objData ) Then ret.Add objData , obj.text
        End If

    Next
End Function


%>
