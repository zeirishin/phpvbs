<%
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

%>
