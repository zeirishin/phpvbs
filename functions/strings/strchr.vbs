<%
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
%>
