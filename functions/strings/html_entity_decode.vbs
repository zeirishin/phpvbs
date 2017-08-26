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
