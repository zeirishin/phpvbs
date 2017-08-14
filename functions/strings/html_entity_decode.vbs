
function htmlspecialchars_decode(str,quote_style)

    Dim I
    Dim sText

    if empty_(quote_style) then quote_style = ENT_COMPAT
    sText = str

    if quote_style <> ENT_NOQUOTES then
        sText = Replace(sText, """, Chr(34))
    end if

    sText = Replace(sText, "<"  , Chr(60))
    sText = Replace(sText, ">"  , Chr(62))
    sText = Replace(sText, "&" , Chr(38))
    sText = Replace(sText, " ", Chr(32))

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
