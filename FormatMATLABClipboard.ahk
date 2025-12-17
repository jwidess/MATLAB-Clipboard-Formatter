#Requires AutoHotkey v2.0

; Start Notification
MsgBox("`nPress F2 to fix and paste.", "MATLAB Formatter", 64)

; Main Hotkey F2
F2::
{
    ; Get HTML from Clipboard
    fullHtml := GetClipboardHTML()
    if (fullHtml = "")
    {
        ToolTip("Error: No HTML on clipboard")
        SetTimer () => ToolTip(), -2000
        return
    }

    ; Extract CSS and Body
    cssContent := ""
    if RegExMatch(fullHtml, "s)<style[^>]*>(.*?)</style>", &match)
        cssContent := match[1]
    
    bodyContent := ""
    if RegExMatch(fullHtml, "s)<body[^>]*>(.*?)</body>", &match)
        bodyContent := match[1]
    else
        bodyContent := fullHtml 

    ; Clean up fragment markers
    bodyContent := RegExReplace(bodyContent, "<!--StartFragment-->")
    bodyContent := RegExReplace(bodyContent, "<!--EndFragment-->")
    bodyContent := RegExReplace(bodyContent, "<!--\[if mso\].*?<!\[endif\]-->")

    ; Remove BR tags
    bodyContent := RegExReplace(bodyContent, 'i)<br\s*/?>', '')
    
    ; Define styles
    forcedStyles := "font-size: 8pt; line-height: 6pt; margin: 0; padding: 0;"
    
    ; Define quotes
    q := Chr(34)  ; Double Quote (")
    sq := Chr(39) ; Single Quote (')
    
    ; Handle Empty Lines ---
    ; MATLAB has empty lines that are often: <div class="S..."><span></span></div>
    ; Replace the inner empty span with a styled non breaking space
    
    ; Force style on 'inlineWrapper' containers (lines of code)
    quoteClass := "[" . q . sq . "]"
    wrapperNeedle := "i)class\s*=\s*" . quoteClass . "inlineWrapper" . quoteClass
    wrapperReplacement := "style=" . q . forcedStyles . q
    bodyContent := RegExReplace(bodyContent, wrapperNeedle, wrapperReplacement)

    ; Fix pure empty lines (that might not be inlineWrapper)
    ; Replace inner empty spans with a sized space to ensure the line has height.
    ; BEFORE inlining other classes.
    
    ; Find <span ...></span> with nothing inside
    emptySpanNeedle := "i)<span[^>]*>\s*</span>"
    ; Replace with a span containing a space, styled to 6pt
    filledSpanReplacement := "<span style=" . q . forcedStyles . q . ">&nbsp;</span>"
    
    bodyContent := RegExReplace(bodyContent, emptySpanNeedle, filledSpanReplacement)


    ; Parse CSS classes from the header into a Map
    classMap := Map()
    Loop Parse, cssContent, "}"
    {
        if RegExMatch(A_LoopField, "s)\.([a-zA-Z0-9_\-]+)\s*\{(.*)", &cssMatch)
        {
            className := cssMatch[1]
            styleRules := cssMatch[2]
            
            styleRules := RegExReplace(styleRules, "\s+", " ")
            styleRules := Trim(styleRules)
            styleRules := StrReplace(styleRules, '"', "'") ; Sanitize quotes
            
            ; Remove existing font and line-height rules
            styleRules := RegExReplace(styleRules, "i)font-size:[^;]+;?", "")
            styleRules := RegExReplace(styleRules, "i)line-height:[^;]+;?", "")
            
            classMap[className] := styleRules
        }
    }

    ; Inline the styles for the specific code spans
    For className, rules in classMap
    {
        newStyle := rules . "; " . forcedStyles
        newStyle := StrReplace(newStyle, ";;", ";")
        
        needle := "i)class\s*=\s*" . quoteClass . className . quoteClass
        replacement := "style=" . q . newStyle . q
        
        bodyContent := RegExReplace(bodyContent, needle, replacement)
    }
    
    ; Wrap in a master div
    finalBody := '<div style="font-family: Consolas, ' . sq . 'Courier New' . sq . ', monospace; color: rgb(33, 33, 33); ' . forcedStyles . '">' . bodyContent . '</div>'

    ; Set clipboard content and paste
    if SetClipboardHTML(finalBody)
    {
        Sleep 100 
        Send "^v"
    }
    else
    {
        MsgBox("Failed to update clipboard.", "Error", 16)
    }
}

; Helper Functions ---

GetClipboardHTML() {
    cf_html := DllCall("RegisterClipboardFormat", "Str", "HTML Format", "UInt")
    if !DllCall("IsClipboardFormatAvailable", "UInt", cf_html)
        return ""
    if !DllCall("OpenClipboard", "Ptr", 0)
        return ""
    hData := DllCall("GetClipboardData", "UInt", cf_html, "Ptr")
    if !hData {
        DllCall("CloseClipboard")
        return ""
    }
    pData := DllCall("GlobalLock", "Ptr", hData, "Ptr")
    html := StrGet(pData, "UTF-8")
    DllCall("GlobalUnlock", "Ptr", hData)
    DllCall("CloseClipboard")
    return html
}

SetClipboardHTML(HtmlBody, HtmlHead := "") {
    Local CF_HTML := DllCall("RegisterClipboardFormat", "Str", "HTML Format", "UInt")
    Local Html := "Version:0.9`r`nStartHTML:000000000`r`nEndHTML:000000000`r`nStartFragment:000000000`r`nEndFragment:000000000`r`n<!DOCTYPE>`r`n<html>`r`n<head>`r`n" . HtmlHead . "`r`n</head>`r`n<body>`r`n<!--StartFragment-->" . HtmlBody . "<!--EndFragment-->`r`n</body>`r`n</html>"
    
    Html := StrReplace(Html, "StartHTML:000000000", Format("StartHTML:{:09}", InStr(Html, "<html>")-1))
    Html := StrReplace(Html, "EndHTML:000000000", Format("EndHTML:{:09}", InStr(Html, "</html>")-1))
    Html := StrReplace(Html, "StartFragment:000000000", Format("StartFragment:{:09}", InStr(Html, "<!--StartFrag")-1))
    Html := StrReplace(Html, "EndFragment:000000000", Format("EndFragment:{:09}", InStr(Html, "<!--EndFragme")-1))
    
    BufHtml := Buffer(StrPut(Html, "UTF-8"), 0)
    StrPut(Html, BufHtml, "UTF-8")
    
    if !DllCall("OpenClipboard", "Ptr", 0)
        return 0
    DllCall("EmptyClipboard")
    
    hMemHtml := DllCall("GlobalAlloc", "UInt", 0x42, "UPtr", BufHtml.Size, "Ptr")
    pMemHtml := DllCall("GlobalLock", "Ptr", hMemHtml, "Ptr")
    DllCall("RtlMoveMemory", "Ptr", pMemHtml, "Ptr", BufHtml.Ptr, "UPtr", BufHtml.Size)
    DllCall("GlobalUnlock", "Ptr", hMemHtml)
    DllCall("SetClipboardData", "UInt", CF_HTML, "Ptr", hMemHtml)
    DllCall("CloseClipboard")
    return 1
}
