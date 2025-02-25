on run {input, parameters}
    set theOutput to {}
    
    repeat with i in input
        set t to i as string
        
        if not (t ends with ".ppt" or t ends with ".pptx") then
            set fileName to (do shell script "basename " & quoted form of (POSIX path of t))
display alert "Invalid File" message "The file '" & fileName & "' is not a PowerPoint (.ppt or .pptx) file. Exiting script."
            return
        end if
    end repeat
    
    tell application "Microsoft PowerPoint"
        launch
        
        repeat with i in input
            set t to i as string
            set pdfPath to my makeNewPath(i)
            
            try
                open i
                set activePres to active presentation
                
                save activePres in (POSIX file pdfPath) as save as PDF
                close active presentation saving no
                set end of theOutput to pdfPath
            on error errMsg
                set fileName to (do shell script "basename " & quoted form of (POSIX path of t))
display alert "Error" message "Failed to convert '" & fileName & "': " & errMsg
                return
            end try
        end repeat
        
    end tell
    
    tell application "Microsoft PowerPoint"
        quit
    end tell
    
    return theOutput
end run

on makeNewPath(f)
    set t to f as string
    if t ends with ".pptx" then
        return (POSIX path of (text 1 thru -6 of t)) & ".pdf"
    else
        return (POSIX path of (text 1 thru -5 of t)) & ".pdf"
    end if
end makeNewPath
