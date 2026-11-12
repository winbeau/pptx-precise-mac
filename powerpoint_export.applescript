on normalizedPath(rawPathValue)
    set pathText to rawPathValue as text
    if pathText starts with "/" then return pathText
    return POSIX path of (pathText as alias)
end normalizedPath

on run argv
    if (count of argv) is not 4 then error "Expected snapshot, PDF, original path and timeout"
    set inputPath to item 1 of argv
    set outputPath to item 2 of argv
    set originalPath to item 3 of argv
    set waitSeconds to (item 4 of argv) as integer
    set inputAlias to POSIX file inputPath as alias
    set outputFile to POSIX file outputPath as text
    set openedPresentation to missing value
    with timeout of waitSeconds seconds
        tell application "Microsoft PowerPoint"
            repeat with existingIndex from 1 to (count of presentations)
                set existingPresentation to presentation existingIndex
                set existingPath to ""
                try
                    set existingFile to full name of existingPresentation
                    set existingPath to my normalizedPath(existingFile)
                end try
                if existingPath is originalPath and saved of existingPresentation is false then
                    error "请先保存 PowerPoint 中的修改，再重新转换：" & originalPath
                end if
            end repeat
            try
                open inputAlias
                repeat with candidateIndex from 1 to (count of presentations)
                    set candidatePresentation to presentation candidateIndex
                    set candidateFile to full name of candidatePresentation
                    if my normalizedPath(candidateFile) is inputPath then
                        set openedPresentation to candidatePresentation
                        exit repeat
                    end if
                end repeat
                if openedPresentation is missing value then error "PowerPoint 未打开指定的临时副本，已停止导出。"
                save openedPresentation in outputFile as save as PDF
                do shell script "/bin/test -s " & quoted form of outputPath
                close openedPresentation saving no
                set openedPresentation to missing value
            on error errorText number errorNumber
                if openedPresentation is not missing value then
                    try
                        close openedPresentation saving no
                    end try
                end if
                error errorText number errorNumber
            end try
        end tell
    end timeout
    return outputPath
end run
