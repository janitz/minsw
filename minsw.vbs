'parameters
const MAXSIZE = 18
const MINSIZE = 5
const MINES = 10 '%

const BOMB = "*"
const NOBOMB = "0"
const EMPTYFIELD = " "

'enable console by changing the interpreter from wscript to cscript
set oWSH = CreateObject("WScript.Shell")
vbsInterpreter = "cscript.exe"
if inStr(LCase(WScript.FullName), vbsInterpreter) = 0 then
    oWSH.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
    WScript.Quit
end if

sub printf(txt)
    WScript.StdOut.Write txt
end sub

function scanf(txt)
    printf txt
    scanf = LCase(WScript.StdIn.ReadLine)
end function





call main()

sub main()
    'get the playground size
    set pgSize = getPlaygroundSize()

    'create the playground
    set pg = new playground
    pg.init(pgSize)

    startTime = 0
    gameOver = false
    info = "naechster zug"
    do until gameOver
        if info = "verlorn" then gameOver = true
        if info = "gewonne" then gameOver = true
        printf pg.toString()
        printf info & vbCrLf
        if gameOver then
            scanf(timer - startTime & "s " & vbCrLf)
        else
            set inp = getInput()
            if startTime = 0 then startTime = timer
            info = pg.show(inp.x, inp.y)
            if pg.won then info = "gewonne"
        end if
    loop

    pg.showAll
    scanf(pg.toString())

end sub

'asks for the playground size and returns a vector object
function getPlaygroundSize()

    'create the vector object
    set getPlaygroundSize = new vector

    'ask for the size
    text = scanf("sag mal die groesse" & vbCrLf & "(std is 10x10)" & vbCrLf)

    'split text in two strings which only contain numbers
    foundNum = 0
    lastCharNum = false
    for cnt = 1 to len(text)
        if isNumeric(mid(text, cnt, 1)) then
            if not lastCharNum then foundNum = foundNum + 1

            if foundNum = 1 then xStr = xStr & mid(text, cnt, 1)
            if foundNum = 2 then yStr = yStr & mid(text, cnt, 1)

            lastCharNum = true
        else
            lastCharNum = false
        end if
    next

    'convert string to long
    xSize = CLng(xStr)
    ySize = CLng(yStr)

    'set to std
    if xSize = 0 then xSize = 10
    if ySize = 0 then ySize = 10

    'if not valid, set to valid size
    if xSize < MINSIZE then xSize = MINSIZE
    if xSize > MAXSIZE then xSize = MAXSIZE
    if ySize < MINSIZE then ySize = MINSIZE
    if ySize > MAXSIZE then ySize = MAXSIZE

    'save size to the vector object to return
    getPlaygroundSize.x = xSize
    getPlaygroundSize.y = ySize

end function

'asks for a field and returns a vector object
function getInput()
    'create the vector object
    set getInput = new vector
    getInput.x = 0
    getInput.y = 0

    'ask for the field
    text = LCase(scanf("wo? "))

    foundNum = 0
    lastCharNum = false
    foundValidChar = false
    for cnt = 1 to len(text)
        c = mid(text, cnt, 1)
        if isNumeric(c) then
            if not lastCharNum then	foundNum = foundNum + 1

            if foundNum = 1 then yStr = yStr & c

            lastCharNum = true
        else
            lastCharNum = false

            if not foundValidChar then
                for cCnt = 1 to 26
                    if c = chr(96 + cCnt) then
                        foundValidChar = true
                        getInput.x = cCnt
                    end if
                next
            end if

        end if
    next

    getInput.y = CLng(yStr)

end function

class vector
    dim x
    dim y
end class

class playground

    dim fields()
    dim visible()
    dim size

    sub init(size_)
        'init the class varaibles
        set size = size_
        redim fields(size.y - 1, size.x - 1)
        redim visible(size.y - 1, size.x - 1)

        for yCnt = 0 to size.y - 1
            for xCnt = 0 to size.x -1
                visible(yCnt, xCnt) = false
                fields(yCnt, xCnt) = EMPTYFIELD
            next
        next

        setMines
        calcNumbers
    end sub

    sub setMines()

        'calculate how many mines
        amountFields = size.x * size.y
        amountMines = amountFields * MINES / 100

        Randomize

        'place the mines
        for cnt = 1 to amountMines
            found = false
            do until found
                xPos = int(rnd * size.x)
                yPos = int(rnd * size.y)

                if not fields(yPos, xPos) = BOMB then
                    fields(yPos, xPos) = BOMB
                    found = true
                end if
            loop
        next

    end sub

    sub calcNumbers()

        for yCnt = 0 to size.y - 1
            for xCnt = 0 to size.x -1
                if fields(yCnt, xCnt) = EMPTYFIELD then
                    mineCnt = 0
                    for yOff = -1 to 1
                        for xOff = -1 to 1
                            if yCnt + yOff >= 0 and yCnt + yOff < size.y then
                                if xCnt + xOff >= 0 and xCnt + xOff < size.x then
                                    if fields(yCnt + yOff, xCnt + xOff) = BOMB then
                                        mineCnt = mineCnt + 1
                                    end if
                                end if
                            end if
                        next
                    next
                    if mineCnt = 0 then
                        fields(yCnt, xCnt) = NOBOMB
                    else
                        fields(yCnt, xCnt) = "" & mineCnt
                    end if
                end if
            next
        next

    end sub

    function show(x, y)
        show = "naechster zug"
        fail = false

        'if not valid
        if x < 1 or x > size.x then fail = true
        if y < 1 or y > size.y then fail = true

        if fail then
            show = "ausm feld drausen"
        else
            if visible(y - 1, x - 1) then
                show = "hasch schon getestet"
            else
                visible(y - 1, x - 1) = true
                if fields(y - 1, x - 1) = BOMB then
                    show = "verlorn"
                elseif fields(y - 1, x - 1) = NOBOMB then
                    showMore x - 1, y - 1
                end if
            end if
        end if
    end function

    sub showMore(x,y)
        for yOff = -1 to 1
            for xOff = -1 to 1
                if y + yOff >= 0 and y + yOff < size.y then
                    if x + xOff >= 0 and x + xOff < size.x then
                        if not visible(y + yOff, x + xOff) then
                            visible(y + yOff, x + xOff) = true
                            if fields(y + yOff, x + xOff) = NOBOMB then showMore x + xOff, y + yOff
                        end if
                    end if
                end if
            next
        next
    end sub

    sub showAll()
        for yCnt = 0 to size.y - 1
            for xCnt = 0 to size.x -1
                visible(yCnt, xCnt) = true
            next
        next
    end sub

    function won()
        won = true
        for yCnt = 0 to size.y - 1
            for xCnt = 0 to size.x -1
                if visible(yCnt, xCnt) and fields(yCnt, xCnt) = BOMB then won = false
                if not visible(yCnt, xCnt) and not fields(yCnt, xCnt) = BOMB then won = false
                if not won then exit for
            next
        next
    end function

    function toString()
        ret = vbCrLf & vbCrLf & vbCrLf & vbCrLf

        'firstline
        ret = ret & " "
        for xCnt = 0 to size.x - 1
            ret = ret & " " & chr(xCnt + 97) & " "
        next
        ret = ret & vbCrLf


        'splitline
        splitline = " +"
        for xCnt = 0 to size.x - 1
            splitline = splitline & "---+"
        next

        ret = ret & splitline & vbCrLf

        for yCnt = 0 to size.y - 1
            ret = ret & yCnt + 1
            if yCnt < 9 then
                ret = ret & " "
            else
                ret = ret & " "
            end if

            ret = ret & "| "
            for xCnt = 0 to size.x - 1
                if visible(yCnt, xCnt) then
                    ret = ret & fields(yCnt, xCnt) & " | "
                else
                    ret = ret & EMPTYFIELD & " | "
                end if
            next
            ret = ret & vbCrLf & splitline & vbCrLf
        next
        toString = ret
    end function

end class

