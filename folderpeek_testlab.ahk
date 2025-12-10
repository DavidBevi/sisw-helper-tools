; ██████  █████  ██     █████  ██████ █████  █████  ██████ ██████ ██  ██                ██████ ██████  █████ ██████ ██      ████  █████  
; ██     ██   ██ ██     ██  ██ ██     ██  ██ ██  ██ ██     ██     ██ ██      ██  ██       ██   ██     ██       ██   ██     ██  ██ ██  ██ 
; █████  ██   ██ ██     ██  ██ ████   █████  █████  ████   ████   ████         ██         ██   ████    ████    ██   ██     ██████ █████  
; ██     ██   ██ ██     ██  ██ ██     ██  ██ ██     ██     ██     ██ ██      ██  ██       ██   ██         ██   ██   ██     ██  ██ ██  ██ 
; ██      █████  ██████ █████  ██████ ██  ██ ██     ██████ ██████ ██  ██                  ██   ██████ █████    ██   ██████ ██  ██ █████  
                                                                                                                                               
;2025-12-10: this script now contains many different utilities, organized in subsections inside logical sections (settings, hotkeys, funcs...)
;            they should be independent, but I'm not sure

;FOLDERPEEK: periodically checks if the mouse cursor is over a file explorer window, if so it displays the content of the folder under the cursor
;            txt files are also displayed, unless they are a diag report, in which case only the table with frame and modules SN is displayed
;            if 7-Zip is installed it also displays the content of 7z files

;CALLAUNCHER: a little window in the bottom right corner with buttons to open calibration programs and reports
;             the window should auto-open when a new ethernet device (a system) is detected (buggy)
;             manually open with F3, manually close with ESC

;HELPERGUIS: 3 little transparent boxes with orange text will appear over Adjustment+Calibration when it's open and focused
;            if you drag them over the NEXT and PREV buttons you'll be able to use the keyboard to click said buttons
;            it's a hack because of how Adj+Cal is made, but it's nice and it should be robust

;SYSTEMSDB: a database of every TAC/Order, used to show the systems details in a small gui when you copy a serial number in microsoft word
;           systems have to be imported manually by copying the tables from emails from Sabrina, then clicking on tray icon → Systems DB tools
;           also in tray menu → Systems DB tools → you can open a gui containing the entire systems db to see it / edit it

;MISC: in tray menu → pick one of the folders open in file explorer and change its icon to folder-with-arrow, to mark it as sent




;▼ (DOUBLE-CLICK) RELOAD THIS SCRIPT
~F2::(A_ThisHotkey=A_PriorHotkey && A_TimeSincePriorHotkey<200)? Reload(): {}




;▼ STATIC SETTINGS ###################################################################################
#Requires AutoHotkey v2.0
#SingleInstance Force
CoordMode("Mouse")
OnError((*)=>-1)
; Try TraySetIcon("Shell32.dll", 309)
FileExist(pathof_7z:=A_ProgramFiles "\7-Zip\7z.exe")? {} :
FileExist(pathof_7z:=A_Programs     "\7-Zip\7z.exe")? {} : 0




;▼ HOTKEYS ###################################################################################

;▼ summon / hide CALLAUNCHER gui
F3::CalLauncher("show"), AdjCalHotkeys()
~Esc::CalLauncher("hide")

;▼ use and manage HELPERGUIS
#HotIf WinActive("ahk_exe TestSoftwareNext.exe")
    PgUp:: Send("^a")
    PgDn:: AdjCalHotkeys(1)
    End::  AdjCalHotkeys(2)
    ~LButton:: MoveHelperGuis()

;▼ when you copy a serial number in ms-word this show its details (from SYSTEMSDB)
#HotIf WinActive("ahk_exe WINWORD.EXE")
    ~^c:: Sleep(20), EditSingleSystem(A_Clipboard)




;▼ MAIN ###################################################################################

;▼ tray icon 
_f := FileOpen(A_Temp "\f", "w")
For ch in StrSplit("ƉŐŎŇčĊĚĊĀĀĀčŉňńŒĀĀĀĒĀĀĀĒĈĆĀĀĀŖǎƎŗĀĀĀāųŒŇłĀƮǎĜǩĀĀĀĄŧŁōŁĀĀƱƏċǼšąĀĀĀĉŰňřųĀĀĐǪĀĀĐǪāƂǓĊƘĀĀĀƉŉńŁŔĸŏǍǒǁčƀĠČąǐƖƩǜǀƳěŸŤĤƶŰčƧĒĩũēƔƂōĤǆŷǺĚǽƴőĘąţǂƹċēƎƪŜǤłǠǋƶǃŻŎĺųőďĝƒƋŰŚǸŖčǗƙœƛƩǈĢǮěĸǎƯŝĦƢǦĻǋƴǕńǴƒƼŘǦƒŶĘęƶƚƹƈĦƑũǊĬǌŅŏūſƿƚǐľĀőƋŚďǷČśĭǿƐƜǿĀǠĄƧƑĿƴǈƐǾĨĀĀĀĀŉŅŎńƮłŠƂ")
    _f.RawWrite(StrPtr(Chr(Mod(Ord(ch),256))),1)
_f.Close(), TraySetIcon(A_Temp "\f")
A_IconTip:="𝐅𝐎𝐋𝐃𝐄𝐑𝐏𝐄𝐄𝐊 ⨉ TESTLAB"

;▼ tray menu
A_TrayMenu.Delete()
msaio(m,s?,a?,i?,o?)=>(m.Add(s?,a?,o?), (i??0)? m.SetIcon(s,"Shell32.dll",i) :{})
msapio(m,s?,a?,p?,i?,o?)=>(m.Add(s?,a?,o?), (i??0)? m.SetIcon(s,p,i) :{})

msapio(A_TrayMenu, "Change folder icon...", SetSentIcon, "imageres.dll", 280)
msaio(A_TrayMenu)
menuDB:=Menu(), msaio(menuDB, "Copy clipboard in DB", ImportFromClipboad, 298)
                msaio(menuDB, "Show Systems DB", ShowDbGui, 98)
msaio(A_TrayMenu, "Systems DB tools", menuDB, 335)
msaio(A_TrayMenu)
msapio(A_TrayMenu, "Run Diagnostic", run1, "Shell32.dll", 262)
msaio(A_TrayMenu, "Run Adj+Cal", run2)
msaio(A_TrayMenu)
msapio(A_TrayMenu, "Diag Reports", run3, "imageres.dll", 246)
msaio(A_TrayMenu, "AC Reports", run4)
msaio(A_TrayMenu)
msaio(A_TrayMenu, "Reload", (*)=>Reload())
msaio(A_TrayMenu, "Exit", (*)=>ExitApp())

;▼ FOLDERPEEK listener
SetTimer(FolderPeek, 60)

;▼ CALLAUNCHER listener
SetTimer(CalLauncher, 1000)

;▼ HELPERGUIS: make variables to manage the guis
ACguis:=[]
ACpos:=GetPosFromIni()

;▼ HELPERGUIS: hook function to EVENT_SYSTEM_FOREGROUND (0x3) so it runs when focused-window changes
DllCall("SetWinEventHook", "UInt",0x3, "UInt",0x3, "Ptr",0, "Ptr",CallbackCreate(AdjCalHotkeys), "UInt",0, "UInt",0, "UInt",0)




;▼ FUNCS ###################################################################################

;▼ FOLDERPEEK: preview contents of folders and some filetypes in file explorer
FolderPeek(*) {
    Static oldx:=0, oldy:=0, contentof:=Map(), cache:=["",""] ;[path,contents]
    Static dif:= [Ord("𝟎")-Ord("0"), Ord("𝐚")-Ord("a"), Ord("𝐀")-Ord("A"), 0]
    If !WinActive("ahk_exe explorer.exe") && cache="" {
        Return
    }
    MouseGetPos(&x,&y)
    If x=oldx && y=oldy {
        Return
    } Else oldx:=x, oldy:=y
    path:=ExplorerGetHoveredItem()

    ;▼ folder
    If cache[1]!=path && DirExist(path) {
        dirs:="", contentof[path]:=""
        Loop Parse StrSplit(path,"\")[-1]  ;► 𝐟𝐚𝐧𝐜𝐲 𝐛𝐨𝐥𝐝 𝐟𝐨𝐥𝐝𝐞𝐫-𝐧𝐚𝐦𝐞
            dirs.=Chr(Ord(A_LoopField)+dif[A_LoopField~="[0-9]"?1: 
                A_LoopField~="[a-z]"?2: A_LoopField~="[A-Z]"?3: 4])
        Loop Files, path "\*.*", "DF"
            ; If A_LoopFileName="desktop.ini"
            ;     Continue
            If A_Index<46
                DirExist(path "\" A_LoopFileName)? dirs.="`n🖿 " A_LoopFileName:
                    contentof[path].="`n     " A_LoopFileName
            Else overflow:="`n`n    (+ " A_Index-45 ")"
        cache:=[path, dirs contentof[path] (IsSet(overflow)? overflow: "")]

    ;▼ 7z
    } Else If path~="\.7z$" && IsSet(pathof_7z) {
        If !(contentof.Has(path)) {
            guipeek("LOADING 7z...","x" x+16 " y" y+20)
            files:="", contentof[path]:=""
            Try contents:=ComObject("WScript.Shell").Exec('"' pathof_7z '" l -ba "' path '"').StdOut.ReadAll()
            Loop Parse StrSplit(path,"\")[-1]  ;► 𝐟𝐚𝐧𝐜𝐲 𝐛𝐨𝐥𝐝 𝐚𝐫𝐜𝐡𝐢𝐯𝐞-𝐧𝐚𝐦𝐞
                contentof[path].=Chr(Ord(A_LoopField)+dif[A_LoopField~="[0-9]"?1: 
                    A_LoopField~="[a-z]"?2: A_LoopField~="[A-Z]"?3: 4])
            Loop Parse contents, "`n"
                StrLen(A_LoopField)<2? {}: files.="`n • " SubStr(A_LoopField,54)
            contentof[path].= Sort(files)
        }
        cache:=[path,contentof[path]]

    ;▼ txt
    } Else If FileExist(path) && path~="\.txt$" {
        (!contentof.Has(path) && FileRead(path)~="^\QFront-end Connection")? ;----- DEV'S CUSTOM FILTER --- ;
            contentof[path]:=SubStr(FileRead(path),start:=RegExMatch(FileRead(path),"F IO C.*UNIQUEN.*hv"), ;
            RegExMatch(FileRead(path),"\| *\R*F IO C.*backpla")-start+1) :{} ;---- can be deleted safely -- ;
        contentof.Has(path)? {}: contentof[path]:=SubStr(FileRead(path),1,1500)
        cache:=[path,contentof[path]]
    } Else If !DirExist(path) {
        cache:=["",""]
    }

    ;▼ router
    DirExist(cache[1]) or path~="\.7z$" ? ToolTip(cache[2]) : ToolTip("")
    cache[1]~="\.txt$" ? guipeek(cache[2],"x" x+16 " y" y+20) : guipeek()

    ;▼ gui
    guipeek(text:=0, opts:="") {
        Static g, gt
        guiexists:=0
        Try guiexists:=g.Hwnd
        If !guiexists && text ;► create gui, populate, show
            a:=WinActive("A"), g:=Gui("-Caption +ToolWindow +AlwaysOnTop -DPIScale"),
            g.SetFont(,"Consolas"), g.SetFont(,"Monospace"), g.SetFont(,"Monoid"),
            gt:=g.AddText(,text), g.Show(opts), WinSetTransparent(240,g.hwnd),
            WinActive("A")=a?{}:WinActivate(a)
        Else If guiexists && text && gt.Value!=text ;► destroy gui, recreate
            g.Destroy(),guipeek(text,opts)
        Else If guiexists && !text ;► destroy gui
            g.Destroy()
    }
}

;▼ FOLDERPEEK: get path of the item under mouse cursor in file explorer
ExplorerGetHoveredItem() {
    static h:=DllCall('LoadLibrary','str','oleacc','ptr')
    DllCall('GetCursorPos', 'int64*', &pt:=0)
    hwnd:=DllCall('GetAncestor','ptr',DllCall('user32.dll\WindowFromPoint','int64',pt),'uint',2)
    If RegExMatch(WinGetClass(hwnd),'^(?:(?<desktop>Progman|WorkerW)|(?:Cabinet|Explore)WClass)$',&M)
        shellWindows:=ComObject('Shell.Application').Windows,
        M.Desktop? shellWindow:=shellWindows.Item(ComValue(0x13,0x8)): GetFromActiveTab()
    GetFromActiveTab() {
        Try activeTab:=ControlGetHwnd('ShellTabWindowClass1',hwnd)
        For w in shellWindows ; https://learn.microsoft.com/en-us/windows/win32/shell/shellfolderview
            (w.hwnd=hwnd && IsSet(activeTab))? ; https://www.autohotkey.com/boards/viewtopic.php?f=83&t=109907
                (ComCall(3,ComObjQuery(w,id:="{000214E2-0000-0000-C000-000000000046}",id), 'uint*',&thisTab:=0),
                thisTab=activeTab? shellWindow:=w :{}) :{}
    }
    If !IsSet(shellWindow)
        Return
    If DllCall('oleacc\AccessibleObjectFromPoint', 'int64',pt, 'ptr*',&pAcc:=0, 'ptr',buf:=Buffer(8+2*A_PtrSize))=0
        idChild:=NumGet(buf,8,'uint'), accObj:=ComValue(9,pAcc,1)
    Switch accObj.accRole[idChild] {
        Case 42: Try name:=accObj.accParent.accName[idChild]    ; editable text
        Case 34: Try name:=accObj.accName[idChild]              ; list item
    }
    Return(IsSet(name)? (RTrim(shellWindow.Document.Folder.Self.Path, '\') '\' name) :"")
}

;▼ CALLAUNCHER (& TRAY MENU): button functions
run1(*)=>(Run("C:\Program Files (x86)\SISW Instruments\Hardware Diagnostics\SimcenterScadasDiagnostics.exe"),
    id:=("ahk_id " WinWait("ahk_exe SimcenterScadasDiagnostics.exe")), ControlClick("Button6", id,,,,"NA"),
    ControlClick("Button7", id,,,,"NA"), ControlClick("Button12",id,,,,"NA"), ControlClick("Button15",id,,,,"NA"))
run2(*)=>Run("C:\Program Files (x86)\SISW Instruments\Adjustment and Calibration\TestSoftwareNext.exe")
run3(*)=>Run("C:\Users\" A_UserName "\Documents\Simcenter Diagnostics")
run4(*)=>Run("C:\Users\" A_UserName "\Documents\SISW Instruments\Adjustment and Calibration\Reports")

;▼ CALLAUNCHER: show a window to quickly launch calibration programs and reports
CalLauncher(action:="auto") {
    Static g1:=Gui("+ToolWindow") ;GUI
    If !IsSet(t1) { ;TRAY MENU
        g1.title:= "Testlab + Folderpeek [F3]",  g1.SetFont("s10")
    }
    Static t1:= g1.AddText(,"No device`t`t`t`t`t")
    t1.OnEvent("Click",Menu_AdaptersList)
    Static b1:= g1.AddButton("r2 w140",     "Launch`nDiagnostics").OnEvent("Click",run1)
    Static b2:= g1.AddButton("x+10 r2 w140","Launch`nAdj + Calib").OnEvent("Click",run2)
    Static b3:= g1.AddButton("xm y+4 w140", "Diag Reports").OnEvent("Click",run3)
    Static b4:= g1.AddButton("x+10 w140","Adj+Cal Reports").OnEvent("Click",run4)
    Static oldIPs:="No device"

    newIPs:=GetEthernetIPs()
    oldIPs=newIPs? {}: (t1.Text:=newIPs, t1.Redraw(), oldIPs:=newIPs, action:="show")
    
    action="hide"?  g1.Hide(): 
    action="show"?  g1.Show("x" A_ScreenWidth-400 " y" A_ScreenHeight-200) :{}

    GetEthernetIPs(*) {
        wmi:=ComObjGet("winmgmts://./root/cimv2"), IPs:=""
        For adapter in wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE Name = '" IniRead(A_ScriptName,A_ComputerName,"eth") "'")
            For conf in wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE Index = " adapter.Index " AND IPEnabled = TRUE")
                For ip in conf.IPAddress
                    IPs.= (IPs?", ":"") ip
        Return(IPs? IPs: "No device")
    }

    Menu_AdaptersList(*) {
        wmi:=ComObjGet("winmgmts://./root/cimv2"), m:=Menu()
        For adapter in wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapter")
            m.Add(adapter.Name,(p*)=>(IniWrite(p[1],A_ScriptFullPath,A_ComputerName,"eth").Bind(adapter.Name)))
        Try m.Default:=IniRead(A_ScriptFullPath,A_ComputerName,"eth")
        m.Show()
    }
}

;▼ HELPERGUIS: create helper guis when Adj+Cal exists
AdjCalHotkeys(p*) {
    If (ACid:=WinExist("ahk_exe TestSoftwareNext.exe")) && !ACguis.Length {
        For el in ["Pag↓", "Fine", " "] {
            ACguis.Push(Gui("-Caption +Owner" ACid, "gui" A_Index))
            ACguis[A_Index].SetFont("s12 caa5500 q3", "Consolas")
            ACguis[A_Index].BackColor:="cccccc"
            WinSetTransColor("cccccc", ACguis[A_Index])
            ACguis[A_Index].AddText("x0 y-2", "⌈")
            ACguis[A_Index].AddText("x5 y2", el)
            ACguis[A_Index].Show(ACpos[A_Index])
        }
    }
    ;destroy helper guis when Adj+Cal doesn't exist
    Loop (!ACid)*ACguis.Length
        ACguis[1].Destroy(), ACguis.RemoveAt(1)
    ;hHotkeys, move+click in the location of helper guis
    If p[1]=1 or p[1]=2
        Try WinGetPos(&x,&y,,,ACguis[p[1]]), Click(,x,y-1), WinGetPos(&x,&y,,,ACguis[3]), Click(,x,y-1)
}

;▼ HELPERGUIS: extract helper guis' pos from this script's INI section ("variable settings")
GetPosFromIni(*) {
    arr:=[]
    For g in ["gui1","gui2","gui3"] {
        RegExMatch((raw:=IniRead(A_ScriptName,A_ComputerName,g)),"^\d+\D+\d+$")? beg:=RegExMatch(raw,"\D+",&m) :{}
        arr.Push("x" SubStr(raw,1,beg-1) " y" SubStr(raw,beg+m.Len) " NoActivate")
    }
    Return(arr)
}

;▼ HELPERGUIS: moves guis and stores new positions in this script's INI section ("variable settings")
MoveHelperGuis(*) {
        MouseGetPos(&mx,&my)
        For el in ACguis {
            WinGetPos(&x,&y,&w,&h,el)
            If i := ((x<=mx && mx<=x+w) && (y<=my && my<=y+h)) * A_Index {
                ACguis[i].BackColor:="ceac49e"
                dx:=mx-x, dy:=my-y
                While GetKeyState("LButton","P") {
                    MouseGetPos(&mx,&my), ToolTip(WinGetTitle(ACguis[i]))
                    ACguis[i].Show(ACpos[i]:=("x" mx-dx " y" my-dy " NoActivate"))
                }
                ACguis[i].BackColor:="cccccc"
                For g in ACguis {
                    WinGetPos(&x,&y,&w,&h,g), IniWrite(x "-" y,A_ScriptName,A_ComputerName,"gui" A_Index)
                }
                Return
            }
        }
    }

;▼ SYSTEMSDB: internal db alignment
AlignDbTable(*) {
    ; Fetch table, prepare variable to store maxes
    table := IniRead(A_ScriptName, "db")
    table_maxes := [0,0,0,0,0]
    ; Read pass, calculate the max len of each element
    For row in StrSplit(table, "`n") {
        For el in StrSplit(SubStr(row, StrLen(StrSplit(row, "=")[1])+2), "| ") {
            table_maxes[A_Index]<(el_len := StrLen(Trim(el, " ")))? table_maxes[A_Index]:=el_len :{}
        }
    }
    ; Write pass, apply padding to non-max-len elements
    For row in StrSplit(table, "`n") {
        newVal := " "
        For el in StrSplit(SubStr(row, StrLen(StrSplit(row, "=")[1])+2), "| ") {
            newVal .= Trim(el, " ") (A_Index=5? "": Format("{:-" table_maxes[A_Index]-StrLen(Trim(el, " ")) "}","") " | ")
        }
        newVal := StrReplace(newVal, "�", "-")
        IniWrite(newVal, A_ScriptName, "db", StrSplit(row, "=")[1])
    }
}

;▼ SYSTEMSDB: makes a gui to show db
ShowDbGui(p*) {
    ; Open a Gui with a ListView with the SystemsDB
    ; Clicking on rows opens EditSystem()
    Static g2:=0, lv:=0, vis:=0
    If !p.Length or Type(p[1])!="String" or p[1]!="refresh"
        vis := !vis

    If vis && !g2 {
        g2 := Gui(,"SystemsDB")
        g2.SetFont(,"Consolas")
        lv := g2.AddListView("x0 y0 w660 h800 Backgroundeeeeee ",["SN","TAC","SOLD-TO","TYPE","EXP","CUSTOMER"])
    }
    
    If vis or (p.Length && Type(p[1])="String" && p[1]="refresh") {
        lv.Delete()
        For el in StrSplit(IniRead(A_ScriptName, "db"), "`n") {
            kv := StrSplit(el, "=")
            e := StrSplit(kv[2], " | ")
            lv.Add(, Trim(kv[1]), Trim(e[1]), Trim(e[2]), Trim(e[3]), Trim(e[4]), Trim(e[5]))
        }
        lv.ModifyCol()
        g2.OnEvent("Size",(g,e,w,h)=>(lv.Move(,,w,h)))
        g2.OnEvent("Close",ShowDbGui)
        lv.OnEvent("DoubleClick",EditSingleSystem)
    }

    vis? (g2.Show("NoActivate")): (g2.Destroy(), g2:=0)
}

;▼ SYSTEMSDB: make a gui to show/edit a single system
EditSingleSystem(controlOrKey, row?) {
    ; Make GUI to show, edit, save all the elements of a system
    ; Param(s) can be SN or ListView+Row
    Static g3:=[]
    If Type(controlOrKey)="Gui.ListView" && Type(row)="Integer" {
        col(i, k:=controlOrKey, r:=row)=>k.GetText(r,i)
        els := [col(1), col(2), col(3), col(4), col(5), col(6)]
    } Else If controlOrKey~="^\s*\d{8}\s*$" { 
        Try rawLine := IniRead(A_ScriptName, "db", key:=RegExReplace(controlOrKey, "\D"))
        Catch
            Return
        e := StrSplit(rawLine, " | ")
        els := [key, Trim(e[1]), Trim(e[2]), Trim(e[3]), Trim(e[4]), Trim(e[5])]
    }
    If !IsSet(els) ;When CtrC in Word without a SN selected
        Return
    g3.Push(Gui("+AlwaysOnTop +ToolWindow","Edit row"))
    g3[-1].SetFont("s9","Consolas")  
    edits := []
    For el in els {
        g3[-1].AddText("x4 y+2 w30 Right",["SN", "TAC", "S-T", "TYPE", "EXP", "CUST"][A_Index])
        edits.Push(g3[-1].AddEdit("x+2 yp-2 w200", el))
    }
    (g3[-1].AddButton("h20","Save")).OnEvent("Click", SaveRow)

    MouseGetPos(&x,&y)
    g3[-1].Show("NoActivate x" x+20 " y" y-100)

    Hotkey("~Esc", (*)=>Close(), "On"), Close(g:=g3)=>(g.Length? (g[-1].Destroy(), g.Pop(), Close(g)): 0)

    SaveRow(*) {
        Try key := edits[1].Value
        Catch
            Return
        val := " " edits[2].Value " | " edits[3].Value " | " edits[4].Value " | " edits[5].Value " | " edits[6].Value
        IniWrite(val, A_ScriptName, "db", key)
        AlignDbTable(), ShowDbGui("refresh")
    }
}

;▼ SYSTEMSDB: import from clipboard to this script's INI section ("variable settings")
ImportFromClipboad(*) {
    ; Parse clipboard, trim carriage-return
    ; Expect 5 columns [TAC  CUSTOMER  SOLDTO  SN  EXP]
    ; Convert to table [SN = TAC | SOLD-TO | TYPE | EXP | CUSTOMER]
    i := 0
    For el in StrSplit(StrReplace(A_Clipboard, Chr(13)), "`n") {
        If StrLen(el)<4
            Continue
        i := Mod(++i,5)
        i=1? tac:=Trim(el): i=2? cus:=Trim(StrReplace(el, "�", "-")):
        i=3? st:=Trim(el): i=4? key:=Trim(el):
        (   oldVal := IniRead(A_ScriptName, "db", key, " |  | _type_"),
            typ := Trim(StrSplit(oldVal, " | ")[3]), exp := Fmt(el),
            newVal := (" " tac " | " st " | " typ " | " exp " | " cus),
            IniWrite(newVal, A_ScriptName, "db", key))
    }
    AlignDbTable()

    Fmt(date) {
        ; Convert date format: 31-Jan-25 → 2025-01-31
        Try date := Trim(date)
        If !(date~="\d\d-\D\D\D-\d{2,4}")
            GoTo BadEnding

        dmy:=StrSplit(date, "-"),  d:=dmy[1],  m:=StrLower(dmy[2]),  y:=Mod(dmy[3],2000)+2000
        m:=("genjan"~=m?1: "feb"=m?2: "mar"=m?3: "apr"=m?4: "magmay"~=m?5: "giujun"~=m?6:
            "lugjul"~=m?7: "agoaug"~=m?8: "setsep"~=m?9: "ottoct"~=m?10: "nov"=m?11: "dicdec"~=m?12: 0)
            
        If d>0 && d<32 && m && y>2000 && y<2100
            Return(y "-" Format("{:02}", m) "-" Format("{:02}", d))

        BadEnding:
        MsgBox("Invalid date:'" date "'")
        Return(date)
    }
}

;▼ MISC: asks you to pick one of the folders open in file explorer, changes its icon to "sent"
SetSentIcon(*) {
    m := Menu()
    m.Add(t:="Which folder?",(*)=>0), m.Default:=t, m.Disable(t)

    dirs := []
    For Tab in ComObject("Shell.Application").Windows {
        dirs.Push(dir:=Tab.Document.Folder.Self.Path), m.Add(dir,(d:=dir,*)=>set(d))
    }

    m.Show()

    set(d) {
        Try (FileOpen(A_Temp "\changeicon.ahk", "w")).Write('#Requires AutoHotkey v2.0`nFileSetAttrib("+S","' d '")`nTry r:=FileSetAttrib("-RS",f:="' d '\desktop.ini")`nIsSet(r)?{}:FileAppend("",f)`nIniWrite("IconResource=C:\Windows\System32\imageres.dll,279",f,".ShellClassInfo")`nFileSetAttrib("+RSH-A",f)`nDllCall("Shell32\SHChangeNotify","UInt",0x08000000,"UInt",0,"Ptr",0,"Ptr",0)`nExitApp()'), Run('*RunAs "' A_Temp '\changeicon.ahk"')
        Catch
            MsgBox("Aborted", "Change icon", "Iconi")
    }
}




/* VARIABLE SETTINGS (INI SECTION) ###################################################################################

[default values]
gui1=1520-250
gui2=1520-310
gui3=1300-450
eth=

[DLE-G10-W11]
gui1=1520-250
gui2=1520-310
gui3=1300-450
eth=ASIX AX88772B USB2.0 to Fast Ethernet Adapter

[DLE-ALE]
gui1=1138-130
gui2=1141-183
gui3=958-300
eth=

[▼ DATABATE OF TAC/ORDERS AND SYSTEMS' DETAILS ▼]

[db]
SN      = TAC      | SOLD-TO  | TYPE        | EXP        | CUSTOMER
61125037= 10552327 | 1371920  | SCL20-ENV   | 2026-12-31 | THALES ALENIA SPACE ITALIA SPA
61130701= 10552328 | 1371920  | SCL21S      | 2026-12-31 | THALES ALENIA SPACE ITALIA SPA
61125038= 10552329 | 1371920  | SCL21S      | 2026-12-31 | THALES ALENIA SPACE ITALIA SPA
25180210= 10552330 | 1371920  | SCM205      | 2026-12-31 | THALES ALENIA SPACE ITALIA SPA
46083225= 10552348 | 1371898  | SCM05       | 2026-09-30 | IVECO SPA
25222508= 10552349 | 1371898  | SCR205      | 2026-09-30 | IVECO SPA
53103918= 10552350 | 1371898  | SCR05       | 2026-09-30 | IVECO SPA
25232817= 10552351 | 1371898  | SCR205      | 2026-09-30 | IVECO SPA
42042505= 12050925 | 1371917  | SC316-UTP   | 2026-08-31 | LEONARDO SPA
42033607= 12084207 | 1371917  | SC310V      | 2026-08-31 | LEONARDO SPA
25193312= 10551386 | 1371949  | SCR205      | 2025-12-31 | CONTINENTAL BRAKES ITALY SPA - Blengio
11152111= 10551398 | 1406728  | XS12-A      | 2026-06-30 | MASERATI SPA - Aversano
25161625= 10551399 | 1406728  | SCM205      | 2026-06-30 | MASERATI SPA - Aversano
22162527= 10551412 | 1600011  | SCM202      | 2026-06-30 | MASERATI SPA - Brancato _Grassigli
22162526= 10551413 | 1600011  | SCM202      | 2026-06-30 | MASERATI SPA - Brancato _Grassigli
25152513= 10551415 | 1695245  | SCR205      | 2026-06-30 | MASERATI SPA - Bonazzi
25185102= 10551416 | 1695245  | SCR205      | 2026-06-30 | MASERATI SPA - Bonazzi
12225001= 10551509 | 10133088 | SC-XS-12    | 2028-04-30 | AUTOMOBILI LAMBORGHINI SPA - Franceschini
22225002= 10551510 | 10133088 | SCM03S      | 2028-04-30 | AUTOMOBILI LAMBORGHINI SPA - Franceschini
22215105= 10551511 | 10133088 | SCM2E02     | 2028-04-30 | AUTOMOBILI LAMBORGHINI SPA - Franceschini
29210116= 10551512 | 10133088 | SCR209      | 2028-04-30 | AUTOMOBILI LAMBORGHINI SPA - Franceschini
25240101= 10551429 | 10135716 | SCR2E05     | 2026-06-30 | AUTOMOBILI LAMBORGHINI SPA - Formisano
11163197= 10551457 | 1371949  | SC-XS06-E   | 2025-12-31 | CONTINENTAL BRAKES ITALY SPA - Blengio
46083702= 10551464 | 1371979  | SCR05       | 2027-12-31 | WASS SUBMARINE SYSTEMS SRL - Lazzaretti
29204212= 10551590 | 1371961  | SCM10S      | 2026-07-31 | BRIDGESTONE EUROPE NV/SA - Emiliani
25165204= 10551591 | 1371961  | SCR205      | 2026-07-31 | BRIDGESTONE EUROPE NV/SA - Emiliani
53121107= 10551609 | 1372000  | SCM05V      | 2026-07-31 | SERMS SRL - Alvino
47092901= 10551610 | 1372000  | SCR01       | 2026-07-31 | SERMS SRL - Alvino
12232822= 10552164 | 10095682 | XS          | 2026-09-30 | HITACHI RAIL STS SPA
22232823= 10552165 | 10095682 | SC02?       | 2026-09-30 | HITACHI RAIL STS SPA
25213503= 10552175 | 1924149  | SCM2E05V    | 2026-09-30 | BE CAE & TEST SRL
22205001= 10552196 | 1866968  | _type_      | 2026-07-31 | ITT ITALIA S.R.L.
53093302= 10551271 | 1371890  | SCM05       | 2025-12-31 | LEONARDO SPA
22211905= 10551274 | 1900943  | SCM202      | 2025-12-31 | LEONARDO SPA
53115018= 10551283 | 1371947  | SCM05       | 2025-12-31 | YAMAHA MOTOR RACING SRL
46075115= 10551290 | 1477266  | SCM05       | 2025-12-31 | MARELLI EUROPE SPA
46073013= 10551291 | 1477267  | SCR205      | 2025-12-31 | MARELLI EUROPE SPA
46084823= 10551303 | 10133090 | SCR05       | 2025-12-31 | AUTOMOBILI LAMBORGHINI SPA
25225254= 10551304 | 10133090 | SCR2E05     | 2025-12-31 | AUTOMOBILI LAMBORGHINI SPA
11191915= 10550664 | 1461918  | SC-XS12-A   | 2025-12-31 | FERRARI SPA
47072205= 10550665 | 1461918  | SCM01       | 2025-12-31 | FERRARI SPA
46074210= 10550666 | 1461918  | SCM05       | 2025-12-31 | FERRARI SPA
46074209= 10550667 | 1461918  | SCR209      | 2025-12-31 | FERRARI SPA
22182502= 10550668 | 1461918  | SCM202      | 2025-12-31 | FERRARI SPA
22202502= 10550669 | 1461918  | SCM202      | 2025-12-31 | FERRARI SPA
25182501= 10550670 | 1461918  | SCM205      | 2025-12-31 | FERRARI SPA
25202503= 10550671 | 1461918  | SCM205      | 2025-12-31 | FERRARI SPA
25162307= 10550673 | 1461918  | SCR205      | 2025-12-31 | FERRARI SPA
25215014= 10550674 | 1461918  | SCR205      | 2025-12-31 | FERRARI SPA
25213525= 10550675 | 1461918  | SCR05       | 2025-12-31 | FERRARI SPA
25204911= 10550676 | 1461918  | SCR205      | 2025-12-31 | FERRARI SPA
29195007= 10550677 | 1461918  | SCM209      | 2025-12-31 | FERRARI SPA
22174917= 10550049 | 1619019  | SCR202      | 2025-12-31 | FERRARI SPA
21183714= 10551121 | 1378487  | SCM201V     | 2025-12-31 | MBDA ITALIA SPA
21184908= 10551122 | 1378487  | SCM201V     | 2025-12-31 | MBDA ITALIA SPA
21230419= 10551123 | 1378487  | SCM201-ENV  | 2025-12-31 | MBDA ITALIA SPA
21230420= 10551124 | 1378487  | SCM201-ENV  | 2025-12-31 | MBDA ITALIA SPA
21230422= 10551125 | 1378487  | SCM201-ENV  | 2025-12-31 | MBDA ITALIA SPA
21243616= 10551126 | 1378487  | SCM201      | 2025-12-31 | MBDA ITALIA SPA
22183715= 10551127 | 1378487  | SCM202      | 2025-12-31 | MBDA ITALIA SPA
42112915= 12060420 | 1378487  | SC302VB-UTP | 2025-12-31 | MBDA ITALIA SPA
42082507= 12060420 | 1378487  | SC310V-UTP  | 2025-12-31 | MBDA ITALIA SPA
42101502= 12060420 | 1378487  | SC302VB-UTP | 2025-12-31 | MBDA ITALIA SPA
25134410= 10550860 | 1371899  | SCM205      | 2025-12-31 | ELECTROLUX ITALIA SPA
22185223= 10550861 | 1371899  | SCM202      | 2025-12-31 | ELECTROLUX ITALIA SPA
52111313= 10550862 | 1450772  | SCM02-ENV   | 2025-12-31 | PIAGGIO & C. SPA
47070817= 10550863 | 1459641  | SCM01       | 2025-12-31 | ELECTROLUX ITALIA SPA
47103919= 10550864 | 1459641  | SCM01       | 2025-12-31 | ELECTROLUX ITALIA SPA
29223111= 10550910 | 1870798  | SCM209-ENV  | 2025-12-31 | LEONARDO SPA
11192804= 10550869 | 1371952  | XS06-E      | 2025-12-31 | PRES. DEL CONSIGLIO DEI MINISTRI
11150215= 10550870 | 1371952  | XS06-E      | 2025-12-31 | PRES. DEL CONSIGLIO DEI MINISTRI
11150213= 10550871 | 1371952  | XS06-E      | 2025-12-31 | PRES. DEL CONSIGLIO DEI MINISTRI
11150212= 10550872 | 1371952  | XS06-E      | 2025-12-31 | PRES. DEL CONSIGLIO DEI MINISTRI
53091901= 10550873 | 1371952  | SCR05       | 2025-12-31 | PRES. DEL CONSIGLIO DEI MINISTRI
22242406= 10550899 | 1424157  | SCM02       | 2025-12-31 | YAMAHA MOTOR R&D EUROPE SRL
52121105= 10550926 | 10133090 | SCR202      | 2025-12-31 | AUTOMOBILI LAMBORGHINI SPA
25133510= 10550956 | 1371885  | SCR205      | 2026-01-31 | G.D SPA
22182607= 10550965 | 1661244  | SCM2E02     | 2025-12-31 | MACHINING CENTERS MANUFACTURING SPA
53094708= 10550975 | 1371886  | SCM05       | 2025-12-31 | PIRELLI TYRE SPA
53115019= 10549414 | 1371994  | SCM05       | 2025-12-31 | LOMBARDINI SRL
12240422= 10549415 | 1372019  | SC-XS12-ACL | 2025-12-31 | EMC FIME SRL
52121910= 10549416 | 1372019  | SCM02       | 2025-12-31 | EMC FIME SRL
22194501= 10550483 | 1772737  | SCM202      | 2025-12-31 | FABER SPA
22204922= 10550516 | 1371957  | SCM202-ENV  | 2025-12-31 | LEONARDO SPA
52111401= 10550379 | 1371889  | SCM202-ENV  | 2025-12-31 | VALEO SPA
22181315= 10550380 | 1638555  | SCR202      | 2026-04-30 | RAICAM DRIVELINE SRL
46084827= 10550396 | 1371902  | SCR05       | 2025-12-31 | LEONARDO SPA
53092902= 10550397 | 1371902  | SCM209      | 2025-12-31 | LEONARDO SPA
53092637= 10550398 | 1371902  | SCR05       | 2025-12-31 | LEONARDO SPA
22151610= 10550405 | 1423353  | SCM202      | 2025-12-31 | CECCATO ARIA COMPRESSA SRL
46072630= 10550423 | 1371910  | SCM205      | 2025-12-31 | MARELLI EUROPE SPA
21204102= 10550424 | 1371910  | SCM201V     | 2025-12-31 | MARELLI EUROPE SPA
21204103= 10550425 | 1371910  | SCM201V     | 2025-12-31 | MARELLI EUROPE SPA
25151811= 10549910 | 1461919  | SCR205      | 2025-12-31 | FERRARI SPA
25203004= 10549911 | 1461919  | SCR205      | 2025-12-31 | FERRARI SPA
25210526= 10549912 | 1461919  | SCR205      | 2025-12-31 | FERRARI SPA
22210608= 10549782 | 1371917  | SCM202-ENV  | 2025-12-31 | LEONARDO SPA - NERVIANO / Castellazzi
21183702= 10549783 | 1371958  | SCM201-ENV  | 2025-12-31 | LEONARDO SPA - La Spezia / Meriggi
53132204= 10549791 | 1463941  | SCM05-ENV   | 2025-12-31 | LEONARDO SPA - Aquila / Lozzi
22172875= 10549793 | 1583603  | SCM202-ENV  | 2025-12-31 | LEONARDO SPA - La Spezia / Mornelli
62224507= 10549796 | 1371872  | SCL220-ENV  | 2025-12-31 | THALES ALENIA SPACE ITALIA SPA - Aquila / Chiarieri
62214204= 10549797 | 1371872  | SCL220-ENV  | 2025-12-31 | THALES ALENIA SPACE ITALIA SPA
52124511= 10549798 | 1371872  | SCM02       | 2025-12-31 | THALES ALENIA SPACE ITALIA SPA
21162410= 10549799 | 1371872  | SCM201      | 2025-12-31 | THALES ALENIA SPACE ITALIA SPA
62201103= 10549800 | 1371896  | SCL220-ENV  | 2025-12-31 | THALES ALENIA SPACE ITALIA SPA - RM / Di Pietro
62201104= 10549801 | 1371896  | SCL220-ENV  | 2025-12-31 | THALES ALENIA SPACE ITALIA SPA
62201102= 10549802 | 1371896  | SCL220-ENV  | 2025-12-31 | THALES ALENIA SPACE ITALIA SPA
53092508= 10549803 | 1371896  | SCM05       | 2025-12-31 | THALES ALENIA SPACE ITALIA SPA
25213421= 10549889 | 1371889  | SCR205      | 2025-12-31 | VALEO SPA
22210213= 10549684 | 1874144  | SCR202      | 2025-12-31 | KVERNELAND GROUP RAVENNA SRL
25185216= 10549688 | 1371904  | SCM205      | 2025-12-31 | BREMBO NV
53104866= 10549689 | 1371986  | SCM05       | 2025-12-31 | FERRARI SPA
53113910= 10549690 | 1371986  | SCR05       | 2025-12-31 | FERRARI SPA
29201505= 10549691 | 1371986  | SCR209      | 2025-12-31 | FERRARI SPA
47091804= 10549700 | 1551348  | SCM01       | 2025-12-31 | BRIDGESTONE EUROPE NV/SA
22194914= 10549727 | 1856791  | SCR02       | 2025-09-30 | UNIV. DEGLI STUDI DI TORINO
11181302= 10549465 | 1637893  | SC-XS12-A   | 2025-12-31 | GDM SPA
29181820= 10549485 | 1371902  | SCR209      | 2025-12-31 | LEONARDO SPA
29180806= 10549486 | 1371934  | SCM209-ENV  | 2025-12-31 | LEONARDO SPA
47101501= 10549487 | 1372002  | SCM01       | 2025-12-31 | LEONARDO SPA
22213502= 10549488 | 1372002  | SCM02       | 2025-12-31 | LEONARDO SPA
46072206= 10549501 | 1371874  | SCR205      | 2025-12-31 | LEONARDO SPA
22240206= 10549502 | 1371874  | SCR2E02     | 2025-12-31 | LEONARDO SPA
29184801= 10549503 | 1371874  | SCR209      | 2025-12-31 | LEONARDO SPA
53131119= 10549504 | 1371966  | SCR05       | 2025-12-31 | LEONARDO SPA
25191202= 10549505 | 1371966  | SCR205      | 2025-12-31 | LEONARDO SPA
53122411= 10549310 | 1372028  | SCM05       | 2025-07-31 | CSI SPA
62144405= 10549240 | 1612055  | SCL220      | 2025-10-31 | DUMAREY AUTOMOTIVE ITALIA SPA
11192710= 10549241 | 1742803  | SC-XS12     | 2025-10-31 | DUMAREY AUTOMOTIVE ITALIA SPA
52124718= 10549244 | 10149999 | SCM02       | 2025-11-30 | FPT INDUSTRIAL SPA
21141606= 10549245 | 10149999 | SCM201V     | 2025-11-30 | FPT INDUSTRIAL SPA
47083117= 10549246 | 1371913  | SCM01       | 2025-11-30 | FPT INDUSTRIAL SPA
52124717= 10549247 | 1371913  | SCM02       | 2025-11-30 | FPT INDUSTRIAL SPA
53124716= 10549248 | 1371913  | SCM05       | 2025-11-30 | FPT INDUSTRIAL SPA
53124726= 10549249 | 1371913  | SCM05       | 2025-11-30 | FPT INDUSTRIAL SPA
53131207= 10552620 | 1371964  | SCM05       | 2027-03-31 | HITACHI RAIL STS SPA
11182704= 10552613 | 1658870  | XS          | 2026-06-20 | FERRETTI SPA

[temp]
[]
*/ 
