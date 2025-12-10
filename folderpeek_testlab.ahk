; ██████  █████  ██     █████  ██████ █████  █████  ██████ ██████ ██  ██                ██████ ██████  █████ ██████ ██      ████  █████  
; ██     ██   ██ ██     ██  ██ ██     ██  ██ ██  ██ ██     ██     ██ ██      ██  ██       ██   ██     ██       ██   ██     ██  ██ ██  ██ 
; █████  ██   ██ ██     ██  ██ ████   █████  █████  ████   ████   ████         ██         ██   ████    ████    ██   ██     ██████ █████  
; ██     ██   ██ ██     ██  ██ ██     ██  ██ ██     ██     ██     ██ ██      ██  ██       ██   ██         ██   ██   ██     ██  ██ ██  ██ 
; ██      █████  ██████ █████  ██████ ██  ██ ██     ██████ ██████ ██  ██                  ██   ██████ █████    ██   ██████ ██  ██ █████  
                                                                                                                                               
;2025-12-10: this script contains many different utilities, organized in subsections inside logical sections (settings, hotkeys, funcs...)
;            utilities should be independent, but I'm not sure

;FOLDERPEEK: constantly checks if the mouse cursor is over a file explorer window, if so it displays the content of the folder under the cursor
;            txt files are also displayed, unless they are a diag report, in which case only a chosen section with relevant info is displayed
;            if 7-Zip is installed it also displays the files inside 7z compressed folders

;CALLAUNCHER: a little window in the bottom right corner with buttons to open calibration programs and reports
;             the window should auto-open when a new ethernet device (a scadas) is detected (a bit buggy)
;             manually open with F3, manually close with ESC

;HELPERGUIS: this makes 3 little transparent windows with an orange corner, they will appear over Adjustment+Calibration when it's open and focused
;            if you drag the 2 boxes that also have a key name over the NEXT and PREV buttons you'll be able to use said keys to click the buttons
;            the 3rd box with only the orange corner is where the mouse clicks after clicking NEXT or PREV
;            (it's a hack because of how Adj+Cal is made, but it's very convenient, and it should be robust)

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

[▼ DATABATE OF TAC/ORDERS AND SYSTEMS' DETAILS ▼]

[db]
SN      = TAC      | SOLD-TO  | TYPE        | EXP        | CUSTOMER
12345678= 10002000 | 7654321  | SCR209      | 2025-12-31 | ESEMPIO SPA

[temp]
[]
*/ 
