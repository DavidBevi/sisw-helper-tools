;-----------------------------------------------------------------;
  title:= "Diag Comparer"
;-----------------------------------------------------------------;

#Requires AutoHotkey v2
#SingleInstance Force


; VARIABLES --------------------------------------------------------------------------------------------------


; FUNCTIONS --------------------------------------------------------------------------------------------------

LoadDir(dir_path:="", *) {  ; Finds AsFound+AsLeft, launches SideBySide(AF,AL), updates TX.
    
    Loop {
        If DirExist(dir_path)
            Break
        dir_path:= DirSelect(,6,"Select a directory with 2 Diag reports")
        If dir_path="" 
            ToolTip("`n No folder selected `n Program will close `n "),Sleep(3000),ExitApp()
    }

    diag1_path:= "",  diag2_path:= "",  diags_paths_list:= ""

    Loop Files, dir_path "\*.txt", "F" {
        InStr(A_LoopFileName,"DiagAsFound")? diag1_path:= A_LoopFileFullPath: {}
        InStr(A_LoopFileName,"DiagAsLeft")?  diag2_path:= A_LoopFileFullPath: {}
        If (diag1_path!="" and diag2_path!="")
            Break
        If SubStr(A_LoopFileName,1,5)="LmsDi"
            diags_paths_list.= A_LoopFileFullPath "`n"
    }
    If diags_paths_list="" {
        Loop Files, dir_path "\*.txt", "F" {
            diags_paths_list.= A_LoopFileFullPath "`n"
        }
    }
    If diags_paths_list="" {
        result:= MsgBox("Cannot find certificates","LoadDir",21)
        If result="Retry" {
            LoadDir()
        } Else {
            SetTimer(()=>MsgBox("a"),600)
        }
    } Else {
        If diag1_path=""
            diag1_path:=SubStr(Sort(diags_paths_list),1,InStr(Sort(diags_paths_list),"`n"))
        If diag2_path=""
            diag2_path:=SubStr(Sort(diags_paths_list,"R"),1,InStr(Sort(diags_paths_list,"R"),"`n"))
        SideBySide(Extract_MiniTable(diag1_path),Extract_MiniTable(diag2_path))
        Tx.Value:= dir_path
    }
}

SideBySide(diag1_minitable, diag2_minitable) {  ; Updates LV.
    LV.Delete
    Arr:= [[],[]]
    For i, diag in [diag1_minitable, diag2_minitable]
        Loop Parse, diag, "`n"
            Arr[i].Push(A_LoopField)
    For i, line in Arr[1]
        Arr[1][i]=Arr[2][i]?
            LV.Add(, Arr[1][i], Arr[2][i]):
            LV.Add(, Arr[1][i], Arr[2][i] "◀ ")
}

Extract_MiniTable(diag_path) {  ; Returns a mini-table of modules (string).
    Flag:=0
    Reconstr:=""
    Loop Parse, FileRead(Trim(diag_path,"`n")), "`n"
        If (InStr(A_LoopField,"F IO C") and InStr(A_LoopField,"Lca"))
            Flag:=1
        Else If (Flag and InStr(A_LoopField,"------"))
            Continue
        Else If (Flag and SubStr(A_LoopField,8,1)="|") {
            Val:=[]
            If IsNumber(SubStr(A_LoopField,1,1))
                Subflag:=1
            Loop Parse, A_LoopField, "|"
                If (IsSet(Subflag) and Subflag)
                    Val.Push("$" Trim(A_LoopField)),
                    Subflag:=0
                Else Val.Push(Trim(A_LoopField))
                
                If Val[2]="SYSTEM"
                    Continue
                Else If SubStr(Val[1],1,1)="$"
                    Reconstr=""? {}: Reconstr.="`n",
                    Reconstr.= "`n" Format("{:-14}",Val[2])  SubStr(Val[3],3) "    `nModule........Lca...Flash "
                Else Reconstr.= "`n" Format("{:-14}",Val[2])  Format("{:-6}",Val[7])  Format("{:-6}",Val[8])
        } Else If Flag
            Break
    Return(Reconstr)
}

; MAIN -------------------------------------------------------------------------------------------------------
my_gui:=Gui(,title)
For font in ["Consolas", "Inconsolata", "Monoid"]
    my_gui.SetFont(,font)

Tx:=my_gui.AddText("x10 y10 w386 Right","ShortPath")

Bt:=my_gui.AddButton("x400 y4 w106","Change Dir").OnEvent("Click",LoadDir)

LV:= my_gui.AddListView("x6 y+4 w500 h500 +NoSortHdr",["Diag as Found  ","Diag as Left   "])
LV.ModifyCol(1, "240")
LV.ModifyCol(2, "240")

LoadDir()

my_gui.Show

