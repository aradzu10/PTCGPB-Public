#SingleInstance on

#Include %A_ScriptDir%\..\Scripts\Include\Utils.ahk

;SetKeyDelay, -1, -1
SetMouseDelay, -1
SetDefaultMouseSpeed, 0
;SetWinDelay, -1
;SetControlDelay, -1
SetBatchLines, -1
SetTitleMatchMode, 3

global adbShell, adbPath, adbPorts, winTitle, folderPath, selectedFilePath, mumuFolder
global tradeSectionExpanded, SetDataList, CardDataBySet, SelectedSet, SelectedCard, TradeRadioTrade, TradeRadioShare, tradeBaseY, FriendID, FriendIDInput
global AllPokemonsCSV, AllPokemons
tradeSectionExpanded := false
SetDataList := []
CardDataBySet := {}
AllPokemonsCSV := A_ScriptDir . "\..\Resources\all_pokemon.csv"

IniRead, winTitle, InjectAccount.ini, UserSettings, winTitle, 1
IniRead, fileName, InjectAccount.ini, UserSettings, fileName, name
IniRead, folderPath, InjectAccount.ini, UserSettings, folderPath, C:\Program Files\Netease
IniRead, selectedFilePath, InjectAccount.ini, UserSettings, selectedFilePath, ""
IniRead, FriendID, %A_ScriptDir%\..\Settings.ini, UserSettings, FriendID, none

; Set a custom font and size for better appearance
Gui, Font, s10, Segoe UI
Gui, Color, 1E1E1E  ; Dark background color
Gui, Font, cDCDCDC  ; Light text color

; Add a title with warning styling
Gui, Add, Text, x10 y10 w450 cWhite, This tool is to INJECT (login to) the selected account.
Gui, Add, Text, x10 y+5 w450 cRed, It will LOG OUT OF any current account in that instance.
Gui, Add, Text, x10 y+5 w450 cWhite, Ensure you have the login info of the current account (either a .xml file, nintendo account link, etc.) or you will LOSE it.

; Create a horizontal line for visual separation
Gui, Add, Text, x10 y+15 w450 h1 0x10 c3F3F3F ; Darker separator

; Instance section
instanceList := GetInstanceList(folderPath)
selectedIndex := 1
if (instanceList != "") {
    StringSplit, arr, instanceList, |
    Loop, %arr0%
    {
        if (arr%A_Index% = winTitle) {
            selectedIndex := A_Index
            break
        }
    }
}
Gui, Add, Text, x10 y+15, Instance Name:
Gui, Add, DropDownList, x10 y+5 vwinTitle w200 Choose%selectedIndex%, %instanceList%
Gui, Add, Button, x+10 yp w80 gRefreshInstances, Refresh

; File section
Gui, Add, Text, x10 y+15 cDCDCDC, File Name (without spaces and without .xml):
Gui, Add, Edit, x10 y+5 vfileName w300 c000000 BackgroundFFFFFF, %fileName%
Gui, Add, Button, x+10 yp w80 gBrowseFile, Browse

; Folder section
Gui, Add, Text, x10 y+15 cDCDCDC, MuMu Folder same as main script (C:\Program Files\Netease)
Gui, Add, Edit, x10 y+5 vfolderPath w300 c000000 BackgroundFFFFFF, %folderPath%

; Trade/Share collapsible section
Gui, Add, Text, x10 y+20 w450 h1 0x10 c3F3F3F
Gui, Add, Button, x435 y+10 w25 h20 gToggleTradeSection vToggleTradeBtn, >
Gui, Add, Text, x10 yp+2 cWhite, Trade/Share

; Trade warning — always visible
Gui, Add, Text, x10 y+8 w450 cRed vTradeWarningText, When tab closed, will not auto Trade/Share.

; Expandable trade controls — placed off-screen initially, moved by ApplyExpandState
Gui, Add, Radio, x10 y5000 vTradeRadioTrade Group Hidden Checked, Trade
Gui, Add, Radio, x+20 yp vTradeRadioShare Hidden, Share
Gui, Add, Button, x340 yp w110 h23 Hidden vRunTradeBtn gRunTradeAction, Trade/Share Only
Gui, Add, Text, x10 y5000 cDCDCDC Hidden vSetDropLabel, Set:
Gui, Add, DropDownList, x55 yp vSelectedSet w395 Hidden gOnSetChange,
Gui, Add, Text, x10 y5000 cDCDCDC Hidden vCardDropLabel, Card:
Gui, Add, DropDownList, x55 yp vSelectedCard w395 Disabled Hidden,

; Friend ID row (expandable trade section)
friendIDVal := (FriendID = "" || FriendID = "none") ? "" : FriendID
Gui, Add, Text, x10 y5000 w70 cDCDCDC Hidden vFriendIDLabel, Friend ID:
Gui, Add, Edit, x85 y5000 w365 c000000 BackgroundFFFFFF Hidden vFriendIDInput, %friendIDVal%

; Set cue banner (gray placeholder) for Friend ID input
GuiControlGet, FriendIDHwnd, Hwnd, FriendIDInput
fridHint := "16-digit Friend ID"
SendMessage, 0x1501, 1, % &fridHint,, ahk_id %FriendIDHwnd%

; Bottom separator and action buttons — off-screen initially, positioned by ApplyCollapseState
Gui, Add, Text, x10 y5000 w450 h1 0x10 c3F3F3F vBottomSeparator
Gui, Add, Button, x130 y5000 w100 h40 gSaveSettings cBlue vSubmitBtn, Submit
Gui, Add, Button, x+10 yp w100 h40 gRunInstance cGreen vRunInstanceBtn, Run Instance

; Read warning text position so all layout is computed from actual rendered coordinates
GuiControlGet, warnP, Pos, TradeWarningText
global tradeBaseY
tradeBaseY := warnPY + warnPH

; Open in collapsed state
ApplyCollapseState()
Return

OnGuiClose:
    ExitApp

GuiClose:
    ExitApp

BrowseFile:
    FileSelectFile, selectedFile, 3, , Select XML File, XML Files (*.xml)
    if (selectedFile != "")
    {
        SplitPath, selectedFile, fileNameNoExt, , , fileNameNoExtNoPath
        GuiControl,, fileName, %fileNameNoExtNoPath%
        selectedFilePath := selectedFile
    }
    return

SaveSettings:
    Gui, Submit, NoHide
    ; Removed: Gui, Destroy
    IniWrite, %winTitle%, InjectAccount.ini, UserSettings, winTitle
    IniWrite, %fileName%, InjectAccount.ini, UserSettings, fileName
    IniWrite, %folderPath%, InjectAccount.ini, UserSettings, folderPath
    IniWrite, %selectedFilePath%, InjectAccount.ini, UserSettings, selectedFilePath

mumuFolder := getMumuFolder(folderPath)

adbPath := mumuFolder . "\shell\adb.exe"
if !FileExist(adbPath)
    adbPath := mumuFolder . "\nx_main\adb.exe"
findAdbPorts(mumuFolder)

if(!WinExist(winTitle)) {
    Msgbox, 16, , Can't find instance: %winTitle%. Make sure that instance is running.;'
    ExitApp
}

if !FileExist(adbPath) ;if international mumu file path isn't found look for chinese domestic path
    adbPath := folderPath . "\MuMu Player 12\shell\adb.exe"
if !FileExist(adbPath)
    adbPath := folderPath . "\MuMuPlayer-12.0\shell\adb.exe"
if !FileExist(adbPath) ;MuMu Player 12 v5
    adbPath := folderPath . "\MuMuPlayerGlobal-12.0\nx_main\adb.exe"
if !FileExist(adbPath) ;MuMu Player 12 v5
    adbPath := folderPath . "\MuMu Player 12\nx_main\adb.exe"
if !FileExist(adbPath)
    adbPath := folderPath . "\MuMuPlayer-12.0\nx_main\adb.exe"
if !FileExist(adbPath) ;MuMu Player 12 v5
    adbPath := folderPath . "\MuMuPlayer\nx_main\adb.exe"
if !FileExist(adbPath)
    adbPath := folderPath . "\MuMuPlayer-12\shell\adb.exe"
if !FileExist(adbPath)
    adbPath := folderPath . "\MuMuPlayer-12\nx_main\adb.exe"
if !FileExist(adbPath)
    adbPath := folderPath . "\MuMuPlayer12\shell\adb.exe"
if !FileExist(adbPath)
    adbPath := folderPath . "\MuMuPlayer12\nx_main\adb.exe"

if !FileExist(adbPath) {
    MsgBox, 16, , Double check your folder path! It should be the one that contains the MuMuPlayer 12 folder! `nDefault is just C:\Program Files\Netease
    ExitApp
}

if(!adbPorts) {
    Msgbox, 16, , Invalid port... Check the common issues section in the readme/github guide.
    ExitApp
}

filePath := selectedFilePath
if (filePath = "")
    filePath := A_ScriptDir . "\" . fileName . ".xml"

if(!FileExist(filePath)) {
    Msgbox, 16, , Can't find XML file: %filePath% ;'
    ExitApp
}
RunWait, %adbPath% connect 127.0.0.1:%adbPorts%,, Hide

MaxRetries := 10
    RetryCount := 0
    Loop {
        try {
            if (!adbShell) {
                adbShell := ComObjCreate("WScript.Shell").Exec(adbPath . " -s 127.0.0.1:" . adbPorts . " shell")
                processID := adbShell.ProcessID
                WinWait, ahk_pid %processID%
                WinMinimize, ahk_pid %processID%  ; Minimize immediately after window appears
                adbShell.StdIn.WriteLine("su")
            }
            else if (adbShell.Status != 0) {
                Sleep, 1000
            }
            else {
                Sleep, 1000
                break
            }
        }
        catch {
            RetryCount++
            if(RetryCount > MaxRetries) {
                Pause
            }
        }
        Sleep, 1000
    }

    if (!ValidateTradeParams())
        return

    loadAccount()
    if (adbShell) {
        WinClose, ahk_pid %processID%  ; Force close the window
        adbShell.Terminate()
        adbShell := ""
    }

    ExecuteTradeShare()
return

ToggleTradeSection:
    if (tradeSectionExpanded)
        ApplyCollapseState()
    else
        ApplyExpandState()
    return

ApplyCollapseState() {
    global tradeSectionExpanded, tradeBaseY
    tradeSectionExpanded := false
    GuiControl, Hide, TradeRadioTrade
    GuiControl, Hide, TradeRadioShare
    GuiControl, Hide, RunTradeBtn
    GuiControl, Hide, SetDropLabel
    GuiControl, Hide, SelectedSet
    GuiControl, Hide, CardDropLabel
    GuiControl, Hide, SelectedCard
    GuiControl, Hide, FriendIDLabel
    GuiControl, Hide, FriendIDInput
    sepY  := tradeBaseY + 10
    btnY  := sepY + 25
    winH  := btnY + 40 + 20
    GuiControl, Move, BottomSeparator, y%sepY%
    GuiControl, Move, SubmitBtn, y%btnY%
    GuiControl, Move, RunInstanceBtn, y%btnY%
    Gui, Show, w470 h%winH%, Arturo's Account Injection Tool
    GuiControl, , ToggleTradeBtn, >
}

ApplyExpandState() {
    global tradeSectionExpanded, SetDataList, CardDataBySet, tradeBaseY
    tradeSectionExpanded := true
    if (SetDataList.MaxIndex() = "" || SetDataList.MaxIndex() = 0)
        LoadPokemonData()
    PopulateSetsDropdown()
    radioY  := tradeBaseY + 15
    friendY := radioY + 23 + 10
    setY    := friendY + 24 + 10
    cardY   := setY + 24 + 10
    sepY    := cardY + 24 + 10
    btnY    := sepY + 25
    winH    := btnY + 40 + 20
    GuiControl, Move, TradeRadioTrade, y%radioY%
    GuiControl, Move, TradeRadioShare, y%radioY%
    GuiControl, Move, RunTradeBtn, y%radioY%
    GuiControl, Move, FriendIDLabel, y%friendY%
    GuiControl, Move, FriendIDInput, y%friendY%
    GuiControl, Move, SetDropLabel, y%setY%
    GuiControl, Move, SelectedSet, y%setY%
    GuiControl, Move, CardDropLabel, y%cardY%
    GuiControl, Move, SelectedCard, y%cardY%
    GuiControl, Move, BottomSeparator, y%sepY%
    GuiControl, Move, SubmitBtn, y%btnY%
    GuiControl, Move, RunInstanceBtn, y%btnY%
    GuiControl, Show, TradeRadioTrade
    GuiControl, Show, TradeRadioShare
    GuiControl, Show, RunTradeBtn
    GuiControl, Show, FriendIDLabel
    GuiControl, Show, FriendIDInput
    GuiControl, Show, SetDropLabel
    GuiControl, Show, SelectedSet
    GuiControl, Show, CardDropLabel
    GuiControl, Show, SelectedCard
    Gui, Show, w470 h%winH%, Arturo's Account Injection Tool
    GuiControl, , ToggleTradeBtn, <
}

OnSetChange:
    Gui, Submit, NoHide
    global CardDataBySet
    if (SelectedSet = "")
        return
    selectedSetCode := Trim(SubStr(SelectedSet, 1, InStr(SelectedSet, " - ") - 1))
    cardListStr := ""
    if (CardDataBySet.HasKey(selectedSetCode)) {
        cards := CardDataBySet[selectedSetCode]
        Loop, % cards.MaxIndex() {
            c := cards[A_Index]
            if (cardListStr != "")
                cardListStr .= "|"
            cardListStr .= c.idx . " - " . c.name
        }
    }
    GuiControl, , SelectedCard, |%cardListStr%
    GuiControl, Enable, SelectedCard
    return

RunTradeAction:
    Gui, Submit, NoHide
    if (!ValidateTradeParams())
        return
    ExecuteTradeShare()
    return

LoadPokemonData() {
    global SetDataList, CardDataBySet, AllPokemonsCSV, AllPokemons
    SetDataList := []
    CardDataBySet := {}
    setAdded := {}
    AllPokemons := ReadCSV(AllPokemonsCSV)
    Loop, % AllPokemons.MaxIndex() {
        row := AllPokemons[A_Index]
        sc := row["set_code"]
        sn := row["set_name"]
        if (!setAdded.HasKey(sc)) {
            setAdded[sc] := 1
            SetDataList.Push({code: sc, name: sn})
        }
        if (!CardDataBySet.HasKey(sc))
            CardDataBySet[sc] := []
        CardDataBySet[sc].Push({id: row["id"], idx: row["card_index"], name: row["pokemon_name"]})
    }
}

PopulateSetsDropdown() {
    global SetDataList
    setListStr := ""
    Loop, % SetDataList.MaxIndex() {
        s := SetDataList[A_Index]
        if (setListStr != "")
            setListStr .= "|"
        setListStr .= s.code . " - " . s.name
    }
    GuiControl, , SelectedSet, |%setListStr%
    GuiControl, Disable, SelectedCard
    GuiControl, , SelectedCard, |
}

ValidateTradeParams() {
    global tradeSectionExpanded, FriendIDInput, SelectedSet, SelectedCard
    if (!tradeSectionExpanded)
        return true
    Gui, Submit, NoHide
    if (!RegExMatch(FriendIDInput, "^\d{16}$")) {
        MsgBox, 16, Error, Friend ID must be exactly 16 digits.
        return false
    }
    if (SelectedSet = "" || SelectedCard = "") {
        MsgBox, 16, Error, Select both a set and a card.
        return false
    }
    return true
}

ExecuteTradeShare() {
    global tradeSectionExpanded, CardDataBySet, SelectedCard, SelectedSet, winTitle, TradeRadioShare, FriendIDInput, FriendID

    if (!tradeSectionExpanded)
        return

    selectedSetCode := Trim(SubStr(SelectedSet, 1, InStr(SelectedSet, " - ") - 1))
    selectedCardIdx := Trim(SubStr(SelectedCard, 1, InStr(SelectedCard, " - ") - 1))
    cardId := ""
    if (CardDataBySet.HasKey(selectedSetCode)) {
        cards := CardDataBySet[selectedSetCode]
        Loop, % cards.MaxIndex() {
            c := cards[A_Index]
            if (c.idx = selectedCardIdx) {
                cardId := c.id
                break
            }
        }
    }
    if (cardId = "") {
        MsgBox, 16, Error, Could not find card data for selection.
        return
    }
    tradeMode := TradeRadioShare ? "Share" : "Trade"
    IniWrite, %FriendIDInput%, %A_ScriptDir%\..\Settings.ini, UserSettings, FriendID
    IniWrite, Trade Only, %A_ScriptDir%\..\Settings.ini, UserSettings, deleteMethod
    FriendID := FriendIDInput
    iniPath := A_ScriptDir . "\..\Scripts\" . winTitle . ".ini"
    IniWrite, %cardId%, %iniPath%, UserSettings, cardId
    IniWrite, %tradeMode%, %iniPath%, UserSettings, tradeMode
    targetScript := A_ScriptDir . "\..\Scripts\" . winTitle . ".ahk"
    Run, "%targetScript%"
}

getMumuFolder(folderPath) {
mumuFolder := folderPath . "\MuMuPlayerGlobal-12.0"
if !FileExist(mumuFolder)
    mumuFolder := folderPath . "\MuMu Player 12"
if !FileExist(mumuFolder)
    mumuFolder := folderPath . "\MuMuPlayer-12.0"
if !FileExist(mumuFolder)
    mumuFolder := folderPath . "\MuMuPlayer"
if !FileExist(mumuFolder)
    mumuFolder := folderPath . "\MuMuPlayer-12"
if !FileExist(mumuFolder)
    mumuFolder := folderPath . "\MuMuPlayer12"
return mumuFolder
}


findAdbPorts(mumuFolderParam) {
    global adbPorts, winTitle
    ; Initialize variables
    adbPorts := 0  ; Create an empty associative array for adbPorts
    mumuFolderPath = %mumuFolderParam%\vms\*
    if !FileExist(mumuFolderPath){
        MsgBox, 16, , Double check your folder path! It should be the one that contains the MuMuPlayer 12 folder! `nDefault is just C:\Program Files\Netease
        ExitApp
    }
    ; Loop through all directories in the base folder
    Loop, Files, %mumuFolderPath%, D  ; D flag to include directories only
    {
        folder := A_LoopFileFullPath
        configFolder := folder "\configs"  ; The config folder inside each directory

        ; Check if config folder exists
        IfExist, %configFolder%
        {
            ; Define paths to vm_config.json and extra_config.json
            vmConfigFile := configFolder "\vm_config.json"
            extraConfigFile := configFolder "\extra_config.json"

            ; Check if vm_config.json exists and read adb host port
            IfExist, %vmConfigFile%
            {
                FileRead, vmConfigContent, %vmConfigFile%
                ; Parse the JSON for adb host port
                RegExMatch(vmConfigContent, """host_port"":\s*""(\d+)""", adbHostPort)
                adbPort := adbHostPort1  ; Capture the adb host port value
            }

            ; Check if extra_config.json exists and read playerName
            IfExist, %extraConfigFile%
            {
                FileRead, extraConfigContent, %extraConfigFile%
                ; Parse the JSON for playerName
                RegExMatch(extraConfigContent, """playerName"":\s*""(.*?)""", playerName)
                if(playerName1 = winTitle) {
                    adbPorts := adbPort
                }
            }
        }
    }
}

loadAccount() {
    global adbShell, adbPath, adbPorts, fileName, selectedFilePath

    static UserPreferencesPath := "/data/data/jp.pokemon.pokemontcgp/files/UserPreferences/v1/"
    static UserPreferences := ["BattleUserPrefs"
        ,"FeedUserPrefs"
        ,"FilterConditionUserPrefs"
        ,"HomeBattleMenuUserPrefs"
        ,"MissionUserPrefs"
        ,"NotificationUserPrefs"
        ,"PackUserPrefs"
        ,"PvPBattleResumeUserPrefs"
        ,"RankMatchPvEResumeUserPrefs"
        ,"RankMatchUserPrefs"
        ,"SoloBattleResumeUserPrefs"
        ,"SortConditionUserPrefs"]

    if (!adbShell) {
        adbShell := ComObjCreate("WScript.Shell").Exec(adbPath . " -s 127.0.0.1:" . adbPorts . " shell")
        ; Extract the Process ID
        processID := adbShell.ProcessID

        ; Wait for the console window to open using the process ID
        WinWait, ahk_pid %processID%

        ; Minimize the window using the process ID
        WinMinimize, ahk_pid %processID%
    }

    adbShell.StdIn.WriteLine("am force-stop jp.pokemon.pokemontcgp")
    Sleep, 200

    ; Clear app data to ensure no previous account information remains
    adbShell.StdIn.WriteLine("rm -f /data/data/jp.pokemon.pokemontcgp/shared_prefs/deviceAccount:.xml")
    Sleep, 200

    Loop, % UserPreferences.MaxIndex() {
        adbShell.StdIn.WriteLine("rm -f " . UserPreferencesPath . UserPreferences[A_Index])
        Sleep, 200
    }

    loadDir := selectedFilePath
    if (loadDir = "")
        loadDir := A_ScriptDir . "\" . fileName . ".xml"
    else {
        ; Don't append .xml if the path already ends with it
        SplitPath, loadDir, , , fileExt
        if (fileExt != "xml")
            loadDir := loadDir . ".xml"
    }

    ; Make sure the file exists before trying to push it
    if (!FileExist(loadDir)) {
        MsgBox, 16, Error, Cannot find the XML file: %loadDir%
        ExitApp
    }

    ; Push the file to the device with better error handling
    RunWait, % adbPath . " -s 127.0.0.1:" . adbPorts . " push """ . loadDir . """ /sdcard/deviceAccount.xml",, Hide
    Sleep, 150

    ; Create the shared_prefs directory if it doesn't exist
    adbShell.StdIn.WriteLine("mkdir -p /data/data/jp.pokemon.pokemontcgp/shared_prefs")
    Sleep, 100

    ; Copy the file with proper permissions
    adbShell.StdIn.WriteLine("cp /sdcard/deviceAccount.xml /data/data/jp.pokemon.pokemontcgp/shared_prefs/deviceAccount:.xml")
    Sleep, 100

    ; Set proper permissions and ownership (combined commands with shorter delay)
    adbShell.StdIn.WriteLine("chmod 664 /data/data/jp.pokemon.pokemontcgp/shared_prefs/deviceAccount:.xml && chown system:system /data/data/jp.pokemon.pokemontcgp/shared_prefs/deviceAccount:.xml")
    Sleep, 200

    ; Clean up and launch app (reduced delay between operations)
    adbShell.StdIn.WriteLine("rm /sdcard/deviceAccount.xml")

    ; Launch the app with both commands in quick succession
    adbShell.StdIn.WriteLine("am start -n jp.pokemon.pokemontcgp/jp.pokemon.pokemontcgp.UnityPlayerActivity")
    Sleep, 100

    adbShell.StdIn.WriteLine("am start -n jp.pokemon.pokemontcgp/com.unity3d.player.UnityPlayerActivity")

    ; Close the shell after all operations complete
    adbShell.Terminate()
    adbShell := ""
}

; New function to get instance list
GetInstanceList(baseFolder) {
    instanceList := ""
    mumuFolder := getMumuFolder(baseFolder)

    ; Loop through all VM directories
    Loop, Files, %mumuFolder%\vms\*, D
    {
        folder := A_LoopFileFullPath
        configFolder := folder "\configs"

        if InStr(FileExist(configFolder), "D") {
            extraConfigFile := configFolder "\extra_config.json"

            if FileExist(extraConfigFile) {
                FileRead, fileContent, %extraConfigFile%
                RegExMatch(fileContent, """playerName"":\s*""(.*?)""", playerName)
                if (playerName1 != "") {
                    if (instanceList != "")
                        instanceList .= "|"
                    instanceList .= playerName1
                }
            }
        }
    }

    return instanceList
}

; Refresh button handler
RefreshInstances:
    refreshedList := GetInstanceList(folderPath)
    GuiControl,, winTitle, |%refreshedList%
    return

RunInstance:
    Gui, Submit, NoHide
    mumuFolder := getMumuFolder(folderPath)
    ; Find the instance number matching the selected name
    instanceNum := ""
    Loop, Files, %mumuFolder%\vms\*, D
    {
        folder := A_LoopFileFullPath
        configFolder := folder "\configs"
        if InStr(FileExist(configFolder), "D") {
            extraConfigFile := configFolder "\extra_config.json"
            if FileExist(extraConfigFile) {
                FileRead, fileContent, %extraConfigFile%
                RegExMatch(fileContent, """playerName"":\s*""(.*?)""", playerName)
                if (playerName1 = winTitle) {
                    RegExMatch(folder, "[^-]+$", instanceNum)
                    break
                }
            }
        }
    }
    if (instanceNum != "") {
        mumuExe := mumuFolder . "\shell\MuMuPlayer.exe"
        if !FileExist(mumuExe)
            mumuExe := mumuFolder . "\nx_main\MuMuNxMain.exe"
        if FileExist(mumuExe) {
            Run, "%mumuExe%" -v "%instanceNum%"
        } else {
            MsgBox, 16, Error, Could not find MuMuPlayer.exe at %mumuExe%
        }
    }
    else {
        MsgBox, 16, Error, Could not find instance number for %winTitle%
        }
    return
