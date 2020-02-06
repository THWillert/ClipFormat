; :wrap=none:collapseFolds=1:maxLineLen=80:mode=autoitscript:tabSize=8:folding=indent:
#include <Array.au3>
#include <Clipboard.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>
#include <GuiListView.au3>
#include <Inet.au3>
#include <ListBoxConstants.au3>
#include <ListViewConstants.au3>
#include <SendMessage.au3>
#include <String.au3>
#include <WindowsConstants.au3>

#cs

V0.964:
* Neu: Macros in Texten werden nun ersetzt (ausschaltbar über Prefs-Menü)
* Neu: Macros auch in Preffix usw. einsetzbar
* Neu: RegEx-Liste: Die Anzahl der Ersetzungen wird angezeigt (Spalte "Replaced")
* Neu: RegEx-Liste: Die Breite der Spalten wird nun automatisch angepaßt
* Geändert: Tab: "Convert": Bei Int2Bin wird nun eine zweite Auswhahl für die Anzahl der Stellen eingeblendet
* Korrektur: Einzelne Wörter werden nun auch auch richtig "Formatiert" (Tab: "Format")
* Korrektur: Die Abstände der Satzzeichen werden nun VOR dem Formatieren korrigiert.

V0.963:
* Neu: Angefangen die Oberfläche umzubauen (Automatic On / Off kann nun über Button bedient werden, anstatt des Umwegs über das Menü)
* Neu: Preffix / Suffix Datumsangaben mit nun auch mit zweistelliger Jahreszahl
* Korrektur: Add / Insert ist nun das Up/Down-Control wieder an der richtigen Stelle

V0.962:
* Neu: Auf dem Tab "Format" eine Option zur Korrektur der Leerzeichen vor und nach Satzzeichen eingefügt
* Optimierung: Große RegEx-Listen sind nun wesentlich schneller
* Korrektur: Fehler bei der Datumsauswahl Preffix / Suffix behoben
* Korrektur: Der "Exe"-Parameter der RegEx-Liste wird nun wieder korrekt aus der Ini geladen

#ce

AutoItSetOption("TrayIconHide", 1)
OnAutoItExitRegister("_WriteIni")

Const $sDATA = " |MM-DD|YYYY-MM-DD|YY-MM-DD|YYYY-MM-DD HH:MI:SS|YY-MM-DD HH:MI:SS|MM/DD|MM/DD/YYYY|MM/DD/YY|MM/DD/YYYY HH:MI:SS|MM/DD/YY HH:MI:SS|HH:MI:SS|YYYY|YY|EPOCH"
Global Const $sIni = @ScriptDir & "\" & @ScriptName & ".ini"

Global $aRegEx[1000][3]
Global $aMacro[56][3]
Global $aFunc_List[15]
Global $sTextIn, $sTextOut
Global $sLast
Global $sCLIPBOARD
Global Const $VERSION = "V0.964"

#tidy_off
;==============================================================================
$ClipFormat = GUICreate("ClipFormat",440,355,-1,-1,-1,-1)
   $hMenu = GUICtrlCreateMenu("Prefs")
   $hAuto = GUICtrlCreateMenuItem("Automatic",$hMenu)
   GUICtrlSetState(-1, $GUI_Checked)
   $hMenuMacroCheck = GUICtrlCreateMenuItem("Replace Macros",$hMenu)
   GUICtrlSetState(-1, $GUI_Checked)
   $hMenuHistorySave = GUICtrlCreateMenuItem("Save History",$hMenu)
   GUICtrlSetState(-1, $GUI_Checked)
   GUICtrlCreateMenuItem("",$hMenu)
   $hMenuFuncList = GUICtrlCreateMenuItem("Function List ...",$hMenu)

   $hViewMenu = GUICtrlCreateMenu("View")
   $hTop = GUICtrlCreateMenuItem("Stay On Top",$hViewMenu)
   GUICtrlSetState(-1, $GUI_Checked)
   GUICtrlCreateMenuItem("",$hViewMenu)
   $hViewHistory = GUICtrlCreateMenuItem("History ...",$hViewMenu)
   $hViewRegEx = GUICtrlCreateMenuItem("RegExReplace ...",$hViewMenu)
   $hViewMacros = GUICtrlCreateMenuItem("Macros ...",$hViewMenu)

   $hExtras = GUICtrlCreateMenu("Extras")
   $hExtras_SortLinesUP = GUICtrlCreateMenuItem("Sort Lines UP",$hExtras)
   $hExtras_SortLinesDOWN = GUICtrlCreateMenuItem("Sort Lines DOWN",$hExtras)


   $hInfoMenu = GUICtrlCreateMenu("Info")
   $hAbout = GUICtrlCreateMenuItem("About ...",$hInfoMenu)
   $hIn = GUICtrlCreateEdit("",10,10,420,55,2244,512)
   $hOut = GUICtrlCreateEdit("",10,70,420,55,2244,512)

   $tab = GUICtrlCreateTab(10, 130, 420, 145)

   GUICtrlCreateTabItem("Format")
	  $hComboStyle = GUICtrlCreateCombo("",20,160,400,21,3,-1)
	  GUICtrlSetData(-1,"---|Sentence|Title|Title 2|Proper|Uppercase|Lowercase|Reverse")
	  $hFormat_Check_Mark = GUICtrlCreateCheckbox("Correct Punctuation Space",20,190,200,20,-1,-1)
	  GUICtrlSetState(-1,84)

   GUICtrlCreateTabItem("Cut")
	  $hCheck_CutFirst = GUICtrlCreateCheckbox("First n",20,160,80,20,-1,-1)
	  GUICtrlSetState(-1,84)
	  $hCutFirst_Pos = GUICtrlCreateInput("", 100, 160, 45,20,$ES_NUMBER)
	  GUICtrlCreateUpdown(-1)

	  $hCheck_CutLast = GUICtrlCreateCheckbox("Last n",20,190,80,20,-1,-1)
	  GUICtrlSetState(-1,84)
	  $hCutLast_Pos = GUICtrlCreateInput("", 100, 190, 45,20,$ES_NUMBER)
	  GUICtrlCreateUpdown(-1)

	  $hCheck_CutMid = GUICtrlCreateCheckbox("From",20,220,80,20,-1,-1)
	  GUICtrlSetState(-1,84)
	  $hCutMidStart_Pos = GUICtrlCreateInput("", 100, 220, 45,20,$ES_NUMBER)
	  GUICtrlCreateUpdown(-1)
	  GUICtrlCreateLabel("To",170,220,20,20)
	  $hCutMidEnd_Pos = GUICtrlCreateInput("", 200, 220, 45,20,$ES_NUMBER)
	  GUICtrlCreateUpdown(-1)

   GUICtrlCreateTabItem("Add")
	  $hCheckPreffix = GUICtrlCreateCheckbox("Add Preffix",20,160,70,20,-1,-1)
	  GUICtrlSetState(-1,84)
	  $hPreffix = GUICtrlCreateCombo("",100,160,150,20,3,-1)
	  GUICtrlSetData(-1,$sDATA)
	  $hPreffix2 = GUICtrlCreateInput("", 255, 160, 145,20)
	  $hPreffix1 = GUICtrlCreateInput("", 405, 160, 15,20)

	  $hCheckSuffix = GUICtrlCreateCheckbox("Add Suffix",20,190,70,20,-1,-1)
	  GUICtrlSetState(-1,84)
	  $hSuffix1 = GUICtrlCreateInput("", 100, 190, 15,20)
	  $hSuffix2 = GUICtrlCreateInput("", 120, 190, 135,20)
	  $hSuffix = GUICtrlCreateCombo("",260,190,160,20,3,-1)
	  GUICtrlSetData(-1,$sDATA)

	  $hCheckInsert = GUICtrlCreateCheckbox("Insert",20,220,80,20,-1,-1)
	  GUICtrlSetState(-1,84)
		$hInsert_Pos = GUICtrlCreateInput("", 100, 220, 45,20,$ES_NUMBER)
		 GUICtrlCreateUpdown(-1)
	  $hInsert_Input = GUICtrlCreateInput("", 150, 220, 135,20)


   GUICtrlCreateTabItem("Replace")
	  $hStripWS = GUICtrlCreateCheckbox("Trim WS",20,160,100,20,-1,-1)
	  GUICtrlSetState(-1,81)

	  $sDoubleWS = GUICtrlCreateCheckbox("Strip double WS",20,180,100,20,-1,-1)
	  GUICtrlSetState(-1,81)

	  $hStripCR = GUICtrlCreateCheckbox("Strip CR",20,200,100,20,-1,-1)
	  GUICtrlSetState(-1,81)

	  $hNONAscii = GUICtrlCreateCheckbox("Remove Non ASCII",20,220,150,20,-1,-1)
	  GUICtrlSetState(-1,81)

	  $hRegExReplace = GUICtrlCreateCheckbox("RegExReplace",20,240,100,20,-1,-1)
	  GUICtrlSetState(-1,81)

   GUICtrlCreateTabItem("Convert")
	  $hComboConvert = GUICtrlCreateCombo("",20,160,400,21,3,-1)
	  GUICtrlSetData(-1,"---|URL encode|URL decode|IP To Name|Name To IP|String2Hex|Hex2String|RGB2Hex|Hex2RGB|Int2Bin|Bin2Int")
		
	  	; BIN / INT
	  	$hComboconvertInt2Bin = GUICtrlCreateCombo("",20,190,400,21,3,-1)
		GUICtrlSetData($hComboconvertInt2Bin,"Int2Bin (Auto)|Int2Bin 8|Int2Bin 16|Int2Bin 32|Int2Bin 64")
		GUICtrlSetState($hComboconvertInt2Bin, $GUI_HIDE)

	  ;$hResolveCheck = GUICtrlCreateCheckbox("Resolve IP",20,160,100,20,-1,-1)
	  ;GUICtrlSetState(-1,81)
   GUICtrlCreateTabItem("Comment")
	  GUICtrlCreateLabel("Line Comment",20,160,100,20)
	  $hComboLineComment = GUICtrlCreateCombo("",120,160,300,21,3,-1)
	  GUICtrlSetData(-1," |  //|   #|   ;|   '|   Rem|   --|   %|   *")
	  GUICtrlCreateLabel("Block Comment",20,190,100,20)
	  $hComboBlockComment = GUICtrlCreateCombo("",120,190,300,21,3,-1)
	  GUICtrlSetData(-1," |  /*    */|   (*    *)|   {    }|   #cs    #ce|   <!--    -->|   %{    %}|   {-    -}")

   GUICtrlCreateTabItem("")

   $ihSpace = 280

   $hBut_Auto = GUICtrlCreateButton("Automatic ON", 5, $ihSpace, 100, 25)

	$ihSpace = 310
   $hStatus1 = GUICtrlCreateInput("",5,$ihSpace,200,20,$ES_READONLY)
   $hStatus2 = GUICtrlCreateInput("",205,$ihSpace,160,20,$ES_READONLY+$ES_CENTER)
   $hStatus3 = GUICtrlCreateInput("",365,$ihSpace,40,20,$ES_READONLY+$ES_CENTER)
   $hStatus4 = GUICtrlCreateInput("",405,$ihSpace,30,20,$ES_READONLY+$ES_CENTER)

GUISetState(@SW_SHOW,$ClipFormat)
;==============================================================================
$hWinHist = GUICreate("History, ClipFormat",440,380,-1,-1,$WS_SIZEBOX,-1,$ClipFormat)
   $hLast = GUICtrlCreatelist("",0,0,440,360,-1,-1)
   GUICtrlSetResizing(-1,$GUI_DOCKBORDERS)
   $hlistmenu = GUICtrlCreateContextMenu($hLast)
   $hmenu_lv_edit = GUICtrlCreateMenuItem("Copy to Clipboard",$hlistmenu)
   $hmenu_lv_add = GUICtrlCreateMenuItem("Add to Clipboard",$hlistmenu)
   GUICtrlCreateMenuItem("",$hlistmenu)
   $hmenu_lv_delete = GUICtrlCreateMenuItem("Delete",$hlistmenu)
   $hmenu_lv_clear = GUICtrlCreateMenuItem("Delete All",$hlistmenu)

   _Load_History()
GUISetState(@SW_HIDE,$hWinHist)
;==============================================================================
$hWinFunc = GUICreate("Functionlist, ClipFormat",300,360,-1,-1,-1,-1,$ClipFormat)
   $hFuncList = GUICtrlCreatelist("",0,0,300,320,$LBS_DISABLENOSCROLL,-1)
   GUICtrlSetResizing(-1,$GUI_DOCKBORDERS)
   $hFuncList_Up = GUICtrlCreateButton("Up", 10,325,70,25)
   $hFuncList_Down = GUICtrlCreateButton("Down", 80,325,70,25)
   $hFuncList_Test = GUICtrlCreateButton("Test", 170,325,70,25)
   $hFuncList_Reset = GUICtrlCreateButton("Reset", 240,325,50,25)
   _Load_FuncList()
GUISetState(@SW_HIDE,$hWinFunc)
;==============================================================================
$hWinMacros = GUICreate("Macros, ClipFormat",440,380,-1,-1,$WS_SIZEBOX,-1,$ClipFormat)
   $hMacroView = GUICtrlCreateListView("Type|Macro|Discription",0,0,440,355,$LVS_REPORT+$LVS_SINGLESEL+$LVS_SHOWSELALWAYS,$LVS_EX_FULLROWSELECT+$LVS_EX_GRIDLINES)
   GUICtrlSetResizing(-1,$GUI_DOCKBORDERS)
   _MacroList()
   _GUICtrlListView_AddArray($hMacroView, $aMacro)
   $hMacroMenu = GUICtrlCreateContextMenu($hMacroView)
   $hMacroMenu_Copy = GUICtrlCreateMenuItem("Copy Macro",$hMacroMenu)
   $hMacroMenu_CopyValue = GUICtrlCreateMenuItem("Copy Current Value",$hMacroMenu)

   _ListView_AutoColSize($hMacroView)
GUISetState(@SW_HIDE,$hWinMacros)
;==============================================================================
$WinReplace = GUICreate("RegExReplace, ClipFormat",440,405,-1,-1,$WS_SIZEBOX,-1,$ClipFormat)
   $hReplaceView = GUICtrlCreateListView("Exe|Search|Replace|Comment|Replaced|ERROR",0,0,440,255,$LVS_REPORT+$LVS_SINGLESEL+$LVS_SHOWSELALWAYS+$LVS_NOSORTHEADER,$LVS_EX_FULLROWSELECT+$LVS_EX_GRIDLINES+$LVS_EX_CHECKBOXES)
   GUICtrlSetResizing(-1,$GUI_DOCKBORDERS)

   $hRegExmenu = GUICtrlCreateContextMenu($hReplaceView)
   $hRegEx_Select = GUICtrlCreateMenuItem("Select All",$hRegExmenu)
   $hRegEx_DeSelect = GUICtrlCreateMenuItem("Deselect All",$hRegExmenu)
   GUICtrlCreateMenuItem("",$hRegExmenu)
   $hRegExmenu_delete = GUICtrlCreateMenuItem("Delete",$hRegExmenu)
   GUICtrlCreateMenuItem("",$hRegExmenu)
   $hRegEx_Insert = GUICtrlCreateMenuItem("Insert",$hRegExmenu)
   $hRegEx_Add = GUICtrlCreateMenuItem("Add",$hRegExmenu)

   $hRegEx_Up = GUICtrlCreateButton("Up",10,260,70)
   GUICtrlSetResizing(-1,$GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKSIZE)
   $hRegEx_Down = GUICtrlCreateButton("Down",80,260,70)
   GUICtrlSetResizing(-1,$GUI_DOCKLEFT+$GUI_DOCKBOTTOM+$GUI_DOCKSIZE)

   $hRegEx_Search = GUICtrlCreateInput("Search (Regex)",5,290,430)
   GUICtrlSetResizing(-1,$GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKHEIGHT)
   $hRegEx_Replace = GUICtrlCreateInput("Replace (Regex)",5,320,430)
   GUICtrlSetResizing(-1,$GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKHEIGHT)
   $hRegEx_Comment = GUICtrlCreateInput("Comment",5,350,430)
   GUICtrlSetResizing(-1,$GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKBOTTOM+$GUI_DOCKHEIGHT)

   _Load_RegEx()
   _RegEx_List2Array()
   _ListView_AutoColSize($hReplaceView)
GUISetState(@SW_HIDE,$WinReplace)
#tidy_on
;==============================================================================

Global $hNext = _ClipBoard_SetViewer($ClipFormat)

GUIRegisterMsg($WM_CHANGECBCHAIN, "WM_CHANGECBCHAIN")
GUIRegisterMsg($WM_DRAWCLIPBOARD, "WM_DRAWCLIPBOARD")

; Init
If _IsChecked($hTop) Then WinSetOnTop("ClipFormat", "", 1)
If _IsChecked($hAuto) Then
	AdlibRegister("_GetClipBoard")
EndIf

_GetClipBoard()

;==============================================================================
While Sleep(10)
	$naMsg = GUIGetMsg(1)
	Switch $naMsg[1]
		Case $ClipFormat
			Switch $naMsg[0]
				Case $GUI_EVENT_CLOSE
					Exit
				Case $hViewHistory
					GUISetState(@SW_SHOW, $hWinHist)
				Case $hViewRegEx
					GUISetState(@SW_SHOW, $WinReplace)
				Case $hViewMacros
					GUISetState(@SW_SHOW, $hWinMacros)
				Case $hMenuFuncList
					GUISetState(@SW_SHOW, $hWinFunc)
				Case $hAuto, $hBut_Auto
					If _IsChecked($hAuto) Then
						AdlibUnRegister("_GetClipBoard")
						GUICtrlSetState($hAuto, $GUI_UNCHECKED)
						GUICtrlSetData($hBut_Auto, "Automatic OFF")
					Else
						$sCLIPBOARD = _ClipBoard_GetData($CF_UNICODETEXT)
						AdlibRegister("_GetClipBoard")
						GUICtrlSetState($hAuto, $GUI_Checked)
						GUICtrlSetData($hBut_Auto, "Automatic ON")
					EndIf
				Case $hTop
					If _IsChecked($hTop) Then
						WinSetOnTop("ClipFormat", "", 0)
						GUICtrlSetState($hTop, $GUI_UNCHECKED)
					Else
						WinSetOnTop("ClipFormat", "", 1)
						GUICtrlSetState($hTop, $GUI_Checked)
					EndIf
				Case $hMenuHistorySave
					If _IsChecked($hMenuHistorySave) Then
						GUICtrlSetState($hMenuHistorySave, $GUI_UNCHECKED)
					Else
						GUICtrlSetState($hMenuHistorySave, $GUI_Checked)
					EndIf
				Case $hMenuMacroCheck
					If _IsChecked($hMenuMacroCheck) Then
						GUICtrlSetState($hMenuMacroCheck, $GUI_UNCHECKED)
					Else
						GUICtrlSetState($hMenuMacroCheck, $GUI_Checked)
					EndIf
				Case $hExtras_SortLinesUP
					_SortLines()
				Case $hExtras_SortLinesDOWN
					_SortLines(1)
				Case $hMenuMacroCheck, $hComboStyle, $hFormat_Check_Mark, _
						$hStripWS, $hCheckPreffix, $hCheckSuffix, _
						$sDoubleWS, $hStripWS, $hPreffix, $hPreffix1, $hPreffix2, _
						$hSuffix, $hSuffix1, $hSuffix2, $hRegExReplace, $hStripCR, _
						$hNONAscii, $hCheckInsert, $hInsert_Pos, $hInsert_Input, _
						$hCheck_CutFirst, $hCutFirst_Pos, $hCheck_CutLast, $hCutLast_Pos, _
						$hCheck_CutMid, $hCutMidStart_Pos, $hCutMidEnd_Pos, _
						$hComboConvert, $hComboconvertInt2Bin,  _
						$hComboLineComment, $hComboBlockComment

					$sTMP = GUICtrlRead($hComboConvert)
					If $sTMP = "Int2Bin" Then
						GUICtrlSetState($hComboconvertInt2Bin, $GUI_SHOW)
					Else
						GUICtrlSetState($hComboconvertInt2Bin, $GUI_HIDE)
					EndIf

					If _IsChecked($hAuto) Then
						AdlibUnRegister("_GetClipBoard")
						GUICtrlSetState($hAuto, $GUI_UNCHECKED)
						GUICtrlSetData($hBut_Auto, "Automatic OFF")
					EndIf

					_AutoFormat()
				Case $hAbout
					WinSetOnTop("ClipFormat", "", 0)
					MsgBox(64, "ClipFormat", "(c) 2008-" & @YEAR & " by Thorsten Willert, " & $VERSION & @CRLF & "www.thorsten-willert.de", 0, $ClipFormat)
					If _IsChecked($hTop) Then WinSetOnTop("ClipFormat", "", 1)
			EndSwitch
		Case $hWinHist
			Switch $naMsg[0]
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_HIDE, $hWinHist)
				Case $hmenu_lv_edit
					ClipPut(GUICtrlRead($hLast))
				Case $hmenu_lv_add
					_Clip_Add()
				Case $hmenu_lv_clear
					WinSetOnTop("ClipFormat", "", 0)
					If MsgBox(1 + 32, "ClipFormat", "Delete History?") = 1 Then _GUICtrlListBox_ResetContent($hLast)
					If _IsChecked($hTop) Then WinSetOnTop("ClipFormat", "", 1)
				Case $hmenu_lv_delete
					_GUICtrlListBox_DeleteString($hLast, _GUICtrlListBox_GetCurSel($hLast))
			EndSwitch
		Case $hWinFunc
			Switch $naMsg[0]
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_HIDE, $hWinFunc)
				Case $hFuncList_Up
					_FuncList_Sort(1)
				Case $hFuncList_Down
					_FuncList_Sort(-1)
				Case $hFuncList_Reset
					_FuncList_Sort(0)
				Case $hFuncList_Test
					_AutoFormat()
			EndSwitch
		Case $WinReplace
			Switch $naMsg[0]
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_HIDE, $WinReplace)
				Case $hRegEx_Add
					_AddRegEx()
					_RegEx_List2Array()
				Case $hRegEx_Insert
					_InsertRegEx()
					_RegEx_List2Array()
					_ListView_AutoColSize($hReplaceView)
				Case $hRegExmenu_delete
					_DeleteRegEx()
					_RegEx_List2Array()
					_ListView_AutoColSize($hReplaceView)
				Case $hReplaceView, $GUI_EVENT_PRIMARYDOWN
					_EditRegEx()
				Case $hRegEx_Search, $hRegEx_Replace, $hRegEx_Comment
					_UpdateRegEx()
					_RegEx_List2Array()
					_ListView_AutoColSize($hReplaceView)
				Case $hRegEx_Select
					_SelectRegex(1)
				Case $hRegEx_DeSelect
					_SelectRegex(0)
				Case $hRegEx_Up
					_RexExList_Sort(1)
					_RegEx_List2Array()
				Case $hRegEx_Down
					_RexExList_Sort(-1)
					_RegEx_List2Array()
			EndSwitch
		Case $hWinMacros
			Switch $naMsg[0]
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_HIDE, $hWinMacros)
				Case $hMacroMenu_Copy
					_Macros_Copy()
				Case $hMacroMenu_CopyValue
					_Macros_CopyValue()
			EndSwitch
	EndSwitch
WEnd
;==============================================================================
Func _FuncList_Sort($m)
	Local $h = GUICtrlGetHandle($hFuncList)
	Local $iIndex = _GUICtrlListBox_GetCurSel($h)
	Local $iMax = _GUICtrlListBox_GetCount($h)

	If $m = 1 Then
		If $iIndex = 0 Then Return
		_GUICtrlListBox_SwapString($h, $iIndex, $iIndex - 1)
		_GUICtrlListBox_SetCurSel($h, $iIndex - 1)
	EndIf
	If $m = -1 Then
		If $iIndex + 1 = $iMax Then Return
		_GUICtrlListBox_SwapString($h, $iIndex, $iIndex + 1)
		_GUICtrlListBox_SetCurSel($h, $iIndex + 1)
	EndIf
	If $m = 0 Then
		_GUICtrlListBox_ResetContent($h)
		_Load_FuncList()
	EndIf

EndFunc   ;==>_FuncList_Sort
;==============================================================================
Func _RexExList_Sort($m)
	Local $h = GUICtrlGetHandle($hReplaceView)
	Local $iIndex = _GUICtrlListView_GetSelectedIndices($h)
	Local $iMax = _GUICtrlListView_GetItemCount($h)

	If $m = 1 Then
		If $iIndex = 0 Then Return
		__GUICtrlListView_Swap($h, $iIndex, $iIndex - 1)
		_GUICtrlListView_SetItemSelected($h, $iIndex - 1)
	EndIf
	If $m = -1 Then
		If $iIndex + 1 = $iMax Then Return
		__GUICtrlListView_Swap($h, $iIndex, $iIndex + 1)
		_GUICtrlListView_SetItemSelected($h, $iIndex + 1)
	EndIf
EndFunc   ;==>_RexExList_Sort
;==============================================================================
Func __GUICtrlListView_Swap($h, $iIndex1, $iIndex2)
	Local $sTMP
	$sTMP = _GUICtrlListView_GetItemTextString($h, $iIndex1)
	__GUICtrlListView_SetItemTextString($h, $iIndex1, _GUICtrlListView_GetItemTextString($h, $iIndex2))
	__GUICtrlListView_SetItemTextString($h, $iIndex2, $sTMP)
EndFunc   ;==>__GUICtrlListView_Swap
;==============================================================================
Func __GUICtrlListView_SetItemTextString($h, $iIndex, $sText)
	Local $aTMP = StringSplit($sText, AutoItSetOption("GUIDataSeparatorChar"), 2)
	For $i = 0 To UBound($aTMP) - 1
		_GUICtrlListView_SetItemText($h, $iIndex, $aTMP[$i], $i)
	Next
EndFunc   ;==>__GUICtrlListView_SetItemTextString
;==============================================================================
Func _Macros_Copy()
	Local $h = GUICtrlGetHandle($hMacroView)
	_SetClipBoard(_GUICtrlListView_GetItemText($h, _GUICtrlListView_GetSelectedIndices($h), 1))
EndFunc   ;==>_Macros_Copy
;==============================================================================
Func _Macros_CopyValue()
	Local $h = GUICtrlGetHandle($hMacroView)
	_SetClipBoard(_Macro_Exec(_GUICtrlListView_GetItemText($h, _GUICtrlListView_GetSelectedIndices($h), 1)))
EndFunc   ;==>_Macros_CopyValue
;==============================================================================
Func _Macro_StringReplace($sString)

	For $i = 0 To UBOund($aMacro)-1
		If StringInStr($sString, $aMacro[$i][1]) Then
			$sString = StringReplace($sString, $aMacro[$i][1], _Macro_Exec( $aMacro[$i][1] ) )
		EndIf
	Next

	Return $sString
EndFunc
;=============================================================================
Func _Macro_Exec($sMacro)
	GUICtrlSetData($hStatus1, "Executing Macro ...")

	Switch $sMacro
		Case "@IPAddressPublic"
			GUICtrlSetData($hStatus1, "Getting Puplic IP ...")
			$sMacro = _GetIP()
			If @error Then $sMacro = "0.0.0.0"
		Case "@WNR_ISO"
			$sMacro = _WeekNumberISO()
		Case "@EPOCH"
			$sMacro = _Epoch_encrypt(@YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC)
			ConsoleWrite($sMacro & @CRLF)
		Case Else
			$sMacro = Execute($sMacro)
	EndSwitch

	GUICtrlSetData($hStatus1, "")

	Return String($sMacro)
EndFunc   ;==>_Macro_Exec
;==============================================================================
Func _Clip_Add()
	_SetClipBoard(_ClipBoard_GetData($CF_UNICODETEXT) & GUICtrlRead($hLast))
EndFunc   ;==>_Clip_Add
;==============================================================================
Func _SetClipBoard($s)
	_ClipBoard_SetData($s, $CF_UNICODETEXT)
EndFunc   ;==>_SetClipBoard
;==============================================================================
Func _GetClipBoard()

	$sTextIn = $sCLIPBOARD

	If @error Or _
		$sTextIn = "" Or _
		StringLen($sTextIn) > 100000 Or _
		$sTextIn = $sTextOut Or _
		$sTextIn = $sLast _
	Then Return

	GUICtrlSetData($hIn, $sTextIn)
	If _IsChecked($hMenuHistorySave) Then GUICtrlSetData($hLast, $sTextIn)
	$sLast = $sTextIn

	_AutoFormat()
	GUICtrlSetData($hStatus1, "")
EndFunc   ;==>_GetClipBoard
;==============================================================================
Func _AutoFormat()
	GUICtrlSetBkColor($hStatus4, 0xFF0000)
	GUICtrlSetData($hStatus1, "Formating ...")
	Local $Timer = TimerInit()
	$sTextOut = GUICtrlRead($hIn)
	Local $sSuffix, $sPreffix
	Local $sCS = "", $sCE = ""
	Local $sTMP = ""

	_FuncList_Read()

	For $i = 0 To UBound($aFunc_List) - 1
		Switch $aFunc_List[$i]
			Case "Format"
				If _IsChecked($hFormat_Check_Mark) Then $sTextOut = _Format_Punctuation($sTextOut)
				$sTextOut = _Format($sTextOut)
			Case "Add Preffix"
				If _IsChecked($hCheckPreffix) Then $sTextOut = _FormatReplace(GUICtrlRead($hPreffix)) & GUICtrlRead($hPreffix2) & GUICtrlRead($hPreffix1) & $sTextOut
			Case "Add Suffix"
				If _IsChecked($hCheckSuffix) Then $sTextOut &= GUICtrlRead($hSuffix1) & GUICtrlRead($hSuffix2) & _FormatReplace(GUICtrlRead($hSuffix))
			Case "Add Insert"
				If _IsChecked($hCheckInsert) Then $sTextOut = _StringInsert($sTextOut, GUICtrlRead($hInsert_Input), Int(GUICtrlRead($hInsert_Pos)))
			Case "Cut First"
				If _IsChecked($hCheck_CutFirst) Then $sTextOut = StringRight($sTextOut, StringLen($sTextOut) - GUICtrlRead($hCutFirst_Pos))
			Case "Cut Last"
				If _IsChecked($hCheck_CutLast) Then $sTextOut = StringLeft($sTextOut, StringLen($sTextOut) - GUICtrlRead($hCutLast_Pos))
			Case "Cut Mid"
				If _IsChecked($hCheck_CutMid) Then $sTextOut = StringLeft($sTextOut, GUICtrlRead($hCutMidStart_Pos)) & StringMid($sTextOut, GUICtrlRead($hCutMidEnd_Pos))
			Case "Replace Trim WS"
				If _IsChecked($hStripWS) Then $sTextOut = StringStripWS($sTextOut, 3)
			Case "Replace Strip double WS"
				If _IsChecked($sDoubleWS) Then $sTextOut = StringRegExpReplace($sTextOut, '(\S+)[\x20|\t]{2,}(?=\S+)', '$1 $2')
			Case "Replace CR"
				If _IsChecked($hStripCR) Then $sTextOut = StringStripCR($sTextOut)
			Case "Replace Remove Non ASCII"
				If _IsChecked($hNONAscii) Then $sTextOut = StringRegExpReplace($sTextOut, "[^\x00-\x7F]", " ")
			Case "Replace RegExReplace"
				If _IsChecked($hRegExReplace) Then $sTextOut = _RegReplace($sTextOut)
			Case "Convert"
				$sTextOut = _Convert($sTextOut)
			Case "Comment Line"
				$sTextOut = _Comment($sTextOut, 1)
			Case "Comment Block"
				$sTextOut = _Comment($sTextOut, 2)
		EndSwitch
	Next

	$sTextOut = $sCS & $sTextOut & $sCE
	If _IsChecked($hMenuMacroCheck) Then $sTextOut = _Macro_StringReplace($sTextOut)

	_SetClipBoard($sTextOut)

	GUICtrlSetData($hOut, $sTextOut)
	GUICtrlSetData($hStatus4, Round(TimerDiff($Timer)))
	GUICtrlSetBkColor($hStatus4, 0xf0f0f0)
	GUICtrlSetData($hStatus1, "")

EndFunc   ;==>_AutoFormat
;=============================================================================
Func _SortLines($iMode = 0)
	$sTextOut = GUICtrlRead($hIn)
	Local $aTMP = StringSplit($sTextOut, @CRLF, 2)
	If Not @error Then
		_ArraySort($aTMP, $iMode)
	Else
		Return
	EndIf
	GUICtrlSetData($hOut, _ArrayToString($aTMP, @CRLF))
EndFunc   ;==>_SortLines
;=============================================================================
Func _Format_Punctuation($s) ; asdf asfsdf. asdfdf ( fgsfdgsd ) sdgsdg :fghdgfh
	$s = StringRegExpReplace($s, '\s*(\.|:|;|,|\?|!|\))', '$1 ')
	$s = StringRegExpReplace($s, '(\()\s*', ' $1')
	Return StringStripWS($s, 4)
EndFunc   ;==>_Format_Punctuation
;=============================================================================
Func _Format($s)
	Local $sRet = $s
	Switch GUICtrlRead($hComboStyle)

		Case "Lowercase"
			$sRet = StringLower($s)
		Case "Proper"
			$sRet = _StringProper($s)
		Case "Sentence"
			$sRet = _Format2(1, $s)
		Case "Title"
			$sRet = _Format2(2, $s)
		Case "Title 2"
			$sRet = _Format2(3, $s)
		Case "Uppercase"
			$sRet = StringUpper($s)
		Case "Reverse"
			$sRet = _StringReverse($s)
	EndSwitch

	Return $sRet
EndFunc   ;==>_Format
;=============================================================================
Func _Comment($s, $iMode = 0)
	Local $sLineC = StringStripWS(GUICtrlRead($hComboLineComment), 3)
	Local $sLineBl = GUICtrlRead($hComboBlockComment)
	Local $sRet = $s
	If $iMode = 1 Then
		$s = StringReplace($s, @CRLF, @CRLF & $sLineC & " ")
		$sRet = $sLineC & $s
	ElseIf $iMode = 2 Then
		Local $aTMP = StringSplit(StringStripWS($sLineBl, 7), " ", 2)
		If Not @error Then $sRet = $aTMP[0] & @CRLF & $s & @CRLF & $aTMP[1]
	EndIf

	Return $sRet
EndFunc   ;==>_Comment
;=============================================================================
Func _Convert($s)
	Local $sRet = $s

	Switch GUICtrlRead($hComboConvert)
		Case "URL encode"
			$sRet = _urlencode($s)
		Case "URL decode"
			$sRet = _urldecode($s)
		Case "IP To Name"
			TCPStartup()
			$sRet = _TCPIpToName(StringStripWS($s, 3))
			TCPShutdown()
		Case "Name To IP"
			TCPStartup()
			$sRet = TCPNameToIP(StringStripWS($s, 3))
			TCPShutdown()
		Case "String2Hex"
			$sRet = _StringToHex($s)
		Case "Hex2String"
			$sRet = StringMid(_HexToString($s), 1)
		Case "RGB2Hex"
			$sRet = _RGB2Hex($s)
		Case "Hex2RGB"
			$sRet = _Hex2RGB($s)
		Case "Int2Bin"
			Switch GUICtrlRead($hComboconvertInt2Bin)
				Case "Int2Bin (Auto)"
					$sRet = _Integer2Binary($s)
				Case "Int2Bin 8"
					$sRet = _Integer2Binary($s, 8)
				Case "Int2Bin 16"
					$sRet = _Integer2Binary($s, 16)
				Case "Int2Bin 32"
					$sRet = _Integer2Binary($s, 32)
				Case "Int2Bin 64"
					$sRet = _Integer2Binary($s, 64)
			EndSwitch
		Case "Bin2Int"
			$sRet = _Binary2Integer($s)
	EndSwitch

	Return $sRet
EndFunc   ;==>_Convert
;=============================================================================
Func _RGB2Hex($s)
	Local $sRet = ""
	$s = StringStripWS($s, 8)
	Local $aTMP = StringSplit($s, ",")
	If Not @error Then
		For $i = 1 To $aTMP[0]
			If Int($aTMP[$i]) > 255 Then Return $s
			$sRet &= Hex($aTMP[$i], 2)
		Next
		Return $sRet
	EndIf
	Return $s
EndFunc   ;==>_RGB2Hex
;=============================================================================
Func _Hex2RGB($s)
	Local $sRet = ""
	$s = StringStripWS($s, 8)
	If StringIsXDigit($s) And StringLen($s) < 7 Then
		$sRet &= Dec(StringMid($s, 1, 2)) & ","
		$sRet &= Dec(StringMid($s, 3, 2)) & ","
		$sRet &= Dec(StringMid($s, 5, 2))
		Return $sRet
	EndIf
	Return $s
EndFunc   ;==>_Hex2RGB
;11111111111111111111100001111111111
;=============================================================================
Func _Binary2Integer($in) ;coded by UEZ
	Local $int, $x, $i = 1, $aTMP = StringSplit(_StringReverse($in), "")
	For $x = 1 To UBound($aTMP) - 1
		$int += $aTMP[$x] * $i
		$i *= 2
	Next
	$aTMP = 0
	Return StringFormat('%.0f', $int)
EndFunc   ;==>_Binary2Integer
;=============================================================================
Func _Integer2Binary($in, $iLen = 0) ;coded by UEZ / Thorsten Willert
	If $in = 0 Then Return 0
	Local $bin
	While $in > 0
		$bin &= Mod($in, 2)
		$in = Floor($in / 2)
	WEnd
	$bin = _StringReverse($bin)
	If $iLen = 0 Then
		Return $bin
	Else
		Return StringFormat("%0" & $iLen & "s", $bin)
	EndIf
EndFunc   ;==>_Integer2Binary
;=============================================================================
Func _RegReplace($s)
	Local $h = GUICtrlGetHandle($hReplaceView)
	;Local $iEnd = _GUICtrlListView_GetItemCount($h)
	Local $sRet = $s
	Local $ERR, $EXT

	For $i = 0 To 999
		If $aRegEx[$i][0] = "" Then ExitLoop
		If $aRegEx[$i][0] Then
			$sRet = StringRegExpReplace($sRet, $aRegEx[$i][1], $aRegEx[$i][2])
			$ERR = @error
			$EXT = @extended
			If Not $ERR Then
				_GUICtrlListView_SetItemText($h, $i, $EXT, 4)
				_GUICtrlListView_SetItemText($h, $i, "", 5)
			Else
				_GUICtrlListView_SetItemText($h, $i, "ERR " & $EXT, 5)
				ConsoleWrite(@error & " | " & @extended & @CRLF)
			EndIf
		EndIf
	Next
	Return $sRet
EndFunc   ;==>_RegReplace
;==============================================================================
Func _RegEx_List2Array()
	Local $h = GUICtrlGetHandle($hReplaceView)
	Local $iEnd = _GUICtrlListView_GetItemCount($h)

	For $i = 0 To $iEnd
		$aRegEx[$i][0] = _B2S(_GUICtrlListView_GetItemChecked($h, $i))
		$aRegEx[$i][1] = _GUICtrlListView_GetItemText($h, $i, 1)
		$aRegEx[$i][2] = _GUICtrlListView_GetItemText($h, $i, 2)
	Next
	;_ArrayDisplay($aRegEx)
EndFunc   ;==>_RegEx_List2Array
;==============================================================================
Func _AddRegEx()
	_GUICtrlListView_AddItem($hReplaceView, "")
EndFunc   ;==>_AddRegEx
;==============================================================================
Func _InsertRegEx()
	_GUICtrlListView_InsertItem($hReplaceView, "", _GUICtrlListView_GetSelectedIndices($hReplaceView, False))
EndFunc   ;==>_InsertRegEx
;==============================================================================
Func _UpdateRegEx()
	Local $iSelected = _GUICtrlListView_GetSelectedIndices($hReplaceView, False)
	Local $h = GUICtrlGetHandle($hReplaceView)

	_GUICtrlListView_SetItemText($h, $iSelected, GUICtrlRead($hRegEx_Search), 1)
	_GUICtrlListView_SetItemText($h, $iSelected, GUICtrlRead($hRegEx_Replace), 2)
	_GUICtrlListView_SetItemText($h, $iSelected, GUICtrlRead($hRegEx_Comment), 3)
EndFunc   ;==>_UpdateRegEx
;==============================================================================
Func _EditRegEx()
	Local $iSelected = _GUICtrlListView_GetSelectedIndices($hReplaceView, False)
	Local $h = GUICtrlGetHandle($hReplaceView)

	Local $sTMP = _GUICtrlListView_GetItemText($h, $iSelected, 1)
	If $sTMP <> "" Then
		GUICtrlSetData($hRegEx_Search, $sTMP)
		GUICtrlSetData($hRegEx_Replace, _GUICtrlListView_GetItemText($h, $iSelected, 2))
		GUICtrlSetData($hRegEx_Comment, _GUICtrlListView_GetItemText($h, $iSelected, 3))
	EndIf
EndFunc   ;==>_EditRegEx
;==============================================================================
Func _DeleteRegEx()
	_GUICtrlListView_DeleteItemsSelected(GUICtrlGetHandle($hReplaceView))
EndFunc   ;==>_DeleteRegEx
;==============================================================================
Func _Format2($iMode, $s)
	Local $aTMP
	Local $r = ""

	$aTMP = StringSplit($s, " ")
	If Not @error Then
		Switch $iMode
			Case 1 ; Sentence
				$r = StringUpper(StringLeft($aTMP[1], 1)) & StringMid(StringLower($aTMP[1]), 2) & " "
				For $i = 2 To $aTMP[0]
					$r &= StringLower($aTMP[$i]) & " "
				Next
			Case 2 ; Title
				For $i = 1 To $aTMP[0]
					$r &= StringUpper(StringLeft($aTMP[$i], 1)) & StringMid(StringLower($aTMP[$i]), 2) & " "
				Next
			Case 3 ; Title 2
				For $i = 1 To $aTMP[0]
					If Not StringIsUpper($aTMP[$i]) Then
						$r &= StringUpper(StringLeft($aTMP[$i], 1)) & StringMid(StringLower($aTMP[$i]), 2) & " "
					Else
						$r &= $aTMP[$i] & " "
					EndIf
				Next
		EndSwitch
	Else
		$r = StringUpper(StringLeft($aTMP[1], 1)) & StringMid(StringLower($aTMP[1]), 2)
	EndIf

	Return SetError(0,0,$r)
EndFunc   ;==>_Format2
;==============================================================================
Func _FormatReplace($s)
	Local $sEpoch = @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC

	$s = StringReplace($s, "YYYY", @YEAR)
	$s = StringReplace($s, "YY", StringRight(@year,2) )
	$s = StringReplace($s, "MM", @MON)
	$s = StringReplace($s, "DD", @MDAY)
	$s = StringReplace($s, "HH", @HOUR)
	$s = StringReplace($s, "MI", @MIN)
	$s = StringReplace($s, "SS", @SEC)
	$s = StringReplace($s, "EPOCH", _Epoch_encrypt($sEpoch))

	Return $s
EndFunc   ;==>_FormatReplace

;==============================================================================

Func _MacroList()
	Local $aType[5]
	$aType[0] = "All Users"
	$aType[1] = "Current User"
	$aType[2] = "Other Directories"
	$aType[3] = "System Info"
	$aType[4] = "Current Time and Date"

	$aMacro[0][0] = $aType[0]
	$aMacro[0][1] = "@AppDataCommonDir"
	$aMacro[0][2] = "Path to Application Data"

	$aMacro[1][0] = $aType[0]
	$aMacro[1][1] = "@DesktopCommonDir"
	$aMacro[1][2] = "Path to Desktop"

	$aMacro[2][0] = $aType[0]
	$aMacro[2][1] = "@DocumentsCommonDir"
	$aMacro[2][2] = "Path to Documents"

	$aMacro[3][0] = $aType[0]
	$aMacro[3][1] = "@FavoritesCommonDir"
	$aMacro[3][2] = "All Users: Path to Favorites"

	$aMacro[4][0] = $aType[0]
	$aMacro[4][1] = "@ProgramsCommonDir"
	$aMacro[4][2] = "All Users: Path to Favorites"

	$aMacro[5][0] = $aType[0]
	$aMacro[5][1] = "@StartMenuCommonDir"
	$aMacro[5][2] = "All Users: Path to Start Menu folder"

	$aMacro[6][0] = $aType[0]
	$aMacro[6][1] = "@StartupCommonDir"
	$aMacro[6][2] = "All Users: Path to Startup folder"
	;========================================================

	$aMacro[7][0] = $aType[1]
	$aMacro[7][1] = "@AppDataDir"
	$aMacro[7][2] = "Path to current user's Application Data"

	$aMacro[8][0] = $aType[1]
	$aMacro[8][1] = "@DesktopDir"
	$aMacro[8][2] = "Path to current user's  Desktop"

	$aMacro[9][0] = $aType[1]
	$aMacro[9][1] = "@MyDocumentsDir"
	$aMacro[9][2] = "Path to My Documents target"

	$aMacro[10][0] = $aType[1]
	$aMacro[10][1] = "@FavoritesDir"
	$aMacro[10][2] = "Path to current user's Favorites"

	$aMacro[11][0] = $aType[1]
	$aMacro[11][1] = "@ProgramsDir"
	$aMacro[11][2] = "Path to current user's Programs"

	$aMacro[12][0] = $aType[1]
	$aMacro[12][1] = "@StartMenuDir"
	$aMacro[12][2] = "Startup folder"

	$aMacro[13][0] = $aType[1]
	$aMacro[13][1] = "@UserProfileDir"
	$aMacro[13][2] = "Path to current user's Profile folder"

	;========================================================

	$aMacro[14][0] = $aType[2]
	$aMacro[14][1] = "@HomeDrive"
	$aMacro[14][2] = "Drive letter of drive containing current user's home directory"

	$aMacro[15][0] = $aType[2]
	$aMacro[15][1] = "@HomePath"
	$aMacro[15][2] = "Directory part of current user's home directory"

	$aMacro[16][0] = $aType[2]
	$aMacro[16][1] = "@HomeShare"
	$aMacro[16][2] = "Server and share name containing current user's home directory"

	$aMacro[17][0] = $aType[2]
	$aMacro[17][1] = "@LogonDNSDomain"
	$aMacro[17][2] = "Logon DNS Domain"

	$aMacro[18][0] = $aType[2]
	$aMacro[18][1] = "@LogonDomain"
	$aMacro[18][2] = "Logon Domain"

	$aMacro[19][0] = $aType[2]
	$aMacro[19][1] = "@LogonServer"
	$aMacro[19][2] = "Logon server"

	$aMacro[20][0] = $aType[2]
	$aMacro[20][1] = "@ProgramFilesDir"
	$aMacro[20][2] = "Path to Program Files folder"

	$aMacro[21][0] = $aType[2]
	$aMacro[21][1] = "@CommonFilesDir"
	$aMacro[21][2] = "Path to Common Files folder"

	$aMacro[22][0] = $aType[2]
	$aMacro[22][1] = "@WindowsDir"
	$aMacro[22][2] = "Path to Windows folder"

	$aMacro[23][0] = $aType[2]
	$aMacro[23][1] = "@SystemDir"
	$aMacro[23][2] = "Path to Windows' System (or System32) folder"

	$aMacro[24][0] = $aType[2]
	$aMacro[24][1] = "@TempDir"
	$aMacro[24][2] = "Path to the temporary files folder"

	;========================================================

	$aMacro[25][0] = $aType[3]
	$aMacro[25][1] = "@CPUArch"
	$aMacro[25][2] = "Returns 'X86' when the CPU is a 32-bit CPU and 'X64' when the CPU is 64-bit"

	$aMacro[26][0] = $aType[3]
	$aMacro[26][1] = "@KBLayout"
	$aMacro[26][2] = "Returns code denoting Keyboard Layout"

	$aMacro[27][0] = $aType[3]
	$aMacro[27][1] = "@MUILang"
	$aMacro[27][2] = "Returns code denoting Multi Language if available (Vista is OK by default)"

	$aMacro[28][0] = $aType[3]
	$aMacro[28][1] = "@OSArch"
	$aMacro[28][2] = "Returns one of the following: 'X86', 'IA64', 'X64' - this is the architecture type of the currently running operating system"

	$aMacro[29][0] = $aType[3]
	$aMacro[29][1] = "@OSLang"
	$aMacro[29][2] = "Returns code denoting OS Language"

	$aMacro[30][0] = $aType[3]
	$aMacro[30][1] = "@OSType"
	$aMacro[30][2] = "Returns 'WIN32_NT' for 2000/XP/2003/Vista/2008/Win7/2008R2"

	$aMacro[31][0] = $aType[3]
	$aMacro[31][1] = "@OSVersion"
	$aMacro[31][2] = "Returns one of the following: 'WIN_2008R2', 'WIN_7', 'WIN_8', 'WIN_2008', 'WIN_VISTA', 'WIN_2003', 'WIN_XP', 'WIN_XPe', 'WIN_2000'"

	$aMacro[32][0] = $aType[3]
	$aMacro[32][1] = "@OSBuild"
	$aMacro[32][2] = "Returns the OS build number"

	$aMacro[33][0] = $aType[3]
	$aMacro[33][1] = "@OSServicePack"
	$aMacro[33][2] = "Service pack info in the form of 'Service Pack 3'"

	$aMacro[34][0] = $aType[3]
	$aMacro[34][1] = "@ComputerName"
	$aMacro[34][2] = "Computer's network name"

	$aMacro[35][0] = $aType[3]
	$aMacro[35][1] = "@UserName"
	$aMacro[35][2] = "ID of the currently logged on user"

	$aMacro[36][0] = $aType[3]
	$aMacro[36][1] = "@IPAddress1"
	$aMacro[36][2] = "IP address of first network adapter"

	$aMacro[37][0] = $aType[3]
	$aMacro[37][1] = "@IPAddress2"
	$aMacro[37][2] = "IP address of second network adapter"

	$aMacro[38][0] = $aType[3]
	$aMacro[38][1] = "@IPAddress3"
	$aMacro[38][2] = "IP address of third network adapter"

	$aMacro[39][0] = $aType[3]
	$aMacro[39][1] = "@IPAddress4"
	$aMacro[39][2] = "IP address of fourth network adapter"

	$aMacro[40][0] = $aType[3]
	$aMacro[40][1] = "@IPAddressPublic"
	$aMacro[40][2] = "Public IP address of a network / computer"

	$aMacro[41][0] = $aType[3]
	$aMacro[41][1] = "@DesktopHeight"
	$aMacro[41][2] = "Height of the desktop screen in pixels"

	$aMacro[42][0] = $aType[3]
	$aMacro[42][1] = "@DesktopWidth"
	$aMacro[42][2] = "Width of the desktop screen in pixels"

	$aMacro[43][0] = $aType[3]
	$aMacro[43][1] = "@DesktopDepth"
	$aMacro[43][2] = "Depth of the desktop screen in bits per pixel"

	$aMacro[44][0] = $aType[3]
	$aMacro[44][1] = "@DesktopRefresh"
	$aMacro[44][2] = "Refresh rate of the desktop screen in Hz"

	;========================================================

	$aMacro[45][0] = $aType[4]
	$aMacro[45][1] = "@MSEC"
	$aMacro[45][2] = "Millisecons"

	$aMacro[46][0] = $aType[4]
	$aMacro[46][1] = "@SEC"
	$aMacro[46][2] = "Current Time and Date: Seconds"

	$aMacro[47][0] = $aType[4]
	$aMacro[47][1] = "@MIN"
	$aMacro[47][2] = "Minutes"

	$aMacro[48][0] = $aType[4]
	$aMacro[48][1] = "@HOUR"
	$aMacro[48][2] = "Hour (24-hour format)"

	$aMacro[49][0] = $aType[4]
	$aMacro[49][1] = "@MDAY"
	$aMacro[49][2] = "Day of month"

	$aMacro[50][0] = $aType[4]
	$aMacro[50][1] = "@MON"
	$aMacro[50][2] = "Month"

	$aMacro[51][0] = $aType[4]
	$aMacro[51][1] = "@YEAR"
	$aMacro[51][2] = "Year (four-digit)"

	$aMacro[52][0] = $aType[4]
	$aMacro[52][1] = "@WDAY"
	$aMacro[52][2] = "Day of week"

	$aMacro[53][0] = $aType[4]
	$aMacro[53][1] = "@WNR_ISO"
	$aMacro[53][2] = "Weeknumber ISO"

	$aMacro[54][0] = $aType[4]
	$aMacro[54][1] = "@YDAY"
	$aMacro[54][2] = "Day of year"

	$aMacro[55][0] = $aType[4]
	$aMacro[55][1] = "@EPOCH"
	$aMacro[55][2] = "Unix time, or POSIX time"

EndFunc   ;==>_MacroList
;==============================================================================
Func _FuncList_Read()
	Local $h = GUICtrlGetHandle($hFuncList)

	For $i = 0 To UBound($aFunc_List) - 1
		$aFunc_List[$i] = _GUICtrlListBox_GetText($h, $i)
	Next

EndFunc   ;==>_FuncList_Read
;==============================================================================
Func _Func_List()

	$aFunc_List[0] = "Format"
	$aFunc_List[1] = "Cut First"
	$aFunc_List[2] = "Cut Last"
	$aFunc_List[3] = "Cut Mid"
	$aFunc_List[4] = "Add Preffix"
	$aFunc_List[5] = "Add Suffix"
	$aFunc_List[6] = "Add Insert"
	$aFunc_List[7] = "Replace Remove Non ASCII"
	$aFunc_List[8] = "Replace RegExReplace"
	$aFunc_List[9] = "Replace Trim WS"
	$aFunc_List[10] = "Replace Strip double WS"
	$aFunc_List[11] = "Replace CR"
	$aFunc_List[12] = "Convert"
	$aFunc_List[13] = "Comment Line"
	$aFunc_List[14] = "Comment Block"

EndFunc   ;==>_Func_List
;==============================================================================
Func _Load_FuncList()
	_Func_List()
	Local $h = GUICtrlGetHandle($hFuncList)
	For $i = 0 To UBound($aFunc_List) - 1
		_GUICtrlListBox_AddString($h, $aFunc_List[$i])
	Next
EndFunc   ;==>_Load_FuncList
;==============================================================================
Func _ListView_AutoColSize($h)
	Local $iCol = _GUICtrlListView_GetColumnCount($h)
	For $i = 0 To $iCol
		_GUICtrlListView_SetColumnWidth($h, $i, $LVSCW_AUTOSIZE)
	Next
EndFunc   ;==>_ListView_AutoColSize
;==============================================================================
Func _IsChecked($h)
	Return BitAND(GUICtrlRead($h), $GUI_Checked) = $GUI_Checked
EndFunc   ;==>_IsChecked

;==============================================================================
Func _WriteIni()

	_Save_RegEx()
	If _IsChecked($hMenuHistorySave) Then _Save_History()
	;_Save_FuncList()

	_ClipBoard_ChangeChain($ClipFormat, $hNext)

EndFunc   ;==>_WriteIni
;==============================================================================
Func _SelectRegex($b)
	Local $h = GUICtrlGetHandle($hReplaceView)
	Local $iEnd = _GUICtrlListView_GetItemCount($h)

	For $i = 0 To $iEnd
		_GUICtrlListView_SetItemChecked($h, $i, $b)
	Next
EndFunc   ;==>_SelectRegex
;==============================================================================
Func _Save_FuncList()
	GUICtrlSetData($hStatus1, "SavingRegEx...")
	Local $aData[15][2]

	_FuncList_Read()

	For $i = 0 To 14
		$aData[$i][0] = $i
		$aData[$i][1] = $aFunc_List[$i]
	Next

	IniWriteSection($sIni, "FuncList", $aData, 0)

EndFunc   ;==>_Save_FuncList
;==============================================================================
Func _Save_RegEx()
	GUICtrlSetData($hStatus1, "SavingRegEx...")
	Local $sItem = ""

	Local $h = GUICtrlGetHandle($hReplaceView)
	Local $iEnd = _GUICtrlListView_GetItemCount($h)
	If $iEnd = 0 Then
		IniDelete($sIni, "RegEx")
		Return
	EndIf

	Local $aData[1][2]
	ReDim $aData[$iEnd][2]

	For $i = 0 To $iEnd - 1
		$aData[$i][0] = $i
		$aData[$i][1] = _B2S(_GUICtrlListView_GetItemChecked($h, $i)) & "|"
		$aData[$i][1] &= StringToBinary(_GUICtrlListView_GetItemText($h, $i, 1)) & "|"
		$aData[$i][1] &= StringToBinary(_GUICtrlListView_GetItemText($h, $i, 2)) & "|"
		$aData[$i][1] &= _GUICtrlListView_GetItemText($h, $i, 3)
	Next

	IniWriteSection($sIni, "RegEx", $aData, 0)

EndFunc   ;==>_Save_RegEx
;==============================================================================
Func _Load_RegEx()
	Local $aTMP
	Local $aLIST
	Local $h = GUICtrlGetHandle($hReplaceView)
	Local $iNew

	$aTMP = IniReadSection($sIni, "RegEx")
	If Not @error Then
		For $i = 1 To $aTMP[0][0]
			$aLIST = StringSplit($aTMP[$i][1], "|")
			If Not @error Then
				$iNew = _GUICtrlListView_InsertItem($h, "")
				_GUICtrlListView_SetItemChecked($h, $iNew, $aLIST[1])
				_GUICtrlListView_SetItemText($h, $iNew, BinaryToString($aLIST[2]), 1)
				_GUICtrlListView_SetItemText($h, $iNew, BinaryToString($aLIST[3]), 2)
				_GUICtrlListView_SetItemText($h, $iNew, $aLIST[4], 3)
			EndIf
		Next
	EndIf

EndFunc   ;==>_Load_RegEx
;==============================================================================
Func _Save_History()
	GUICtrlSetData($hStatus1, "SavingHistory...")

	Local $h = GUICtrlGetHandle($hLast)
	Local $iEnd = _GUICtrlListBox_GetCount($h)
	If $iEnd = 0 Then
		IniDelete($sIni, "History")
		Return
	EndIf

	Local $aData[1][2]
	ReDim $aData[$iEnd][2]

	For $i = 0 To $iEnd - 1
		$aData[$i][0] = $i
		$aData[$i][1] = _GUICtrlListBox_GetText($h, $i)
	Next

	IniWriteSection($sIni, "History", $aData, 0)

EndFunc   ;==>_Save_History
;==============================================================================
Func _Load_History()
	Local $aTMP, $sTMP
	Local $aLIST
	Local $h = GUICtrlGetHandle($hLast)
	Local $iNew

	$aTMP = IniReadSection($sIni, "History")
	If Not @error Then
		For $i = 1 To $aTMP[0][0]
			If $sTMP = $aTMP[$i][1] Then ContinueLoop
			If StringLen($aTMP[$i][1]) < 5 Then ContinueLoop

			$sTMP = $aTMP[$i][1]
			_GUICtrlListBox_AddString($h, $aTMP[$i][1])
		Next
	EndIf

EndFunc   ;==>_Load_History
;==============================================================================
Func _B2S($b)
	If $b = "TRUE" Then Return 1
	Return 0
EndFunc   ;==>_B2S
;==============================================================================
; by trancexx, http://www.autoitscript.com/forum/topic/83667-epoch-time/#entry598706
Func _Epoch_encrypt($date)
	Local $main_split = StringSplit($date, " ")
	If $main_split[0] - 2 Then
		Return SetError(1, 0, "") ; invalid time format
	EndIf
	Local $asDatePart = StringSplit($main_split[1], "/")
	Local $asTimePart = StringSplit($main_split[2], ":")
	If $asDatePart[0] - 3 Or $asTimePart[0] - 3 Then
		Return SetError(1, 0, "") ; invalid time format
	EndIf
	If $asDatePart[2] < 3 Then
		$asDatePart[2] += 12
		$asDatePart[1] -= 1
	EndIf
	Local $i_aFactor = Int($asDatePart[1] / 100)
	Local $i_bFactor = Int($i_aFactor / 4)
	Local $i_cFactor = 2 - $i_aFactor + $i_bFactor
	Local $i_eFactor = Int(1461 * ($asDatePart[1] + 4716) / 4)
	Local $i_fFactor = Int(153 * ($asDatePart[2] + 1) / 5)
	Local $aDaysDiff = $i_cFactor + $asDatePart[3] + $i_eFactor + $i_fFactor - 2442112
	Local $iTimeDiff = $asTimePart[1] * 3600 + $asTimePart[2] * 60 + $asTimePart[3]
	Return $aDaysDiff * 86400 + $iTimeDiff
EndFunc   ;==>_Epoch_encrypt
;==============================================================================
Func _urlencode($string)
	$string = StringSplit($string, "")
	For $i = 1 To $string[0]
		If AscW($string[$i]) < 48 Or AscW($string[$i]) > 122 Then
			$string[$i] = "%" & _StringToHex($string[$i])
		EndIf
	Next
	$string = _ArrayToString($string, "", 1)
	Return $string
EndFunc   ;==>_urlencode
;==============================================================================
Func _urldecode($string)
	$string = StringSplit($string, "%")
	For $i = 1 To $string[0]
		If StringIsXDigit(StringLeft($string[$i], 2)) Then ; Otherwise you get stray 0x in the string.
			$string[$i] = _HexToString(StringLeft($string[$i], 2)) & StringTrimLeft($string[$i], 2)
		EndIf
	Next
	$string = _ArrayToString($string, "", 1)
	Return $string
EndFunc   ;==>_urldecode
;==============================================================================
; Handle $WM_CHANGECBCHAIN messages
Func WM_CHANGECBCHAIN($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg

	; If the next window is closing, repair the chain
	If $iwParam = $hNext Then
		$hNext = $ilParam
		; Otherwise pass the message to the next viewer
	ElseIf $hNext <> 0 Then
		_SendMessage($hNext, $WM_CHANGECBCHAIN, $iwParam, $ilParam, 0, "hwnd", "hwnd")
	EndIf
EndFunc   ;==>WM_CHANGECBCHAIN
;==============================================================================
; Handle $WM_DRAWCLIPBOARD messages
Func WM_DRAWCLIPBOARD($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg

	If _IsChecked($hAuto) Then $sCLIPBOARD = _ClipBoard_GetData($CF_UNICODETEXT)

	; Pass the message to the next viewer
	If $hNext <> 0 Then _SendMessage($hNext, $WM_DRAWCLIPBOARD, $iwParam, $ilParam)
EndFunc   ;==>WM_DRAWCLIPBOARD
