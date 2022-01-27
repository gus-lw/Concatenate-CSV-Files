#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Augustin WALTER
 Version:		 3.0

 Script Function:
	Concatenate csv files into a single one

 Changelog:


#ce ----------------------------------------------------------------------------

#include <File.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <Array.au3>
#include <GuiListBox.au3>



Global Const $softVersion = 1
Global Const $author = "Augustin WALTER"
Global Const $softName = "Concatenate csv files"
Global Const $cpYear = 2018
Global $devMode = True

Global $intCSVFilesFOund = 0
Dim $arrayFilesInList[1], $arrayFilesPath[1] ; arrays


$guiMain = GUICreate($softName, 450, 520 )
	GUISetBkColor(0xF2F2F2)

GUICtrlCreateLabel("Directory containing .csv files", 10, 10, 150, 25)
$inputDirectory = GUICtrlCreateInput("", 40, 30, 350, 20)
$buttonChooseDir = GUICtrlCreateButton(" ", 390, 29, 38, 22)
	GUICtrlSetImage(-1, "shell32.dll", "-4", 0)
	GUICtrlSetTip(-1, "Choose the directory")

GUICtrlCreateLabel("CSV files found:", 10, 70, 150, 25)
$listItemsFound = GUICtrlCreateList("", 40, 90, 386, 200)
$buttonAdd = GUICtrlCreateButton("Add item", 227, 280, 100, 20)
$buttonRemove = GUICtrlCreateButton("Remove item", 327, 280, 100, 20)

GUICtrlCreateLabel("Destination file path:", 10, 310, 150, 25)
$inputOutputFile = GUICtrlCreateInput("", 40, 330, 350, 20)
$buttonOutputFile = GUICtrlCreateButton(" ", 390, 329, 38, 22)
	GUICtrlSetImage(-1, "shell32.dll", "-4", 0)
	GUICtrlSetTip(-1, "Selecte the final file")

GUICtrlCreateLabel("Search for string to replace into CSV files (all instance will be replaced in all files):", 10, 370, 410, 25)
GUICtrlCreateLabel("Replace ", 40, 392, 60, 20)
$inputStringToReplace = GUICtrlCreateInput("", 85, 390, 160, 20)
GUICtrlCreateLabel("by", 252, 392, 25, 20)
$inputReplacementString = GUICtrlCreateInput("", 268, 390, 160, 20)

$labelFilesNumber = GUICtrlCreateLabel("Concatenante " & $intCSVFilesFOund &" CSV files:", 10, 420, 150, 25)
$progressConcat = GUICtrlCreateProgress(40, 440, 386, 20, 0x08)
$buttonConcat = GUICtrlCreateButton("Start", 327, 465, 100, 25)
$labelProgress = GUICtrlCreateLabel("Click on " & '"Start" to begin', 50, 469, 200, 20)

GUICtrlCreateLabel('"' & $softName & '"' & " copyright © " & $cpYear & " " & $author & ". All rights reserved.", 11, 506, 430, 20, 0x01)
	GUICtrlSetFont(-1, 7)
	GUICtrlSetColor(-1, 0xE2E2E2)
GUICtrlCreateLabel('"' & $softName & '"' & " copyright © " & $cpYear & " " & $author & ". All rights reserved.", 10, 505, 430, 20, 0x01)
	GUICtrlSetFont(-1, 7)
	GUICtrlSetColor(-1, 0x808080)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT )

If $devMode = True Then

	$x = 1

Else

	$x = -2000

EndIf

$buttonViewArr = GUICtrlCreateButton("View arrays", 1, 483, 60, 20)

If $devMode = False Then GUICtrlSetState($buttonViewArr, $GUI_DISABLE)

GUISetState(@SW_SHOW)


While 1

	Switch GUIGetMsg()

		Case $buttonViewArr

			ConsoleWrite(@CRLF & GUICtrlRead($listItemsFound))
			_ArrayDisplay($arrayFilesInList)
			_ArrayDisplay($arrayFilesPath)

		Case $buttonConcat

			_Concat()

		Case $buttonAdd
			$addItem = FileOpenDialog("Select the CSV file to add", @UserProfileDir, "(*.csv)", 2, "", $guiMain)

			; Ajout au tableau contenant les chemins d'accès complets
			_ArrayAdd($arrayFilesPath, $addItem)

			; Ajout au tableau contenant les noms des fichiers
			$searchBackSlash = StringInStr($addItem, "\", 0, -1)
			$stringLength = StringLen($addItem)
			$returnStr = StringRight($addItem, $stringLength - $searchBackSlash)
			_ArrayAdd($arrayFilesInList, $returnStr)

			$readTemp = GUICtrlRead($listItemsFound)
			_GUICtrlListBox_AddString($listItemsFound, $returnStr)

		Case $buttonRemove
			$fileToRemove = GUICtrlRead($listItemsFound)

			$fileToRemoveIndex = _ArraySearch($arrayFilesInList, $fileToRemove)
			_ArrayDelete($arrayFilesInList, $fileToRemoveIndex)

			;$fileToRemoveIndex = _ArraySearch($arrayFilesPath, $fileToRemove)
			_ArrayDelete($arrayFilesPath, $fileToRemoveIndex)

			_GUICtrlListBox_DeleteString($listItemsFound, _GUICtrlListBox_GetCaretIndex($listItemsFound))


		Case $buttonOutputFile

			$stringOutputFile = FileSaveDialog("Save the final CSV file", @UserProfileDir, "(*.csv)", "", "Concatenated CSV.csv", $guiMain)

			GUICtrlSetData($inputOutputFile,$stringOutputFile)

		Case $buttonChooseDir
			$stringDirChose = FileSelectFolder("Select the directory where .csv files are located:", @UserProfileDir, 2, "", $guiMain)

			GUICtrlSetData($inputDirectory, $stringDirChose)
			If GUICtrlRead($inputOutputFile) = "" Then GUICtrlSetData($inputOutputFile, $stringDirChose & "\Concatenated CSV.csv")

			_looxForFiles($stringDirChose)

		Case -3

			Exit

	EndSwitch

WEnd

Func _looxForFiles($dir)

	$guiLookForFiles = GUICreate($softName, 300, 100, -1, -1, BitOR($WS_POPUP, $WS_THICKFRAME), 0x00000008, $guiMain)
	GUICtrlCreateLabel("Looking for CSV files inside the directory ...", 10, 10, 280, 50, BitOR(0x0200, 0x01))
	GUICtrlCreateProgress(20, 70, 260, 20, 0x08)
	WinSetTrans($guiLookForFiles, "", 0)

	GUISetState(@SW_SHOW)

	For $i = 0 To 255 Step 5

		WinSetTrans($guiLookForFiles, "", $i)
		Sleep(1)

	Next
	WinSetTrans($guiLookForFiles, "", 255)

	$arrayFilesInList = _FileListToArrayRec ($dir, "*.csv", 1, 0, 1, 0) ;get files names
	$arrayFilesPath = _FileListToArrayRec ($dir, "*.csv", 1, 0, 1, 2) ;get files paths

	For $file = 1 To $arrayFilesInList[0] Step 1

		GUICtrlSetData($listItemsFound, $arrayFilesInList[$file])

	Next

	For $i = 255 To 0 Step -5

		WinSetTrans($guiLookForFiles, "", $i)
		Sleep(1)

	Next
	GUIDelete($guiLookForFiles)

EndFunc


Func _Concat()

	ConsoleWrite(@CRLF & "STart!")

EndFunc