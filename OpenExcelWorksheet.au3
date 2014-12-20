#NoTrayIcon
#include <File.au3>
#include <array.au3>

$windowtitle = "Open Excel Worksheet 1.00"

If $CmdLine[0] < 2 Then
	Msgbox (4096,$windowtitle,'Please pass at least two parameters (use quotes): "<Excel filename>" "<Worksheet to activate>" "<Optional cell to activate>"')
	Exit
EndIf

$Path = $CmdLine[1]
$WorkSheet = $CmdLine[2]

; If file does not exist quit
If NOT FileExists($Path) then
  Msgbox (4096,$windowtitle,"File not found: "& $Path)
  Exit
Endif

Dim $szDrive, $szDir, $szFName, $szExt
$PathArray = _PathSplit($Path, $szDrive, $szDir, $szFName, $szExt)

$oExcel = ObjGet($Path)  ; Get an Excel Object from an existing filename

If IsObj($oExcel) Then
		With $oExcel
			.Application.Visible = 1	
			.Windows($szFName & $szExt).Visible = 1
			.Worksheets($WorkSheet).Activate			
		EndWith
		If $CmdLine[0] = 3 Then $oExcel.Worksheets($WorkSheet).Range($CmdLine[3]).Activate
EndIf