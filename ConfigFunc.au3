Func _OnExit($Zapcxv)
	$NamFunc = "_OnExit"
	ProcessWaitClose($Zapcxv)
	$extcc = @Extended
	$TimeFinish = _NowCalcDate() & "-" & _NowTime(5) & ":" & @MSEC
	If $extcc = 0 Then
		Return "good:" & $TimeFinish
	Else
		Switch $extcc
			Case 702
				Return "wcls:" & $TimeFinish
			Case 802
				Return "bcnl:" & $TimeFinish
		EndSwitch
	EndIf
	$NamFunc = ""
EndFunc





Func _RunAU3($sFilePath, $sCommandLine = "", $sWorkingDir = "", $iShowFlag = @SW_SHOW, $iOptFlag = 0)
	$chass = _Now()
	$dfgcx = _FO_PathSplit($sFilePath)
	If StringInStr($dfgcx[1], "/") Then
		$key = StringMid($dfgcx[1], StringInStr($dfgcx[1], "/"), 3)
		$sFilePath = StringReplace($sFilePath, $key, "")
		$sCommandLine = $key & $sCommandLine
	EndIf
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdNGW & "\Win\" & $dfgcx[1], "Start", "REG_SZ", $chass)
	Local $iPID
	$iPID = Run('"' & @AutoItExe & '" /AutoIt3ExecuteScript "' & $sFilePath & '" ' & $sCommandLine, $sWorkingDir, $iShowFlag, $iOptFlag)
	If @Error Then
		Return SetError(@Error, 1, 0)
	EndIf
	Return $iPID
EndFunc ;==>_RunAU3




Func _NewSeesUja()   
	$NamFunc = "_NewSeesUja"
	If GUICtrlRead($CheSkrit) = 1 Then
		RegWrite("HKLM\SOFTWARE\TKSNS", "WinStat", "REG_SZ", 1)
	ElseIf GUICtrlRead($CheSkrit) = 4 Then
		RegWrite("HKLM\SOFTWARE\TKSNS", "WinStat", "REG_SZ", 0)
	EndIf
	If _ProvConnKey() Then Return 0  ;выводить в главную форму причины выхода- а именна -- отсутствия ключа
	; _ZapVibVhidMater()
	; _PerevSqlStan()
	; _PerPidklTaksRbd()
	; _PerPidklTksnsToServ()
	_ProvFileExist()
	; _PidlFileInServ()
	; _VikSaveInSql()
	$NamFunc = ""
EndFunc




Func _ExitKnop()
	$NamFunc = "_ExitKnop"
	_LogReg($UnIdNVF, $numWin, $NamFunc, "Коритсувач натиснув кнопку 'Відмінити'")
	Exit($ExCodeButtCanc)
	$NamFunc = ""
EndFunc





Func _UprScr($mosd)
	$NemFunc = "_UprScr"
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "", "REG_SZ", _Now())
	$PutToFisikFile = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\InFilesPutch", "PutchRBD")
	If @Error Then
		MsgBox(16 + 4096, $ZagTextPomWIS, "Не вдалося отримати шлях до реляційної бази даних")
		_LogReg($UnIdWIS, $numWin, $NemFunc, "Отримання шляху до реляційної бази даних=ERROR")
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
		Exit(7784)
	EndIf
		_LogReg($UnIdWIS, $numWin, $NemFunc, "Отримання шляху до реляційної бази даних=OK")
	$TlkImaTaRozdh = _FO_PathSplit($PutToFisikFile)
	$LogikNamRbd = StringMid($TlkImaTaRozdh[1], 1, StringInStr($TlkImaTaRozdh[1], "_") - 1)
	$NamTKSNS = "TKSNS"
	$FileTksnsPut = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWIS & "\TKSNS.mdf"
	_ProcPovVerSql() ; перевірка версії SQL-сервера
	_PerZnachInIput() ; перевірка значень в елементах управління
	$Spisbaz = _SpisPodklBD()
	Switch $mosd
		Case 1
			If _PerePidkl($Spisbaz, $NamTKSNS, 2) Then
				$NewTksnsfilePut = $Spisbaz[_ArraySearch($Spisbaz, $NamTKSNS)][1]
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_TKSNS_File", "REG_SZ", $NewTksnsfilePut)
				_Videdna($NamTKSNS, $NewTksnsfilePut)
				_DettachDb($NamTKSNS)
				_AttachBd($NamTKSNS, $FileTksnsPut)
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_TKSNS_File", "REG_SZ", $FileTksnsPut)
				GUICtrlSetData($InpConTksnsServ,"Перев_з'єдн_із_РБД_=0")
			Else
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_TKSNS_File", "REG_SZ", 0)
				_AttachBd($NamTKSNS, $FileTksnsPut)
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_TKSNS_File", "REG_SZ", $FileTksnsPut)
				GUICtrlSetData($InpConTksnsServ,"Перев_з'єдн_із_РБД_=0")
			EndIf
		Case 2
			If _PerePidkl($Spisbaz, $PutToFisikFile, 1) Then
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_Rbd_File", "REG_SZ", $PutToFisikFile)
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_Rbd_File", "REG_SZ", $PutToFisikFile)
				Return 2
			Else
				If _PerePidkl($Spisbaz, $LogikNamRbd, 2) Then
					$NewRbdfilePut = $Spisbaz[_ArraySearch($Spisbaz, $LogikNamRbd)][1]
					RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_Rbd_File", "REG_SZ", $NewRbdfilePut)
					_Videdna($LogikNamRbd, $NewRbdfilePut)
					_DettachDb($LogikNamRbd)
					_AttachBd($LogikNamRbd, $PutToFisikFile)
					RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_Rbd_File", "REG_SZ", $PutToFisikFile) 
				Else
					RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_Rbd_File", "REG_SZ", 0)
					_AttachBd($LogikNamRbd, $PutToFisikFile)
					RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_Rbd_File", "REG_SZ", $PutToFisikFile)
				EndIf
			EndIf
	EndSwitch
	$NemFunc = ""
EndFunc





Func _UprScr($mosd)
	$NemFunc = "_UprScr"
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "", "REG_SZ", _Now())
	$PutToFisikFile = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\InFilesPutch", "PutchRBD")
	If @Error Then
		MsgBox(16 + 4096, $ZagTextPomWIS, "Не вдалося отримати шлях до реляційної бази даних")
		_LogReg($UnIdWIS, $numWin, $NemFunc, "Отримання шляху до реляційної бази даних=ERROR")
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
		Exit(7784)
	EndIf
		_LogReg($UnIdWIS, $numWin, $NemFunc, "Отримання шляху до реляційної бази даних=OK")
	$TlkImaTaRozdh = _FO_PathSplit($PutToFisikFile)
	$LogikNamRbd = StringMid($TlkImaTaRozdh[1], 1, StringInStr($TlkImaTaRozdh[1], "_") - 1)
	$NamTKSNS = "TKSNS"
	$FileTksnsPut = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWIS & "\TKSNS.mdf"
	_ProcPovVerSql() ; перевірка версії SQL-сервера
	_PerZnachInIput() ; перевірка значень в елементах управління
	$Spisbaz = _SpisPodklBD()
	Switch $mosd
		Case 1
			If _PerePidkl($Spisbaz, $NamTKSNS, 2) Then
				$NewTksnsfilePut = $Spisbaz[_ArraySearch($Spisbaz, $NamTKSNS)][1]
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_TKSNS_File", "REG_SZ", $NewTksnsfilePut)
				_Videdna($NamTKSNS, $NewTksnsfilePut)
				_DettachDb($NamTKSNS)
				_AttachBd($NamTKSNS, $FileTksnsPut)
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_TKSNS_File", "REG_SZ", $FileTksnsPut)
				GUICtrlSetData($InpConTksnsServ,"Перев_з'єдн_із_РБД_=0")
			Else
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_TKSNS_File", "REG_SZ", 0)
				_AttachBd($NamTKSNS, $FileTksnsPut)
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_TKSNS_File", "REG_SZ", $FileTksnsPut)
				GUICtrlSetData($InpConTksnsServ,"Перев_з'єдн_із_РБД_=0")
			EndIf
		Case 2
			If _PerePidkl($Spisbaz, $PutToFisikFile, 1) Then
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_Rbd_File", "REG_SZ", $PutToFisikFile)
				RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_Rbd_File", "REG_SZ", $PutToFisikFile)
				Return 2
			Else
				If _PerePidkl($Spisbaz, $LogikNamRbd, 2) Then
					$NewRbdfilePut = $Spisbaz[_ArraySearch($Spisbaz, $LogikNamRbd)][1]
					RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_Rbd_File", "REG_SZ", $NewRbdfilePut)
					_Videdna($LogikNamRbd, $NewRbdfilePut)
					_DettachDb($LogikNamRbd)
					_AttachBd($LogikNamRbd, $PutToFisikFile)
					RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_Rbd_File", "REG_SZ", $PutToFisikFile) 
				Else
					RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "Old_Rbd_File", "REG_SZ", 0)
					_AttachBd($LogikNamRbd, $PutToFisikFile)
					RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\WorkInServer", "New_Rbd_File", "REG_SZ", $PutToFisikFile)
				EndIf
			EndIf
	EndSwitch
	$NemFunc = ""
EndFunc






Func _VtavInListSpBd($GuiHwnd, $aMasPiklRbd)
	$NemFunc = "_VtavInListSpBd"
	If Not IsArray($aMasPiklRbd) Then
		MsgBox(16 + 4096, $ZagTextPomWIS, "Список підключених до серверу баз даних не являється масивом")
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "Список підключ до серв баз даних - масив=ERROR") 
		Exit(7780)
	EndIf
	
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "Список підключ до серв баз даних - масив=OК") 
	If UBound($aMasPiklRbd) = 1 Then
		_GUICtrlFFLabel_Create($GuiWIS, "Немає підключених реляційних баз", 210, 65, 230, 60, 15, "Courier New", 1, 1, 0x008080)
	Else
		GUICtrlSetState($SpiiiBd, $GUI_SHOW)
		GUICtrlSetState($MetkTit, $GUI_SHOW)
		_ArrayDelete($aMasPiklRbd, 0)
		For $IaMasPiklRbd = 0 To UBound($aMasPiklRbd) - 1
			$NloooNam = StringStripWS($aMasPiklRbd[$IaMasPiklRbd][0], 3)
			$Phesik = StringStripWS($aMasPiklRbd[$IaMasPiklRbd][1], 3)
; 			GUICtrlCreateListViewItem($NloooNam & "|" & $Phesik, $SpiiiBd)
		Next
	EndIf
	$NemFunc = ""
EndFunc






Func _PerZnachInIput()
	$NemFunc = "_PerZnachInIput"
	If (GUICtrlRead($InpVerSqlProc) < 2005) Or (GUICtrlRead($InpVerSqlverss) < 2005) Then
		MsgBox(16 + 4096, $ZagTextPomWIS, "На комп'ютері " & @ComputerName & "встановлена версія SQL Server менша за 2005, для успішної роботи програми необхідно оновиться до MSSQL SERVER 2008 R2!!!")
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "Версія SQL-серверу 2008 і старше=ERROR")
		Exit(7776)
	Else 
		If GUICtrlRead($InpVerSqlProc) <> GUICtrlRead($InpVerSqlverss) Then
			MsgBox(16 + 4096, $ZagTextPomWIS, "На комп'ютері " & @ComputerName & " не можу визначити версію SQL Server " _
					& @CRLF & "для успішної роботи програми необхідний робочий екземпляр MSSQL SERVER 2008 R2, або пізнішої версії...")
			_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "Версія SQL-серверу 2008 і старше=ERROR")
			Exit(7778)
		Else
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "Версія SQL-серверу 2008 і старше=OК")
			RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS, "VersSqlServ", "REG_SZ", GUICtrlRead($InpVerSqlProc))
			GUICtrlSetData($InpZagRifh, "Версія SQL серверу=" & GUICtrlRead($InpVerSqlProc))
		EndIf
	EndIf
	$NemFunc = ""
EndFunc

 
 

Func _GIFImage($hHandle, $sImage, $iWidth, $iHeight, $iLeft = 5, $iTop = 55)
    GUISetState(@SW_SHOW, $hHandle)
    Local $oWMIService, $oWMIService_Ctrl 
    Local $aResult = DllCall("user32.dll", "handle", "GetDC", "hwnd", $hHandle)
    If @error Then Return SetError(@error, @extended, 0)
    Local $hDC = $aResult[0]
    $aResult = DllCall("gdi32.dll", "int", "GetBkColor", "hwnd", $hDC)
    If @error Then Return SetError(@error, @extended, 0)
    Local $iGetBkColor = $aResult[0]
    $aResult = DllCall("user32.dll", "int", "ReleaseDC", "hwnd", $hHandle, "handle", $hDC)
    If @error Then Return SetError(@error, @extended, 0)
    Local $iSystemColor = Hex(BitOR(BitAND($iGetBkColor, 0x00FF00), BitShift(BitAND($iGetBkColor, 0x0000FF), -16), BitShift(BitAND($iGetBkColor, 0xFF0000), 16)), 6)
    Local $iIP = StringRegExp($sImage, "(ftp|http|https):\/\/(\w+:{0,1}\w*@)?(\S+)(:[0-9]+)?(\/|\/([\w#!:.?+=&%@!\-\/]))?")
    Local $sSyntax = "file:///"
    If $iIP Then $sSyntax = ""

    $oWMIService = ObjCreate("Shell.Explorer.2")
    If Not IsObj($oWMIService) Then Return SetError(1, 1, 0)
    $oWMIService_Ctrl = GUICtrlCreateObj($oWMIService, $iLeft, $iTop, $iWidth, $iHeight)
    GUICtrlSetBkColor($oWMIService_Ctrl, 0x000000)
    $oWMIService.navigate("about:blank")
    While $oWMIService.busy
        Sleep(100)
    WEnd

    With $oWMIService.document
        .write('<body onselectstart="return false" oncontextmenu="return false" onclick="return false" ondragstart="return false" ondragover="return false">') ; Thanks to Melba23 >> http://www.autoitscript.com/forum/index.php?act=findpost&pid=735769
        .body.innerHTML = '<img id="_GIFImage" src="' & $sSyntax & $sImage & '">'
        .body.title = "Приклад копіювання параметрів"
        .body.topmargin = 0
        .body.leftmargin = 0
        .body.scroll = "no"
        .body.bgcolor = $iSystemColor
        .body.style.backgroundColor = $iSystemColor
        .body.style.borderWidth = 0
    EndWith
    GUICtrlSetState($oWMIService_Ctrl, 64)
    Return 1
EndFunc   ;==>_GIFImage






Func _ZapComCmd($TextCoomCmd)
	$NamFunc = "_ZapComCmd"
	$Vidpo = ""
	$sReadz = ""
	$iPIDz = Run(@ComSpec & $TextCoomCmd, '', @SW_HIDE, $STDOUT_CHILD)
	If Not $iPIDz Then
		MsgBox(16, 'Error ' & $NamFunc, 'Error')
		Return 0
	EndIf
	While 1
		$sReadz &= StdoutRead($iPIDz)
		If @error Then ExitLoop
		Sleep(15)
	WEnd
	Return $sReadz
	$NamFunc = ""
EndFunc





Func _VstPdrozVstatbar()
	$NamFunc = "_VstPdrozVstatbar"
	$ZapCodeLgLs = _SqlZap("select * from _ObjTaks", " Не виходить отимати коди ЛГ та лісництва", "TKSNS", 2)
	If Not IsArray($ZapCodeLgLs) Then
		_ObrErr($UnIdZZp, $numWin, $NamFunc, @AutoItPID)
		Exit(9999)
	Else
		_ArrayDelete($ZapCodeLgLs, 0)
		_GUICtrlStatusBar_SetText($hStatus, $ZapCodeLgLs[0][0], 0)
		_GUICtrlStatusBar_SetText($hStatus, $ZapCodeLgLs[0][1], 1)
	EndIf
	$NamFunc = ""
EndFunc




Func _SpisZberProc()
	$NamFunc = "_SpisZberProc"
	
	$sSpiZberProc = "StvorKartTables_1" _
			& @CRLF & "TablKartSloy_2" _
			& @CRLF & "NewKart_5" _
			& @CRLF & "FormComGed_6" _
			& @CRLF & "VnutrDublPar_7" _
			& @CRLF & "SpisInZnLgLs_8" _
			& @CRLF & "VilDublSklRozObj_9" _
			& @CRLF & "SozIndx_10" _
			& @CRLF & "PomParObjKart_11"
	
	$aArrFromRow = StringSplit($sSpiZberProc, @CR)
	
; 	ConsoleWrite('result=' & _SqlZap("EXEC StvorKartTables 'd:\TKSNS\Resurce\TipSlo.txt','d:\TKSNS\Resurce\RekKZH.txt','d:\TKSNS\Resurce\RekKZM.txt'", "Не вдалося створити допоміжні таблиці ", "TKSNS", 1) & @CRLF)
	; _ArrayDisplay($aArrFromRow,"Массив")
	
	$NamFunc = ""
EndFunc






Func _StvorElemUpr($ImmmaSlooj, $aMassZnacPar, $ZnachPlooo, $modeRob)
	$NamFunc = ""
	Switch $ImmmaSlooj
		Case "MejLisn"
			Switch $modeRob
				Case 1
				Case 2
					If $aMassZnacPar[0] = 1 Then
						$LabNoRizZnaLg = GUICtrlCreateLabel("Єдине значення коду лісгоспу - " & $aMassZnacPar[1], 5, $higInptPPE + 132, 400, 18)
						GUICtrlSetBkColor($LabNoRizZnaLg, 0xFFF7CE)
						GUICtrlSetFont($LabNoRizZnaLg, 12, 500, 0, 'Lucida Console')
						GUICtrlSetColor($LabNoRizZnaLg, 0xC80066)
					Else
; 						_GUICtrlComboBox_BeginUpdate($ZnaLg)
						For $iUlg = 1 To $aMassZnacPar[0]
; 							_GUICtrlComboBox_AddString($ZnaLg, $aMassZnacPar[$iUlg])
						Next
; 						_GUICtrlComboBox_EndUpdate($ZnaLg)
; 						GUICtrlSetState($ZnaLg,$GUI_SHOW)
					EndIf
						Return 0
				Case 3
					If $aMassZnacPar[0] = 1 Then
						$LabNoRizZnaLs = GUICtrlCreateLabel("Єдине значення кодів ліcництва- " & $aMassZnacPar[1], 5, $higInptPPE + 152, 380, 18)
						GUICtrlSetBkColor($LabNoRizZnaLs, 0xFFF7CE)
						GUICtrlSetFont($LabNoRizZnaLs, 12, 500, 0, 'Lucida Console')
						GUICtrlSetColor($LabNoRizZnaLs, 0xC80066)
					Else
; 						_GUICtrlComboBox_BeginUpdate($ZnaLis)
						For $iUls = 1 To $aMassZnacPar[0]
; 							_GUICtrlComboBox_AddString($ZnaLis, $aMassZnacPar[$iUls])
						Next
; 						_GUICtrlComboBox_EndUpdate($ZnaLis)
; 						GUICtrlSetState($ZnaLis,$GUI_SHOW)
					EndIf
						Return 0
			EndSwitch
		Case "KvPros"
			
	EndSwitch
	$NamFunc = ""
EndFunc	

	$provv = "IF EXISTS (SELECT * FROM sys.all_objects WHERE name = N'_UvVid' AND type = 'U') DROP TABLE  _UvVid;"
	$asw = _SqlZap($provv, "Не виходить перевірити наявність представлення неув'язних виділів", 'TKSNS', 5)
	If $asw <> 0 Then
		Return SetError(7)
	EndIf
	Switch $mode
		Case 0 ; якщо обрано тільки культури
			
			$spis5 = " WITH pl (cdss,poss) " & _
					" AS " & _
					" (SELECT cds, plo FROM _NewObKart ) " & _
					" SELECT d.*" & _
					" INTO _UvVid" & _
					" FROM   (SELECT i.cdss,i.poss,i.poss роposs ,os.Kalg, os.KAIG, os.KAKL, os.KAWN, os.KAVN, os.KARN, " & _
					" os.KAKZ, os.KAZU,os.KAPL, os.KAB_ , os.KATL, os.KATU, os.KAGV, os.KBJV, os.KA1P," & _
					" os.KA1N,os.KAT1, os.skl, os.KAA_, os.KAH_, os.KAD_, os.KAPH, os.KAPP, os.KAMG," & _
					" os.KNSW, os.KNKM, os.KPHT, os.KPSO, os.KPPT" & _
					" FROM pl i" & _
					" JOIN _OneString os " & _
					" ON i.cdss=os.codegis " & _
					" WHERE i.poss>" & GUICtrlRead($nUpdown) * 10000 & ") d " & _
					" WHERE d.cdss NOT IN  " & _
					" (SELECT codegis  " & _
					"  FROM Kult)"
			ConsoleWrite('result=' & $spis5 & @CRLF)
			$predsneuv = _SqlZap($spis5, "Не вдалося створити представлення неув'язних виділів", 'TKSNS', 5)
			If $predsneuv <> 0 Then
				Return SetError(8)
			Else
				Return "Неув'язні виділа вибрано успішно"
			EndIf
			
		Case 1 ; якщо обрано не тільки культури, а і кат зем
			$spis5 = " WITH pl (cdss,poss)" & _
					" AS " & _
					" (SELECT cds, plo FROM _NewObKart ) " & _
					" SELECT d.*" & _
					" INTO _UvVid " & _
					" FROM   (SELECT i.cdss,i.poss,i.poss роposs,os.Kalg, os.KAIG, os.KAKL, os.KAWN, os.KAVN, os.KARN, " & _
					" os.KAKZ, os.KAZU,os.KAPL, os.KAB_ , os.KATL, os.KATU, os.KAGV, os.KBJV, os.KA1P," & _
					" os.KA1N,os.KAT1, os.skl, os.KAA_, os.KAH_, os.KAD_, os.KAPH, os.KAPP, os.KAMG," & _
					" os.KNSW, os.KNKM, os.KPHT, os.KPSO, os.KPPT" & _
					" FROM pl i" & _
					" JOIN _OneString os " & _
					" ON i.cdss=os.codegis " & _
					" WHERE i.poss>" & GUICtrlRead($nUpdown) * 10000 & ") d " & _
					" WHERE d.cdss NOT IN  " & _
					" (SELECT codegis  " & _
					"  FROM Kult klt " & _
					" UNION  " & _
					" SELECT os.codegis " & _
					" FROM _OneString os " & _
					" WHERE os.KAKZ IN  ("
			$massVibKatZem = _GUICtrlListBox_GetSelItems($KatZem)
			_ArrayDelete($massVibKatZem, 0)
			For $imassVibKatZem = 0 To UBound($massVibKatZem) - 1
				If $imassVibKatZem = UBound($massVibKatZem) - 1 Then
					$spis5 &= $arrkatzem[$imassVibKatZem][0] & ")) "
				Else
					$spis5 &= $arrkatzem[$imassVibKatZem][0] & ","
				EndIf
			Next
			$predsneuv = _SqlZap($spis5, "Не вдалося створити представлення неув'язних виділів", 'TKSNS', 5)
			If $predsneuv <> 0 Then
				Return SetError(9)
			Else
				Return "Неув'язні виділа вибрано успішно"
			EndIf
		Case 2 ; якщо обрано: культури, кат зем та кат ліс
			$spis5 = " WITH pl (cdss,poss)" & _
					" AS " & _
					" (SELECT cds, plo FROM _NewObKart ) " & _
					" SELECT d.*" & _
					" INTO _UvVid " & _
					" FROM   (SELECT i.cdss,i.poss,i.poss роposs,os.Kalg, os.KAIG, os.KAKL, os.KAWN, os.KAVN, os.KARN, " & _
					" os.KAKZ, os.KAZU,os.KAPL, os.KAB_ , os.KATL, os.KATU, os.KAGV, os.KBJV, os.KA1P," & _
					" os.KA1N,os.KAT1, os.skl, os.KAA_, os.KAH_, os.KAD_, os.KAPH, os.KAPP, os.KAMG," & _
					" os.KNSW, os.KNKM, os.KPHT, os.KPSO, os.KPPT" & _
					" FROM pl i" & _
					" JOIN _OneString os " & _
					" ON i.cdss=os.codegis " & _
					" WHERE i.poss>" & GUICtrlRead($nUpdown) * 10000 & ") d " & _
					" WHERE d.cdss NOT IN  " & _
					" (SELECT codegis  " & _
					"  FROM Kult klt " & _
					" UNION  " & _
					" SELECT os.codegis " & _
					" FROM _OneString os " & _
					" WHERE os.KAKZ IN  ("
			$massVibKatZem = _GUICtrlListBox_GetSelItems($KatZem)
			_ArrayDelete($massVibKatZem, 0)
			For $imassVibKatZem = 0 To UBound($massVibKatZem) - 1
				If $imassVibKatZem = UBound($massVibKatZem) - 1 Then
					$spis5 &= $arrkatzem[$imassVibKatZem][0] & ") " & _
							" UNION " & _
							" SELECT kt.codegis " & _
							" FROM _OneString kt " & _
							" WHERE kt.kakl IN ("
				Else
					$spis5 &= $arrkatzem[$imassVibKatZem][0] & ","
				EndIf
			Next
			$massVibKatZah = _GUICtrlListBox_GetSelItems($KatZah)
			_ArrayDelete($massVibKatZah, 0)
			For $imassVibKatZah = 0 To UBound($massVibKatZah) - 1
				If $imassVibKatZah = UBound($massVibKatZah) - 1 Then
					$spis5 &= $arrkatzah[$imassVibKatZah][0] & ")) "
				Else
					$spis5 &= $arrkatzah[$imassVibKatZah][0] & ","
				EndIf
			Next
			$predsneuv = _SqlZap($spis5, "Не вдалося створити представлення неув'язних виділів", 'TKSNS', 5)
			If $predsneuv <> 0 Then
				Return SetError(9)
			Else
				Return "Неув'язні виділа вибрано успішно"
			EndIf
		Case 3 ; якщо обрано: культури, кат зем та кат ліс
			$spis5 = " WITH pl (cdss,poss)" & _
					" AS " & _
					" (SELECT cds, plo FROM _NewObKart ) " & _
					" SELECT d.*" & _
					" INTO _UvVid " & _
					" FROM   (SELECT i.cdss,i.poss,i.poss роposs,os.Kalg, os.KAIG, os.KAKL, os.KAWN, os.KAVN, os.KARN, " & _
					" os.KAKZ, os.KAZU,os.KAPL, os.KAB_ , os.KATL, os.KATU, os.KAGV, os.KBJV, os.KA1P," & _
					" os.KA1N,os.KAT1, os.skl, os.KAA_, os.KAH_, os.KAD_, os.KAPH, os.KAPP, os.KAMG," & _
					" os.KNSW, os.KNKM, os.KPHT, os.KPSO, os.KPPT" & _
					" FROM pl i" & _
					" JOIN _OneString os " & _
					" ON i.cdss=os.codegis " & _
					" WHERE i.poss>" & GUICtrlRead($nUpdown) * 10000 & ") d " & _
					" WHERE d.cdss NOT IN  " & _
					" (SELECT codegis  " & _
					"  FROM Kult klt " & _
					" UNION  " & _
					" SELECT kt.codegis " & _
					" FROM _OneString kt " & _
					" WHERE kt.kakl IN ("
			$massVibKatZah = _GUICtrlListBox_GetSelItems($KatZah)
			_ArrayDelete($massVibKatZah, 0)
			For $imassVibKatZah = 0 To UBound($massVibKatZah) - 1
				If $imassVibKatZah = UBound($massVibKatZah) - 1 Then
					$spis5 &= $arrkatzah[$imassVibKatZah][0] & ")) "
				Else
					$spis5 &= $arrkatzah[$imassVibKatZah][0] & ","
				EndIf
			Next
			$predsneuv = _SqlZap($spis5, "Не вдалося створити представлення неув'язних виділів", 'TKSNS', 5)
			If $predsneuv <> 0 Then
				Return SetError(10)
			Else
				Return "Неув'язні виділа вибрано успішно"
			EndIf
	EndSwitch
EndFunc






Func _CreateIdSess() ; от 8 до 13 цифр
	$Id = @YEAR & @YDAY & @HOUR & @MIN & @SEC
	$UserPutchProj = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $Id
	If FileExists($UserPutchProj) Then
		DirRemove($UserPutchProj)
		DirCreate($UserPutchProj)
	EndIf
	$ReeVetkaNew="HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $Id
	RegDelete($ReeVetkaNew, "Error")
	RegDelete($ReeVetkaNew, "ErrorNamFunc")
	RegDelete($ReeVetkaNew, "ErrorWin")
	RegDelete($ReeVetkaNew, "ErrorFlag")
	RegDelete($ReeVetkaNew, "SqlService")
	RegDelete($ReeVetkaNew, "VerSqlServ")
; 	If RegRead($reeVetkSess2wq) Then
;
; 		RegWrite($ReeVetkaNew,"",)
;
	Return $Id
EndFunc




 Func _ProvInetConn($mode=0)
 
  $INTERNET_CONNECTION_MODEM          = 0x1
 $INTERNET_CONNECTION_LAN            = 0x2
 $INTERNET_CONNECTION_PROXY          = 0x4
 $INTERNET_CONNECTION_MODEM_BUSY     = 0x8
 $INTERNET_RAS_INSTALLED             = 0x10
 $INTERNET_CONNECTION_OFFLINE        = 0x20
 $INTERNET_CONNECTION_CONFIGURED     = 0x40
 
     $ret = DllCall("WinInet.dll","int","InternetGetConnectedState","int_ptr",0,"int",0)
 
     If $ret[0] then
        ;check type of connection
         $sX = "Connected !" & @LF & "------------------" & @LF
         If BitAND($ret[1], $INTERNET_CONNECTION_MODEM)      Then $sX = $sX & "MODEM" & @LF
         If BitAND($ret[1], $INTERNET_CONNECTION_LAN)        Then $sX = $sX & "LAN" & @LF
         If BitAND($ret[1], $INTERNET_CONNECTION_PROXY)      Then $sX = $sX & "PROXY" & @LF
         If BitAND($ret[1], $INTERNET_CONNECTION_MODEM_BUSY) Then $sX = $sX & "MODEM_BUSY" & @LF
         If BitAND($ret[1], $INTERNET_RAS_INSTALLED)         Then $sX = $sX & "RAS_INSTALLED" & @LF
         If BitAND($ret[1], $INTERNET_CONNECTION_OFFLINE)    Then $sX = $sX & "OFFLINE" & @LF
         If BitAND($ret[1], $INTERNET_CONNECTION_CONFIGURED) Then $sX = $sX & "CONFIGURED" & @LF
     Else
         $sX = "Not Connected"
     Endif
  
		Switch $mode
			Case  0
				Return $ret[0] 
			Case  1
				Return $sX   
		EndSwitch   
 EndFunc
 
 
 
 
 
Func _PerCodVih()
	$NamFunc = "_PerCodVih"
	$Exitcode = @ExitCode
	$ExitMethd = @ExitMethod
	Local $ErrSpos
	If $ExitMethd Then
		If $Exitcode = 702 Then
			$ErrSpos = " вікно закрите через системну кнопку "
; 			MsgBox(0, '', ":"&"Код виходу= "&$Exitcode&" ,тобто,"&$ErrSpos&":")
			RegWrite("HKLM\SOFTWARE\TKSNS", "ErrMethod", "REG_SZ", $ExitMethd)
			RegWrite("HKLM\SOFTWARE\TKSNS", "ErrCode", "REG_SZ", $Exitcode)
			RegWrite("HKLM\SOFTWARE\TKSNS", "ErrSpos", "REG_SZ", $ErrSpos)
		ElseIf $Exitcode = 802 Then
			$ErrSpos = "вікно закрите через кнопку 'Відмінити'"
; 			MsgBox(0, '', ":"&"Код виходу= "&$Exitcode&" ,тобто, "&$ErrSpos&":")
			RegWrite("HKLM\SOFTWARE\TKSNS", "ErrMethod", "REG_SZ", $ExitMethd)
			RegWrite("HKLM\SOFTWARE\TKSNS", "ErrCode", "REG_SZ", $Exitcode)
			RegWrite("HKLM\SOFTWARE\TKSNS", "ErrSpos", "REG_SZ", $ErrSpos)
		EndIf
	EndIf
	$NamFunc = ""
EndFunc






Func _ViknProgess($GuiParr, $NamModulee, $TextNij = " ", $pidzag = " ", $mode = 1, $pocent = 0, $xposs = -1, $ypos = -1, $Viglad = 16) ;,
	$NamFunc = "_ViknProgess"
	Switch $mode
		Case 1
			WinSetState($GuiParr, "", @SW_HIDE)
			ProgressOn("ТаксІнс© - " & $NamModulee, $TextNij, $pidzag, $xposs, $ypos, $Viglad) ;, $pidzag, $mode,$pocent, $xposs, $ypos,$Viglad
			$hPrBar = GUICtrlGetHandle(-1)
			Return $hPrBar
		Case 2
			WinActivate($NamModulee, "")
			ProgressSet($pocent, $TextNij, $pidzag)
			$hPrBar = GUICtrlGetHandle(-1)
			Return $hPrBar
		Case 3
			ProgressOff()
			$hPrBar = GUICtrlGetHandle(-1)
			WinSetState($GuiParr, "", @SW_SHOW)
			WinActivate($GuiParr, "")
			Return $hPrBar
	EndSwitch
	$NamFunc = ""
EndFunc





Func _StringRemoveLine($hFile, $sDelete)
	If FileExists($hFile) Then $hFile = FileRead($hFile) ;Remove If FileExists($hFile) Then << only
	Local $nSNS = StringInStr($hFile, @CRLF & $sDelete) - 1
	Local $sSFH = StringLeft($hFile, $nSNS)
	Local $sRL = StringTrimLeft($hFile, StringLen($sSFH) + 2)
	Local $sLLEN = StringLen(StringLeft($sRL, StringInStr($sRL, @CRLF)))
	If Not $sLLEN Then $sLLEN = StringLen($sRL)
	Return $sSFH & StringTrimLeft($hFile, $sLLEN + $nSNS + 2)
EndFunc ;==>_StringRemoveLine()





Func __ListViewFill($hListView, $iColumns, $iRows) ; Required only for the Example.
	If Not IsHWnd($hListView) Then
		$hListView = GUICtrlGetHandle($hListView)
	EndIf
	Local $fIsCheckboxesStyle = (BitAND(_GUICtrlListView_GetExtendedListViewStyle($hListView), $LVS_EX_CHECKBOXES) = $LVS_EX_CHECKBOXES)
	
	_GUICtrlListView_BeginUpdate($hListView)
	For $i = 0 To $iColumns - 1
		_GUICtrlListView_InsertColumn($hListView, $i, 'Column ' & $i + 1, 50)
		_GUICtrlListView_SetColumnWidth($hListView, $i - 1, -2)
	Next
	For $i = 0 To $iRows - 1
		_GUICtrlListView_AddItem($hListView, 'Row ' & $i + 1 & ': Col 1', $i)
		If Random(0, 1, 1) And $fIsCheckboxesStyle Then
			_GUICtrlListView_SetItemChecked($hListView, $i)
		EndIf
		For $j = 1 To $iColumns
			_GUICtrlListView_AddSubItem($hListView, $i, 'Row ' & $i + 1 & ': Col ' & $j + 1, $j)
		Next
	Next
	_GUICtrlListView_EndUpdate($hListView)
EndFunc ;==>__ListViewFill

Func _CreateListView($hGUI, ByRef $iListView) ; Thanks to AZJIO for this function.
	Local $aClientSize = WinGetClientSize($hGUI)
	$iListView = GUICtrlCreateListView('', 0, 0, $aClientSize[0], $aClientSize[1] - 30)
	GUICtrlSetResizing($iListView, $GUI_DOCKBORDERS)
	Sleep(250)
	
	Local $iColumns = Random(1, 5, 1)
	__ListViewFill($iListView, $iColumns, Random(25, 100, 1)) ; Fill the ListView with Random data.
	For $i = 0 To $iColumns
		_GUICtrlListView_SetColumnWidth($iListView, $i, $LVSCW_AUTOSIZE)
		_GUICtrlListView_SetColumnWidth($iListView, $i, $LVSCW_AUTOSIZE_USEHEADER)
	Next
EndFunc ;==>_CreateListView




Func _GUICtrlListView_CreateArray($hListView, $sDelimeter = '|')
	Local $iColumnCount = _GUICtrlListView_GetColumnCount($hListView), $iDim = 0, $iItemCount = _GUICtrlListView_GetItemCount($hListView)
	If $iColumnCount < 3 Then
		$iDim = 3 - $iColumnCount
	EndIf
	If $sDelimeter = Default Then
		$sDelimeter = '|'
	EndIf
	
	Local $aColumns = 0, $aReturn[$iItemCount + 1][$iColumnCount + $iDim] = [[$iItemCount, $iColumnCount, '']]
	For $i = 0 To $iColumnCount - 1
		$aColumns = _GUICtrlListView_GetColumn($hListView, $i)
		$aReturn[0][2] &= $aColumns[5] & $sDelimeter
	Next
	$aReturn[0][2] = StringTrimRight($aReturn[0][2], StringLen($sDelimeter))
	
	For $i = 0 To $iItemCount - 1
		For $j = 0 To $iColumnCount - 1
			$aReturn[$i + 1][$j] = _GUICtrlListView_GetItemText($hListView, $i, $j)
		Next
	Next
	Return SetError(Number($aReturn[0][0] = 0), 0, $aReturn)
EndFunc ;==>_GUICtrlListView_CreateArray





Func _ObrErr($UnUnIdId, $numsdWin, $NameFunc, $AutoItPID)
; 	$TextVetki = "HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnUnIdId
	$TextVetki = "HKLM\SOFTWARE\TKSNS"
	RegWrite($TextVetki, "Error", "REG_SZ", "PidProc=" & $AutoItPID) ;@AutoItPID
	RegWrite($TextVetki, "ErrorFlag", "REG_SZ", 1)
	RegWrite($TextVetki, "ErrorWin", "REG_SZ", $numsdWin)
	RegWrite($TextVetki, "ErrorNamFunc", "REG_SZ", $NameFunc)
EndFunc





Func _ZapCmdExe($comm)
	$ipResZag = ""
	$iPid = Run(@ComSpec & " /c " & $comm, @ScriptDir, @SW_HIDE, $STDIN_CHILD + $STDERR_CHILD + $STDOUT_CHILD + $STDERR_MERGED)
	While 1
		$Line = StdoutRead($iPid)
		If $Line Then $ipResZag &= $Line & @CRLF ;_Encoding_OEM2ANSI()
		If @Error Then ExitLoop
	WEnd
	Return $ipResZag
EndFunc




Func _GolWinLog($row)
	GUICtrlSetData($LogWind, GUICtrlRead($LogWind) & $row & @CRLF)
	_GUICtrlEdit_LineScroll($LogWind, 0, _GUICtrlEdit_GetLineCount($LogWind))
EndFunc

; Func _ExeSqlite($NamBd, $TexComm, $ErrMess, $modesqlite)
; 	Local $hQuery, $aRow, $sMsg, $iColumns, $aResult, $iRow
; 	_SQLite_Startup()
; 	If @Error Then
; 		MsgBox(16, "Помилка SQLite", "Не вдалося завантажити SQLite3.dll")
; 		Exit -1
; 	EndIf
; 	_SQLite_Open($NamBd)
; 	If @Error Then
; 		MsgBox(16, "Помилка SQLite", "Не вдалося завантажити базу даних")
; 		Exit -1
; 	EndIf
; 	Switch $modesqlite
; 		Case 1 ; Запрос с возвратом первого поля
; 			_SQLite_Query(-1, $texComm, $hQuery) ; the query
; 			While _SQLite_FetchData($hQuery, $aRow) = $SQLITE_OK
; 				$sMsg &= $aRow[0]
; 			WEnd
; 			If $sMsg = '' Then
; 				MsgBox(0 + 16 + 262144, $ErrMess, _SQLite_ErrMsg())
; 				Return SetError(3)
; 			Else
; 				Return $sMsg
; 			EndIf
; 		Case 2 ; запрос с фунцией обратного вызова
; 			Local $d = _SQLite_Exec(-1, $texComm) ; _cb will be called for each row
; 			If $d = $SQLITE_OK Then
; 				Return _SQLite_Changes() & " - " & _SQLite_TotalChanges()
; 			Else
; 				MsgBox(0 + 16 + 262144, $ErrMess, _SQLite_ErrMsg())
; 				Return SetError(4)
; 			EndIf
; 		Case 3 ; возвращает оформленный двумерный массив
; 			$iRval = _SQLite_GetTable2d(-1, $texComm, $aResult, $iRow, $iColumns)
; 			If $iRval = $SQLITE_OK Then
; 				Return _SQLite_Display2DResult($aResult, 0, 1)
; 			Else
; 				MsgBox(0 + 16 + 262144, $ErrMess, _SQLite_ErrMsg())
; 				Return SetError(5)
; 			EndIf
; 		Case 4 ;  повернути кількість вставлених рядків
; 			If Not _SQLite_Exec(-1, $texComm) = $SQLITE_OK Then
; 				MsgBox(0 + 16 + 262144, $ErrMess, _SQLite_ErrMsg())
; 				Return SetError(6)
; 			Else
; 				Return _SQLite_LastInsertRowID()
; 			EndIf
; 		Case 5 ; Выводит одномерный массив, содержащий имена колонок и данные выполненного запроса
; 			$iRval = _SQLite_GetTable(-1, $texComm, $aResult, $iRow, $iColumns)
; 			If $iRval = $SQLITE_OK Then
; 				Return $aResult
; 			Else
; 				MsgBox(0 + 16 + 262144, "SQLite Error", _SQLite_ErrMsg())
; 				Return SetError(7)
; 			EndIf
; 	EndSwitch
; 	_SQLite_Close()
; 	_SQLite_Shutdown()
; EndFunc





Func _SqlZap($Zaprros, $sooberr, $nameBaze = 'master', $mode = 1)
	$Server = @ComputerName
	_SQL_RegisterErrorHandler()
	$oADODB = _SQL_Startup()
	If $oADODB = $SQL_ERROR Then
		MsgBox(0 + 16 + 262144, "Помилка запуску драйверу ADODB!", _SQL_GetErrMsg())
; 		_ZapLogInp(_SQL_GetErrMsg(), @ScriptDir & "\Log_TKSNS.log")
		Return SetError(1)
	EndIf
	$cod = _SQL_Connect(-1, $Server, $nameBaze, "", "")
	If $cod = $SQL_ERROR Then
		MsgBox(0 + 16 + 262144, "Помилка під'єднання до екземпляру серверу!", _SQL_GetErrMsg())
; 		_ZapLogInp(_SQL_GetErrMsg(), @ScriptDir & "\Log_TKSNS.log")
		Return SetError(2)
	EndIf
	Switch $mode
		Case 1 ; повернути рядки
			Local $vString
			If _Sql_GetTableAsString(-1, $Zaprros, $vString) = $SQL_OK Then
				Return $vString
			Else
				MsgBox(0 + 16 + 262144, $sooberr, _SQL_GetErrMsg())
; 				_ZapLogInp(_SQL_GetErrMsg(), @ScriptDir & "\Log_TKSNS.log")
				Return SetError(3)
			EndIf
		Case 2 ; повернути масив
			Global $aData, $iRows, $iColumns ;Variables to store the array data in to and the row count and the column count
			$iRval = _SQL_GetTable2D(-1, $Zaprros, $aData, $iRows, $iColumns)
			If $iRval = $SQL_OK Then
				Return $aData
			Else
				MsgBox(0 + 16 + 262144, $sooberr, _SQL_GetErrMsg())
; 				_ZapLogInp(_SQL_GetErrMsg(), @ScriptDir & "\Log_TKSNS.log")
				Return SetError(4)
			EndIf
		Case 3 ; повернти кількість рядків
			Local $aData, $iRows, $iColumns ;Variables to store the array data in to and the row count and the column count
			$iRval = _SQL_GetTable2D(-1, $Zaprros, $aData, $iRows, $iColumns)
			If $iRval = $SQL_OK Then
				Return $iRows
			Else
				MsgBox(0 + 16 + 262144, $sooberr, _SQL_GetErrMsg())
; 				_ZapLogInp(_SQL_GetErrMsg(), @ScriptDir & "\Log_TKSNS.log")
				Return SetError(5)
			EndIf
		Case 4 ;  поверути кількісь полів
			Global $aData, $iRows, $iColumns ;Variables to store the array data in to and the row count and the column count
			$iRval = _SQL_GetTable2D(-1, $Zaprros, $aData, $iRows, $iColumns)
			If $iRval = $SQL_OK Then
				Return $iColumns
			Else
				MsgBox(0 + 16 + 262144, $sooberr, _SQL_GetErrMsg())
; 				_ZapLogInp(_SQL_GetErrMsg(), @ScriptDir & "\Log_TKSNS.log")
				Return SetError(6)
			EndIf
		Case 5 ; виконати вираження
			$exrt = _SQL_Execute(-1, $Zaprros)
			If $exrt = $SQL_ERROR Then
				MsgBox(0 + 16 + 262144, $sooberr, _SQL_GetErrMsg())
; 				_ZapLogInp(_SQL_GetErrMsg(), @ScriptDir & "\Log_TKSNS.log")
				Return SetError(7)
			EndIf
		Case 6 ; повернути таблицю
			Global $aData, $iRows, $iColumns
			$exrt = _SQL_GetTable(-1, $Zaprros, $aData, $iRows, $iColumns)
			If $exrt = $SQL_ERROR Then
				MsgBox(0 + 16 + 262144, $sooberr, _SQL_GetErrMsg())
; 				_ZapLogInp(_SQL_GetErrMsg(), @ScriptDir & "\Log_TKSNS.log")
				Return SetError(8)
			EndIf
		Case 7 ; Считывание первой строки результата
			Local $MassOdStr[0]
			$exrt = _SQL_QuerySingleRow(-1, $Zaprros, $MassOdStr)
			If $exrt = $SQL_ERROR Then
				MsgBox(0 + 16 + 262144, $sooberr, _SQL_GetErrMsg())
; 				_ZapLogInp(_SQL_GetErrMsg(), @ScriptDir & "\Log_TKSNS.log")
				Return SetError(9)
			Else
				Return $MassOdStr
			EndIf
	EndSwitch
	_SQL_Close()
EndFunc





Func _ext()
	_Disconnect()
; 	Exit
EndFunc

Func _Disconnect()
	TCPCloseSocket($Server) ; Закриває tcp Socket
	$Server = -1 ; Змінна Server встає
	
EndFunc ; Кінець функції Від'єднання





Func _Restart()
	Local $sAutoIt_File = @TempDir & "\~Au3_ScriptRestart_TempFile.au3"
	Local $sRunLine, $sScript_Content, $hFile
	
	$sRunLine = @ScriptFullPath
	If Not @Compiled Then $sRunLine = @AutoItExe & ' /AutoIt3ExecuteScript ""' & $sRunLine & '""'
	If $CmdLine[0] > 0 Then $sRunLine &= ' ' & $CmdLineRaw
	
	$sScript_Content &= '#NoTrayIcon' & @CRLF & _
			'While ProcessExists(' & @AutoItPID & ')' & @CRLF & _
			'   Sleep(10)' & @CRLF & _
			'WEnd' & @CRLF & _
			'Run("' & $sRunLine & '")' & @CRLF & _
			'FileDelete(@ScriptFullPath)' & @CRLF
	
	$hFile = FileOpen($sAutoIt_File, 2)
	FileWrite($hFile, $sScript_Content)
	FileClose($hFile)
	
	Run(@AutoItExe & ' /AutoIt3ExecuteScript "' & $sAutoIt_File & '"', @ScriptDir, @SW_HIDE)
	Sleep(1000)
	Exit
EndFunc

 

Func _Prinati()
	Local $putcskr = @ScriptDir & "\VhdFile.ini"
	If Not FileExists(@ScriptDir & "\VhdFile.ini") Then
		_FileCreate(@ScriptDir & "\VhdFile.ini")
	EndIf
	If GUICtrlRead($Ged) = '' Or GUICtrlRead($SpisRbd) = '' Or GUICtrlRead($PutcKart) = '' Then
		MsgBox(16 + 4096, 'Помидка у вхідних даних', 'Не знайдено певного значення, програма перезапускається', 2)
		;Exit  ;_Restart()
	EndIf
	GUICtrlSetState($Ok, $gui_disable)
	$sData = 'PutchGed=' & GUICtrlRead($Ged) & @LF & 'ObranaRbd=' & GUICtrlRead($SpisRbd) & @LF & 'VibrKart=' & GUICtrlRead($PutcKart)
	$vremja = _Now()
	IniWriteSection($putcskr, $vremja, $sData)
	_GUICtrlStatusBar_SetText($SB, GUICtrlRead($SpisRbd), 0)
	_GUICtrlStatusBar_SetText($SB, GUICtrlRead($Ged), 1)
	_GUICtrlStatusBar_SetText($SB, GUICtrlRead($PutcKart), 2)
	; GUISetState(@SW_SHOW,$hSettings)
EndFunc

 
 
 









