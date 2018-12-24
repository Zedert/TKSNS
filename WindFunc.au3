
#CS
Func _ProvFileExist()
	$NamFunc = "_ProvFileExist"
	
	Local $XPosLabl = 5
	Local $YPosLabl = 90
	Local $SHirLabl = 160
	Local $VisotLabl = 15
	
	; ConsoleWrite('$TextStrtyr=' &$a_GFFL_GDILbls[$TextFileMod][$s_GFFL_Restore] & @CRLF)
	$SpsFil =  "1_NewVibFiles.au3" _; 			"2_WorkInServer/pv.au3" _ & @CRLF & 
			& @CRLF & "2_WorkInServer/tb.au3" _
			& @CRLF & "2_WorkInServer/tp.au3" _
			& @CRLF & "3_PervObrobFiles.au3" _
			& @CRLF & "2_WorkInServer/ab.au3" _
			& @CRLF & "4_WorkWithDigit.au3" _
			& @CRLF & "5_GifAnimInGui.au3" _
			& @CRLF & "6_PostSelPar.au3" _
			& @CRLF & "7_ZberKartSql.au3"
	$aSpisFil = StringSplit($SpsFil, @LF, 1)
	For $iwe = $NomPoch To UBound($aSpisFil) - 1
		$TextFileMod = _GUICtrlFFLabel_Create($GuiNGW, "", $XPosLabl, $YPosLabl, $SHirLabl, $VisotLabl)
		$VstaTextFileMod = _GUICtrlFFLabel_SetData($TextFileMod, $aSpisFil[$iwe], 0xE0FED8)
		$TextStart = _GUICtrlFFLabel_Create($GuiNGW, "", $XPosLabl + $SHirLabl + 5, $YPosLabl, $SHirLabl, $VisotLabl)
		$VstaTextStart = _GUICtrlFFLabel_SetData($TextStart, _NowCalcDate() & "-" & _NowTime(5) & ":" & @MSEC, 0xE0FED8)
		If RegRead("HKLM\SOFTWARE\TKSNS", "Error") Then Return 1
		$NazvVikonFile = StringStripWS($a_GFFL_GDILbls[$TextFileMod][$s_GFFL_Restore], 3)
		$Zap = _RunAU3($NazvVikonFile, $UnIdNGW)
		$resVik = _OnExit($Zap)
		$endTiii = StringMid($resVik, 6, -1)
		$TextEnd = _GUICtrlFFLabel_Create($GuiNGW, "", ($XPosLabl + $SHirLabl * 2) + 10, $YPosLabl, $SHirLabl, $VisotLabl)
		$VstaText = _GUICtrlFFLabel_SetData($TextEnd, $endTiii, 0xE0FED8)
		$ResEnd = StringMid($resVik, 1, 4)
		$TextStatVik = _GUICtrlFFLabel_Create($GuiNGW, "", ($XPosLabl + $SHirLabl * 3) + 15, $YPosLabl, $SHirLabl, $VisotLabl)
		$VstaText = _GUICtrlFFLabel_SetData($TextStatVik, $ResEnd, 0xE0FED8)
		_ZchitLog()
		If $ResEnd <> "good" Then ExitLoop
		$YPosLabl += 20
		
; 	MsgBox(64+4096, 'Info', $aSpisFil[$iwe])
	Next
	
	; _ArrayDisplay($aSpisFil,"Массив")
	
	$NamFunc = ""
EndFun
#CE





Func _ZapStrPap()
	$NamFunc = "_ZapStrPap"
	$ProjFold = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdNVF
	$ProgrDataFold = @ProgramFilesDir & "\TKSNS"
	$TemplMap = $ProgrDataFold & "\UvPlo.dmf"
	$RbdFileMdf = $ProgrDataFold & "\TKSNS.mdf"
	$RbdFileLdf = $ProgrDataFold & "\TKSNS_log.ldf"
	$GedIniezFold = $ProjFold & "\BeckUpGedIni\Ged.ini"
	$DigFoldIni = GUICtrlRead($DigitInput2) & "\Ged.ini"
	$DigFoldLibr = GUICtrlRead($DigitInput2) & "\Library\TKSNS_SaveUvj.dsf"
	$DigFoldFavor = GUICtrlRead($DigitInput2) & "\Favorites"
	$DigFoldTempl = GUICtrlRead($DigitInput2) & "\Templates"
	$DriveSql = GUICtrlRead($DigitInput2) & "\Sql.udl"
	$StructFold = "HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdNVF
	$VssiShlah = 'ProjFold:' & $ProjFold & @CRLF & _
			'ProgrDataFold:' & $ProgrDataFold & @CRLF & _
			'TemplMap:' & $TemplMap & @CRLF & _
			'RbdFileMdf:' & $RbdFileMdf & @CRLF & _
			'RbdFileLdf:' & $RbdFileLdf & @CRLF & _
			'GedIniezFold:' & $GedIniezFold & @CRLF & _
			'DigFoldIni:' & $DigFoldIni & @CRLF & _
			'DigFoldLibr:' & $DigFoldLibr & @CRLF & _
			'DriveSql:' & $DriveSql & @CRLF & _
			'DigFoldFavor:' & $DigFoldFavor & @CRLF
	$zapShlyahInReestr = RegWrite($StructFold, "StructFold", "REG_MULTI_SZ", $VssiShlah)
	If $zapShlyahInReestr = 0 Then
		MsgBox(16 + 4096, $ZagTextPomNVF, "Не вдалося записати в реєстр параметр зі значенням інших шляхів")
		_ObrErr($UnIdNVF, $numWin, $NamFunc, @AutoItPID)
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Створення стуктури каталогів=ERROR")
		Return 1
	EndIf
	_LogReg($UnIdNVF, $numWin, $NamFunc, "Створення стуктури каталогів=ОК")
	$NamFunc = ""
	Return 0
EndFunc





Func _VstZnacReestr()
	$NamFunc = "_VstZnacReestr"
	$VEtkaReestr = "HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdNVF & "\InFilesPutch"
	$ZapShlykart = RegWrite($VEtkaReestr, "PutchKart", "REG_SZ", GUICtrlRead($KartInput1))
	$ZapShlyBase = RegWrite($VEtkaReestr, "PutchRBD", "REG_SZ", GUICtrlRead($RdbInput3))
	$ZapShlyDigit = RegWrite($VEtkaReestr, "PutchDigit", "REG_SZ", GUICtrlRead($DigitInput2))
	If $ZapShlykart = 0 Or $ZapShlyBase = 0 Or $ZapShlyDigit = 0 Then
		MsgBox(16 + 4096, $ZagTextPomNVF, "Не вдалося записати в реєстр шляхи до вхідних файлів")
		_ObrErr($UnIdNVF, $numWin, $NamFunc, @AutoItPID)
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Створення записів шялхів =ERROR")
		Return 1
	EndIf
	_LogReg($UnIdNVF, $numWin, $NamFunc, "Створення записів шялхів =ОК")
	$NamFunc = ""
	Return 0
EndFunc





Func _PerKatDig($ewqas)
	$NamFunc = "_PerKatDig"
	If Not (FileExists($ewqas & "\ged.exe")) Or Not (FileExists($ewqas & "\ged.ini")) Or Not (FileExists($ewqas & "\sql.udl")) Then
		MsgBox(16 + 4096, $ZagTextPomNVF, "У вказаному каталозі Digitals не вистачає всіх необхідних файлів")
		_ObrErr($UnIdNVF, $numWin, $NamFunc, @AutoItPID)
		Return 1
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Перевірка каталогу Digitals=ERROR")
	Else
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Перевірка каталогу Digitals=ОК")
		$NamFunc = ""
		Return 0
	EndIf
EndFunc







Func _PerRbdLdf($PutchFileMdf)
	$NamFunc = "_PerRbdLdf"
	$NewNazvFileLdf=""
	$ArrSplitPutchMdf = _FO_PathSplit($PutchFileMdf)
 	$NewwNumFilMdfBezRoz = StringRegExpReplace($ArrSplitPutchMdf[1], '(?i)(.*)(_)(data)', '\1')  
  	$PutFileMdf = StringMid($ArrSplitPutchMdf[0], 1, StringInStr($ArrSplitPutchMdf[0], '\', 0, -1) - 1)
  	$PoiskLdfZaPutch = _FO_FileSearch($PutFileMdf, '', True, 1, 0, 2)
  	$FlagNzaj = 0
	$rtffg =""
	For $iS = 0 To UBound($PoiskLdfZaPutch) - 1 
		$rtffg = StringInStr($PoiskLdfZaPutch[$iS],$NewwNumFilMdfBezRoz)
		If $rtffg Then  
			If StringInStr($PoiskLdfZaPutch[$iS],"ldf") Then
			$NewNazvFileLdf=$ArrSplitPutchMdf[0]&$PoiskLdfZaPutch[$iS]
			RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_"&$UnIdNVF&"\InFilesPutch","PutchRBDLog","REG_SZ",$NewNazvFileLdf)
			$FlagNzaj=1 
			EndIf 
		EndIf 
	Next 
	If $FlagNzaj = 0 Then
		GUICtrlSetData($Input5, "не_знайд_файл_ldf")
		MsgBox(16 + 4096, $ZagTextPomNVF, "За вказаним шляхом до файлу РБД, не знайдено файл журнала бази даних")
		_ObrErr($UnIdNVF, $numWin, $NamFunc, @AutoItPID)
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Прерв наявн файлу журн БД=ERROR")
		Return 1
	Else
		GUICtrlSetData($Input5, "Наявн_файл_журн_РБД=ОК")
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Прерв наявн файлу журн БД=OK") 
		$NamFunc = ""
		Return 0
	EndIf
EndFunc







Func _OtkitDmf()
	$NamFunc = "_OtkitDmf"
	_PriZapModlule($UnIdNVF)  
	Local $StructFodsfsdf = ""
	Local $VibFilGe = ''
	$VibFilGe = FileSelectFolder('Виберіть каталог програми Digitals', "", 4, @DocumentsCommonDir, $GuiNVF)
	If $VibFilGe = "" Or (Not StringInStr(FileGetAttrib($VibFilGe), "D")) Then
		MsgBox(16 + 4096, 'Error', 'Не вибрано каталог Digitals')
		_ObrErr($UnIdNVF, $numWin, $NamFunc, @AutoItPID)
		
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Вказання папки з Digitals=ERROR")
		$PolzCloseFormGed = 1
	Else
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Вказання папки з Digitals=OK")
		$StructFodsfsdf = _PerKatDig($VibFilGe)
		If Not $StructFodsfsdf Then 
			GUICtrlSetData($DigitInput2, "")
			GUICtrlSetData($DigitInput2, $VibFilGe)
			GUICtrlSetState($ButtPapDig, $gui_disable)
			GUICtrlSetData($Input4, "Перев_файл_в_кат_digit=ОК")
			$PolzCloseFormGed = 0
			If Not _VstZnacReestr() Then
				GUICtrlSetData($Input6, "Зап_вхід_даних_в_реєстр=ОК")
				If Not _ZapStrPap() Then
					GUICtrlSetData($Input7, "Зап_парам_з_інш_шлях=ОК")
				Else
					GUICtrlSetData($Input7, "Зап_парам_з_інш_шлях_не_викон")
				EndIf
			Else
				GUICtrlSetData($Input6, "Зап_знач_в_реєстр_не_викон")
			EndIf
		Else
			$PolzCloseFormGed = 2
			GUICtrlSetData($Input4, "не_вистач_всіх_файл_в_кат_digit")
		EndIf
	EndIf
	$NamFunc = ""
EndFunc





Func _ZmiRBD()
	$NamFunc = "_ZmiRBD"
	_PriZapModlule($UnIdNVF)  
	$VibFilRBD = FileOpenDialog('Виберіть файл БД в форматі *.mdf', @DocumentsCommonDir & '\', '(*.mdf)', 1 + 2)
	If @Error Then
		MsgBox(16 + 4096, $ZagTextPomNVF, 'Не вибрано файлу бази, повторіть спробу ще раз')
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Вибір РБД =ERROR")
		$PolzCloseFormBaza = 1
	Else
		If _PerRbdLdf($VibFilRBD) Then
			$PolzCloseFormKart = 2
		GUICtrlSetState($RdbInput3, $gui_enable)
		Else
		GUICtrlSetData($RdbInput3, '')
		GUICtrlSetData($RdbInput3, $VibFilRBD)
		GUICtrlSetState($RdbInput3, $gui_disable)
		GUICtrlSetState($ButtFileBd, $gui_disable)
		_LogReg($UnIdNVF, $numWin, $NamFunc, "Вибір РБД =OK")
		GUICtrlSetState($ButtPapDig, $gui_enable)
		$PolzCloseFormBaza = 0
		_GUICtrlButton_Click($ButtPapDig)
		EndIf
	EndIf
	$NamFunc = ""
EndFunc
 
 
 
 
 
 
 
 
 
 Func _PerePidkl($SpisBDPar, $PatchProvBDf, $mode)
	; перевірка на наявность
	$NemFunc = "_PerePidkl"
	Local $aMasPiklRbd[0]
	$WhatChek = "" ; що саме перевіряти
	Switch $mode
		Case 1
			$WhatChek = 1 ; шлях до файлу МДФ
		Case 2
			$WhatChek = 0 ; логічy назву бази даних
	EndSwitch
	$IndexZnai = ""
	For $iPar = 1 To UBound($SpisBDPar) - 1
		_ArrayAdd($aMasPiklRbd, $SpisBDPar[$iPar][$WhatChek])
	Next
	$oiskBd = _ArraySearch($aMasPiklRbd, $PatchProvBDf)
	If $oiskBd = -1 Then
		$flagPodkl = -1
		_LogReg($UnIdWIS, $numWin, $NemFunc, "Перев підключ БД до серв=ERROR")
		Return False
	Else
		$flagPodkl = 1
		_LogReg($UnIdWIS, $numWin, $NemFunc, "Перев підключ БД до серв=OК")
		Return True
	EndIf
	$NemFunc = ""
EndFunc	




Func _VkazParFun($UnikIden)
	$NemFunc = "_VkazParFun" 
	$textVetk= $GtextVetk  & $UnikIden 
	$vetkaVhiFil = "HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnikIden & "\InFilesPutch" 
	$RBDMdf = RegRead($vetkaVhiFil, "PutchRBD")
	$RBDLdf = RegRead($vetkaVhiFil, "PutchRBDLog") 
	$Kart = RegRead($vetkaVhiFil, "PutchKart") 
	$ReadParFromReg = RegRead($textVetk, "StructFold") 
	$NewReadParFromReg = "VhidFileKarta:" & $Kart & @CRLF & "VhidFileRelacMdf:" & $RBDMdf & @CRLF &"VhidFileRelacLdf:" & $RBDLdf & @CRLF & $ReadParFromReg 
 	$aPutja = StringSplit(StringStripWS($NewReadParFromReg, 2), @LF) 
	   
	 ; структура массива:
		
		
#CS

		Row|Col  
		[0]|12
		[1]|VhidFileKarta:D:\###_Таблица топокарты\Вхідні файли карта и база\Наступні\КБД\13742202_ДП Городнянський\02_Городнянське UTM.dmf
		[2]|VhidFileRelac:D:\###_Таблица топокарты\Вхідні файли карта и база\Наступні\КБД\13742202_ДП Городнянський\7402-16_Data.MDF
		[3]|ProjFold:C:\Users\User\Documents\TKSNS\Project\TKSNS_2018039103523 
		[4]|ProgrDataFold:C:\Program Files\TKSNS 
		[5]|TemplMap:C:\Program Files\TKSNS\UvPlo.dmf 
		[6]|RbdFileMdf:C:\Program Files\TKSNS\TKSNS.mdf 
		[7]|RbdFileLdf:C:\Program Files\TKSNS\TKSNS_log.ldf 
		[8]|GedIniezFold:C:\Users\User\Documents\TKSNS\Project\TKSNS_2018039103523\BeckUpGedIni\Ged.ini 
		[9]|DigFoldIni:C:\Digitals\Ged.ini 
		[10]|DigFoldLibr:C:\Digitals\Library\TKSNS_SaveUvj.dsf 
		[11]|DriveSql:C:\Digitals\Sql.udl 
		[12]|DigFoldFavor:C:\Digitals\Favorites

#CE

	
	
	For $i = 0 To UBound($aPutja) - 1
		$nampar = StringMid($aPutja[$i], 1, StringInStr($aPutja[$i], ":") - 1)
		$NamParPutc = StringMid($aPutja[$i], StringInStr($aPutja[$i], ":") + 1, -1) 
		Switch $nampar
			Case "VhidFileKarta"
				$FilKartFrom = StringStripWS($NamParPutc, 3) 
			Case "VhidFileRelacMdf"
				$FilRelacMDFrom = StringStripWS($NamParPutc, 3)
			Case "VhidFileRelacLdf"
				$FilRelacLDFrom = StringStripWS($NamParPutc, 3)
			Case "ProjFold"
				$WorkPatchUser = StringStripWS($NamParPutc, 3)
			Case "ProgrDataFold"
				$WorkPatchProgram = StringStripWS($NamParPutc, 3)
			Case "TemplMap"
				$FilTemlateFrom = StringStripWS($NamParPutc, 3)
			Case "RbdFileMdf"
				$FilTksnsMFrom = StringStripWS($NamParPutc, 3)
			Case "RbdFileLdf"
				$FilTksnsLFrom = StringStripWS($NamParPutc, 3)
			Case "GedIniezFold"
				$FilIniWhere = StringStripWS($NamParPutc, 3)
			Case "DigFoldIni"
				$Gedputch = StringStripWS($NamParPutc, 3)
			Case "DigFoldLibr"
				$GedLibrPutch = StringStripWS($NamParPutc, 3)
			Case "DriveSql"
				$GedSqlDriver = StringStripWS($NamParPutc, 3)
			Case "DigFoldFavor"
				$GedFavPutch = StringStripWS($NamParPutc, 3)
		EndSwitch
		
	Next
; 	ConsoleWrite( _
; 			' $FilKartFrom: ' & $FilKartFrom & @CRLF & _
; 			' $FilRelacFrom: ' & $FilRelacMFrom & @CRLF & _
; 			' $WorkPatchProgram: ' & $WorkPatchProgram & @CRLF & _
; 			' $WorkPatchUser: ' & $WorkPatchUser & @CRLF & _
; 			' $FilTemlateFrom: ' & $FilTemlateFrom & @CRLF & _
; 			' $FilTksnsMFrom: ' & $FilTksnsMFrom & @CRLF & _
; 			' $FilTksnsLFrom: ' & $FilTksnsLFrom & @CRLF & _
; 			' $FilIniWhere: ' & $FilIniWhere & @CRLF & _
; 			' $Gedputch: ' & $Gedputch & @CRLF & _
; 			' $GedLibrPutch: ' & $GedLibrPutch & @CRLF & _
; 			' $GedSqlDriver: ' & $GedSqlDriver & @CRLF & _
; 			' $GedFavPutch: ' & $GedFavPutch & @CRLF)
; 	
	$NemFunc = ""
EndFunc






Func _CopKartFile($From, $Where) 
	$NemFunc = "_CopKartFile"    
	$aPathKart= _FO_PathSplit($From)	
	$zvdk=$Where&"\"&$aPathKart[1]&$aPathKart[2]
	If FileCopy($From, $zvdk,9)=0 Then
		MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося копіювання файлу карти")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання файлу карти=ERROR")
		Exit(7785)
	Else
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання файлу карти=OK")
		GUICtrlSetData($KopOsnFile, "Коп_кари_в_папк_проекту=ОК")
	EndIf
	$NemFunc = "" 
EndFunc




Func _CopIniFile($PapkDigIni,$papkBeckIni)
	$NemFunc = "_CopIniFile"  
	If FileCopy($PapkDigIni, $papkBeckIni,9)=0 Then
		MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося створити резервну копію іні-файлу")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Створення копію іні-файлу=ERROR")
		Exit(7786)
	Else
		GUICtrlSetData($CopIniFile, "Рез_коп_іні-файлу=ОК")
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Створення копію іні-файлу=ОК")
	EndIf 
	$NemFunc = "" 
EndFunc


Func _CopTemplKart($PapkFromTempl,$PapkWhereTempl)
	$NemFunc = "_CopTemplKart" 
	If FileCopy($PapkFromTempl, $PapkWhereTempl,9)=0 Then 
		MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося копіювати шаблон карти")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Kопіювання шаблона карти=ERROR")
		Exit(7787)
	Else
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Kопіювання шаблона карти=ОК")
		GUICtrlSetData($CopTemplKart, "Коп_шабл_карти=ОК")
	EndIf 
	$NemFunc = ""
	
EndFunc




Func _SkopPolzRbd()
$NemFunc="_SkopPolzRbd" 
	$FromMdfRbd=RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_"&$UnIdPOF&"\InFilesPutch","PutchRBD")
	$FromLdfRbd=RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_"&$UnIdPOF&"\InFilesPutch","PutchRBDLog")
	$WhereMdfRbd=@MyDocumentsDir&"\TKSNS\Project\TKSNS_"&$UnIdPOF 
	If FileCopy($FromMdfRbd, $WhereMdfRbd,9)=0 Then 
		MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося копіювати MDF-файл таксаційної бази даних ")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання MDF-файлу таксаційної бази даних =ERROR")
		Exit(7788)
	Else
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання MDF-файлу таксаційної бази даних =OК")
		GUICtrlSetData($CopTksnsBdMDF, "Коп_файлу_такац_бази_MDF=ОК")
	EndIf  
	If FileCopy($FromLdfRbd, $WhereMdfRbd,9)=0 Then 
		MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося копіювати LDF-файл таксаційної бази даних ")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання LDF-файлу  таксаційної бази даних =ERROR")
		Exit(7789)
	Else
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання LDF-файлу  таксаційної бази даних =ОК")
		GUICtrlSetData($CopTksnsBdLDF, "Коп_файлу_ таксац_бази_LDF_TKSN=ОК")
	EndIf  
$NemFunc=""
EndFunc 






unc _CopTksnsBd($FromTksnsM,$FromTksnsL,$PaspkNaznTksns)
	  
	$NemFunc = "_CopTksnsBd"
	$aPathTksnsM= _FO_PathSplit($FromTksnsM)	
  	$WhreTksnsM=$PaspkNaznTksns&"\"&$aPathTksnsM[1]&$aPathTksnsM[2]
	If FileCopy($FromTksnsM, $WhreTksnsM,9)=0 Then 
		MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося копіювати MDF-файл бази даних ")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання MDF-файлу бази даних =ERROR")
		Exit(7788)
	Else
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання MDF-файлу бази даних =OК")
		GUICtrlSetData($CopTksnsBdMDF, "Коп_файлу_бази_MDF_TKSN=ОК")
	EndIf 
		$aPathTksnsL= _FO_PathSplit($FromTksnsL)	
  	$WhreTksnsL=$PaspkNaznTksns&"\"&$aPathTksnsL[1]&$aPathTksnsL[2] 
	If FileCopy($FromTksnsL, $WhreTksnsL,9)=0 Then 
		MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося копіювати LDF-файл бази даних ")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання LDF-файлу бази даних =ERROR")
		Exit(7789)
	Else
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Копіювання LDF-файлу бази даних =ОК")
		GUICtrlSetData($CopTksnsBdLDF, "Коп_файлу_бази_LDF_TKSN=ОК")
	EndIf 
	$NemFunc = "" 
EndFunc

#CE





Func _StvorKatTempl()
	$NemFunc = "_StvorKatTempl"
	$filkenam = _FO_PathSplit(RegRead($GtextVetk & $UnIdWWD & "\InFilesPutch", "PutchKart"))
	$PutBezFileKart = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWWD
	$PutIzFileKart = $PutBezFileKart & "\" & $filkenam[1] & $filkenam[2]
	$KatZberInTempl = $PutBezFileKart & "\OnTemplate"
	If DirCreate($KatZberInTempl) = 0 Then
		MsgBox(16 + 4096, $ZagTextPomWWD, " ")
		_ObrErr($UnIdWWD, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Створення каталогу для копії іні-файлу=ERROR")
		Exit(7794)
	Else
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Створення каталогу для копії іні-файлу=ОК")
; 		GUICtrlSetData($StvorPapku, "Створ_катл_для_карт_в_шаблон="&$KatZberInTempl)
	EndIf
	$NemFunc = ""
	Return $KatZberInTempl
; 	ConsoleWrite('result=' & $PutBezFileKart & @CRLF & $PutIzFileKart & @CRLF & $KatZberInTempl & @CRLF)
EndFunc



Func _OtrimParResu()
	$NamFunc = "_OtrimParResu"
	$sStringParVidVid = @ScriptDir & "\Resurce\TipSlo.txt" _
			& @CRLF & @ScriptDir & "\Resurce\RekKZH.txt" _
			& @CRLF & @ScriptDir & "\Resurce\RekKZM.txt"
	$Vidil = FileExists(@ScriptDir & "\Resurce\TipSlo.txt")
	$GrKZH = FileExists(@ScriptDir & "\Resurce\RekKZH.txt")
	$GrKZV = FileExists(@ScriptDir & "\Resurce\RekKZM.txt")
	If ($Vidil = 0) Or ($GrKZH = 0) Or ($GrKZV = 0) Then
		MsgBox(16 + 4096, $ZagTextPomZZP, "Не вистачає файлу ресурсів")
		_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
		Exit(999)
	Else
		GUICtrlSetData($InpPerVsiVid, "Exist SpSlo=ОК")
		GUICtrlSetData($InpPerRekomKZH, "Exist SpKzh=ОК")
		GUICtrlSetData($InpPerRekomKZM, "Exist SpKzm=ОК")
	EndIf
	$aResPutch = StringSplit($sStringParVidVid, @CR)
	$sTexComSql = "EXEC StvorKartTables_1 "
	For $iRes = 1 To UBound($aResPutch) - 1
		$sTexComSql &= "'" & StringStripWS($aResPutch[$iRes], 3) & "'"
		If $iRes = UBound($aResPutch) - 1 Then ContinueLoop
		$sTexComSql &= ","
	Next
	Return $sTexComSql
	$NamFunc = ""
EndFunc

Func _PriZapModlule($UnIdIspFuncc)
	$NamFunc = "_PriZapModule"
	RegDelete("HKLM\SOFTWARE\TKSNS", "ErrMethod")
	RegDelete("HKLM\SOFTWARE\TKSNS", "ErrCode")
	RegDelete("HKLM\SOFTWARE\TKSNS", "ErrSpos")
	RegDelete("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdIspFuncc, "ErrorWin")
	RegDelete("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdIspFuncc, "ErrorNamFunc")
	RegDelete("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdIspFuncc, "ErrorFlag")
	RegDelete("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdIspFuncc, "LogString")
	
	$NamFunc = ""
EndFunc

Func _OnEndScr($RegVetkaPutchhh, $NomWinnn, $namewinnn, $NamewModul, $NowE)
; 	ConsoleWrite('result=' & $RegVetkaPutchhh & @CRLF)
	$dfs = $RegVetkaPutchhh & "\Win\" & $NomWinnn
	RegWrite($dfs, "NameModul", "REG_SZ", $NamewModul)
	RegWrite($dfs, "NumberWin", "REG_SZ", $namewinnn)
	RegWrite($dfs, "End", "REG_SZ", $NowE)
EndFunc

Func _CopyTksnsBd($id)
	;-----------параметри-------------
	$Sufix = "TKSNS_" & $id
; 	$FromTksns = _IzReg("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Project\TKSNS_2018016021519\Etap\WorkWithServer", "TksnsBdFrom")
; 	$FromTksnsMdf = StringMid($FromTksns, 1, StringInStr($FromTksns, '|') - 1)
	$FromTksnsLdf = StringReplace($FromTksnsMdf, '.mdf', '_log.ldf')
; 	$whereTksnsBd = _IzReg("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Project\TKSNS_2018016021519\Etap\WorkWithServer", "TksnsBdWhere")
	$wheMdfCop = @MyDocumentsDir & "\TKSNS\Project\" & $Sufix & "\TKSNS.mdf"
	$wheLdfCop = @MyDocumentsDir & "\TKSNS\Project\" & $Sufix & "\TKSNS_log.ldf"
	$FlagErr = 0
	$NamFuncErr = ""
	;---------------------------------
	$KopMdf = FileCopy($FromTksnsMdf, $wheMdfCop, 1)
	$KopLdf = FileCopy($FromTksnsLdf, $wheLdfCop, 1)
	If BitAND($KopMdf, $KopLdf) = 0 Then
		MsgBox(16 + 4096, "Помилка копіювання баз", "не вдалося скопіювати фізичні файли бази даних")
		$NamFuncErr = "Копіюв_файлів_БД"
		$FlagErr = 1
		Return 'Copy_File_Rbd-bed'
	EndIf
	Return 'Copy_File_Rbd-good'
EndFunc

Func _PoperObrGed($Id)
	;-----------параметри-------------
	$ppa = "TKSNS_" & $Id
	$PutcGedd = _IzReg("HKLM\SOFTWARE\TKSNS\Project\" & $ppa, "PutchGed")
	$ZapInsss = _IzReg("HKLM\SOFTWARE\TKSNS\Project\" & "TKSNS_" & $id & "\Etap\PidgotZapus", "PutchBeckupIni")
	$putchxx = StringMid($ZapInsss, StringInStr($ZapInsss, '\\') + 2)
	$WhereBeckup = @MyDocumentsDir & "\TKSNS\Project\" & $ppa & "\" & $putchxx
	$StvorPapkBeckIni = @MyDocumentsDir & "\TKSNS\Project\" & $ppa & "\" & StringMid($putchxx, 1, StringInStr($putchxx, "\Ged.ini") - 1)
	$FromGrdInibeckup = $PutcGedd & "\Ged.ini"
	$drsqlfil = $PutcGedd & "\" & StringMid(_IzReg("HKLM\SOFTWARE\TKSNS\Project\TKSNS_2018016021519\Etap\PidgotZapus", "FileDrivSql"), StringInStr(_IzReg("HKLM\SOFTWARE\TKSNS\Project\TKSNS_2018016021519\Etap\PidgotZapus", "FileDrivSql"), '\\') + 2)
	$rbdVik = _IzReg("HKLM\SOFTWARE\TKSNS\Project\" & $ppa, "NameWorkRbd")
	$FromTemplkart = _IzReg("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Project\TKSNS_2018016021519\Etap\PidgotZapus", "PutchFromTepmplFile")
	$WhereTemplkart = @MyDocumentsDir & "\TKSNS\Project\" & $ppa & "\" & StringMid(_IzReg("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Project\TKSNS_2018016021519\Etap\PidgotZapus", "PutchInTepmplFile"), StringInStr(_IzReg("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Project\TKSNS_2018016021519\Etap\PidgotZapus", "PutchInTepmplFile"), '\\') + 2)
	$FlagErr = 0
	$NamFunkErr = ""
	;---------------------------------
	If Not FileExists($StvorPapkBeckIni) Then
		$Sretds = DirCreate($StvorPapkBeckIni)
		If $Sretds = 0 Then
			MsgBox(16 + 4096, 'Помилка створення каталогу ', "Не вдалося  створити каталог для резервної копію Ged.ini")
			$FlagErr = 1
			$NamFunkErr = "Створ_катал_для_рез_копії_іні"
			Return 'Craete_Putch-bed'
		EndIf
	EndIf

	FileCopy($FromTemplkart, $WhereTemplkart)
	If @Error Then
		MsgBox(16 + 4096, 'Error', "Не вдалося  скопіювати шаблон карти")
		$FlagErr = 1
		$NamFunkErr = "Коп_шабл_карти"
		Return 'bebckupini-bed'
	EndIf

	FileCopy($FromGrdInibeckup, $WhereBeckup)
	If @Error Then
		MsgBox(16 + 4096, 'Error', "Не вдалося  зберегти резервну копію Ged.ini")
		$FlagErr = 1
		$NamFunkErr = "Рез_коп_іні_дігі"
		Return 'bebckupini-bed'
	EndIf

	Global $Port = IniRead($FromGrdInibeckup, 'FormOptions.edtTCPCommand', 'Text', 0)
	$sAvtoz = IniWrite($FromGrdInibeckup, "FormOptions.chbLog", "Checked", "0")
	$sSqlmod = IniWrite($FromGrdInibeckup, "Constants", "SQLMode", "0")
	$sSqlServBd = IniWrite($FromGrdInibeckup, "Constants", "SQLVvodParameters", @ComputerName & " " & "TKSNS")
	$sSqlprov = IniWrite($FromGrdInibeckup, "Constants", "DataLink", "SQL.udl")
	$RozParList = IniWrite($FromGrdInibeckup, "FormGed.spbInfoDown", "Down", "1")

	If BitAND($Port, $sAvtoz, $sSqlmod, $sSqlServBd, $sSqlprov, $RozParList) <> 1 Then
		MsgBox(16 + 4096, "Помилка ", "не можу записати іні файли")
		$FlagErr = 1
		$NamFunkErr = "Запис_іні_файлів"
		Return 'WriteIni-bed'
	EndIf

	Const $sInput = "Provider=SQLNCLI10.1;Integrated Security=SSPI;Persist Security Info=False;User ID="""";Initial Catalog="
	Const $sInput2 = ";Data Source="
	Const $sInput3 = ";Initial File Name="""";Server SPN="""""
	$hFileSQL = FileOpen($drsqlfil, 3)
	If $hFileSQL = 0 Then
		MsgBox(16 + 4096, "Помилка рядка з'днання з SQL Server", "Не знайдено драйвер SQL", 3)
		$FlagErr = 1
		$NamFunkErr = "Немає_файлу_драйвера"
		Return 'NotExistsSqlUdl-bed'
		Exit
	EndIf
	$str = $sInput & "TKSNS" & $sInput2 & @ComputerName & $sInput3
	_ZapLogInp("Визначено рядок підключення: " & $str, @ScriptDir & "\Log_TKSNS.log", 1, @ScriptName)
	_FileWriteToLine($drsqlfil, _FileCountLines($drsqlfil), $str, 1)
	If @Error Then
		MsgBox(16 + 4096, "Помилка рядка з'днання з SQL Server", "Не вийшло записати рядок підключення SQL", 3)
		$FlagErr = 1
		$NamFunkErr = "Помил_запису_в_файлу_драйвера"
		Return 'Write_in_sql_udl-bed'
	EndIf
	FileClose($hFileSQL)
	Return "Poper_obr-good"
EndFunc




