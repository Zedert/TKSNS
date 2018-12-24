 Func _ReadPortTcpGed($PutchKIniFile)  
	$NemFunc = "_ReadPortTcpGed"
	$Port=IniRead($PutchKIniFile, 'FormOptions.edtTCPCommand', 'Text', 'NO')
		If  $Port='NO' Then
			MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося прочитати номер TCP-порту з іні-файлу")
			_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
			_LogReg($UnIdPOF, $numWin, $NemFunc, "Читання номеру TCP-порту=ERROR")
			Exit(7790)
		Else
			_LogReg($UnIdPOF, $numWin, $NemFunc, "Читання номеру TCP-порту=ОК")
			GUICtrlSetData($ReadNumTcpportGed, "Зчит_номер_TCP-порт="&$Port)
		EndIf
	RegWrite($GtextVetk&$UnIdPOF,"GedTcpPort","REG_SZ",$Port) 
	$NemFunc = "" 
EndFunc




Func _ZapParIni($PutchKIniFile,$RelacBd) 
	$NemFunc = "_ZapParIni" 
	$aParIniGed=_FO_PathSplit($RelacBd) 
	$sAvtoz = IniWrite($PutchKIniFile, "FormOptions.chbDesktop","Checked","0")
	$sSqlmod = IniWrite($PutchKIniFile, "Constants", "SQLMode", "0")
	$sSqlServBd = IniWrite($PutchKIniFile, "Constants", "SQLVvodParameters", @ComputerName & " " & $aParIniGed[1])
	$sSqlprov = IniWrite($PutchKIniFile, "Constants", "DataLink", "SQL.udl")
	$RozParList = IniWrite($PutchKIniFile, "FormGed.spbInfoDown", "Down", "1")
	If ($sAvtoz=0 Or $sSqlmod=0 Or $sSqlServBd=0 Or $sSqlprov=0 Or $RozParList=0) Then 
		MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося змінити налаштування в іні-файлі")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Зміна налаштувань в іні-файлі=ERROR")
		Exit(7791)
	Else
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Зміна налаштувань в іні-файлі=OК")
		GUICtrlSetData($WritParInIni, "Зап_парам_в_іні-файл=ОК")
	EndIf 
	$NemFunc = ""
EndFunc





Func _VstNamBdSqlDriver($PutKPapDig)
  _ViknProgess($GuiPOF,$ModName,"Встановлення рядка підключення...",$ZagPrBarWin,2,100,-1,-1,16)    
	$NemFunc = "_VstNamBdSqlDriver"
	Const $sInput = "Provider=SQLNCLI10.1;Integrated Security=SSPI;Persist Security Info=False;User ID="""";Initial Catalog="
	Const $sInput2 = ";Data Source="
	Const $sInput3 = ";Initial File Name="""";Server SPN="""""
	$hFileSQL = FileOpen($PutKPapDig, 3)
	If $hFileSQL =-1 Then 
		MsgBox(16+4096, $ZagTextPomPOF, "Не знайдено драйвер SQL")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Пошук  SQL-драйвера=ERROR")
		Exit(7792)
	EndIf 
	_LogReg($UnIdPOF, $numWin, $NemFunc, "Пошук  SQL-драйвера=ОК")
	$str = $sInput & "TKSNS" & $sInput2 & @ComputerName & $sInput3 
	_FileWriteToLine($PutKPapDig, _FileCountLines($PutKPapDig), $str, 1)
	If @error Then
		MsgBox(16+4096, $ZagTextPomPOF, "Не вдалося перезаписати значення імені БД в драйвері SQL")
		_ObrErr($UnIdPOF, $numWin, $NemFunc,@AutoItPID)
		_LogReg($UnIdPOF, $numWin, $NemFunc, "Встановлення імені БД в SQL-драйвера=ERROR")
		Exit(7792)
	Else
	_LogReg($UnIdPOF, $numWin, $NemFunc, "Встановлення імені БД в SQL-драйвера=ОК")
		GUICtrlSetData($VstanSqldrive, "Втан_знач_в_SQL-драйвер=ОК")
	EndIf
	FileClose($hFileSQL)
	$NemFunc = "" 
EndFunc





Func _StatBarDig()
	$WiknDig = WinActivate("[CLASS:TFormGed;Title:Digitals XE]", "")
	$Stabar = ControlGetHandle($WiknDig, "", "TStatusBar1")
	$TextStatBar = ControlGetText($WiknDig, "", $Stabar)
EndFunc





Func _PriVih()
	If ProcessExists("ged.exe") Then ProcessClose("ged.exe")
EndFunc






Func _PerZapDig()
	$NemFunc = "_PerZapDig"
	If Not ProcessExists("ged.exe") Then
		MsgBox(16 + 4096, $ZagTextPomWWD, "Було закрито Digitals, відбудеться вихід із програми", 1)
		_ObrErr($UnIdWWD, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Процес Digitals=ERROR")
		Exit(7797)
	EndIf
	$NemFunc = ""
EndFunc






Func _SaveMapInTempl()
	$NemFunc = "_SaveMapInTempl"
	$putDoShabl = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWWD & "\UvPlo.dmf"
	$TilkImaFil=_FO_PathSplit($NamKart) 
	$putDoScript = $putDoDig & "\Library\TKSNS_SaveUvj.dsf" 
	$TextCommDig = $TextSavKartScript1 & $putDoShabl & @CRLF & $TextSavKartScript2 &  $TextSavKartScript3 &"TKSNS_"&$UnIdWWD& @CRLF & $TextSavKartScript4 
	If FileExists($putDoScript) Then FileDelete($putDoScript)
	_FileCreate($putDoScript)
	FileWrite($putDoScript, $TextCommDig)
	$NamFileOpnKrt = _ComDigitals("@ExecuteScript TKSNS_SaveUvj", $NumTcpPort)
	If @Error Then
		MsgBox(16 + 4096, $ZagTextPomWWD, "Не вдалося під'єднати карти до шаблону...")
		_ObrErr($UnIdWWD, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Під'єднання карти до шаблону=ERROR")
		Exit(7796)
	Else
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Під'єднання карти до шаблону=ОК")
		GUICtrlSetData($SaveinShablon, "Карт_підє_до_шабл=ОК")
	$NewPatch= @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWWD & "\"&$TilkImaFil[1]&$TilkImaFil[2]
	$VersKartFile = FileGetVersion($NamKart)
	$RozmKartFile = FileGetSize($NamKart)
	$VremFileKart = FileGetTime($NamKart)
	If @Error Then
		MsgBox(16 + 4096, $ZagTextPomWWD, "не вдалося отримати інформацію про  файл карти")
		_ObrErr($UnIdWWD, $numWin, $NemFunc, @AutoItPID)
		Exit(9994)
	Else
		$yyyymd = $VremFileKart[0] & '/' & $VremFileKart[1] & '/' & $VremFileKart[2]
		$hhmmss = $VremFileKart[3] & ':' & $VremFileKart[4] & ':' & $VremFileKart[4]
		$infoForMap = "DateModif:" & $yyyymd &@CRLF& "TimeFileMod:" & $hhmmss &@CRLF& "SizeFilKart:" & $RozmKartFile
		RegWrite($GtextVetk&$UnIdWWD,"ActualKart","REG_MULTI_SZ",$infoForMap)
	EndIf
	
	
	
	
	
	EndIf
	$NemFunc = ""
EndFunc






Func _OpenMapFil()
	$NemFunc = "_OpenMapFil"
	$TilkNamFilKart = _FO_PathSplit($NamKart)
	$NumTcpPort = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWWD, "GedTcpPort")
	$PutDoKart = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWWD & "\" & $TilkNamFilKart[1] & $TilkNamFilKart[2]
	$FilKartOpen = _ComDigitals("@FileOpen " & $PutDoKart, $NumTcpPort)
	If $FilKartOpen = 'no_proc_ged' Then
		_ObrErr($UnIdWWD, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Відриття карти=ERROR")
		Exit(7795)
	Else
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Відриття карти=OК")
		GUICtrlSetData($VidkrKart, "Відк_карта=" & $TilkNamFilKart[1] & $TilkNamFilKart[2])
	EndIf
	$NemFunc = ""
	Return $PutDoKart
EndFunc





Func _StartDigit($ExeGed)
	$NemFunc = "_StartDigit"
	$hWnd = ""
	$StartFunc = TimerInit()
	$gedd = ShellExecute($ExeGed)
	ProcessWait('ged.exe')
	$DigZapu = TimerDiff($StartFunc)
	$hWnd = WinWaitActive("[CLASS:TFormGed;Title:Digitals XE]", "", 20)
; 	ConsoleWrite('time: ' & Round($DigZapu, -1) & " milisec or " & Round($DigZapu / 1000, 2) & " sec" & @CRLF)
	If Not $hWnd Then
		MsgBox(16 + 4096, $ZagTextPomWWD, "Не знайдено вікно Digitals")
		_ObrErr($UnIdWWD, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Вікно Digitals=ERROR")
		Exit(7795)
	Else
		GUICtrlSetData($ZapuskDig, "Digit_запущ_із_ключем_PID=" & $gedd)
		RegWrite($GtextVetk & $UnIdWWD, "PidGed", "REG_SZ", $gedd)
		_LogReg($UnIdWWD, $numWin, $NemFunc, "PID процесу Digitals=OК")
		Return $hWnd
		$NemFunc = ""
	EndIf
EndFunc






Func _ProvExDig()
	$NemFunc = "_ProvExDig"
	
	If ProcessExists("ged.exe") Then
		MsgBox(16 + 4096, $ZagTextPomWWD, "Знайдено поцес ged.exe. Для прподовження роботи програми ТаксІнс - закрийте програму Digitals, та спробуйте ще раз...")
		_ObrErr($UnIdWWD, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Перевірка наявності прооцесу ged.exe=ERROR")
		Exit(7792)
	Else
		_LogReg($UnIdWWD, $numWin, $NemFunc, "Перевірка наявності прооцесу ged.exe=OК")
		GUICtrlSetData($perExistProcGed, "Перев_на_існув_проц_ged.exe=ОК")
	EndIf
	$NemFunc = ""
	Return "Оk"
EndFunc




Func _StvorDsfFilScr()
	$NamFunc = "_StvorDsfFilScr"
	$SozScrVidMejLisn = ""
	$InfoProGroup = ""
	$SozScrVidMejLisn = _StvorSdfFile("TKSNS_PoznMejiLisn.dsf", "\Library\", "@Map.SelectByParameters 2|-7=81261000|0>0", $UnIdCLD, $numWin, $ZagTextPomCLD)
	If $SozScrVidMejLisn = "" Then
		MsgBox(16 + 4096, $ZagTextPomCLD, "Не вдається створити файл скрипта, що помітить межі лісництва")
		_ObrErr($UnIdCLD, $numWin, $NamFunc, @AutoItPID)
		Return 0
	Else
		GUICtrlSetData($InpFilePutchSdfScr, $SozScrVidMejLisn)
	EndIf
	$NamFunc = ""
	Return $SozScrVidMejLisn
EndFunc




Func _PoznMejiLisn()
	$NamFunc = "_PoznMejiLisn"
	Const $IdSloMejLisn = "81261000"
	$hGedGol = WinActivate("[CLASS:TFormGed;Title:Digitals XE]", "")
	If Not $hGedGol Then
		MsgBox(16 + 4096, $ZagTextPomCLD, "Не вдалося знайти головне вікно Digitals")
		_ObrErr($UnIdCLD, $numWin, $NamFunc, @AutoItPID)
		Exit(9991)
	Else 
		GUICtrlSetData($InpHendlGed, "Дескрипт вікн Ged=" & $hGedGol)
		$CoorGolGed=""
		$CoorGolGed=WinGetPos($hGedGol,"")
		If ($CoorGolGed[0]<>-8) Or ($CoorGolGed[1]<>-8) Then 
			$winSetStt=""
			$winSetStt=WinSetState($hGedGol,"",@SW_MAXIMIZE) 
		EndIf
		; 		_ArrayDisplay($CoorGolGed, "$CoorGolGed=") 
 
 
			
			
			$CommInServer = ""
			$CommInServer = _ComDigitals("@ExecuteScript TKSNS_PoznMejiLisn", $TcpIpPortGed)
			If Not $CommInServer Then 
				MsgBox(16 + 4096, $ZagTextPomCLD, "Не вдалося виконати помітку меж лісництва")
				_ObrErr($UnIdCLD, $numWin, $NamFunc, @AutoItPID)
				Exit(9999)
			Else
				GUICtrlSetData($InpPomMejiLisn,"Поміч_меж_лісниц="&$CommInServer )
			EndIf
			
			$DescStatBar=""
			$DescStatBar=ControlGetHandle($hGedGol, "", "[Class:TStatusBar; Instance:1]")
			If Not $DescStatBar Then
				MsgBox(16 + 4096, $ZagTextPomCLD, "Не вдалося здобути хендл статус-бару")
				_ObrErr($UnIdCLD, $numWin, $NamFunc, @AutoItPID)
				Exit(9993)
			Else
			GUICtrlSetData($InpElemStatBar,"Дескрипт стат-бару отримано="&$DescStatBar)
			EndIf 
			
			$ZachIzStatBar=""
			$ZachIzStatBar=_GUICtrlStatusBar_GetText($DescStatBar,0)
			If  $ZachIzStatBar="" Then 
				MsgBox(16 + 4096, $ZagTextPomCLD, "Не вдалося вичитати інфу із статус бару GED")
				_ObrErr($UnIdCLD, $numWin, $NamFunc, @AutoItPID)
				Exit(9999)
			Else
				$NewKilObjMejLisn=StringMid($ZachIzStatBar,StringInStr($ZachIzStatBar,"/")+1,-1)
				_GUICtrlStatusBar_SetText($hStatus,$NewKilObjMejLisn,0)
			EndIf
			
			
			
			$OtrSpisParML="" 
			$OtrSpisParML=_ComDigitals("@Map.Layers.GetValidParameters ID81261000",$TcpIpPortGed)
			If  $OtrSpisParML="" Then 
				MsgBox(16 + 4096, $ZagTextPomCLD, "Не вдалося отримати список доступних параметрів")
				_ObrErr($UnIdCLD, $numWin, $NamFunc, @AutoItPID)
				Exit(9999)
			Else
				$NewSpiDosPar=StringMid($OtrSpisParML,StringInStr($OtrSpisParML,"ID81261000",0,-1)+11,-1) 
				$aSpisDossPar=StringSplit($NewSpiDosPar," ") 
				$RozmMass=""
				$RozmMass=UBound($aSpisDossPar) - 1
				For $idd =1 To UBound($aSpisDossPar) - 1
					If $aSpisDossPar[$idd]=0 Then 
					$RozmMass&="/"&$idd
					ExitLoop
					EndIf
				Next 
				_GUICtrlStatusBar_SetText($hStatus,$RozmMass,1) 
				GUICtrlSetData($InpSpisDostPar,$OtrSpisParML)
			EndIf 
#CE

			
		$TabSheet=""
		$TabSheet = ControlGetHandle($hGedGol, "", "[Class:TTabSheet; Instance:1]")
		If Not $TabSheet Then
			MsgBox(16 + 4096, $ZagTextPomCLD, "Не вдалося здобути хендл елементу управління ")
			_ObrErr($UnIdCLD, $numWin, $NamFunc, @AutoItPID)
			Exit(9992)
		Else
			GUICtrlSetData($InpElemEdit, "Дескрипт TTabSheet отримано=" & $TabSheet) 
				$aPosTabSheet=ControlGetPos($hGedGol,"",$TabSheet) 
				MouseMove($aPosTabSheet[0],$aPosTabSheet[1]+$aPosTabSheet[2]/2 ,50)  
				
		
		endif 
		
			
			$elemTstringGrid=""
			$elemTstringGrid=ControlGetHandle($hGedGol,"","[Class:TStringGrid; Instance:1]")
			If $elemTstringGrid="" Then  
				MsgBox(16+4096, $ZagTextPomCLD, "Не вдалося отримати дескриптор елементу сітка")
				_ObrErr($UnIdCLD, $numWin, $NamFunc,@AutoItPID)
				Exit(999)
			 Else
				GUICtrlSetData($InpElemTgrid,"Дескр_елем_TStringGrid_знайдено="&$elemTstringGrid) 
				
				
				
				
; 				_ArrayDisplay($aPosTstrGrin,"Массив")
#CS
 $array[0] = X координата 
 $array[1] = Y координата 
 $array[2] = ширина 
 $array[3] = высота 



Позиция елемента управления в Digitals - там где параметры
1122    X
85      Y
238     Width
252		Height


Окно развернутого Digitals
-8 	   X
-8 	   Y
1382   Width 
744    Height




 

			 EndIf
			 
			 
			 
#CS

			 $ElemTGroupBox3="" 
			 $ElemTGroupBox3=ControlGetHandle($hGedGol,"","[Class:TGroupBox; Instance:3]") 
			If $ElemTGroupBox3="" Then  
				MsgBox(16+4096, $ZagTextPomCLD, "Не вдалося отримати дескриптор елементу Гроу-Бокс")
				_ObrErr($UnIdCLD, $numWin, $NamFunc,@AutoItPID)
				Exit(999)
			 Else 
				GUICtrlSetData($InpTGroupBox3,"Дескр_елем_TGroupBox_знайдено="&$ElemTGroupBox3)
			 EndIf
			  
			  
			  
			  
			  
			  
			$pokazSpisPar=""
			$pokazSpisPar=_ComDigitals("@ExecuteMenu spbDirectory",$TcpIpPortGed)
			If $pokazSpisPar="" Then  
				MsgBox(16+4096, $ZagTextPomCLD, "Не вдалося запустити скриптову команду на відкриття списку параметрів")
				_ObrErr($UnIdCLD, $numWin, $NamFunc,@AutoItPID)
				Exit(999) 
			 EndIf
			
#CE

			
			
#CS

			$SumVidObj=""
			$SumVidObj=ControlClick($hGedGol,"",$DeskElemEdit,"main" ) ;"right", "middle", "main", "menu", "primary", "secondary". 
			If $SumVidObj="" Then  
				MsgBox(16+4096, $ZagTextPomCLD, "Не вдалося визначити сумарну площу виділених об'єктів меж лісництва")
				_ObrErr($UnIdCLD, $numWin, $NamFunc,@AutoItPID)
				Exit(999)
			Else
				ConsoleWrite('result=' & $SumVidObj & @CRLF)
; 				GUICtrlSetData($InpElemTgrid,"Дескр_елем_TStringGrid_знайдено="&$elemTstringGrid)
			 EndIf  
			  
#CE

; 			$aPosPerRyadPar=ControlGetPos($hGedGol,"",$DeskElemEdit)
; 			_ArrayDisplay($aPosPerRyadPar,"Массив")
; 			$aPosMishk=MouseGetPos()
; 			MouseMove($aPosPerRyadPar[0],$aPosPerRyadPar[1],70) 
; 			MouseMove($aPosMishk[0],$aPosMishk[1],70)

; 			$ContrlCommaaa=""
; 			$ContrlCommaaa=ControlCommand($hGedGol,"",$DeskElemEdit,"EditPaste", 'Хуйня;Шинема-хуйня');"GetCurrentCol", "";"GetCurrentLine", ""
; 			If @Error Then 
; 				MsgBox(16+4096, $ZagTextPomCLD, "Не вдалося повернути кість рядків елементу")
; 				_ObrErr($UnIdCLD, $numWin, $NamFunc,@AutoItPID)
; 				Exit(999)
; 			Else
; 				MsgBox(0, '', ":"&$ContrlCommaaa&":")
; 			EndIf
; 			
			
			
			
			
			
		EndIf
; 	EndIf
	$NamFunc = ""
EndFunc
 
 
 
 
 
 
 
 Func _AddToGrooup($nameGroup, $mode)
	Switch $mode
		Case 1
			_ComDigitals("@Map.DeleteGroup " & $nameGroup, $NomPort)
			_ComDigitals("@Map.Selected.AddToGroup " & $nameGroup, $NomPort)
		Case 2
			_ComDigitals("@Map.DeleteGroup " & $nameGroup, $NomPort)
	EndSwitch
	
EndFunc





Func _CheckPar()
	$FlagEstOshib=0  
	_PoznPustParam1()
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів",2,25)
	_PoznPustParam2()
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів",2,50)
	_PoznPustParam3()
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів",2,75)
	_PoznPustParam4() 
	If $FlagEstOshib=0 Then
		_ViknProgess($GuiPSP,$ModName,$ZagPrBarWin,$PidZagPrBarWin,4)
	Else
		_ViknProgess($GuiPSP,$ModName,$ZagPrBarWin,$PidZagPrBarWin,3)
	EndIf
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів",2,100)
; 	_ViknProgess($GuiPSP,$ModName, "Перевірка наявновсті незаповнених параметрів" ,3)
; 	If $FlagEstOshib>0 Then GUICtrlSetState($GuiPSP,@SW_SHOW) 
EndFunc






Func _StatBarDig($dsa)
	$NemFunc = "_StatBarDig"
	$WiknDig = WinActivate("[CLASS:TFormGed;Title:Digitals XE]", "")
	$Stabar = ControlGetHandle($WiknDig, "", "TStatusBar1")
	$sdr=""
; 	$sdr = _GUICtrlStatusBar_GetText($Stabar, 0)
	
	If @Error Then 
		MsgBox(16 + 4096, $ZagTextPomPSP, "Не вдалося повернути значення кількості помічених об'єктів")
		_ObrErr($UnIdPSP, $numWin, $NemFunc, @AutoItPID)
		Exit(8000)
	Else
		Switch $dsa
			Case 1 
				GUICtrlSetData($InKilLg,StringMid(_GUICtrlStatusBar_GetText($Stabar, 0),StringInStr(_GUICtrlStatusBar_GetText($Stabar, 0), "/") + 1)); StringMid($sdr, StringInStr($sdr, "/") + 1)
			Case 2
				GUICtrlSetData($InKilLs, StringMid(_GUICtrlStatusBar_GetText($Stabar, 0),StringInStr(_GUICtrlStatusBar_GetText($Stabar, 0), "/") + 1))
			Case 3
				GUICtrlSetData($InKilKw, StringMid(_GUICtrlStatusBar_GetText($Stabar, 0),StringInStr(_GUICtrlStatusBar_GetText($Stabar, 0), "/") + 1))
			Case 4
				GUICtrlSetData($InKilVd, StringMid(_GUICtrlStatusBar_GetText($Stabar, 0),StringInStr(_GUICtrlStatusBar_GetText($Stabar, 0), "/") + 1))
		EndSwitch
	EndIf
	$NemFunc = ""
EndFunc

 




Func _ComDigInCtrl()
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів") 
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів",2,0)
	$putKartZregistr = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\InFilesPutch", "PutchKart")
	$RozEdn = _FO_PathSplit($putKartZregistr)
	;шлях до карти в папці "мої документи" 
	$PutchDlaPoriv = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdPSP & "\" & $RozEdn[1] & $RozEdn[2]
	; вставити в елемент управління "номер параметра"  - що здобутий як відповідь Дігі на команду -- Номер за Именем 
; 	_ComDigitals("@Map.BeginUpdate", $NomPort)
	GUICtrlSetData($InNom1, _ComDigitals("@Map.Parameters.FindByName Lg", $NomPort))
	GUICtrlSetData($InNom2, _ComDigitals("@Map.Parameters.FindByName Ls", $NomPort))
	GUICtrlSetData($InNom3, _ComDigitals("@Map.Parameters.FindByName Kw", $NomPort))
	GUICtrlSetData($InNom4, _ComDigitals("@Map.Parameters.FindByName Vd", $NomPort))
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів",2,25)

    ;визначає об'єкти, які виділені за допомогою скриптової функцію """"позначити за параметрами (2|плоша>0|ном_пар_код_ЛГ="" - тобто відсутність значення параметру)"""
	$ESADSd1 = "@Map.SelectByParameters 2|0>0|" & GUICtrlRead($InNom1) & "="
	$ESADSd2 = "@Map.SelectByParameters 2|0>0|" & GUICtrlRead($InNom2) & "="
	$ESADSd3 = "@Map.SelectByParameters 2|0>0|" & GUICtrlRead($InNom3) & "="
	$ESADSd4 = "@Map.SelectByParameters 2|0>0|" & GUICtrlRead($InNom4) & "=" 
; 	Sleep(500)
;  _ComDigitals("@Map.EndUpdate", $NomPort)
	
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів",2,50)\

	;зчитує з реєстру шлях до папки редактор
	$WhereDig = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\InFilesPutch", "PutchDigit")
	
	
	;Якщо в папці Дігі/лібраре є файл скрипта ==>>  видаляємо файл  
	If FileExists($WhereDig & "\Library\SelLg.dsf") Then FileDelete($WhereDig & "\Library\SelLg.dsf")
	;записуємо текст команди, який == заборонити оновлення карти на екрані - зняти зі всіх позначку 
	; - позначити об'єкти, що мають два параметри  (площа більше нуля, та номер параметру має порожнє значення)
	FileWrite($WhereDig & "\Library\SelLg.dsf", "@Map.BeginUpdate" & @CRLF & "@Map.DeselectAll" & @CRLF & $ESADSd1 & @CRLF & "@ExecuteMenu ViewShowSelected"  & @CRLF & "@Map.EndUpdate")
	If FileExists($WhereDig & "\Library\SelLs.dsf") Then FileDelete($WhereDig & "\Library\SelLs.dsf")
	FileWrite($WhereDig & "\Library\SelLs.dsf", "@Map.BeginUpdate" & @CRLF & "@Map.DeselectAll" & @CRLF & $ESADSd2 & @CRLF & "@ExecuteMenu ViewShowSelected"  & @CRLF & "@Map.EndUpdate")
	If FileExists($WhereDig & "\Library\SelKw.dsf") Then FileDelete($WhereDig & "\Library\SelKw.dsf")
	FileWrite($WhereDig & "\Library\SelKw.dsf", "@Map.BeginUpdate" & @CRLF & "@Map.DeselectAll" & @CRLF & $ESADSd3 & @CRLF & "@ExecuteMenu ViewShowSelected"  & @CRLF & "@Map.EndUpdate")
	If FileExists($WhereDig & "\Library\SelVd.dsf") Then FileDelete($WhereDig & "\Library\SelVd.dsf")
	FileWrite($WhereDig & "\Library\SelVd.dsf", "@Map.BeginUpdate" & @CRLF & "@Map.DeselectAll" & @CRLF & $ESADSd4 & @CRLF & "@ExecuteMenu ViewShowSelected" & @CRLF & "@Map.EndUpdate")
	
; 	Sleep(500)
	
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів",2,75)
	
	$FlagEstOshib=0
	_PoznPustParam1()
	_PoznPustParam2()
	_PoznPustParam3()
	_PoznPustParam4() 
	If $FlagEstOshib>0 Then
		_ViknProgess($GuiPSP,$ModName,$ZagPrBarWin,$PidZagPrBarWin,3)
	Else
		_ViknProgess($GuiPSP,$ModName,$ZagPrBarWin,$PidZagPrBarWin,4)
	EndIf
	
; 	Sleep(500)
	
; 	_ViknProgess($GuiPSP, $ModName, "Перевірка наявновсті незаповнених параметрів",2,100)
; 	_ViknProgess($GuiPSP,$ModName, "Перевірка наявновсті незаповнених параметрів" ,3)
EndFunc





Func _PoznPustParam1()
	;налаштування звуку
    _ViknProgess($GuiPSP,$ModName,$ZagPrBarWin,$PidZagPrBarWin,1,3,1,1 )
   _ViknProgess($GuiPSP,$ModName,"Перевірка параметру коду лісгоспу...",$ZagPrBarWin,2,35,20,20,16) 
	$urov = _GetMasterVolumeLevelScalar()
	_SetMasterVolumeLevelScalar(0)
	;виконання скриптa SelLg.dsf
	_ComDigitals("@ExecuteScript SelLg.dsf", $NomPort)
	;зчитування із статус-бару Дігіталс, де відображається кількість позначених об'єктів\та їх загальна кількість картіпершо  так ставка Лугова елементи якості і кількості об'єкті
	 WinSetState($ddssf,"",@SW_MAXIMIZE)
	 _StatBarDig(1)
	If GUICtrlRead($InKilLg) > 0 Then
		$FlagEstOshib=1
		_AddToGrooup("Lg_TKSNS", 1)
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\ControlPar", "ParLg", "REG_SZ", GUICtrlRead($InKilLg))
	Else
		
		_AddToGrooup("Lg_TKSNS", 2)
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\ControlPar", "ParLg", "REG_SZ", 0)
	EndIf
	$CountParNot += GUICtrlRead($InKilLg)
	_SetMasterVolumeLevelScalar($urov)
   _ViknProgess($GuiPSP,$ModName,"Перевірка параметру коду лісгоспу...",$ZagPrBarWin,4) 
EndFunc





Func _PoznPustParam2()
   _ViknProgess($GuiPSP,$ModName,"Перевірка параметру коду лісництва...",$ZagPrBarWin,2,75,20,20,16) 
	$urov = _GetMasterVolumeLevelScalar()
	_SetMasterVolumeLevelScalar(0)
	_ComDigitals("@ExecuteScript SelLS.dsf", $NomPort)
	 WinSetState($ddssf,"",@SW_MAXIMIZE)
	_StatBarDig(2)
	If GUICtrlRead($InKilLs) > 0 Then
		$FlagEstOshib=1
		_AddToGrooup("Ls_TKSNS", 1)
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\ControlPar", "ParLs", "REG_SZ", GUICtrlRead($InKilLg))
	Else
		_AddToGrooup("Ls_TKSNS", 2)
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\ControlPar", "ParLs", "REG_SZ", 0)
	EndIf
	$CountParNot += GUICtrlRead($InKilLs)
	_SetMasterVolumeLevelScalar($urov)
   _ViknProgess($GuiPSP,$ModName,"Перевірка параметру коду лісництва...",$ZagPrBarWin,4) 
EndFunc







Func _PoznPustParam3()
   _ViknProgess($GuiPSP,$ModName,"Перевірка параметру номеру кварталу...",$ZagPrBarWin,2,95,20,20,16) 
	$urov = _GetMasterVolumeLevelScalar()
	_SetMasterVolumeLevelScalar(0)
	_ComDigitals("@ExecuteScript SelKw.dsf", $NomPort)
	 WinSetState($ddssf,"",@SW_MAXIMIZE)
	_StatBarDig(3)
	If GUICtrlRead($InKilKw) > 0 Then
		$FlagEstOshib=1
		_AddToGrooup("Kw_TKSNS", 1)
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\ControlPar", "ParKw", "REG_SZ", GUICtrlRead($InKilKw))
		
	Else
		_AddToGrooup("Kw_TKSNS", 2)
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\ControlPar", "ParKw", "REG_SZ", 0)
	EndIf
	$CountParNot += GUICtrlRead($InKilKw)
	_SetMasterVolumeLevelScalar($urov)
   _ViknProgess($GuiPSP,$ModName,"Перевірка параметру номеру кварталу...",$ZagPrBarWin,4) 
EndFunc






Func _PoznPustParam4()
   _ViknProgess($GuiPSP,$ModName,"Перевірка параметру номеру виделу...",$ZagPrBarWin,2,100,20,20,16) 
	$urov = _GetMasterVolumeLevelScalar()
	_SetMasterVolumeLevelScalar(0)
	_ComDigitals("@ExecuteScript SelVd.dsf", $NomPort)
	 WinSetState($ddssf,"",@SW_MAXIMIZE)
	_StatBarDig(4)
	If GUICtrlRead($InKilVd) > 0 Then
		$FlagEstOshib=1
		_AddToGrooup("Vd_TKSNS", 1)
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\ControlPar", "ParVd", "REG_SZ", GUICtrlRead($InKilVd))
	Else
		_AddToGrooup("Vd_TKSNS", 2)
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdPSP & "\ControlPar", "ParVd", "REG_SZ", 0)
	EndIf
	$CountParNot += GUICtrlRead($InKilVd)
	_SetMasterVolumeLevelScalar($urov)
   _ViknProgess($GuiPSP,$ModName,"Перевірка параметру номеру виделу...",$ZagPrBarWin,4) 
;    _ViknProgess($GuiPSP,$ModName,"Перевірка параметру номеру виделу...",$ZagPrBarWin,3 )
EndFunc





Func _StatBarDig()
	$NemFunc = "_StatBarDig" 
	$WiknDig = WinActivate("[CLASS:TFormGed;Title:Digitals XE]", "")
	$Stabar = ControlGetHandle($WiknDig, "", "TStatusBar1")
	$sdr = _GUICtrlStatusBar_GetText($Stabar, 0)
	If @Error Then
		MsgBox(16 + 4096, $ZagTextPomZKS, "Не вдалося повернути значення кількості об'єктів карти")
		_ObrErr($UnIdZKS, $numWin, $NemFunc,@AutoItPID)
	_LogReg($UnIdZKS, $numWin, $NemFunc, "Знач кільк об'єкт карти=ERROR")
		Exit(8000)
	Else  
	_LogReg($UnIdZKS, $numWin, $NemFunc, "Знач кільк об'єкт карти=OК")
		Return StringMid($sdr,1,StringInStr($sdr,"/")-1)
	EndIf
	$NemFunc = ""
EndFunc






Func _ExpFromSql()
	$NamFunc = "_ExpFromSql"
	$NameKatProj = @MyDocumentsDir & "\TKSNS\Project\" & $UnIdZZP
	$papTempl = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\InFilesPutch", "PutchDigit") & "\Library"
	$PerRjad1 = "@Map.DeselectAll" _
			& @CRLF & "@Map.BeginUpdate"
	$PerRjad3 = "@ExecuteMenu ShowSelected1"
	$Retffs = ""
	$JakyTabl = ""
	$JakyTabl = _SqlZap($PokazTablPom, "Не вдалося отримати список назв таблиць з помилками", "TKSNS", 2)
	If $JakyTabl = "" Then
		MsgBox(16 + 4096, $ZagTextPomZZP, " ")
		_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
		Exit(9999)
	EndIf
	_ArrayDelete($JakyTabl, 0)
	$SqlZape1 = " /C cd /d " & $papTempl & " & bcp ""SELECT  [ComDigg] FROM TKSNS.dbo."
	$SqlZape2 = " queryout "
	$SqlZape3 = " -c -T"
	For $iqw = 0 To UBound($JakyTabl) - 1
		$NameScriptDsf = $JakyTabl[0][$iqw] & ".dsf"
		$putchToScr = $papTempl & "\" & $NameScriptDsf
		$PovnZapr = $SqlZape1 & $JakyTabl[0][$iqw] & '"' & $SqlZape2 & $NameScriptDsf & $SqlZape3
		_ZapComCmd($PovnZapr)
		$PerRjad2 = "@Map.EndUpdate" _
				& @CRLF & "@Map.Selected.AddToGroup " & $JakyTabl[0][$iqw] & @CRLF & $PerRjad3
		_FileWriteToLine($putchToScr, 1, $PerRjad1)
		FileWrite($putchToScr, @CRLF & $PerRjad2)
		$digComm = "@ExecuteScript " & $JakyTabl[0][$iqw]
		$commOtt = ""
		$commOtt = _ComDigitals($digComm, $TcpPort)
		If $commOtt = "" Then
			MsgBox(16 + 4096, $ZagTextPomZZP, " ")
			_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
			Exit(9999)
		EndIf
	Next
	$NamFunc = ""
EndFunc





Func _ProvExDig()
	$NemFunc = "_ProvExDig"
	If Not ProcessExists("ged.exe") Then
		MsgBox(16 + 4096, $ZagTextPomZZP, "Не знайдено поцесу ged.exe!")
		_ObrErr($UnIdZZP, $numWin, $NemFunc, @AutoItPID)
		Exit(9999)
	EndIf
	$NemFunc = ""
EndFunc






Func _StvorGroupInDig($TablWith)
	$NamFunc = "_StvorGroupInDig"
	$PutchToDig = RegRead($VetkaPutSess & "\InFilesPutch", "PutchDigit")
	$PaplaLybr = $PutchToDig & "\Library\TKSNS_Group" & $TablWith & ".dsf"
	If FileExists($PaplaLybr) Then
		FileDelete($PaplaLybr)
		_FileCreate($PaplaLybr)
	EndIf
	$Shabl = "@Map.DeselectAll" & @CRLF
	$Shabl1 = "$uncod=@UTF8ToString "
	$Shabl2 = "@Map.Selected.AddToGroup $uncod" & @CRLF & "@ExecuteMenu ShowSelected1"
	$ZapInBd = "Select ComDigg FROM " & $TablWith
	$OtvetNaZapr = _SqlZap($ZapInBd, "Не вдалося оримати рядки для створення групи", "TKSNS", 2)
	If $OtvetNaZapr = 1 Then
		_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
		Exit(999)
	Else
		_ArrayDelete($OtvetNaZapr, 0)
		$deskrFileScr = FileOpen($PaplaLybr, 2)
		FileWrite($deskrFileScr, $Shabl)
		For $issd = 0 To UBound($OtvetNaZapr) - 1
			FileWrite($deskrFileScr, $OtvetNaZapr[$issd][0] & @CRLF)
		Next
		FileWrite($deskrFileScr, $Shabl1 & $TablWith & @CRLF & $Shabl2)
		FileClose($deskrFileScr)
	EndIf
	WinActivate("[CLASS:TFormGed;Title:Digitals XE]", "")
	$NameScriptFile = _FO_PathSplit($PaplaLybr)
	$FormCommDig = "@ExecuteScript " & $NameScriptFile[1]
	$RobWithDig = _ComDigitals($FormCommDig, $TcpPort)
	If $RobWithDig = "$RESULT" Then
		MsgBox(16 + 4096, $ZagTextPomZZP, "Продовження роботи неможливе, тому що в обраній карті, трапляються помилки обєктів. Для скорішого їх  знаходження та усунення скористайтесь створеною групою, що має префікс 'TKSNS'...")
		_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
		Exit(9999)
	EndIf
	$NamFunc = ""
EndFunc






Func _PoznPomilLg()
	$NamFunc = "_PoznPomilLg"
	Local $sTexStrokaLg
	_GUICtrlComboBox_GetLBText($ZnaLg, _GUICtrlComboBox_GetCurSel($ZnaLg), $sTexStrokaLg)
	$FormCommDigSelectLg = _StvorSdfFile("PoznPomStrLg.dsf", '\Library\', "@Map.BeginUpdate" & @CRLF & "@Map.FindByParameters 2|9=" & $sTexStrokaLg & "|0>0" & @CRLF & "@ExecuteMenu ViewShowSelected" & @CRLF & "@Map.EndUpdate", $UnIdPPE, $numWin, $ZagTextPomPPE)
	$RozbShlyaLg = _FO_PathSplit($FormCommDigSelectLg)
	$texttcommmLg = "@ExecuteScript " & $RozbShlyaLg[1]
	$GedLg = WinActivate("[CLASS:TFormGed;Title:Digitals XE]", "")
	$vidpVidScrLg = _ComDigitals($texttcommmLg, $NumTcpPort)
	WinActivate($GuiPPE)
	$NamFunc = ""
EndFunc




Func _Perv($SLooo)
	$NamFunc = "_Perv"
	$GedWin = WinActivate("[CLASS:TFormGed;Title:Digitals XE]", "")
	If Not $GedWin Then
		MsgBox(16 + 4096, $ZagTextPomPPE, "Не знайдено вікна Digitals")
		_ObrErr($UnIdPPE, $numWin, $NamFunc, @AutoItPID) 
		Exit(9991)
	Else
		GUICtrlSetData($InpGedWin, $GedWin)
		$PutchOut = $PutchDig & "\Exchange\Config.ini"
		$PunchConf = $PutchDig & "\Exchange\Output.txt"
		Switch $SLooo
			Case 1
				$NazzvSll = "MejLisn"
				$Parammms = "0,9,10"
			Case 2
				$NazzvSll = "KvPros"
				$Parammms = "0,12"
		EndSwitch
		$ZapFil = _ZapExchIni($PutchOut, $PunchConf, $NazzvSll, $Parammms)
	EndIf
	$NamFunc = ""
	Return $ZapFil
EndFunc
Func _Perv($SLooo)
	$NamFunc = "_Perv"
	$GedWin = WinActivate("[CLASS:TFormGed;Title:Digitals XE]", "")
	If Not $GedWin Then
		MsgBox(16 + 4096, $ZagTextPomPPE, "Не знайдено вікна Digitals")
		_ObrErr($UnIdPPE, $numWin, $NamFunc, @AutoItPID) 
		Exit(9991)
	Else
		GUICtrlSetData($InpGedWin, $GedWin)
		$PutchOut = $PutchDig & "\Exchange\Config.ini"
		$PunchConf = $PutchDig & "\Exchange\Output.txt"
		Switch $SLooo
			Case 1
				$NazzvSll = "MejLisn"
				$Parammms = "0,9,10"
			Case 2
				$NazzvSll = "KvPros"
				$Parammms = "0,12"
		EndSwitch
		$ZapFil = _ZapExchIni($PutchOut, $PunchConf, $NazzvSll, $Parammms)
	EndIf
	$NamFunc = ""
	Return $ZapFil
EndFunc




Func _ZapExchDig($aSvorConfigFile)
	$NamFunc = "_ZapExchDig"
	$ImmaaKnooop = $aSvorConfigFile[1]
	$SloyDig = $aSvorConfigFile[2]
	$CommDigPom = "@Map.FindByParameters 2|-7=" & StringStripWS($SloyDig, 3) & "|0>0"
	$StvorSdfDig = _StvorSdfFile("TKSNS_PoznMejiLisn.dsf", "\Library\", $CommDigPom, $UnIdPPE, $numWin, $ZagTextPomPPE)
	If $StvorSdfDig = "" Then
		MsgBox(16 + 4096, $ZagTextPomPPE, "Не вдається створити файл скрипта, що помітить межі лісництва")
		_ObrErr($UnIdPPE, $numWin, $NamFunc, @AutoItPID)
		SetError(1)
		Return 'ERROR'
	EndIf
	$PutchIma = _FO_PathSplit($StvorSdfDig)
	$Vipoln = "@ExecuteScript " & $PutchIma[1]
	$OtvServDig = _ComDigitals($Vipoln, $NumTcpPort)
	$ProchStaBar = _GUICtrlStatusBar_GetText(ControlGetHandle(WinActivate("[CLASS:TFormGed;Title:Digitals XE]", ""), "", "TStatusBar1"), 0)
	$kilPozObj = StringMid($ProchStaBar, StringInStr($ProchStaBar, "/") + 1, -1)
	$NazvSoooyy = ""
	Switch $ImmaaKnooop
		Case "MejLisn" 
			GUICtrlSetData($InpKilPozMejLs, $kilPozObj)
			RegWrite($VetkaReest, "KilPoznObjMejaLisn", "REG_SZ", $kilPozObj)
			$NazvSoooyy = " 'межі лісництва' "
		Case "KvPros" 
			$NazvSoooyy = " 'квартальні просіки' "
	EndSwitch
	If $kilPozObj = 0 Then
		MsgBox(16 + 4096, $ZagTextPomPPE, "Не вдається запуск скрипта помітки шару" & $NazvSoooyy & ", так як - відсутні виділені обєкти")
		_ObrErr($UnIdPPE, $numWin, $NamFunc, @AutoItPID)
		SetError(2)
		Return 'ERROR'
	EndIf
	_ComDigitals("@ExecuteMenu ToolsExchange", $NumTcpPort)
	$NamFunc = ""
	_ArrayAdd($aSvorConfigFile, $kilPozObj)
	Return $aSvorConfigFile
EndFunc





Func _ObrobOtrDan($aMassPoperFunc, $mode)
	$NamFunc = ""
	$ExchFileNam = $aMassPoperFunc[3]
	$ExchFilePutch = $PutchDig & '\Exchange\' & $ExchFileNam
	Do
		Sleep(20)
	Until FileExists($ExchFilePutch)
	Switch StringStripWS($aMassPoperFunc[1], 3)
		Case "MejLisn"
			$sProchExchDig = _FileCountLines($ExchFilePutch)
			GUICtrlSetData($InpKilPozMejLs, $sProchExchDig)
			RegWrite($VetkaReest, "$InpKilPozMejLs", "REG_SZ", $sProchExchDig)
			$SumPloLisn = 0.0
			Local $aPomLg[0]
			Local $aPomLs[0]
			$ZnLG = ""
			$ZnLisn = ""
			For $dsa = 1 To $sProchExchDig
				$strr = FileReadLine($ExchFilePutch, $dsa)
				$Tab1 = StringReplace(StringMid($strr, 1, StringInStr($strr, @TAB) - 1), ',', '.')
				$SumPloLisn += $Tab1
				$Tab2 = StringMid($strr, StringInStr($strr, @TAB) + 1, -1)
				$TLg = StringMid($Tab2, 1, StringInStr($Tab2, @TAB) - 1)
				$TLs = StringMid($Tab2, StringInStr($Tab2, @TAB) + 1, -1)
				_ArrayAdd($aPomLg, $TLg)
				_ArrayAdd($aPomLs, $TLs)
			Next
			Switch $mode
				Case 1 ; Площадь меж лесничества
					GUICtrlSetData($InpPloLs, $SumPloLisn)
					RegWrite($VetkaReest, "PloMapMejaLisn", "REG_SZ", $SumPloLisn)
					$NamFunc = ""
					Return $SumPloLisn
				Case 2 ; Помилки параметра коду лісхозу
					$aPomLgUniq = _ArrayUnique($aPomLg)
					$NamFunc = ""
					Return $aPomLgUniq
				Case 3 ; Помилки параметра коду лісництва
					$aPomLsUniq = _ArrayUnique($aPomLs)
					$NamFunc = ""
					Return $aPomLsUniq
			EndSwitch
		Case "KvPros"
		Case 1 
			MsgBox(64+4096, 'Info',"Plo")
	EndSwitch
	$NamFunc = ""
EndFunc






 Func _StvorSdfFile($NamSdfFile, $PutchToTempl, $InfileTextCode, $UniqSess, $nammmmWin, $TextError)
	$NamFunc = "_StvorSdfFile"
	$KatWhereDigit = RegRead("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UniqSess & "\InFilesPutch", "PutchDigit")
	If @Error Then
		MsgBox(16 + 4096, $TextError, "Не вдається прочитати шлях до каталогу Digitals із параметра рєєстру")
		_ObrErr($UniqSess, $nammmmWin, $NamFunc, @AutoItPID)
		Return 0
	EndIf
	$PovnPutchtoFilSdf = $KatWhereDigit & $PutchToTempl & $NamSdfFile
	If FileExists($PovnPutchtoFilSdf) Then
		FileDelete($PovnPutchtoFilSdf)
		_FileCreate($PovnPutchtoFilSdf)
	EndIf
	$DesFilSdf = FileOpen($PovnPutchtoFilSdf, 2)
	If $DesFilSdf = -1 Then
		MsgBox(16 + 4096, $TextError, "Не вдається відкрити sdf-файл скрипта")
		_ObrErr($UniqSess, $nammmmWin, $NamFunc, @AutoItPID)
		Return 0
	Else
		If $nammmmWin = 7 Then
			FileWrite($DesFilSdf, $MaketDsfFile1 & @CRLF)
			FileWrite($DesFilSdf, $InfileTextCode & @CRLF)
			FileWrite($DesFilSdf, $MaketDsfFile2)
		Else
			FileWrite($DesFilSdf, $InfileTextCode & @CRLF)
; 			FileWrite($DesFilSdf, $MaketDsfFile2)
		EndIf
		FileClose($DesFilSdf)
	EndIf
	$NamFunc = ""
	Return $PovnPutchtoFilSdf
EndFunc







Func _ProvConfirm()
	$NamFunc = "_ProvConfirm"
	$wiwssdx = WinActivate("[Title:Confirm;Class:#32770]", "")
	$oParent = _UIA_GetElementFromHandle($wiwssdx)
	$oElement = _UIA_GetControlTypeElement($oParent, "UIA_TextControlTypeId", "Перезаписать базу данных" & @CRLF & "TKSNS на " & @ComputerName & " (Microsoft SQL Server) dbo?")
	If Not IsObj($oElement) Then
		$oElement = _UIA_GetControlTypeElement($oParent, "UIA_TextControlTypeId", "Користуємось реляційною базою?")
		If IsObj($oElement) Then
			WinClose($wiwssdx, "")
		EndIf
	Else
		Send("{ENTER}")
		$new = WinWait("[Title:Saving to SQL-server;Class:TFormProgress]", "", 5)
		If Not $new Then
			MsgBox(16 + 4096, "Помилка команди Digitals", "Не знайдено вікно збереження карти!")
			Exit(999)
		Else
			WinSetOnTop($new, "", 1)
			WinSetState($new, "", @SW_DISABLE)
; 			GUICtrlSetData($InpWinSaveProgZKS, "Вікно_збер_карти=ОК")
; 			ToolTip( 'изменение иконки', Default, Default, 'Статистика', 1, 4)
		EndIf
	EndIf
; 		Return 2
	$NamFunc = ""
EndFunc





$gedd = Run($sdew)
; 		_ZapLogInp("Процес Ged.exe успішно запущений, дескриптор: " & $gedd, @ScriptDir & "\Log_TKSNS.log",1,@ScriptName)
; 		$hWnd = WinWait ("[CLASS:TFormGed;Title:Digitals XE]", "")
; 		_ZapLogInp("Вікно Digitals відкрито, дескриптор: " & $hWnd, @ScriptDir & "\Log_TKSNS.log",1,@ScriptName)
; 		$Nazv = WinGetTitle($hWnd)
; 		_ZapLogInp("Назва вікна Digitals: " & $Nazv, @ScriptDir & "\Log_TKSNS.log",1,@ScriptName)
; 		$demo=StringRegExp($Nazv, 'Demo')
; 		if $demo Then
; 		_ZapLogInp("Перевірка на Демо-версію Digitals" , @ScriptDir & "\Log_TKSNS.log",1,@ScriptName)
; 		MsgBox(16+4096,"ТаксІнс©      Помилка. Demo-режим!", "                  Робота програми ТаксІнс - "&@CRLF & _
; 		 "може виконуватися тільки на ліцензованій копії Digitals "&@CRLF & _
; 		  "              Вставте ключ захисту і спробуйте ще раз..."  	)
; 		ProcessClose($gedd)
; 		_ZapLogInp("Digitals в Демо-режимі" , @ScriptDir & "\Log_TKSNS.log",3,@ScriptName)
; 		EndIf
; 	Exit
; EndIf

; EndFunc







Func _asza()
	If WinExists("[Title:Аналіз карти;Class:TFormProgress]", "") Then
		$hWnd = WinActivate("[Title:Аналіз карти;Class:TFormProgress]", "")
		$knop = ControlClick($hWnd, "", "[Class:TBitBtn]")
	Else
		Return SetError(99)
	EndIf
EndFunc






Func _PoiskDigit($mode = 0)
	Switch $mode
		Case 0
			Local $JKMBTUJ = '\Ged.exe'
			$aArraydisk = DriveGetDrive("fixed")
			$tyui1 = ""
			For $i = 1 To $aArraydisk[0]
				$aFoldDig = _FO_FolderSearch($aArraydisk[$i], "Digitals", True, 1)
				For $r = 0 To UBound($aFoldDig) - 1
					
					If Not FileExists($aFoldDig[$r] & $JKMBTUJ) Then
						ContinueLoop
					EndIf
					_GUICtrlListView_AddItem($Ged, $aFoldDig[$r] & $JKMBTUJ)
				Next
			Next
			$app = StringRegExp($tyui1, '(.*\Digitals)', 3)
			Dim $ertfd[0]
			Local $sfs = ""
			If UBound($app) - 1 = 0 Then
				GUICtrlSetData($Ged, $app[0] & $JKMBTUJ & "\ged.exe")
				Return $app[0] & $JKMBTUJ
			EndIf
			For $k = 0 To UBound($app) - 1
				$sfs &= $app[$K] & $JKMBTUJ & "|"
			Next
			_GUICtrlListView_ClickItem($Ged, 0)
			Return $sfs
		Case 1
; 	Перший попавшийся папка ДИГИ
			Local $JKMBTUJ = '\Ged.exe'
			$aArraydisk = DriveGetDrive("fixed")
			$tyui1 = ""
			For $i = 1 To $aArraydisk[0]
				$aFoldDig = _FO_FolderSearch($aArraydisk[$i], "Digitals", True, 1)
				For $r = 0 To UBound($aFoldDig) - 1
					If Not FileExists($aFoldDig[$r] & $JKMBTUJ) Then
						ContinueLoop
					EndIf
					Return $aFoldDig[$r]
				Next
			Next
		Case 2
; 			Существует ли процес "ged.exe"
			Local $ssadf
			Local $iPid, $paramm[0]
			$iPid = ProcessExists("ged.exe")
			If $iPid Then
				Return _ProcessGetPath($iPid)
			Else
				Local $JKMBTUJ = '\Ged.exe'
				$aArraydisk = DriveGetDrive("fixed")
				$tyui1 = ""
				For $i = 1 To $aArraydisk[0]
					$aFoldDig = _FO_FolderSearch($aArraydisk[$i], "Digitals", True, 1)
					For $r = 0 To UBound($aFoldDig) - 1
						If Not FileExists($aFoldDig[$r] & $JKMBTUJ) Then
							ContinueLoop
						EndIf
						$Starrr = TimerInit()
						$Ima = Run($aFoldDig[$r] & "\Ged.exe")
						$Nfdas =TimerDiff($Starrr) 
						$hWnd = WinWaitActive("[CLASS:TFormGed]", "",40) ; без таймаута (5) ожидание бесконечно  
						$rfhj=  TimerDiff($Starrr)-$Nfdas  &" разн сек"
						Return _ProcessGetPath($Ima) 
					Next
				Next
				
			EndIf
		Case 3
			
			$Starrr = TimerInit()
			$NewProcc = Run(_PoiskDigit(1) & "\Ged.exe")
			
			ConsoleWrite('result=' & $NewProcc & @CRLF)
			MsgBox(64 + 4096, 'Info', 'good Time:' & TimerDiff($Starrr))
			
; 		Ждать появления Окна которое е имеет ПІД - параметр
			
	EndSwitch
	
EndFunc






Func _ProcessGetPath($vProcess)
	Local $iPID = ProcessExists($vProcess)
	If Not $iPID Then Return SetError(1, 0, -1)
	
	Local $aProc = DllCall('kernel32.dll', 'hwnd', 'OpenProcess', 'int', BitOR(0x0400, 0x0010), 'int', 0, 'int', $iPID)
	If Not IsArray($aProc) Or Not $aProc[0] Then Return SetError(2, 0, -1)
	
	Local $vStruct = DllStructCreate('int[1024]')
	
	Local $hPsapi_Dll = DllOpen('Psapi.dll')
	If $hPsapi_Dll = -1 Then $hPsapi_Dll = DllOpen(@SystemDir & '\Psapi.dll')
	If $hPsapi_Dll = -1 Then $hPsapi_Dll = DllOpen(@WindowsDir & '\Psapi.dll')
	If $hPsapi_Dll = -1 Then Return SetError(3, 0, '')
	
	DllCall($hPsapi_Dll, 'int', 'EnumProcessModules', _
			'hwnd', $aProc[0], _
			'ptr', DllStructGetPtr($vStruct), _
			'int', DllStructGetSize($vStruct), _
			'int_ptr', 0)
	Local $aRet = DllCall($hPsapi_Dll, 'int', 'GetModuleFileNameEx', _
			'hwnd', $aProc[0], _
			'int', DllStructGetData($vStruct, 1), _
			'str', '', _
			'int', 2048)
	
	DllClose($hPsapi_Dll)
	
	If Not IsArray($aRet) Or StringLen($aRet[3]) = 0 Then Return SetError(4, 0, '')
	Return $aRet[3]
EndFunc







Func _ParsTextKart($Nbd, $Pk)
	$zaoed = "SELECT cast(right(o.NAME,LEN(o.NAME)-1 ) AS INT)" & @CRLF & _
			"FROM  [sys].[all_columns] i" & @CRLF & _
			"JOIN sysobjects o " & @CRLF & _
			"ON o.id=i.object_id " & @CRLF & _
			"JOIN dbo._Layers AS t " & @CRLF & _
			"ON cast(right(o.NAME,LEN(o.NAME)-1 ) AS INT) =t.layerid  " & @CRLF & _
			"WHERE i.NAME IN ('p013')"
	$dfc = _SqlZap($zaoed, 'Не вдалося отримати таблицю видільних шарів карти', $Nbd, 2) ;_GUICtrlStatusBar_GetText($SB, 0)
	_ArrayDelete($dfc, 0)
	Dim $shab[0]
	For $i = 0 To UBound($dfc) - 1
		_ArrayAdd($shab, "//Layer " & $dfc[$i][0])
	Next
; 	_ArrayDisplay($shab,"Массив")
	$massfilekart = StringSplit(FileRead(StringLeft($Pk, StringLen($Pk) - 4) & ".asc"), @CR) ; Створення масиву із прочитаного вхідного файлу
; 	 _ArrayDisplay($massfilekart,"Массив")
	FileDelete($wertghuj) ; Видалення файлу параметрів об'єктів
	_FileCreate($wertghuj) ; Створення файлу параметрів об'єктів
	Dim $Schetchik[0] ; Оголошення масиву лічильника
	For $imassfilekart = 0 To UBound($massfilekart) - 1 ; Перебор массива із прочитаного вхідного файлу
		$postdr = StringReplace($massfilekart[$imassfilekart], @LF, '')
		If $postdr = '#' Then _ArrayAdd($Schetchik, $imassfilekart) ; Якщо значення рядка є шарп
	Next
	For $iSchetchik = 0 To UBound($Schetchik) - 1
		_ArraySearch($shab, StringMid($massfilekart[$Schetchik[$iSchetchik] + 1], 2, 16))
		If _ArraySearch($shab, StringMid($massfilekart[$Schetchik[$iSchetchik] + 1], 2, 16)) <> -1 Then
			$nach = $Schetchik[$iSchetchik]
			$kon = $Schetchik[$iSchetchik + 1]
			Local $strvid = ''
			Local $Sid = ''
			Local $Oid = ''
			Local $Clg = ''
			Local $Cls = ''
			Local $Ckv = ''
			Local $Cvd = ''
			For $ikon = $nach To $kon
				$postdr1 = StringReplace($massfilekart[$ikon], @LF, '')
				If StringInStr($postdr1, "//Layer ") <> 0 Then
					$Sid = StringMid($postdr1, 9, 8)
					$strvid &= $Sid & ";"
				EndIf
				
				If StringInStr($postdr1, "//ObjectID ") <> 0 Then
					$Oid = StringMid($postdr1, 12, StringLen($postdr1))
					$strvid &= $Oid & ";"
				EndIf
				If StringInStr($postdr1, "//P[21] ") <> 0 Then
					$Clg = StringMid($postdr1, 9, StringLen($postdr1))
					$strvid &= $Clg & ";"
				EndIf
				If StringInStr($postdr1, "//P[22] ") <> 0 Then
					$Cls = StringMid($postdr1, 9, StringLen($postdr1))
					$strvid &= $Cls & ";"
				EndIf
				If StringInStr($postdr1, "//P[24] ") <> 0 Then
					$Ckv = StringMid($postdr1, 9, StringLen($postdr1))
					$strvid &= $Ckv & ";"
				EndIf
				If StringInStr($postdr1, "//P[13] ") <> 0 Then
					$Cvd = StringMid($postdr1, 9, StringLen($postdr1))
					$strvid &= $Cvd & ";"
				EndIf
				
			Next
			FileWrite($wertghuj, $Sid & ";" & $Oid & ";" & $Clg & ";" & $Cls & ";" & $Ckv & ";" & $Cvd & @CRLF)
		Else
			ContinueLoop
		EndIf
	Next
	Return 'good'
EndFunc

; $aStrPapok[3] = @ScriptDir & "\LOG"
; $aStrPapok[4] = @ScriptDir & "\FORM"
; $aStrPapok[5] = @ScriptDir & "\CONFIG"
; $aStrPapok[6] = @ScriptDir & "\HDBK"
; $aStrPapok[7] = @ScriptDir & "\RESURSE"

#cs

;   If $KopMdf = 0 Then
; 	_ZapLogInp("Помилка при копіювання файлу TKSNS.mdf", @ScriptDir & "\Log_TKSNS.log", 1, @ScriptName)
; Else
; 	_ZapLogInp("Файл TKSNS.mdf успішно скопійовано", @ScriptDir & "\Log_TKSNS.log", 1, @ScriptName)
; EndIf

;
; If $KopLdf = 0 Then
; 	_ZapLogInp("Помилка при копіювання файлу TKSNS_log.ldf", @ScriptDir & "\Log_TKSNS.log", 1, @ScriptName)
; Else
; 	_ZapLogInp("Файл  TKSNS_log.ldf успішно скопійовано", @ScriptDir & "\Log_TKSNS.log", 1, @ScriptName)
; EndIf

#ce

#Region old_func



Func _Start($commm, $portt) ; Р¤СѓРЅРєС†С–СЏ Р·Р°РїСѓСЃРєСѓ
	If Not ProcessExists("ged.exe") Then
		MsgBox(16 + 4096 + 524288, 'РџРѕРјРёР»РєР° С„СѓРЅРєС†С–С—', "РќРµ Р·РЅР°Р№РґРµРЅРѕ РїСЂРѕС†РµСЃСѓ Ged.exe")
		Return SetError(99, 0, 'no_proc_ged')
	EndIf
	$Server = TCPConnect(@IPAddress1, $portt) ;РЎС‚РІРѕСЂРµРЅРЅСЏ Socket Is IP Р°РґСЂРµСЃРѕРј С‚Р° РїРѕСЂС‚РѕРј
	If $Server = -1 Then ; РЇРєС‰Рѕ Р—РјС–РЅР° СЃРѕРєРµС‚Сѓ Р”РѕСЂС–РІРЅСЋС” -1 TKSNS_2018037121444
		;  		ConsoleWrite('Err=' & @Error & @CRLF)
  		Sleep(100) ;РЎРїР°С‚Рё 100 РјС–Р»С–СЃРµРєСѓРЅРґ
		MsgBox(16, 'Р¤Р°С‚Р°Р»СЊРЅР° РїРѕРјРёР»РєР°', 'РќРµРјРѕР¶Р»РёРІРѕ РїС–РґРєР»СЋС‡РёС‚РёСЃСЏ РґРѕ СЃРµСЂРІРµСЂР°. Р—РјС–РЅС–С‚СЊ СЃРІРѕС— РЅР°Р»Р°С€С‚СѓРІР°РЅРЅСЏ С‚Р° РїРѕРІС‚РѕСЂС–С‚СЊ СЃРїСЂРѕР±Сѓ. ')
		Return @Error ; РџРѕРІРµСЂРЅСѓС‚Рё С„СѓРЅРєС†С–С”СЋ РњР°РєСЂРѕСЃ РџРѕРјРёР»РєРё
	EndIf
	    	Sleep(150)
	TCPSend($Server, $commm)
	
	$rtde = "Digitals Map Server is ready" & @CRLF & "Version 1.0 Beta" & @CRLF & "Send HELP to obtain command list"
	While 1
		_ProvConfirm()
; 		If =2 Then ExitLoop
; 		$OknoSaveBd = WinActivate("[Title:Confirm;Class:#32770]")
; 		$oknoUnikIdObk = WinActivate("[Title:Digitals XE;Class:#32770]")
; 		If $OknoSaveBd <> ''  Then ;Or $oknoUnikIdObk <> ''
; 			WinKill($OknoSaveBd)
; 		EndIf
		Sleep(15)
		If $Server <> -1 Then ;РЇРєС‰Рѕ СЃРµСЂРІРµСЂ РЅРµ РґРѕСЂС–РІРЅСЋС” -1 С‚РѕРґС–
			$Recv = TCPRecv($Server, 1000000) ;Р—С‡РёС‚СѓРІР°РЅРЅСЏ РІС–РґРїРѕРІС–РґС–
		EndIf
		$wsf = StringReplace($Recv, $rtde, '')
		If $wsf <> "" Then ExitLoop
	WEnd
	_Disconnect()
	Return $wsf
EndFunc





Func _ComDigitals($com, $NumPort)
	$urov = _GetMasterVolumeLevelScalar()
	_SetMasterVolumeLevelScalar(0)
	Local $vidpovizserveru
	TCPStartup() ; Р—Р°РїСѓСЃРє СЃРµСЂРІРµСЂР°
	$Server = -1
	$vidpovizserveru = _Start($com, $NumPort)
	_SetMasterVolumeLevelScalar($urov)
	Return $vidpovizserveru
EndFunc


