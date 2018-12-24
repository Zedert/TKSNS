Func _AttachBd($PutcToRbd, $NewLogNam)
	$NemFunc = "_AttachBd"
	$pidkluch = StringReplace(StringReplace($CraetBdForAtt, StringMid($CraetBdForAtt, StringInStr($CraetBdForAtt, '-----'), 5), $PutcToRbd), StringMid($CraetBdForAtt, StringInStr($CraetBdForAtt, '+++++'), 5), $NewLogNam)
	$PutcToRbddsd=_FO_PathSplit($NewLogNam)
	$pidkluch1211=StringReplace($pidkluch,"_____",$PutcToRbddsd[0]&$PutcToRbddsd[1]&"_log.ldf")	
	$ZapPidkl = _SqlZap($pidkluch1211, "��������� �������� ��'������� ��������� ����", 'master', 5) 
	If $ZapPidkl <> 0 Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ��'�� ����� ����=ERROR") 
		Exit(7783)
	EndIf
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ��'�� ����� ����=O�") 
	$NemFunc = ""
EndFunc






Func _DettachDb($LogNameBd)
	$NemFunc = "_DettachDb"
	$ZamStrToImBd = StringReplace($VidklBd, StringMid($VidklBd, StringInStr($VidklBd, "+++++"), 5), $LogNameBd)
	$VidpSqlZap = _SqlZap($ZamStrToImBd, "�� ������� �������� ��������� ��'������� ���� " & $LogNameBd, 'master', 5)
	If $VidpSqlZap <> 0 Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ��'��� ����=ERROR") 
		Exit(7782)
	EndIf
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ��'��� ����=O�") 
	$NemFunc = ""
EndFunc





Func _Videdna($NameBDD, $Filedsa)
	$NemFunc = "_Videdna"
	$StrToZamLig = StringMid($DelConToServ, StringInStr($DelConToServ, '+++++'), 5)
	$sxcds = StringReplace($DelConToServ, $StrToZamLig, $NameBDD)
	$ZapUdalPidkl = StringReplace($sxcds, "-----", $Filedsa)
	$VidalActiveConn = _SqlZap($ZapUdalPidkl, "��������� �������� ����� �� ��������� �������� �'������", 'master', 5)
	If $VidalActiveConn <> 0 Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID) 
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ����� �� ����� ����� �'���=ERROR") 
		Exit(7781)
	EndIf
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ����� �� ����� ����� �'���=O�") 
	$NemFunc = ""
EndFunc




Func _PerePidkl($SpisBDPar, $PatchProvBDf, $mode)
	$NemFunc="_PerePidkl"
	; ����������  ��'�������
	Local $aMasPiklRbd[0]
	$dvfdx = ""
	Switch $mode
		Case 1
			$dvfdx = 1
		Case 2
			$dvfdx = 0
	EndSwitch
	$IndexZnai = ""
	For $iPar = 1 To UBound($SpisBDPar) - 1
		_ArrayAdd($aMasPiklRbd, $SpisBDPar[$iPar][$dvfdx])
	Next
	$oiskBd = _ArraySearch($aMasPiklRbd, $PatchProvBDf)
	If $oiskBd = -1 Then
		$flagPodkl = -1 
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "ϳ����� �� �� ����=ERROR") 
		Return False
	Else
		$flagPodkl = 1
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "ϳ����� �� �� ����=O�") 
		Return True
	EndIf
	$NemFunc=""
EndFunc






Func _SpisPodklBD() ;$PatchProvBDf, $namRbdBaza = '', $mode = 0
	$NemFunc = "_WorkWithRbd"
	$SpisBD = _SqlZap($ZapSpisPidklBd, "�� ���� ��������� ������� ��������� ���", 'master', 2) 
	For $q = 1 To 4
		_ArrayDelete($SpisBD, _ArraySearch($SpisBD, "master"))
		_ArrayDelete($SpisBD, _ArraySearch($SpisBD, "model"))
		_ArrayDelete($SpisBD, _ArraySearch($SpisBD, "msdb"))
		_ArrayDelete($SpisBD, _ArraySearch($SpisBD, "tempdb"))
	Next
	If IsArray($SpisBD) = 0 Then
		MsgBox(16 + 4096, $ZagTextPomWIS, "�� ������� ��������� ������ ��� ����������")
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID) 
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "���������� ������ ������ �� ���� ��� �����=ERROR") 
		Exit(7779)  
	EndIf
	 _LogReg($UnIdWIS, $numWin, $NemFunc, "���������� ������ ������ �� ���� ��� �����=O�")
	Return $SpisBD
EndFunc





Func _ProcPovVerSql()
	$NemFunc = "_ProcPovVerSql"
	
	Global $PerProc = ""
	Global $PerScrVer = ""
	Global $PerSluj = ""
	
	$ProvProc = _SqlZap($ProvProcVersSql, "�� ���� ��������� �������� ��������� ���������� ���� SQL SERVER", 'master', 5)
	If $ProvProc = 1 Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
		 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ����� ������ ���� ���� SQL SERVER=ERROR")
		Exit(7771)
	EndIf
	_LogReg($UnIdWIS, $numWin, $NemFunc, "����� ����� ������ ���� ���� SQL SERVER=��")
	$RestrProc = _SqlZap($ProcVerSql, "�� ���� �������� ��������� ���������� ���� SQL SERVER", 'master', 5)
	If $RestrProc = 1 Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
		 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ���� ���� SQL SERVER=ERROR")
		Exit(7772)
	EndIf
		 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ������ ���� SQL SERVER=OK")
	$VkonProc = _SqlZap("EXEC dbo.sp_SQLBinList", "�� ���� �������� ��������� ���������� ���� SQL SERVER ", 'master', 1)
	If Not $VkonProc Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
		 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ���� ���� SQL SERVER=ERROR") 
		Exit(7773)
	Else
		 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ���� ���� SQL SERVER=O�") 
		$UdalperStrok = StringMid(StringStripWS(_StringRemoveLine($VkonProc, 'type'), 3), 1, StringInStr(StringStripWS(_StringRemoveLine($VkonProc, 'type'), 3), "|") - 1)
		$PerProc = $UdalperStrok
		GUICtrlSetData($InpVerSqlProc, $UdalperStrok)
	EndIf
	$VresionMacros = _SqlZap('SELECT @@Version', "�� ������� �������� �������� ���������, ����� ��� ������� ����� �������", 'master', 1)
	If $VresionMacros = 1 Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
		 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ���� ���� ����� �� ���� ����=ERROR")  
		Exit(7774)
	Else
		 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ���� ���� ����� �� ���� ����=O�") 
		$GodIzNazv = StringStripWS(StringMid($VresionMacros, 1, 37), 3)
		$tilkRik = StringRegExpReplace($GodIzNazv, '.*([\d]{4}).*', '\1')
		GUICtrlSetData($InpVerSqlverss, $tilkRik)
		$PerScrVer = $tilkRik
	EndIf 
	$NemFunc = ""
	Return  $PerProc & ";" & $PerScrVer  
EndFunc





Func _SqlControl()
	$NemFunc = "_SqlControl"
	$StanSlush = ""
	$VersSqlServ = ""
	$Pido = Run(@ComSpec & ' /C sc query MSSQLSERVER', "", @SW_HIDE, $STDIN_Child + $STDERR_Merged)
	ProcessWaitClose($Pido)
	Local $Data
	While 1
		$Datao = StdoutRead($Pido)
		$Rewo = _Encoding_OEM2ANSI($Datao)
		$Data &= $Rewo
		If $Rewo = "" Then ExitLoop
	WEnd
	$aSearch = StringSplit($Data, ":")
	$VersSqlServ = StringMid(StringStripWS($aSearch[2], 3), 1, StringInStr(StringStripWS($aSearch[2], 3), " ") - 1)
	$StanSlush = StringMid(StringStripWS($aSearch[4], 3), 1, 1)
	If (Not $VersSqlServ) Or (Not $StanSlush) Or ($StanSlush <> 4) Then
		MsgBox(16 + 4096, $ZagTextPomWIS, "�� ������� �������� ���������� ��� ������ MSSQL")
		_ObrErr($UnIdWIS, $numWin, $NemFunc,@AutoItPID)
		 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ��� ������ MSSQL=ERROR") 
		Exit(7778)
	EndIf
		 _LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ��� ������ MSSQL=O�") 
	$NemFunc = ""
	Return  $VersSqlServ&";"&$StanSlush
EndFunc  






Func _AttachBd($mode)
	$NemFunc = "_AttachBd"
	$IzReePutchmdf = ""
	$IzReePutchldf = ""
	$LogicNaaamBd = ""
	$CommInReg = ""
	Switch $mode
		Case 1 ; ���� TKSNS
			$MdfFilPutch = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWIS & "\TKSNS.mdf"
			$LdfRFilPutch = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWIS & "\TKSNS_log.ldf"
			If (FileExists($MdfFilPutch) = 0) Or (FileExists($LdfRFilPutch) = 0) Then
				MsgBox(16 + 4096, $ZagTextPomWIS, "�� ������� ��'����� ���� TKSNS, ����� �� �� - �� �������� ����� � ����� �������")
				_ObrErr($UnIdWIS, $numWin, $NamFunc, @AutoItPID)
				Return 0
			EndIf
			$LogicNaaamBd = "TKSNS"
			$CommInReg = "���� 'TKSNS' ������ ��'������"
		Case 2 ; ���� ����������
			$IzReePutchmdf = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\InFilesPutch", "PutchRBD")
			$IzReePutchldf = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\InFilesPutch", "PutchRBDLog")
			If ($IzReePutchmdf = "") Or ($IzReePutchldf = "") Then
				MsgBox(16 + 4096, $ZagTextPomWIS, "�� ������� ��'����� ���������� ���,���� �� - ���� ������ ����� � �����")
				_ObrErr($UnIdWIS, $numWin, $NamFunc, @AutoItPID)
				Return 0
			EndIf
			$aNazzPutchmdf = _FO_PathSplit($IzReePutchmdf)
			$aNazzPutchldf = _FO_PathSplit($IzReePutchldf)
			If StringInStr($aNazzPutchmdf[1], "_data") Then
				$LogicNaaamBd = StringMid($aNazzPutchmdf[1], 1, StringInStr($aNazzPutchmdf[1], "_data") - 1)
			Else
				$LogicNaaamBd = $aNazzPutchmdf[1]
			EndIf
			$MdfFilPutch = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWIS & "\" & $aNazzPutchmdf[1] & $aNazzPutchmdf[2]
			$LdfRFilPutch = @MyDocumentsDir & "\TKSNS\Project\TKSNS_" & $UnIdWIS & "\" & $aNazzPutchldf[1] & $aNazzPutchldf[2]
			$CommInReg = "���������� ��� ������ ��'������"
	EndSwitch
	$pidkluch = StringReplace(StringReplace($CraetBdForAtt, StringMid($CraetBdForAtt, StringInStr($CraetBdForAtt, '-----'), 5), $LogicNaaamBd), StringMid($CraetBdForAtt, StringInStr($CraetBdForAtt, '+++++'), 5), $MdfFilPutch)
	$pidkluch1211 = StringReplace($pidkluch, "_____", $LdfRFilPutch)
	$ZapPidkl = _SqlZap($pidkluch1211, "��������� �������� ��'������� ��������� ����", 'master', 5)
	If $ZapPidkl <> 0 Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWIS, $numWin, $NemFunc, "����� ��'�� ����� ����=ERROR")
		Exit(7783)
	EndIf
	_LogReg($UnIdWIS, $numWin, $NemFunc, "����� ��'�� ����� ����=O�")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdWIS & "\FilConInServ", "Attach-" & $mode, "REG_SZ", $CommInReg)
	$NemFunc = ""
	Return 1
EndFunc





Func _DettachDb($LogNameBd)
	$NemFunc = "_DettachDb"
	$ZamStrToImBd = StringReplace($VidklBd, StringMid($VidklBd, StringInStr($VidklBd, "+++++"), 5), $LogNameBd)
	$VidpSqlZap = _SqlZap($ZamStrToImBd, "�� ������� �������� ��������� ��'������� ���� " & $LogNameBd, 'master', 5)
	If $VidpSqlZap <> 0 Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ��'��� ����=ERROR")
		Return 0
; 		Exit(7782)
	EndIf
	_LogReg($UnIdWIS, $numWin, $NemFunc, "����� ������ ��'��� ����=O�")
	Return 1
	$NemFunc = ""
EndFunc




Func _Videdna($NameBDD, $Filedsa)
	$NemFunc = "_Videdna"
	$StrToZamLig = StringMid($DelConToServ, StringInStr($DelConToServ, '+++++'), 5)
	$sxcds = StringReplace($DelConToServ, $StrToZamLig, $NameBDD)
	$ZapUdalPidkl = StringReplace($sxcds, "-----", $Filedsa)
	$VidalActiveConn = _SqlZap($ZapUdalPidkl, "��������� �������� ����� �� ��������� �������� �'������", 'master', 5)
	If $VidalActiveConn <> 0 Then
		_ObrErr($UnIdWIS, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWIS, $numWin, $NemFunc, "����� ����� �� ����� ����� �'���=ERROR")
		Return 0
; 		Exit(7781)
	EndIf
	_LogReg($UnIdWIS, $numWin, $NemFunc, "����� ����� �� ����� ����� �'���=O�")
	Return 1
	$NemFunc = ""
EndFunc





Func _SpisPodklBD() ;$PatchProvBDf, $namRbdBaza = '', $mode = 0
	$NemFunc = "_WorkWithRbd"
	$SpisBD = _SqlZap($ZapSpisPidklBd, "�� ���� ��������� ������� ��������� ���", 'master', 2)
	For $q = 1 To 4
		_ArrayDelete($SpisBD, _ArraySearch($SpisBD, "master"))
		_ArrayDelete($SpisBD, _ArraySearch($SpisBD, "model"))
		_ArrayDelete($SpisBD, _ArraySearch($SpisBD, "msdb"))
		_ArrayDelete($SpisBD, _ArraySearch($SpisBD, "tempdb"))
	Next
	If IsArray($SpisBD) = 0 Then
		MsgBox(16 + 4096, $ZagTextPomWIS, "�� ������� ��������� ������ ��� ����������")
		_ObrErr($UnIdWIS, $numWin, $NemFunc, @AutoItPID)
		_LogReg($UnIdWIS, $numWin, $NemFunc, "���������� ������ ������ �� ���� ��� �����=ERROR")
		Exit(7779)
	EndIf
	_LogReg($UnIdWIS, $numWin, $NemFunc, "���������� ������ ������ �� ���� ��� �����=O�")
	Return $SpisBD
EndFunc






Func _ProvWinSaveSql()
$NamFunc="_ProvWinSaveSql"
$ZapKilobjInSql=_SqlZap("SELECT sum(Rows)   FROM   KolZapFromKart","�� ������� ��������� ������� ����� ������� �����","TKSNS",2)
If $ZapKilobjInSql = 1 Then   
  	_ObrErr($UnIdZKS, $numWin, $NamFunc,@AutoItPID)
	_LogReg($UnIdZKS, $numWin, $NamFunc, "���� ����  ���� ���� �����=ERROR")
  	Exit(9999)
Else 
	_LogReg($UnIdZKS, $numWin, $NamFunc, "���� ����  ���� ���� �����=O�")
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\TKSNS\Session\TKSNS_"&$UnIdZKS,"KilObjMapInSql","REG_SZ",$ZapKilobjInSql[1][0])
GUICtrlSetData($InpKilObjSaveZKS,$ZapKilobjInSql[1][0])	
  EndIf 
$NamFunc=""
EndFunc







Func _ZapZberProc()
	$NamFunc="_ZapZberProc"
	
$VidalLogta=_SqlZap($UdalLogTabl,"�� ������� ��������� �� ���� ��������� ������� ���� ","TKSNS",5)
If $VidalLogta=1 Then 
	_ObrErr($UnIdZZP, $numWin, $NamFunc,@AutoItPID)
	Exit(999) 
EndIf

$VkonPro1="Exec PerParametSloy_01"

$VidVikProc1=_SqlZap($VkonPro1,"�� ������� �������� ��������� �������� ���� '��� �������'","TKSNS",5)
If $VidVikProc1=1 Then 
	_ObrErr($UnIdZZP, $numWin, $NamFunc,@AutoItPID)
	Exit(999) 
EndIf

$VkonPro2="EXEC PerZnaParKwPros_02"
$VidVikProc2=_SqlZap($VkonPro2,"�� ������� �������� ��������� �������� ���� '��������� ������ �� 5�'","TKSNS",5)
If $VidVikProc2=1 Then 
	_ObrErr($UnIdZZP, $numWin, $NamFunc,@AutoItPID)
	Exit(999) 
EndIf 


$ComandPerZbeProc= _OtrimParResu()
$VidpovSertv=_SqlZap($ComandPerZbeProc,"�� ������� �������� �������� ","TKSNS",2)
ConsoleWrite('result=' & $VidpovSertv & @CRLF)
	
	
	$NamFunc=""
	EndFunc

	
	
	

Func _PerKilStrokInSql($Uslov, $TablNam)
	$NamFunc = ""
	$textZaoo = $KilStrMeta & @CRLF & $Uslov & "'" & $TablNam & "'"
	$OtvFromSql = _SqlZap($textZaoo, "�� ������� �������� ������� ����� �������" & $TablNam, "TKSNS", 2)
	If $OtvFromSql = 1 Then
		_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
		SetError(1)
		Return 0
	EndIf
	_ArrayDelete($OtvFromSql, 0)
	If UBound($OtvFromSql) > 1 Then
		_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
		SetError(1)
		Return 0
	Else
		Return $OtvFromSql[0][1]
	EndIf
	$NamFunc = ""
EndFunc	






Func _Start2Etap()
	$NamFunc = "_Start2Etap"
	$LogNewNam = _FO_PathSplit(RegRead($VetkaPutSess & "\InFilesPutch", "PutchRBD"))
	If StringInStr($LogNewNam[1], "_data") Then $LogNewNam[1] = StringMid($LogNewNam[1], 1, StringInStr($LogNewNam[1], "_data") - 1)
	$SqlSpisPodklBd = _SqlZap("EXEC  [sp_databases]", "�� ������� �������� ������ ��������� ���", "master", 2)
	If Not IsArray($SqlSpisPodklBd) Then
		_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
		Exit(9999)
	Else
		_ArrayDelete($SqlSpisPodklBd, 0)
		$poiskBd = _ArraySearch($SqlSpisPodklBd, $LogNewNam[1], 0, UBound($SqlSpisPodklBd), 0, 0, 1, 0)
		If $poiskBd = -1 Then
			MsgBox(16 + 4096, $ZagTextPomZZP, "��������� ��� �� ��������� �� �������")
			_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
			Exit(9999)
		Else
			_GUICtrlStatusBar_SetText($hStatus, $SqlSpisPodklBd[$poiskBd][0], 2)
			$ParNamBd = "'[" & _GUICtrlStatusBar_GetText($hStatus, 2) & "]'"
			$ParCodLg = "'" & _GUICtrlStatusBar_GetText($hStatus, 0) & "'"
			$ParCodeLs = "'" & _GUICtrlStatusBar_GetText($hStatus, 1) & "'"
			$TextZapSql = "EXEC ZapKart_13 " & $ParNamBd & "," & $ParCodLg & "," & $ParCodeLs
			_ZapAndPerev($TextZapSql, $InpCopMak, 0, "���_����_������_", 6)
		EndIf
	EndIf
	$NamFunc = ""
EndFunc






Func _ZapAndPerev($ComTextSql, $elemResss, $ElemPovern = 0, $CommPrefiks = " ", $cherga = 0)
	$NamFunc = "_ZapAndPerev"
	$sVidpSqlRes = _SqlZap($ComTextSql, "�� ������� �������� ��������� " & $ComTextSql, "TKSNS", 5)
	If $sVidpSqlRes = 1 Then
		_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
		Exit(9991)
	Else
		Local $textInRess
		Switch $cherga
			Case 1
				$textInRess = ">>������ ��������� StvorKartTables "
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 2
				$textInRess = ">>������ ��������� NewKart "
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 3
				$textInRess = ">>������ ��������� FormComGed "
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 4
				$textInRess = ">>������ ��������� VnutrDublPar "
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 5
				$textInRess = ">>������ ��������� PomParObjKart"
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 6
				$textInRess = ">>������ ��������� ZapKart"
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 7
				$textInRess = ">>������ ��������� ZkopiuvDovid"
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 8
				$textInRess = ">>������ ���������  VstKwProsKbd"
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 9
				$textInRess = ">>������ ���������  AfterSyncContrl"
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 10
				$textInRess = ">>������ ��������� VstavLinObj"
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
			Case 11
				$textInRess = ">>������ ��������� SformUvjVid"
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
				
		EndSwitch
		GUICtrlSetData($elemResss, $CommPrefiks & $sVidpSqlRes)
		
		; ConsoleWrite('result=' & StringInStr($ComTextSql,"EXEC StvorKartTables") & @CRLF)
		If StringInStr($ComTextSql, "EXEC StvorKartTables_1") Then
			$VidpSqlZap = StringStripWS(_SqlZap("SELECT COUNT(*) FROM  _ErrVstaTabl", "�� ������� �������� �������� ������� �����, � ������� �������", "TKSNS", 1), 3)
			If $VidpSqlZap > 0 Then
				MsgBox(16 + 4096, $ZagTextPomZZP, "�� ���� �������! �������� ������� _ErrVstaTabl")
				_ObrErr($UnIdZZP, $numWin, $NamFunc, @AutoItPID)
				Exit(9992)
			Else
				$textInRess = ">>������ �������� ������� ������� ���� ��������� ��������� " & $ComTextSql
				RegWrite("HKLM\SOFTWARE\TKSNS\Session\TKSNS_" & $UnIdZZP & "\ZapZberProc", $cherga & "_proc", "REG_SZ", _Now() & $textInRess)
				GUICtrlSetData($ElemPovern, $CommPrefiks & "��")
			EndIf
		EndIf
	EndIf
	Return "good"
	$NamFunc = ""
EndFunc






Func _NeUvKZah()
	$ZapAgrKZah = "SELECT DISTINCT m.KAKL, " & _
			"(SELECT  ltrim(rtrim(d.name))" & _
			" FROM _DovKZah  d" & _
			" WHERE m.kakl = d.KAkl )  katzah" & _
			" FROM _onestring m ORDER BY katzah"
	$RecKatZah = _SqlZap("SELECT RCZ.Zap FROM _RecKatZah RCZ", "�� ������� ��������� ������ �������������� �� ������� ��������", 'TKSNS', 2)
	_ArrayDelete($RecKatZah, 0)
	$arrkatzah = _SqlZap($ZapAgrKZah, "�� ������� ��������� ������ �������� ����������", 'TKSNS', 2)
	_ArrayDelete($arrkatzah, 0)
	_GUICtrlListBox_BeginUpdate($KatZah)
	_GUICtrlListBox_Destroy($KatZah)
	$KatZah = GUICtrlCreateList("", 480, 115, 440, 150, BitOR($WS_VSCROLL, $LBS_MULTIPLESEL))
	For $iarrkatzah = 0 To UBound($arrkatzah) - 1
		_GUICtrlListBox_AddString($KatZah, $arrkatzah[$iarrkatzah][1])
; 	_ArrayDisplay($RecKatZah,"������")
		If _ArraySearch($RecKatZah, $arrkatzah[$iarrkatzah][0]) = -1 Then ContinueLoop
		_GUICtrlListBox_SetSel($KatZah, $iarrkatzah)
	Next
	If _GUICtrlListBox_GetSelCount($KatZah) > 0 Then
		GUICtrlSetState($CheckKatzah, $GUI_CHECKED)
	Else
		GUICtrlSetState($KatZah, $GUI_DISABLE)
	EndIf
	_GUICtrlListBox_EndUpdate($KatZah)
EndFunc

Func _NeUvKZem()
	Global $arrkatzem
	$ZapAgrKZem = "SELECT DISTINCT m.KAKZ" & _
			" ,(SELECT  ltrim(rtrim(d.name)) " & _
			" FROM _DovKZem  d  " & _
			" WHERE m.kakz = d.KAKZ )  [������в� ������] " & _
			" FROM _onestring m  WHERE KAKZ NOT IN(1201,1202,1203,1204,1205,1206,1207)"
	$RecKatZem = _SqlZap("SELECT RCZ.Zap FROM _RecKatZem RCZ", "�� ������� ��������� ������ �������������� �� ������� ��������", 'TKSNS', 2)
	_ArrayDelete($RecKatZem, 0)
; 	_ArrayDisplay($RecKatZem,"������")
	$arrkatzem = _SqlZap($ZapAgrKZem, "�� ������� ��������� ������ �������� ������              ", 'TKSNS', 2)
	_ArrayDelete($arrkatzem, 0)
	_GUICtrlListBox_BeginUpdate($KatZem)
	_GUICtrlListBox_Destroy($KatZem)
	$KatZem = GUICtrlCreateList("", 5, 115, 440, 150, BitOR($WS_VSCROLL, $LBS_MULTIPLESEL))
	For $iarrkatzem = 0 To UBound($arrkatzem) - 1
		_GUICtrlListBox_AddString($KatZem, $arrkatzem[$iarrkatzem][1])
; 		ConsoleWrite('result=' & _ArraySearch($RecKatZem, $arrkatzem[$iarrkatzem][0]) & @CRLF)
		If _ArraySearch($RecKatZem, $arrkatzem[$iarrkatzem][0]) = -1 Then
			ContinueLoop
		EndIf
		_GUICtrlListBox_SetSel($KatZem, $iarrkatzem,$LVS_LIST)
	Next
	_GUICtrlListBox_EndUpdate($KatZem)
EndFunc






Func _VidkKart($nbd, $pged)
	$stvribo = "IF OBJECT_ID('_GotTalKart', 'U') IS NOT NULL DROP TABLE _GotTalKart; CREATE TABLE _GotTalKart (P VARCHAR(50));" ;  ����� ��� ��������� ������� � ���
	$StvrTablPat = _SqlZap($stvribo, "������� ����������� ������� c�������� ������� ��������� ��'����!", $nbd, 5)
	
	$zapnazaliv = "BULK INSERT _GotTalKart FROM  " & "'" & $wertghuj & "'"
	$MassImport = _SqlZap($zapnazaliv, "������� ����������� ������� c�������� ������� ��������� ����� ��'���� �����!", $nbd, 5)
	
	$ferrgh = @ScriptDir & '\Pars.txt'
	$parscds = FileRead($ferrgh)
	
	_SqlZap($parscds, '�� ������� "���������" ����� ������� ��������� �����', $nbd, 5)
	$ZnachNullVPar = "SELECT count(*) FROM [_ParKartIzFile] WHERE lg IS NULL OR ls IS NULL OR kw IS NULL OR vd IS NULL"
	If _SqlZap($ZnachNullVPar, "�� ������� �������� ������� ������� ��������� ��'����", $nbd, 1) <> 0 Then
		$TablPomPar = "SELECT '@Map.SelectByParameters 2|-7='+CAST( [IdSloy] AS VARCHAR(100))+'|-6='+CAST( [IdObj] AS VARCHAR(50))" & @CRLF & _
				"FROM [3306-16_Data].[dbo].[_ParKartIzFile]" & @CRLF & _
				"WHERE lg IS NULL OR ls IS NULL OR kw IS NULL OR vd IS NULL"
		$TablPom = _SqlZap($TablPomPar, "�� ������� �������� ������� ������� ��������� ��'����", $nbd, 1)
		$namgroupom = '���_TAKSINS' & StringMid(StringReplace(_Now(), ' ', '_'), 4)
		FileDelete($pged & "\Library\_PomParKart.dsf")
		FileWrite($pged & "\Library\_PomParKart.dsf", '@Map.BeginUpdate' & $TablPom & '$uncod=@UTF8ToString ' & $namgroupom & @CRLF & '@Map.Selected.AddToGroup $uncod' & @CRLF & '@Map.EndUpdate' & @CRLF & "@Map.DeselectAll")
		; _ComDigitals("@ExecuteScript _PomParKart.dsf")
		MsgBox(64 + 4096, 'Info', "OK!")
	EndIf
EndFunc





Func _SpisBD()
	_SQL_RegisterErrorHandler()
	$oADODB = _SQL_Startup()
	If $oADODB = $SQL_ERROR Then
		MsgBox(0 + 16 + 262144, "������� ������� �������� ADODB!", _SQL_GetErrMsg())
		Return SetError(1)
	EndIf
	$con = _SQL_Connect(-1, @ComputerName, "master", "", "")
	If $con = $SQL_ERROR Then
		MsgBox(0 + 16 + 262144, "������� ��'������� �� ���������� �������!", _SQL_GetErrMsg())
		Return SetError(2)
	EndIf
	Local $sSpis = ''
	Local $hData
	$hData = _SQL_Execute(-1, "SELECT  d.name AS DBName,m.physical_name AS FileName FROM sys.databases d JOIN sys.master_files m ON d.database_id=m.database_id WHERE m.[type]=0 ORDER BY d.name;")
	If $hData = $SQL_ERROR Then
		MsgBox(0 + 16 + 262144, "��������� ������� ����� �� �������!", _SQL_GetErrMsg())
		Return SetError(3)
	EndIf
	Local $aRow
	While _SQL_FetchData($hData, $aRow) = $SQL_OK ; Read Out the next Row
		If $aRow[0] == 'master' Then ContinueLoop
		If $aRow[0] == 'tempdb' Then ContinueLoop
		If $aRow[0] == 'model' Then ContinueLoop
		If $aRow[0] == 'msdb' Then ContinueLoop
		$sSpis &= $aRow[0] & "|"
	WEnd
	_SQL_Close()
	Return $sSpis
EndFunc










	