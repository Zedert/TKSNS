;Данная библиотека должна выполняться во время инсталляции...






Func _SozdNewProj($Uniq, $pKart, $pDig, $pNwRbd)
	$nazvPap = "TKSNS_" & $Uniq
	$ErrFlag = _RegWrite("HKLM\SOFTWARE\TKSNS", "Error", 1, 0)
	$ddSkor = "HKLM\SOFTWARE\TKSNS\Project\"
	$NewProg = _RegWrite($ddSkor & $nazvPap)
	$VstFlagErr = _RegWrite($ddSkor & $nazvPap, "FlagErr", 1, 0)
	$VstFlagRemoveExtit = _RegWrite($ddSkor & $nazvPap, "FlagRemExit", 1, 1)
	$PapkKart = _RegWrite($ddSkor & $nazvPap, "PutchKart", 1, $pKart)
	$PapkDigit = _RegWrite($ddSkor & $nazvPap, "PutchGed", 1, $pDig)
	$PapRbd = _RegWrite($ddSkor & $nazvPap, "NameWorkRbd", 1, $pNwRbd)
	$TksnsOsnRbd = _RegWrite($ddSkor & $nazvPap, "PutchDobBdOsn", 1, @ScriptDir & "\Tepmlate_DB\TKSNS.mdf")
	$TksnsLogRbd = _RegWrite($ddSkor & $nazvPap, "PutchDobBdLog", 1, @ScriptDir & "\Tepmlate_DB\TKSNS_log.ldf")
	$gilEtap = _RegWrite($ddSkor & $nazvPap & "\Etap")
	$EtapPidg = _RegWrite($ddSkor & $nazvPap & "\Etap\PidgotZapus", "", 1, 0)
	$EtapPidgPerFlagErr = _RegWrite($ddSkor & $nazvPap & "\Etap\PidgotZapus", "FlagErr", 1, 0)
	$EtapPidgCopyIni = _RegWrite($ddSkor & $nazvPap & "\Etap\PidgotZapus", "PutchOrinIni", 1, 0)
	$EtapPidgPerWriteIni = _RegWrite($ddSkor & $nazvPap & "\Etap\PidgotZapus", "ZapInIni", 1, 0)
	$EtapPidgPerReadIni = _RegWrite($ddSkor & $nazvPap & "\Etap\PidgotZapus", "ReadFromIni", 1, 0)
	$EtapPidgPerSaveIni = _RegWrite($ddSkor & $nazvPap & "\Etap\PidgotZapus", "PutchBeckupIni", 1, "~Project\\BeckUpGedIni\Ged.ini")
	$EtapPidgPerCopyTempKart = _RegWrite($ddSkor & $nazvPap & "\Etap\PidgotZapus", "PutchTemlKart", 1, "~Project\\UvPlo.dmf")
	$EtapZapDig = _RegWrite($ddSkor & $nazvPap & "\Etap\ZapDig")
	$EtapZapDigFlagErr = _RegWrite($ddSkor & $nazvPap & "\Etap\ZapDig", "FlagErr", 1, 0)
	$EtapZapDigExistProc = _RegWrite($ddSkor & $nazvPap & "\Etap\ZapDig", "ExistWorkGed", 1, 0)
	$EtapZapDigPutcTempl = _RegWrite($ddSkor & $nazvPap & "\Etap\ZapDig", "PutchTembTksns", 1, 0)
	$EtapZapDigPidnNewProc = _RegWrite($ddSkor & $nazvPap & "\Etap\ZapDig", "PidProc", 1, 0)
	$EtapZapDigSearchWin = _RegWrite($ddSkor & $nazvPap & "\Etap\ZapDig", "TemplWin", 1, "[Title:Digitals XE;Class:TFormGed]")
	$EtapZapDigTextGedCom = _RegWrite($ddSkor & $nazvPap & "\Etap\ZapDig", "TablWithTextComm", 1, "/DB-TKSNS /T-DigitScript /C-TextScript")
	If BitAND($NewProg, $gilEtap, $EtapPidg, $VstFlagErr, $PapkKart, $PapkDigit, $PapRbd, $TksnsOsnRbd, $TksnsLogRbd, $gilEtap, $EtapPidg, $EtapPidgCopyIni, $EtapPidgPerWriteIni, $EtapPidgPerReadIni, $EtapPidgPerSaveIni, $EtapZapDig, $EtapZapDigExistProc, $EtapZapDigPutcTempl, $EtapZapDigPidnNewProc, $EtapZapDigSearchWin, $EtapZapDigTextGedCom, $EtapPidgPerCopyTempKart, $EtapPidgPerFlagErr, $VstFlagRemoveExtit, $ErrFlag) <> 1 Then
		MsgBox(16 + 4096, "Помилка ", "не вдалося записатиз начення в реєстр імені робочої бази даних")
		Exit(78)
		Return 'bed'
	EndIf
	Return 'good'
EndFunc
 