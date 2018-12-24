
#include <xml.au3>
#include <encoding.au3>
 
#CS

---- �������������� ������� FILE ----
createnew �������� ����� ��������� �������
findbysid ����� ����� �� �������������� ������������
queryallocranges ������ ���������� ���������� ��� �����
queryextents ������ �������� ������ ��� �����
queryfileid ������ �������������� ���������� �����
queryfilenamebyid ����������� ������������� ����� ������ ��� 
�������������� �����
queryvaliddata ������ ���������� ����� ������ ��� �����
setshortname �������� ��������� ����� ��� �����
setvaliddata �������� ���������� ����� ������ ��� �����
setzerodata ��������� ������ � �����

#CE
; 0x000500000003ccd4 
ConsoleWrite(_FileID(@ScriptFullPath) & @CRLF) ;@ScriptFullPath
; ��������� ������������� ����� ���� � �����.
Func _FileID($sFilePath)
    Local $iPID = Run(@ComSpec & ' /c fsutil file queryfileid "' & $sFilePath & '"', @SystemDir, @SW_HIDE, $STDOUT_CHILD + $STDERR_CHILD), $sReturn = ''  
    While 1
        $sReturn &= StdoutRead($iPID)
        If @error Then
            ExitLoop
        EndIf
    WEnd
    Return StringStripWS( _Encoding_866To1251($sReturn) , 8)
EndFunc   ;==>_FileID