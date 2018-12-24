
#include <xml.au3>
#include <encoding.au3>
 
#CS

---- ПОДДЕРЖИВАЕМЫЕ КОМАНДЫ FILE ----
createnew Создание файла заданного размера
findbysid Поиск файла по идентификатору безопасности
queryallocranges Запрос выделенных диапазонов для файла
queryextents Запрос областей памяти для файла
queryfileid Запрос идентификатора указанного файла
queryfilenamebyid Отображение произвольного имени ссылки для 
идентификатора файла
queryvaliddata Запрос допустимой длины данных для файла
setshortname Указание короткого имени для файла
setvaliddata Указание допустимой длины данных для файла
setzerodata Обнуление данных в файле

#CE
; 0x000500000003ccd4 
ConsoleWrite(_FileID(@ScriptFullPath) & @CRLF) ;@ScriptFullPath
; Повторить идентификатор файла пути к файлу.
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