' 교통량 API 서버 자동 시작 스크립트
' 콘솔 창 없이 백그라운드에서 pythonw.exe 로 api_server.py 를 실행합니다.

Dim WshShell, scriptDir, scriptPath

Set WshShell = CreateObject("WScript.Shell")

' 이 .vbs 파일이 위치한 폴더를 작업 디렉토리로 사용
scriptDir  = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
scriptPath = scriptDir & "\api_server.py"

' window style 0 = 창 완전 숨김 / False = 완료를 기다리지 않음(비동기)
WshShell.Run "pythonw.exe """ & scriptPath & """", 0, False

Set WshShell = Nothing
