echo githubからダウンロードして比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\SyncGithubToCodes.vbs" "%~dp0" "https://github.com/draemonash2/other/archive/master.zip" "other-master"

echo %MYDIRPATH_CODES%内の_localフォルダを比較します。
"%MYDIRPATH_CODES%\vbs\tools\win\file_ope\SyncCodesToLocal.vbs" "%~dp0"

