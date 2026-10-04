!macro NSIS_HOOK_POSTINSTALL
  MessageBox MB_YESNO|MB_ICONQUESTION "Install the local AI now? This installs Ollama and downloads the Qwen3 8B model (about 5.2 GB). It needs an internet connection and several GB of free disk space." IDNO skip_ollama
  nsExec::ExecToLog '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\resources\install-ollama.ps1" -InstallModel'
  Pop $0
  ${If} $0 != 0
    MessageBox MB_ICONEXCLAMATION "The local AI could not be installed. You can finish setup later from Settings in Church Presenter."
  ${EndIf}
  skip_ollama:
!macroend
