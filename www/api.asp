<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' ASPPY_AI_Builder - REST JSON API
Response.ContentType = "application/json"
Response.Charset = "utf-8"

' =================== UTILITY FUNCTIONS ===================

Function SqlEscape(val)
    If IsNull(val) Or IsEmpty(val) Then
        SqlEscape = "NULL"
    Else
        SqlEscape = "'" & Replace(val, "'", "''") & "'"
    End If
End Function

Function GetDBPath()
    GetDBPath = Server.MapPath("app.db")
End Function

Function OpenDB()
    Set OpenDB = Server.CreateObject("ADODB.Connection")
    OpenDB.Open GetDBPath()
End Function

Function JsonOk(data)
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "success", True
    If Not IsEmpty(data) And Not IsNull(data) Then
        result.Add "data", data
    End If
    JsonOk = ASPPY.JSON.Encode(result)
End Function

Function JsonError(msg, code)
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "success", False
    result.Add "error", msg
    If Not IsEmpty(code) And Not IsNull(code) Then
        result.Add "code", code
    End If
    JsonError = ASPPY.JSON.Encode(result)
End Function

Function UrlDecodeVB(s)
    On Error Resume Next
    Dim out, i, ch, hex
    out = ""
    i = 1
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch = "+" Then
            out = out & " "
            i = i + 1
        ElseIf ch = "%" And i + 2 <= Len(s) Then
            hex = Mid(s, i + 1, 2)
            If IsNumeric("&H" & hex) Then
                out = out & Chr(CLng("&H" & hex))
                i = i + 3
            Else
                out = out & ch
                i = i + 1
            End If
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    UrlDecodeVB = out
    On Error Goto 0
End Function

Function ParseJsonBody()
    On Error Resume Next
    Dim totalBytes, bin, strBody
    ' Try reading from form field "json" first (workaround for JSON body issues)
    If Request.Form.Exists("json") Then
        strBody = Request.Form("json")
        If Len(strBody) > 0 Then
            Set ParseJsonBody = ASPPY.JSON.Decode(strBody)
            If Err.Number <> 0 Then
                Err.Clear
                ParseJsonBody = Null
            End If
            On Error Goto 0
            Exit Function
        End If
    End If
    ' Fallback: read raw body
    totalBytes = Request.TotalBytes
    If totalBytes > 0 Then
        bin = Request.BinaryRead(totalBytes)
        strBody = CStr(bin)
        If Len(strBody) > 0 Then
            Set ParseJsonBody = ASPPY.JSON.Decode(strBody)
            If Err.Number <> 0 Then
                Err.Clear
                ParseJsonBody = Null
            End If
            On Error Goto 0
            Exit Function
        End If
    End If
    ParseJsonBody = Null
    On Error Goto 0
End Function

Function IsAdmin()
    IsAdmin = (Session("role") = "admin")
End Function

Function IsLoggedIn()
    IsLoggedIn = (Session("user_id") <> "" And Not IsNull(Session("user_id")))
End Function

Function RequireAuth()
    If Not IsLoggedIn() Then
        Response.Status = "401 Unauthorized"
        Response.Write JsonError("Authentication required", 401)
        Response.End
    End If
End Function

Function RequireAdmin()
    RequireAuth()
    If Not IsAdmin() Then
        Response.Status = "403 Forbidden"
        Response.Write JsonError("Admin access required", 403)
        Response.End
    End If
End Function

Function IsValidAsciiName(name)
    Dim i, ch, code
    If Len(name) = 0 Then
        IsValidAsciiName = False
        Exit Function
    End If
    For i = 1 To Len(name)
        ch = Mid(name, i, 1)
        code = Asc(ch)
        If code < 32 Or code > 126 Then
            IsValidAsciiName = False
            Exit Function
        End If
    Next
    IsValidAsciiName = True
End Function

Function SafeFileNamePart(name)
    Dim i, ch, code, result
    result = ""
    For i = 1 To Len(name)
        ch = Mid(name, i, 1)
        code = Asc(ch)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or code = 45 Or code = 95 Or code = 46 Then
            result = result & ch
        End If
    Next
    If Len(result) = 0 Then
        result = "untitled"
    End If
    SafeFileNamePart = result
End Function

' Physical directory of a project by its random folder name (never username based)
Function GetProjectDir(folderName)
    Dim root
    root = Server.MapPath("/")
    GetProjectDir = root & "\sites\" & SafeFileNamePart(folderName) & "\"
End Function

' Public URL prefix (root-relative) of a project by its random folder name
Function GetProjectUrlPrefix(folderName)
    GetProjectUrlPrefix = "/sites/" & SafeFileNamePart(folderName) & "/"
End Function

' Generate a random, unique folder name for a new project (no username exposure)
Function GenerateProjectFolder(conn)
    Dim chars, i, r, candidate, rs, tries
    chars = "abcdefghijklmnopqrstuvwxyz0123456789"
    Randomize
    For tries = 1 To 25
        candidate = ""
        For i = 1 To 16
            r = Int(Rnd() * Len(chars)) + 1
            candidate = candidate & Mid(chars, r, 1)
        Next
        Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM projects WHERE folder=" & SqlEscape(candidate))
        If CLng(rs("cnt").Value) = 0 Then
            rs.Close
            GenerateProjectFolder = candidate
            Exit Function
        End If
        rs.Close
    Next
    GenerateProjectFolder = candidate
End Function

' Very light email format validation
Function IsValidEmail(email)
    Dim atPos, dotPos
    IsValidEmail = False
    If Len(email) < 5 Or Len(email) > 254 Then Exit Function
    If InStr(1, email, " ", 1) > 0 Then Exit Function
    atPos = InStr(1, email, "@", 1)
    If atPos < 2 Then Exit Function
    If InStr(atPos + 1, email, "@", 1) > 0 Then Exit Function
    dotPos = InStrRev(email, ".")
    If dotPos < atPos + 2 Then Exit Function
    If dotPos >= Len(email) Then Exit Function
    IsValidEmail = True
End Function

' Replace long dashes (em dash etc.) in AI output with a plain hyphen
Function SanitizeAiText(s)
    Dim t
    If IsNull(s) Or IsEmpty(s) Then
        SanitizeAiText = ""
        Exit Function
    End If
    t = CStr(s)
    t = Replace(t, ChrW(8212), "-") ' em dash
    t = Replace(t, ChrW(8213), "-") ' horizontal bar
    SanitizeAiText = t
End Function

' Rewrite absolute-style image references in AI output to relative ones (img/...)
Function MakeImagePathsRelative(s, folderName)
    Dim t
    t = s
    t = Replace(t, "/sites/" & folderName & "/img/", "img/")
    t = Replace(t, "src=""/img/", "src=""img/")
    t = Replace(t, "src='/img/", "src='img/")
    t = Replace(t, "url('/img/", "url('img/")
    t = Replace(t, "url(""/img/", "url(""img/")
    t = Replace(t, "url(/img/", "url(img/")
    MakeImagePathsRelative = t
End Function

' Send a plain-text email using the SMTP settings from the admin config.
' Returns "" on success, otherwise an error description.
Function SendMail(toAddr, subject, bodyText)
    Dim host, port, smtpUser, smtpPass, fromAddr, useSsl, msg
    host = GetConfigValue("smtp_host")
    port = GetConfigValue("smtp_port")
    smtpUser = GetConfigValue("smtp_user")
    smtpPass = GetConfigValue("smtp_password")
    fromAddr = GetConfigValue("smtp_from")
    useSsl = GetConfigValue("smtp_ssl")
    
    If Len(host) = 0 Then
        SendMail = "SMTP server is not configured (admin panel > Email settings)"
        Exit Function
    End If
    If Len(port) = 0 Then port = "25"
    If Len(fromAddr) = 0 Then fromAddr = smtpUser
    
    On Error Resume Next
    Set msg = Server.CreateObject("CDO.Message")
    msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = host
    msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(port)
    msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
    If Len(smtpUser) > 0 Then
        msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = smtpUser
        msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = smtpPass
    End If
    If useSsl = "1" Then
        msg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    End If
    msg.Configuration.Fields.Update
    
    msg.From = fromAddr
    msg.To = toAddr
    msg.Subject = subject
    msg.TextBody = bodyText
    msg.Send
    
    If Err.Number <> 0 Then
        SendMail = "Email send failed: " & Err.Description
        Err.Clear
    Else
        SendMail = ""
    End If
    On Error Goto 0
End Function

Function EnsureDir(path)
    Dim fso, cleanPath, parent
    cleanPath = path
    ' Strip trailing backslash
    If Right(cleanPath, 1) = "\" Then
        cleanPath = Left(cleanPath, Len(cleanPath) - 1)
    End If
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(cleanPath) Then
        parent = fso.GetParentFolderName(cleanPath)
        If Len(parent) > 0 Then
            If Not fso.FolderExists(parent) Then
                EnsureDir parent
            End If
        End If
        fso.CreateFolder cleanPath
		if right(cleanPath,4)<>"\img" then 'not for the image folder
			fso.CopyFile server.htmlencode("favicon.ico"), cleanPath
			fso.CopyFile server.htmlencode("robots.txt"), cleanPath		
		end if
    End If
End Function

' Single-file mode: only index.html is written by the AI. Legacy
' style.css/script.js files on disk are still served/restored for old
' projects, but new generations never create them.
Function ValidateFileExtension(filename)
    Dim lname
    lname = LCase(filename)
    If lname = "index.html" Then
        ValidateFileExtension = True
        Exit Function
    End If
    ValidateFileExtension = False
End Function

Function ContainsPathTraversal(path)
    ContainsPathTraversal = (InStr(1, path, "..", 1) > 0 Or _
                              InStr(1, path, "~", 1) > 0 Or _
                              InStr(1, path, ":", 1) > 0)
End Function

Function MakeTimestamp()
    MakeTimestamp = Year(Date) & "-" & _
                    Right("0" & Month(Date), 2) & "-" & _
                    Right("0" & Day(Date), 2) & "T" & _
                    Right("0" & Hour(Time), 2) & "-" & _
                    Right("0" & Minute(Time), 2) & "-" & _
                    Right("0" & Second(Time), 2)
End Function

' =================== SETUP ===================

' Create all tables, seed config, and add the default admin when missing.
' Returns True when the default admin was created.
Function InitSchema(conn)
    conn.Execute "CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL UNIQUE, email TEXT NOT NULL DEFAULT '', password_hash TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'user', created_at TEXT NOT NULL DEFAULT (datetime('now')))"
    
    conn.Execute "CREATE TABLE IF NOT EXISTS projects (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, project_name TEXT NOT NULL, folder TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL DEFAULT (datetime('now')), updated_at TEXT NOT NULL DEFAULT (datetime('now')), FOREIGN KEY (user_id) REFERENCES users(id), UNIQUE(user_id, project_name))"
    
    conn.Execute "CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value TEXT DEFAULT '')"
    
    conn.Execute "CREATE TABLE IF NOT EXISTS prompts (id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, prompt TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT (datetime('now')))"
    
    EnsureJobsTable conn
    
    conn.Execute "CREATE TABLE IF NOT EXISTS password_resets (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, token TEXT NOT NULL, used INTEGER NOT NULL DEFAULT 0, expires_at TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT (datetime('now')))"
    
    conn.Execute "CREATE TABLE IF NOT EXISTS chat_messages (id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, role TEXT NOT NULL, content TEXT NOT NULL, reasoning TEXT NOT NULL DEFAULT '', tokens_in INTEGER NOT NULL DEFAULT 0, tokens_out INTEGER NOT NULL DEFAULT 0, duration_seconds INTEGER NOT NULL DEFAULT (datetime('now')), FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE)"
    
    ' Migrations for databases created by older versions (errors are ignored when the column already exists)
    On Error Resume Next
    conn.Execute "ALTER TABLE users ADD COLUMN email TEXT NOT NULL DEFAULT ''"
    Err.Clear
    conn.Execute "ALTER TABLE projects ADD COLUMN folder TEXT NOT NULL DEFAULT ''"
    Err.Clear
    conn.Execute "ALTER TABLE projects ADD COLUMN notes TEXT NOT NULL DEFAULT ''"
    Err.Clear
    conn.Execute "ALTER TABLE jobs ADD COLUMN folder TEXT NOT NULL DEFAULT ''"
    Err.Clear
    conn.Execute "ALTER TABLE jobs ADD COLUMN notify INTEGER NOT NULL DEFAULT 0"
    Err.Clear
    conn.Execute "ALTER TABLE jobs ADD COLUMN email TEXT NOT NULL DEFAULT ''"
    Err.Clear
    conn.Execute "ALTER TABLE jobs ADD COLUMN cancel INTEGER NOT NULL DEFAULT 0"
    Err.Clear
    conn.Execute "ALTER TABLE chat_messages ADD COLUMN reasoning TEXT NOT NULL DEFAULT ''"
    Err.Clear
    conn.Execute "ALTER TABLE chat_messages ADD COLUMN tokens_in INTEGER NOT NULL DEFAULT 0"
    Err.Clear
    conn.Execute "ALTER TABLE chat_messages ADD COLUMN tokens_out INTEGER NOT NULL DEFAULT 0"
    Err.Clear
    conn.Execute "ALTER TABLE chat_messages ADD COLUMN duration_seconds INTEGER NOT NULL DEFAULT 0"
    Err.Clear
    On Error Goto 0
    
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('internal_url', 'http://127.0.0.1:8080')"
    
    ' Seed default config
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('llm_endpoint', '')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('llm_model', '')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('llm_api_key', '')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('system_prompt', '')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('enable_wikipedia', '0')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('llm_parameters', '')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('enable_sections', '1')"
    
    ' Email / SMTP settings (manageable from the admin panel)
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('public_url', 'http://localhost:8080')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('smtp_host', '')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('smtp_port', '587')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('smtp_user', '')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('smtp_password', '')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('smtp_from', '')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('smtp_ssl', '1')"
    
    ' Editable email templates
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('reset_email_subject', 'Reset your Autonome password')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('reset_email_body', 'Hello {username},' || char(10) || char(10) || 'Someone (hopefully you) requested a password reset for your Autonome account.' || char(10) || 'Click the link below to set a new password:' || char(10) || char(10) || '{link}' || char(10) || char(10) || 'This link is valid for 1 hour. If you did not request this, you can safely ignore this email.')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('notify_email_subject', 'Your website generation has finished')"
    conn.Execute "INSERT OR IGNORE INTO config (key, value) VALUES ('notify_email_body', 'Hello {username},' || char(10) || char(10) || 'The AI generation for your project ""{project}"" has finished with status: {status}.' || char(10) || char(10) || 'Open your project: {link}')"
    
    ' Check if admin exists
    Dim rs
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM users WHERE role='admin'")
    Dim adminExists
    adminExists = False
    If Not rs.EOF Then
        If CLng(rs("cnt").Value) > 0 Then
            adminExists = True
        End If
    End If
    rs.Close
    
    If Not adminExists Then
        Dim defaultHash
        defaultHash = ASPPY.Crypto.Hash("admin123", 10)
        conn.Execute "INSERT INTO users (username, email, password_hash, role) VALUES ('admin', 'admin@example.com', " & SqlEscape(defaultHash) & ", 'admin')"
        InitSchema = True
    Else
        InitSchema = False
    End If
End Function

Function HandleSetup()
    Dim conn, adminCreated
    Set conn = OpenDB()
    adminCreated = InitSchema(conn)
    conn.Close
    
    If adminCreated Then
        Response.Write JsonOk("Database initialized. Default admin: admin / admin123")
    Else
        Response.Write JsonOk("Database already initialized")
    End If
End Function

' Run lightweight migrations on every request so existing databases get new columns/tables
' without needing a full InitSchema (which only runs on fresh install).
Sub EnsureMigrations(conn)
    On Error Resume Next
    conn.Execute "ALTER TABLE projects ADD COLUMN notes TEXT NOT NULL DEFAULT ''"
    Err.Clear
    conn.Execute "CREATE TABLE IF NOT EXISTS chat_messages (id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, role TEXT NOT NULL, content TEXT NOT NULL, reasoning TEXT NOT NULL DEFAULT '', tokens_in INTEGER NOT NULL DEFAULT 0, tokens_out INTEGER NOT NULL DEFAULT 0, duration_seconds INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL DEFAULT (datetime('now')), FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE)"
    Err.Clear
    conn.Execute "ALTER TABLE chat_messages ADD COLUMN reasoning TEXT NOT NULL DEFAULT ''"
    Err.Clear
    conn.Execute "ALTER TABLE chat_messages ADD COLUMN tokens_in INTEGER NOT NULL DEFAULT 0"
    Err.Clear
    conn.Execute "ALTER TABLE chat_messages ADD COLUMN tokens_out INTEGER NOT NULL DEFAULT 0"
    Err.Clear
    conn.Execute "ALTER TABLE chat_messages ADD COLUMN duration_seconds INTEGER NOT NULL DEFAULT 0"
    Err.Clear
    On Error Goto 0
End Sub

' Auto-initialize the database on the first request (e.g. fresh install or
' after all data was deleted), so no manual setup call is needed.
Sub EnsureDatabase()
    On Error Resume Next
    Dim conn, rs, needsInit
    needsInit = False
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT id FROM users LIMIT 1")
    If Err.Number <> 0 Then
        Err.Clear
        needsInit = True
    Else
        rs.Close
    End If
    On Error Goto 0
    If needsInit Then
        InitSchema conn
    Else
        EnsureMigrations conn
    End If
    conn.Close
End Sub

' =================== AUTH ===================

Function HandleLogin()
    Dim body, username, password, conn, rs, sql, storedHash, valid
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    username = body("username")
    password = body("password")
    
    If Len(username) = 0 Or Len(password) = 0 Then
        Response.Write JsonError("Username and password required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    sql = "SELECT id, username, email, password_hash, role FROM users WHERE username=" & SqlEscape(username)
    Set rs = conn.Execute(sql)
    
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Invalid credentials", 401)
        Exit Function
    End If
    
    storedHash = rs("password_hash").Value
    valid = ASPPY.Crypto.Verify(password, storedHash)
    
    If valid Then
        Session("user_id") = CLng(rs("id").Value)
        Session("username") = rs("username").Value
        Session("role") = rs("role").Value
        Session("email") = rs("email").Value
        
        Dim userData
        Set userData = Server.CreateObject("Scripting.Dictionary")
        userData.Add "id", CLng(rs("id").Value)
        userData.Add "username", rs("username").Value
        userData.Add "email", rs("email").Value
        userData.Add "role", rs("role").Value
        
        rs.Close
        conn.Close
        Response.Write JsonOk(userData)
    Else
        rs.Close
        conn.Close
        Response.Write JsonError("Invalid credentials", 401)
    End If
End Function

Function HandleLogout()
    Session.Abandon
    Response.Write JsonOk("Logged out")
End Function

Function HandleSession()
    If IsLoggedIn() Then
        Dim sessData
        Set sessData = Server.CreateObject("Scripting.Dictionary")
        sessData.Add "user_id", Session("user_id")
        sessData.Add "username", Session("username")
        sessData.Add "email", Session("email")
        sessData.Add "role", Session("role")
        Response.Write JsonOk(sessData)
    Else
        Response.Write JsonOk(Null)
    End If
End Function

Function HandleRegister()
    Dim body, username, email, password, conn, rs, hash
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    username = body("username")
    password = body("password")
    email = ""
    If body.Exists("email") Then email = Trim(CStr(body("email")))
    
    If Len(username) = 0 Or Len(password) = 0 Then
        Response.Write JsonError("Username and password required", 400)
        Exit Function
    End If
    
    If Len(email) = 0 Then
        Response.Write JsonError("Email address required", 400)
        Exit Function
    End If
    
    If Not IsValidEmail(email) Then
        Response.Write JsonError("Invalid email address", 400)
        Exit Function
    End If
    
    If Len(password) < 6 Then
        Response.Write JsonError("Password must be at least 6 characters", 400)
        Exit Function
    End If
    
    If Not IsValidAsciiName(username) Then
        Response.Write JsonError("Username must be ASCII characters only", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM users WHERE username=" & SqlEscape(username))
    If Not rs.EOF Then
        If CLng(rs("cnt").Value) > 0 Then
            rs.Close
            conn.Close
            Response.Write JsonError("Username already exists", 409)
            Exit Function
        End If
    End If
    rs.Close
    
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM users WHERE email=" & SqlEscape(email))
    If Not rs.EOF Then
        If CLng(rs("cnt").Value) > 0 Then
            rs.Close
            conn.Close
            Response.Write JsonError("Email address already in use", 409)
            Exit Function
        End If
    End If
    rs.Close
    
    hash = ASPPY.Crypto.Hash(password, 10)
    conn.Execute "INSERT INTO users (username, email, password_hash, role) VALUES (" & SqlEscape(username) & ", " & SqlEscape(email) & ", " & SqlEscape(hash) & ", 'user')"
    
    Set rs = conn.Execute("SELECT id, username, email, role FROM users WHERE username=" & SqlEscape(username))
    If Not rs.EOF Then
        Session("user_id") = CLng(rs("id").Value)
        Session("username") = rs("username").Value
        Session("role") = rs("role").Value
        Session("email") = rs("email").Value
        
        Dim userData
        Set userData = Server.CreateObject("Scripting.Dictionary")
        userData.Add "id", CLng(rs("id").Value)
        userData.Add "username", rs("username").Value
        userData.Add "email", rs("email").Value
        userData.Add "role", rs("role").Value
        
        rs.Close
        conn.Close
        Response.Write JsonOk(userData)
    Else
        rs.Close
        conn.Close
        Response.Write JsonError("Registration failed", 500)
    End If
End Function

Function HandleChangePassword()
    RequireAuth()
    
    Dim body, currentPassword, newPassword, conn, rs, storedHash, valid
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    currentPassword = body("current_password")
    newPassword = body("new_password")
    
    If Len(currentPassword) = 0 Or Len(newPassword) = 0 Then
        Response.Write JsonError("Current and new password required", 400)
        Exit Function
    End If
    
    If Len(newPassword) < 6 Then
        Response.Write JsonError("New password must be at least 6 characters", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT password_hash FROM users WHERE id=" & Session("user_id"))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("User not found", 404)
        Exit Function
    End If
    
    storedHash = rs("password_hash").Value
    valid = ASPPY.Crypto.Verify(currentPassword, storedHash)
    
    If Not valid Then
        rs.Close
        conn.Close
        Response.Write JsonError("Current password is incorrect", 401)
        Exit Function
    End If
    
    rs.Close
    Dim newHash
    newHash = ASPPY.Crypto.Hash(newPassword, 10)
    conn.Execute "UPDATE users SET password_hash=" & SqlEscape(newHash) & " WHERE id=" & Session("user_id")
    conn.Close
    
    Response.Write JsonOk("Password changed successfully")
End Function

' Change the email address of the logged-in user (requires the current password)
Function HandleUpdateEmail()
    RequireAuth()
    
    Dim body, password, newEmail, conn, rs, storedHash, valid
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    password = body("password")
    newEmail = ""
    If body.Exists("email") Then newEmail = Trim(CStr(body("email")))
    
    If Len(password) = 0 Or Len(newEmail) = 0 Then
        Response.Write JsonError("Password and new email required", 400)
        Exit Function
    End If
    
    If Not IsValidEmail(newEmail) Then
        Response.Write JsonError("Invalid email address", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT password_hash FROM users WHERE id=" & Session("user_id"))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("User not found", 404)
        Exit Function
    End If
    storedHash = rs("password_hash").Value
    rs.Close
    
    valid = ASPPY.Crypto.Verify(password, storedHash)
    If Not valid Then
        conn.Close
        Response.Write JsonError("Password is incorrect", 401)
        Exit Function
    End If
    
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM users WHERE email=" & SqlEscape(newEmail) & " AND id<>" & Session("user_id"))
    If Not rs.EOF Then
        If CLng(rs("cnt").Value) > 0 Then
            rs.Close
            conn.Close
            Response.Write JsonError("Email address already in use", 409)
            Exit Function
        End If
    End If
    rs.Close
    
    conn.Execute "UPDATE users SET email=" & SqlEscape(newEmail) & " WHERE id=" & Session("user_id")
    conn.Close
    
    Session("email") = newEmail
    
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "message", "Email address updated"
    result.Add "email", newEmail
    Response.Write JsonOk(result)
End Function

' Send a password reset link by email. Always responds neutrally so email
' addresses cannot be probed, unless the mail server itself fails.
Function HandleForgotPassword()
    Dim body, email, conn, rs, userId, username, token, publicUrl, link, subject, bodyTpl, sendErr
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    email = ""
    If body.Exists("email") Then email = Trim(CStr(body("email")))
    If Len(email) = 0 Then
        Response.Write JsonError("Email address required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT id, username FROM users WHERE email=" & SqlEscape(email))
    If rs.EOF Then
        rs.Close
        conn.Close
        ' Neutral answer: do not reveal whether the email exists
        Response.Write JsonOk("If that email address is registered, a reset link has been sent.")
        Exit Function
    End If
    
    userId = CLng(rs("id").Value)
    username = rs("username").Value
    rs.Close
    
    token = RandomToken()
    conn.Execute "INSERT INTO password_resets (user_id, token, expires_at) VALUES (" & userId & ", " & SqlEscape(token) & ", datetime('now', '+1 hour'))"
    conn.Close
    
    publicUrl = GetConfigValue("public_url")
    If Len(publicUrl) = 0 Then publicUrl = "http://localhost:8080"
    If Right(publicUrl, 1) = "/" Then publicUrl = Left(publicUrl, Len(publicUrl) - 1)
    link = publicUrl & "/#/reset/" & token
    
    subject = GetConfigValue("reset_email_subject")
    If Len(subject) = 0 Then subject = "Reset your Autonome password"
    bodyTpl = GetConfigValue("reset_email_body")
    If Len(bodyTpl) = 0 Then bodyTpl = "Hello {username}," & vbLf & vbLf & "Click the link below to set a new password:" & vbLf & "{link}" & vbLf & vbLf & "This link is valid for 1 hour."
    bodyTpl = Replace(bodyTpl, "{username}", username)
    bodyTpl = Replace(bodyTpl, "{link}", link)
    
    sendErr = SendMail(email, subject, bodyTpl)
    If Len(sendErr) > 0 Then
        Response.Write JsonError(sendErr, 500)
        Exit Function
    End If
    
    Response.Write JsonOk("If that email address is registered, a reset link has been sent.")
End Function

' Set a new password using a valid reset token from the email link
Function HandleResetPassword()
    Dim body, token, newPassword, conn, rs, userId, newHash
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    token = body("token")
    newPassword = body("password")
    
    If Len(token) = 0 Or Len(newPassword) = 0 Then
        Response.Write JsonError("Token and new password required", 400)
        Exit Function
    End If
    
    If Len(newPassword) < 6 Then
        Response.Write JsonError("Password must be at least 6 characters", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT id, user_id FROM password_resets WHERE token=" & SqlEscape(token) & " AND used=0 AND expires_at > datetime('now')")
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Invalid or expired reset link. Please request a new one.", 400)
        Exit Function
    End If
    
    Dim resetId
    resetId = CLng(rs("id").Value)
    userId = CLng(rs("user_id").Value)
    rs.Close
    
    newHash = ASPPY.Crypto.Hash(newPassword, 10)
    conn.Execute "UPDATE users SET password_hash=" & SqlEscape(newHash) & " WHERE id=" & userId
    conn.Execute "UPDATE password_resets SET used=1 WHERE id=" & resetId
    conn.Close
    
    Response.Write JsonOk("Password has been reset. You can now sign in with your new password.")
End Function

' =================== CONFIG (ADMIN ONLY) ===================

Function HandleGetConfig()
    RequireAdmin()
    
    Dim conn, rs, cfg
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT key, value FROM config ORDER BY key")
    
    Set cfg = Server.CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        cfg.Add rs("key").Value, rs("value").Value
        rs.MoveNext
    Loop
    rs.Close
    conn.Close
    
    Response.Write JsonOk(cfg)
End Function

Function HandleSaveConfig()
    RequireAdmin()
    
    Dim body, conn, key, val
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    
    Dim configKeys, i
    configKeys = Array("llm_endpoint", "llm_model", "llm_api_key", "system_prompt", "enable_wikipedia", "llm_parameters", "enable_sections", "internal_url", _
                       "public_url", "smtp_host", "smtp_port", "smtp_user", "smtp_password", "smtp_from", "smtp_ssl", _
                       "reset_email_subject", "reset_email_body", "notify_email_subject", "notify_email_body")
    For i = 0 To UBound(configKeys)
        key = configKeys(i)
        If body.Exists(key) Then
            ' Upsert: also works when the key is missing from an older database
            conn.Execute "INSERT OR REPLACE INTO config (key, value) VALUES ('" & key & "', " & SqlEscape(body(key)) & ")"
        End If
    Next
    
    conn.Close
    Response.Write JsonOk("Configuration saved")
End Function

' List the model names available on the LLM endpoint (OpenAI-compatible
' /v1/models, supported by LM Studio and Ollama). The endpoint/key from the
' request body take precedence so unsaved form values can be probed.
Function HandleListModels()
    RequireAdmin()
    
    Dim body, endpoint, apiKey
    Set body = ParseJsonBody()
    endpoint = ""
    apiKey = ""
    If Not IsNull(body) Then
        If body.Exists("endpoint") Then endpoint = Trim("" & body("endpoint"))
        If body.Exists("api_key") Then apiKey = "" & body("api_key")
    End If
    If Len(endpoint) = 0 Then endpoint = GetConfigValue("llm_endpoint")
    If Len(apiKey) = 0 Then apiKey = GetConfigValue("llm_api_key")
    
    If Len(endpoint) = 0 Then
        Response.Write JsonError("Set the LLM endpoint URL first", 400)
        Exit Function
    End If
    If Right(endpoint, 1) = "/" Then endpoint = Left(endpoint, Len(endpoint) - 1)

    ' Anthropic does not expose an OpenAI-compatible /v1/models endpoint.
    ' Return a curated list of known Anthropic models so the UI remains usable.
    ' Note: claude-sonnet-4-20250514 retired 2026-06-15 (use claude-sonnet-5 / claude-sonnet-4-6).
    If IsAnthropicEndpoint(endpoint) Then
        Dim anthModels, mIdx
        Set anthModels = Server.CreateObject("Scripting.Dictionary")
        anthModels.Add 0, "claude-sonnet-5"
        anthModels.Add 1, "claude-sonnet-4-6"
        anthModels.Add 2, "claude-opus-4-6"
        anthModels.Add 3, "claude-sonnet-4-20250514"
        anthModels.Add 4, "claude-3-5-sonnet-20241022"
        anthModels.Add 5, "claude-3-opus-20240229"
        anthModels.Add 6, "claude-3-haiku-20240307"
        ' Include the user-supplied name so it stays selectable even if not in the curated list
        Dim cfgModel
        cfgModel = GetConfigValue("llm_model")
        If Len(cfgModel) > 0 Then
            Dim alreadyThere, kk
            alreadyThere = False
            For kk = 0 To anthModels.Count - 1
                If anthModels(kk) = cfgModel Then alreadyThere = True
            Next
            If Not alreadyThere Then
                anthModels.Add anthModels.Count, cfgModel
            End If
        End If
        Response.Write JsonOk(anthModels)
        Exit Function
    End If
    
    On Error Resume Next
    Dim xmlhttp
    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "GET", endpoint & "/v1/models", False
    If Len(apiKey) > 0 Then
        xmlhttp.setRequestHeader "Authorization", "Bearer " & apiKey
    End If
    xmlhttp.setTimeouts 10000, 10000, 10000, 10000
    xmlhttp.Send
    
    If Err.Number <> 0 Then
        Dim reachErr
        reachErr = Err.Description
        Err.Clear
        On Error Goto 0
        Response.Write JsonError("Could not reach the LLM endpoint: " & reachErr, 502)
        Exit Function
    End If
    If xmlhttp.Status < 200 Or xmlhttp.Status >= 300 Then
        On Error Goto 0
        Response.Write JsonError("LLM endpoint returned HTTP " & xmlhttp.Status, 502)
        Exit Function
    End If
    
    Dim parsed, models, idx, entry
    Set models = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    Set parsed = ASPPY.JSON.Decode(xmlhttp.responseText)
    If IsObject(parsed) Then
        If parsed.Exists("data") Then
            For Each entry In parsed("data")
                If IsObject(entry) Then
                    If entry.Exists("id") Then
                        models.Add idx, "" & entry("id")
                        idx = idx + 1
                    End If
                End If
            Next
        End If
    End If
    On Error Goto 0
    
    Response.Write JsonOk(models)
End Function

' =================== PROJECTS ===================

Function HandleGetProjects()
    RequireAuth()
    
    Dim conn, rs, projects, idx
    Set conn = OpenDB()
    On Error Resume Next
    Set rs = conn.Execute("SELECT id, project_name, folder, notes, created_at, updated_at FROM projects WHERE user_id=" & Session("user_id") & " ORDER BY updated_at DESC")
    If Err.Number <> 0 Then
        Err.Clear
        Set rs = conn.Execute("SELECT id, project_name, folder, created_at, updated_at FROM projects WHERE user_id=" & Session("user_id") & " ORDER BY updated_at DESC")
    End If
    On Error Goto 0
    
    Set projects = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    Do While Not rs.EOF
        Dim p
        Set p = Server.CreateObject("Scripting.Dictionary")
        p.Add "id", CLng(rs("id").Value)
        p.Add "name", rs("project_name").Value
        p.Add "folder", rs("folder").Value
        On Error Resume Next
        Dim nval
        nval = rs("notes").Value
        If Err.Number = 0 Then
            If IsNull(nval) Then nval = ""
            p.Add "notes", CStr(nval)
        Else
            Err.Clear
            p.Add "notes", ""
        End If
        On Error Goto 0
        p.Add "created_at", rs("created_at").Value
        p.Add "updated_at", rs("updated_at").Value
        
        projects.Add idx, p
        idx = idx + 1
        rs.MoveNext
    Loop
    rs.Close
    conn.Close
    
    Response.Write JsonOk(projects)
End Function

Function HandleCreateProject()
    RequireAuth()
    
    Dim body, projectName, folderName, conn, rs, projectDir
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    ' No ASCII restriction anymore: the project name is only a label; the
    ' folder on disk is a random name and never derived from it.
    projectName = Trim("" & body("name"))
    If Len(projectName) = 0 Then
        Response.Write JsonError("Project name required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM projects WHERE user_id=" & Session("user_id") & " AND project_name=" & SqlEscape(projectName))
    If Not rs.EOF Then
        If CLng(rs("cnt").Value) > 0 Then
            rs.Close
            conn.Close
            Response.Write JsonError("Project already exists", 409)
            Exit Function
        End If
    End If
    rs.Close
    
    ' Random folder name: usernames are never exposed in URLs or file paths
    folderName = GenerateProjectFolder(conn)
    
    conn.Execute "INSERT INTO projects (user_id, project_name, folder) VALUES (" & Session("user_id") & ", " & SqlEscape(projectName) & ", " & SqlEscape(folderName) & ")"
    
    Set rs = conn.Execute("SELECT id, project_name, folder, created_at, updated_at FROM projects WHERE user_id=" & Session("user_id") & " AND project_name=" & SqlEscape(projectName))
    If Not rs.EOF Then
        ' Create directory structure
        projectDir = GetProjectDir(folderName)
        
        Dim fso
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(projectDir) Then
            EnsureDir projectDir
            EnsureDir projectDir & "img"
            EnsureDir projectDir & "backups"
            WriteStarterFiles projectDir, projectName
        End If
        
        Dim pData
        Set pData = Server.CreateObject("Scripting.Dictionary")
        pData.Add "id", CLng(rs("id").Value)
        pData.Add "name", rs("project_name").Value
        pData.Add "folder", rs("folder").Value
        pData.Add "notes", ""
        pData.Add "created_at", rs("created_at").Value
        pData.Add "updated_at", rs("updated_at").Value
        
        rs.Close
        conn.Close
        Response.Write JsonOk(pData)
    Else
        rs.Close
        conn.Close
        Response.Write JsonError("Failed to create project", 500)
    End If
End Function

' Write the default onboarding file (single-file index.html, self-contained)
' for a project. Used on project creation and on project reset.
Sub WriteStarterFiles(projectDir, projectName)
    Dim tpl
    tpl = "<!DOCTYPE html>" & vbCrLf
    tpl = tpl & "<html lang=""en"">" & vbCrLf
            tpl = tpl & "<head>" & vbCrLf
            tpl = tpl & "    <meta charset=""UTF-8"">" & vbCrLf
            tpl = tpl & "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbCrLf
            tpl = tpl & "    <title>" & projectName & " - ready to start</title>" & vbCrLf
            tpl = tpl & "    <link href=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css"" rel=""stylesheet"">" & vbCrLf
            tpl = tpl & "    <link href=""https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"" rel=""stylesheet"">" & vbCrLf
            tpl = tpl & "    <style>" & vbCrLf
            tpl = tpl & "    /* Custom styles on top of Bootstrap 5 */" & vbCrLf
            tpl = tpl & "    .starter-bg { background: linear-gradient(160deg, #f8fafc 0%, #eef2ff 55%, #fdf2f8 100%); }" & vbCrLf
            tpl = tpl & "    </style>" & vbCrLf
            tpl = tpl & "</head>" & vbCrLf
            tpl = tpl & "<body class=""starter-bg"">" & vbCrLf
            tpl = tpl & "    <div class=""container min-vh-100 d-flex align-items-center justify-content-center py-5"">" & vbCrLf
            tpl = tpl & "        <div class=""text-center"" style=""max-width: 720px;"">" & vbCrLf
            tpl = tpl & "            <div class=""display-1 mb-3"">🚀</div>" & vbCrLf
            tpl = tpl & "            <h1 class=""display-5 fw-bold mb-2"">" & projectName & "</h1>" & vbCrLf
            tpl = tpl & "            <p class=""lead text-secondary mb-5"">Your website doesn't exist yet — but that won't take long. Here's how it works:</p>" & vbCrLf
            tpl = tpl & "            <div class=""row g-4 text-start"">" & vbCrLf
            tpl = tpl & "                <div class=""col-md-4"">" & vbCrLf
            tpl = tpl & "                    <div class=""card h-100 shadow-sm border-0"">" & vbCrLf
            tpl = tpl & "                        <div class=""card-body"">" & vbCrLf
            tpl = tpl & "                            <div class=""fs-1 mb-2"">📷</div>" & vbCrLf
            tpl = tpl & "                            <h5 class=""card-title"">1. Upload photos</h5>" & vbCrLf
            tpl = tpl & "                            <p class=""card-text text-secondary small"">Click <strong>Upload images</strong> in the top right. The AI automatically uses your photos in the design.</p>" & vbCrLf
            tpl = tpl & "                        </div>" & vbCrLf
            tpl = tpl & "                    </div>" & vbCrLf
            tpl = tpl & "                </div>" & vbCrLf
            tpl = tpl & "                <div class=""col-md-4"">" & vbCrLf
            tpl = tpl & "                    <div class=""card h-100 shadow-sm border-0"">" & vbCrLf
            tpl = tpl & "                        <div class=""card-body"">" & vbCrLf
            tpl = tpl & "                            <div class=""fs-1 mb-2"">💬</div>" & vbCrLf
            tpl = tpl & "                            <h5 class=""card-title"">2. Describe your site</h5>" & vbCrLf
            tpl = tpl & "                            <p class=""card-text text-secondary small"">Type what you want in the chat, e.g. <em>&quot;Build a modern site for my bakery with opening hours&quot;</em>.</p>" & vbCrLf
            tpl = tpl & "                        </div>" & vbCrLf
            tpl = tpl & "                    </div>" & vbCrLf
            tpl = tpl & "                </div>" & vbCrLf
            tpl = tpl & "                <div class=""col-md-4"">" & vbCrLf
            tpl = tpl & "                    <div class=""card h-100 shadow-sm border-0"">" & vbCrLf
            tpl = tpl & "                        <div class=""card-body"">" & vbCrLf
            tpl = tpl & "                            <div class=""fs-1 mb-2"">✨</div>" & vbCrLf
            tpl = tpl & "                            <h5 class=""card-title"">3. AI builds it</h5>" & vbCrLf
            tpl = tpl & "                            <p class=""card-text text-secondary small"">The AI writes the complete website. Not happy? Just ask for changes — or restore a previous version via <strong>History</strong>.</p>" & vbCrLf
            tpl = tpl & "                        </div>" & vbCrLf
            tpl = tpl & "                    </div>" & vbCrLf
            tpl = tpl & "                </div>" & vbCrLf
            tpl = tpl & "            </div>" & vbCrLf
            tpl = tpl & "            <p class=""text-secondary small mt-5 mb-0"">💡 Tip: the more specific your description (goal, audience, colors, style), the better the result.</p>" & vbCrLf
            tpl = tpl & "        </div>" & vbCrLf
            tpl = tpl & "    </div>" & vbCrLf
            tpl = tpl & "    <footer class=""py-4 mt-5 border-top bg-white"">" & vbCrLf
            tpl = tpl & "        <div class=""container text-center small text-secondary"">" & vbCrLf
            tpl = tpl & "            <span>&copy; " & Year(Date) & " " & projectName & "</span> &middot; " & vbCrLf
            tpl = tpl & "            <a href=""#"" class=""text-decoration-none"" data-bs-toggle=""modal"" data-bs-target=""#privacyModal"">Privacy</a> &middot; " & vbCrLf
            tpl = tpl & "            <a href=""#"" class=""text-decoration-none"" data-bs-toggle=""modal"" data-bs-target=""#cookieModal"">Cookies</a>" & vbCrLf
            tpl = tpl & "        </div>" & vbCrLf
            tpl = tpl & "    </footer>" & vbCrLf
            tpl = tpl & "    <!-- Cookie notice - shows once via localStorage (Bootstrap modal) -->" & vbCrLf
            tpl = tpl & "    <div class=""modal fade"" id=""cookieModal"" tabindex=""-1"" aria-labelledby=""cookieModalLabel"" aria-hidden=""true"">" & vbCrLf
            tpl = tpl & "      <div class=""modal-dialog modal-dialog-centered"">" & vbCrLf
            tpl = tpl & "        <div class=""modal-content"">" & vbCrLf
            tpl = tpl & "          <div class=""modal-header"">" & vbCrLf
            tpl = tpl & "            <h5 class=""modal-title"" id=""cookieModalLabel"">Cookies</h5>" & vbCrLf
            tpl = tpl & "            <button type=""button"" class=""btn-close"" data-bs-dismiss=""modal"" aria-label=""Close""></button>" & vbCrLf
            tpl = tpl & "          </div>" & vbCrLf
            tpl = tpl & "          <div class=""modal-body"">" & vbCrLf
            tpl = tpl & "            <p class=""small text-secondary mb-2"">We use only necessary cookies to make this site work (functional / localStorage for your preferences). No tracking or marketing cookies are used. If embedded content (maps, video, analytics) is added later, it may set its own cookies — see Privacy for details.</p>" & vbCrLf
            tpl = tpl & "            <p class=""small text-secondary mb-0"">You can change your choice anytime via the Cookies link in the footer.</p>" & vbCrLf
            tpl = tpl & "          </div>" & vbCrLf
            tpl = tpl & "          <div class=""modal-footer"">" & vbCrLf
            tpl = tpl & "            <button type=""button"" class=""btn btn-secondary btn-sm"" id=""cookieReject"" data-bs-dismiss=""modal"">Reject</button>" & vbCrLf
            tpl = tpl & "            <button type=""button"" class=""btn btn-primary btn-sm"" id=""cookieAccept"">Accept</button>" & vbCrLf
            tpl = tpl & "          </div>" & vbCrLf
            tpl = tpl & "        </div>" & vbCrLf
            tpl = tpl & "      </div>" & vbCrLf
            tpl = tpl & "    </div>" & vbCrLf
            tpl = tpl & "    <!-- Privacy statement (Bootstrap modal) -->" & vbCrLf
            tpl = tpl & "    <div class=""modal fade"" id=""privacyModal"" tabindex=""-1"" aria-labelledby=""privacyModalLabel"" aria-hidden=""true"">" & vbCrLf
            tpl = tpl & "      <div class=""modal-dialog modal-lg modal-dialog-scrollable"">" & vbCrLf
            tpl = tpl & "        <div class=""modal-content"">" & vbCrLf
            tpl = tpl & "          <div class=""modal-header"">" & vbCrLf
            tpl = tpl & "            <h5 class=""modal-title"" id=""privacyModalLabel"">Privacy statement</h5>" & vbCrLf
            tpl = tpl & "            <button type=""button"" class=""btn-close"" data-bs-dismiss=""modal"" aria-label=""Close""></button>" & vbCrLf
            tpl = tpl & "          </div>" & vbCrLf
            tpl = tpl & "          <div class=""modal-body small"">" & vbCrLf
            tpl = tpl & "            <p><strong>Data controller:</strong> " & projectName & ".</p>" & vbCrLf
            tpl = tpl & "            <p><strong>What we collect:</strong> only data you provide (e.g. via forms) and standard server logs (IP, user agent) for security. No data is sold.</p>" & vbCrLf
            tpl = tpl & "            <p><strong>Purpose & legal basis:</strong> to operate and secure the website (legitimate interest / contract) and to respond to your requests (consent).</p>" & vbCrLf
            tpl = tpl & "            <p><strong>Retention:</strong> server logs are kept briefly, form data only as needed to handle your request.</p>" & vbCrLf
            tpl = tpl & "            <p><strong>Your rights:</strong> you may request access, correction or deletion of your data. Contact us via the site.</p>" & vbCrLf
            tpl = tpl & "            <p class=""mb-0 text-secondary"">This is a starter template — adapt this text to your actual processing.</p>" & vbCrLf
            tpl = tpl & "          </div>" & vbCrLf
            tpl = tpl & "          <div class=""modal-footer"">" & vbCrLf
            tpl = tpl & "            <button type=""button"" class=""btn btn-secondary btn-sm"" data-bs-dismiss=""modal"">Close</button>" & vbCrLf
            tpl = tpl & "          </div>" & vbCrLf
            tpl = tpl & "        </div>" & vbCrLf
            tpl = tpl & "      </div>" & vbCrLf
            tpl = tpl & "    </div>" & vbCrLf
            tpl = tpl & "    <script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js""></script>" & vbCrLf
            tpl = tpl & "    <script>" & vbCrLf
            tpl = tpl & "    console.log('" & projectName & " - ready for AI generation');" & vbCrLf
            tpl = tpl & "    document.addEventListener('DOMContentLoaded', function() {" & vbCrLf
            tpl = tpl & "      try {" & vbCrLf
            tpl = tpl & "        if (!localStorage.getItem('cookieConsent')) {" & vbCrLf
            tpl = tpl & "          var el = document.getElementById('cookieModal');" & vbCrLf
            tpl = tpl & "          if (el && window.bootstrap) new bootstrap.Modal(el).show();" & vbCrLf
            tpl = tpl & "        }" & vbCrLf
            tpl = tpl & "      } catch(e) {}" & vbCrLf
            tpl = tpl & "      var acc = document.getElementById('cookieAccept');" & vbCrLf
            tpl = tpl & "      if (acc) acc.addEventListener('click', function() {" & vbCrLf
            tpl = tpl & "        try { localStorage.setItem('cookieConsent','accepted'); } catch(e) {}" & vbCrLf
            tpl = tpl & "        var m = document.getElementById('cookieModal'); if (m && window.bootstrap) { var inst = bootstrap.Modal.getInstance(m); if (inst) inst.hide(); }" & vbCrLf
            tpl = tpl & "      });" & vbCrLf
            tpl = tpl & "      var rej = document.getElementById('cookieReject');" & vbCrLf
            tpl = tpl & "      if (rej) rej.addEventListener('click', function() {" & vbCrLf
            tpl = tpl & "        try { localStorage.setItem('cookieConsent','rejected'); } catch(e) {}" & vbCrLf
            tpl = tpl & "      });" & vbCrLf
            tpl = tpl & "    });" & vbCrLf
            tpl = tpl & "    </script>" & vbCrLf
            tpl = tpl & "</body>" & vbCrLf
            tpl = tpl & "</html>" & vbCrLf
            WriteFileUtf8 projectDir & "index.html", tpl
            ' Cleanup legacy 3-file setup (single-file needs no external css/js)
            On Error Resume Next
            Dim fsoClean
            Set fsoClean = Server.CreateObject("Scripting.FileSystemObject")
            If fsoClean.FileExists(projectDir & "style.css") Then fsoClean.DeleteFile projectDir & "style.css", True
            If fsoClean.FileExists(projectDir & "script.js") Then fsoClean.DeleteFile projectDir & "script.js", True
            On Error Goto 0
End Sub

Function HandleDeleteProject()
    RequireAuth()
    
    Dim body, projectId, conn, rs, folderName, projectDir, fso
    Set body = ParseJsonBody()
    If IsNull(body) Then
        'Response.Write JsonError("Invalid JSON body", 400)
        'Exit Function
    End If
    
    projectId = body("id")
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT folder FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id"))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    folderName = rs("folder").Value
    rs.Close
    
    ' Delete children first to avoid FOREIGN KEY constraint failure (chat_messages has FK to projects)
    On Error Resume Next
    conn.Execute "DELETE FROM chat_messages WHERE project_id=" & CLng(projectId)
    conn.Execute "DELETE FROM prompts WHERE project_id=" & CLng(projectId)
    conn.Execute "DELETE FROM jobs WHERE project_id=" & CLng(projectId)
    Err.Clear
    On Error Goto 0
    ' Now delete parent
    On Error Resume Next
    conn.Execute "DELETE FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id")
    If Err.Number <> 0 Then
        ' FK still blocked (older rows with FK ON) – try disabling FK temporarily
        Err.Clear
        conn.Execute "PRAGMA foreign_keys=OFF"
        conn.Execute "DELETE FROM chat_messages WHERE project_id=" & CLng(projectId)
        conn.Execute "DELETE FROM prompts WHERE project_id=" & CLng(projectId)
        conn.Execute "DELETE FROM jobs WHERE project_id=" & CLng(projectId)
        conn.Execute "DELETE FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id")
        conn.Execute "PRAGMA foreign_keys=ON"
    End If
    On Error Goto 0
    conn.Close
    
    If Len(folderName) > 0 Then
        projectDir = GetProjectDir(folderName)
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(projectDir) Then
            DeleteFolderRecursive fso, projectDir
        End If
    End If
    
    Response.Write JsonOk("Project deleted")
End Function

Function HandleRenameProject()
    RequireAuth()
    
    Dim body, projectId, newName, conn, rs
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    projectId = body("id")
    If IsNull(projectId) Or Len(projectId) = 0 Then
        If body.Exists("project_id") Then projectId = body("project_id")
    End If
    newName = ""
    If body.Exists("name") Then newName = Trim(CStr(body("name")))
    If Len(newName) = 0 And body.Exists("new_name") Then newName = Trim(CStr(body("new_name")))
    If Len(projectId) = 0 Or Len(newName) = 0 Then
        Response.Write JsonError("project_id and new name required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT folder FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id"))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    rs.Close
    
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM projects WHERE user_id=" & Session("user_id") & " AND project_name=" & SqlEscape(newName) & " AND id<>" & CLng(projectId))
    If Not rs.EOF Then
        If CLng(rs("cnt").Value) > 0 Then
            rs.Close
            conn.Close
            Response.Write JsonError("A project with that name already exists", 409)
            Exit Function
        End If
    End If
    rs.Close
    
    conn.Execute "UPDATE projects SET project_name=" & SqlEscape(newName) & ", updated_at=datetime('now') WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id")
    conn.Close
    
    Response.Write JsonOk("Project renamed")
End Function

Function HandleUpdateNotes()
    RequireAuth()
    
    Dim body, projectId, notesVal, conn, rs
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    projectId = body("project_id")
    If Len(projectId) = 0 Then
        If body.Exists("id") Then projectId = body("id")
    End If
    notesVal = ""
    If body.Exists("notes") Then notesVal = CStr(body("notes"))
    
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT id FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id"))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    rs.Close
    
    On Error Resume Next
    conn.Execute "UPDATE projects SET notes=" & SqlEscape(notesVal) & ", updated_at=datetime('now') WHERE id=" & CLng(projectId)
    If Err.Number <> 0 Then
        Err.Clear
        ' column may not exist yet - try to add it then retry
        conn.Execute "ALTER TABLE projects ADD COLUMN notes TEXT NOT NULL DEFAULT ''"
        Err.Clear
        conn.Execute "UPDATE projects SET notes=" & SqlEscape(notesVal) & ", updated_at=datetime('now') WHERE id=" & CLng(projectId)
    End If
    On Error Goto 0
    conn.Close
    
    Dim r
    Set r = Server.CreateObject("Scripting.Dictionary")
    r.Add "message", "Notes saved"
    r.Add "notes", notesVal
    Response.Write JsonOk(r)
End Function

Function HandleGetNotes()
    RequireAuth()
    
    Dim projectId, conn, rs
    projectId = Request.QueryString("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    On Error Resume Next
    Set rs = conn.Execute("SELECT notes FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id"))
    If Err.Number <> 0 Then
        Err.Clear
        Set rs = conn.Execute("SELECT project_name FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id"))
        If rs.EOF Then
            rs.Close
            conn.Close
            Response.Write JsonError("Project not found", 404)
            Exit Function
        End If
        rs.Close
        conn.Close
        Dim emptyRes
        Set emptyRes = Server.CreateObject("Scripting.Dictionary")
        emptyRes.Add "notes", ""
        Response.Write JsonOk(emptyRes)
        On Error Goto 0
        Exit Function
    End If
    On Error Goto 0
    
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    Dim notesVal2, res2
    notesVal2 = rs("notes").Value
    If IsNull(notesVal2) Then notesVal2 = ""
    rs.Close
    conn.Close
    
    Set res2 = Server.CreateObject("Scripting.Dictionary")
    res2.Add "notes", CStr(notesVal2)
    Response.Write JsonOk(res2)
End Function

Function HandleGetChatHistory()
    RequireAuth()
    
    Dim projectId, folderName, conn, rs, msgs, idx
    projectId = Request.QueryString("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    Set conn = OpenDB()
    On Error Resume Next
    conn.Execute "CREATE TABLE IF NOT EXISTS chat_messages (id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, role TEXT NOT NULL, content TEXT NOT NULL, reasoning TEXT NOT NULL DEFAULT '', tokens_in INTEGER NOT NULL DEFAULT 0, tokens_out INTEGER NOT NULL DEFAULT 0, duration_seconds INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL DEFAULT (datetime('now')))"
    On Error Goto 0
    
    Set rs = conn.Execute("SELECT id, role, content, reasoning, tokens_in, tokens_out, duration_seconds, created_at FROM chat_messages WHERE project_id=" & CLng(projectId) & " ORDER BY id ASC")
    
    Set msgs = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    Do While Not rs.EOF
        Dim m
        Set m = Server.CreateObject("Scripting.Dictionary")
        m.Add "id", CLng(rs("id").Value)
        m.Add "role", rs("role").Value
        Dim cval
        cval = rs("content").Value
        If IsNull(cval) Then cval = ""
        m.Add "content", CStr(cval)
        Dim rval
        rval = rs("reasoning").Value
        If IsNull(rval) Then rval = ""
        m.Add "reasoning", CStr(rval)
        On Error Resume Next
        m.Add "tokens_in", CLng(rs("tokens_in").Value)
        m.Add "tokens_out", CLng(rs("tokens_out").Value)
        m.Add "duration_seconds", CLng(rs("duration_seconds").Value)
        If Err.Number <> 0 Then Err.Clear
        On Error Goto 0
        m.Add "created_at", rs("created_at").Value
        msgs.Add idx, m
        idx = idx + 1
        rs.MoveNext
    Loop
    rs.Close
    conn.Close
    
    Response.Write JsonOk(msgs)
End Function

Function HandleKeepAlive()
    RequireAuth()
    Session.Timeout = 60
    Dim d
    Set d = Server.CreateObject("Scripting.Dictionary")
    d.Add "message", "session alive"
    d.Add "timeout", Session.Timeout
    Response.Write JsonOk(d)
End Function

' Reset a project to its onboarding state: starter files, empty prompt
' history and no backups. Uploaded photos (img folder) are KEPT.
Function HandleResetProject()
    RequireAuth()
    
    Dim body, projectId, conn, rs, folderName, projectName, projectDir, fso
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    projectId = body("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT project_name, folder FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id"))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    projectName = rs("project_name").Value
    folderName = rs("folder").Value
    rs.Close
    
    ' Refuse while a generation is running
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM jobs WHERE project_id=" & CLng(projectId) & " AND status IN ('pending','running') AND cancel=0 AND created_at > datetime('now','-3 hours')")
    If CLng(rs("cnt").Value) > 0 Then
        rs.Close
        conn.Close
        Response.Write JsonError("A generation is running for this project. Wait until it finishes (or cancel it) before resetting.", 409)
        Exit Function
    End If
    rs.Close
    
    ' Clear the prompt history and job records so the AI starts fresh
    On Error Resume Next
    conn.Execute "DELETE FROM prompts WHERE project_id=" & CLng(projectId)
    conn.Execute "DELETE FROM jobs WHERE project_id=" & CLng(projectId)
    conn.Execute "DELETE FROM chat_messages WHERE project_id=" & CLng(projectId)
    Err.Clear
    On Error Goto 0
    conn.Execute "UPDATE projects SET updated_at=datetime('now') WHERE id=" & CLng(projectId)
    conn.Close
    
    projectDir = GetProjectDir(folderName)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    EnsureDir projectDir
    EnsureDir projectDir & "img" ' photos are kept
    EnsureDir projectDir & "backups"
    
    ' Remove all backups (each backup is a subfolder of backups\)
    Dim bFolder, subPaths, sf, idx2, k2
    If fso.FolderExists(projectDir & "backups") Then
        Set bFolder = fso.GetFolder(projectDir & "backups")
        Set subPaths = Server.CreateObject("Scripting.Dictionary")
        idx2 = 0
        For Each sf In bFolder.SubFolders
            subPaths.Add idx2, sf.Path
            idx2 = idx2 + 1
        Next
        For k2 = 0 To idx2 - 1
            DeleteFolderRecursive fso, subPaths(k2)
        Next
    End If
    
    ' Remove all generated html/css/js files in the project root
    Dim rootFolder, rf, filePaths, ext2
    Set rootFolder = fso.GetFolder(projectDir)
    Set filePaths = Server.CreateObject("Scripting.Dictionary")
    idx2 = 0
    For Each rf In rootFolder.Files
        ext2 = LCase(fso.GetExtensionName(rf.Name))
        If ext2 = "html" Or ext2 = "htm" Or ext2 = "css" Or ext2 = "js" Then
            filePaths.Add idx2, rf.Path
            idx2 = idx2 + 1
        End If
    Next
    For k2 = 0 To idx2 - 1
        fso.DeleteFile filePaths(k2), True
    Next
    
    ' Lay down the fresh onboarding files
    WriteStarterFiles projectDir, projectName
    
    Response.Write JsonOk("Project reset - your photos were kept")
End Function

' Recursively copy the contents of srcPath into dstPath. Subfolders named
' skipFolderName (top level only) are not copied.
Sub CopyDirContents(fso, srcPath, dstPath, skipFolderName)
    Dim folder, f, sub2
    If Not fso.FolderExists(dstPath) Then fso.CreateFolder dstPath
    Set folder = fso.GetFolder(srcPath)
    For Each f In folder.Files
        fso.CopyFile f.Path, dstPath & "\" & f.Name, True
    Next
    For Each sub2 In folder.SubFolders
        If Len(skipFolderName) = 0 Or LCase(sub2.Name) <> LCase(skipFolderName) Then
            CopyDirContents fso, sub2.Path, dstPath & "\" & sub2.Name, ""
        End If
    Next
End Sub

' Duplicate a project for the same user: files (including photos) and prompt
' history are copied, backups are NOT.
Function HandleCopyProject()
    RequireAuth()
    
    Dim body, projectId, conn, rs, srcName, srcFolder
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    projectId = body("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT project_name, folder FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id"))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    srcName = rs("project_name").Value
    srcFolder = rs("folder").Value
    rs.Close
    
    ' Find a free name: "name (copy)", "name (copy 2)", ...
    Dim newName, suffix, cnt
    newName = srcName & " (copy)"
    suffix = 2
    Do While suffix < 50
        Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM projects WHERE user_id=" & Session("user_id") & " AND project_name=" & SqlEscape(newName))
        cnt = CLng(rs("cnt").Value)
        rs.Close
        If cnt = 0 Then Exit Do
        newName = srcName & " (copy " & suffix & ")"
        suffix = suffix + 1
    Loop
    
    ' New random folder + project row
    Dim newFolder, newId, srcNotes
    srcNotes = ""
    On Error Resume Next
    Set rs = conn.Execute("SELECT notes FROM projects WHERE id=" & CLng(projectId))
    If Err.Number = 0 And Not rs.EOF Then
        If Not IsNull(rs("notes").Value) Then srcNotes = rs("notes").Value
    End If
    rs.Close
    Err.Clear
    On Error Goto 0
    newFolder = GenerateProjectFolder(conn)
    conn.Execute "INSERT INTO projects (user_id, project_name, folder, notes) VALUES (" & Session("user_id") & ", " & SqlEscape(newName) & ", " & SqlEscape(newFolder) & ", " & SqlEscape(srcNotes) & ")"
    Set rs = conn.Execute("SELECT last_insert_rowid() AS id")
    newId = CLng(rs("id").Value)
    rs.Close
    
    ' Copy the prompt history so the AI keeps its context in the copy
    On Error Resume Next
    conn.Execute "INSERT INTO prompts (project_id, prompt, created_at) SELECT " & newId & ", prompt, created_at FROM prompts WHERE project_id=" & CLng(projectId)
    conn.Execute "INSERT INTO chat_messages (project_id, role, content, reasoning, tokens_in, tokens_out, duration_seconds, created_at) SELECT " & newId & ", role, content, reasoning, tokens_in, tokens_out, duration_seconds, created_at FROM chat_messages WHERE project_id=" & CLng(projectId)
    Err.Clear
    On Error Goto 0
    
    ' Copy all files except the backups folder
    Dim fso, srcDir, dstDir
    srcDir = GetProjectDir(srcFolder)
    dstDir = GetProjectDir(newFolder)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(srcDir) Then
        CopyDirContents fso, Left(srcDir, Len(srcDir) - 1), Left(dstDir, Len(dstDir) - 1), "backups"
    End If
    EnsureDir dstDir & "backups"
    EnsureDir dstDir & "img"
    
    ' Return the new project
    Dim pData
    Set pData = Server.CreateObject("Scripting.Dictionary")
    Set rs = conn.Execute("SELECT id, project_name, folder, notes, created_at, updated_at FROM projects WHERE id=" & newId)
    If Err.Number <> 0 Then
        Err.Clear
        Set rs = conn.Execute("SELECT id, project_name, folder, created_at, updated_at FROM projects WHERE id=" & newId)
    End If
    If Not rs.EOF Then
        pData.Add "id", CLng(rs("id").Value)
        pData.Add "name", rs("project_name").Value
        pData.Add "folder", rs("folder").Value
        On Error Resume Next
        Dim cnotes
        cnotes = rs("notes").Value
        If Err.Number = 0 Then
            If IsNull(cnotes) Then cnotes = ""
            pData.Add "notes", CStr(cnotes)
        Else
            Err.Clear
            pData.Add "notes", srcNotes
        End If
        On Error Goto 0
        pData.Add "created_at", rs("created_at").Value
        pData.Add "updated_at", rs("updated_at").Value
    End If
    rs.Close
    conn.Close
    
    Response.Write JsonOk(pData)
End Function

Sub DeleteFolderRecursive(fso, path)
    Dim folder, file, subfolder
    If fso.FolderExists(path) Then
        Set folder = fso.GetFolder(path)
        For Each file In folder.Files
            file.Delete True
        Next
        For Each subfolder In folder.SubFolders
            DeleteFolderRecursive fso, subfolder.Path
        Next
        folder.Delete True
    End If
End Sub

Function HandleGetProjectFiles()
    RequireAuth()
    
    Dim projectId, folderName, projectDir, fso, files, idx
    projectId = Request.QueryString("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    projectDir = GetProjectDir(folderName)
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set files = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    
    If fso.FolderExists(projectDir) Then
        Dim folder, file
        Set folder = fso.GetFolder(projectDir)
        For Each file In folder.Files
            Dim ext
            ext = LCase(fso.GetExtensionName(file.Name))
            If ext = "html" Or ext = "css" Or ext = "js" Then
                Dim fileData, content
                Set fileData = Server.CreateObject("Scripting.Dictionary")
                fileData.Add "name", file.Name
                fileData.Add "size", file.Size
                fileData.Add "modified", CStr(file.DateLastModified)
                
                ' Read content
                content = ReadFileContent(fso, file.Path)
                If Not IsNull(content) Then
                    fileData.Add "content", content
                Else
                    fileData.Add "content", ""
                End If
                
                files.Add idx, fileData
                idx = idx + 1
            End If
        Next
    End If
    
    Response.Write JsonOk(files)
End Function

' Write a text file as UTF-8 (with BOM) via ADODB.Stream
Sub WriteFileUtf8(path, content)
    Dim stm
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.CharSet = "utf-8"
    stm.Open
    stm.WriteText content
    stm.SaveToFile path, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub

' Read a text file as UTF-8; falls back to FSO (legacy files); strips BOM
Function ReadFileContent(fso, path)
    On Error Resume Next
    Dim stm, content, f, bomCode
    content = ""
    
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.CharSet = "utf-8"
    stm.Open
    stm.LoadFromFile path
    content = stm.ReadText(-1)
    stm.Close
    
    If Err.Number <> 0 Then
        Err.Clear
        Set f = fso.OpenTextFile(path, 1, False, 0)
        If Err.Number <> 0 Then
            Err.Clear
            ReadFileContent = Null
            On Error Goto 0
            Exit Function
        End If
        content = f.ReadAll
        f.Close
    End If
    
    ' Strip a leading UTF-8 BOM character (U+FEFF)
    If Len(content) > 0 Then
        bomCode = AscW(Left(content, 1))
        If bomCode = -257 Or bomCode = 65279 Then
            content = Mid(content, 2)
        End If
    End If
    
    ReadFileContent = content
    On Error Goto 0
End Function

' =================== IMAGE UPLOAD ===================

Function HandleUploadImage()
    RequireAuth()
    
    Dim projectId, folderName, projectDir, fso
    projectId = Request.Form("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    projectDir = GetProjectDir(folderName)
    Dim imgDir
    imgDir = projectDir & "img\"
    EnsureDir imgDir
    
    If Request.Files.Count = 0 Then
        Response.Write JsonError("No file uploaded", 400)
        Exit Function
    End If
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    ' URL prefix for displaying thumbnails in the app (random folder, no username)
    Dim urlPrefix
    urlPrefix = GetProjectUrlPrefix(folderName) & "img/"
    
    Dim urls, idx, uploadedFile, origName, ext, safeName, savePath
    Set urls = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    
    For Each uploadedFile In Request.Files
        origName = uploadedFile.FileName
        If Len(origName) = 0 Then
            ' Skip, go to next file
        Else
            ext = LCase(fso.GetExtensionName(origName))
            If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "gif" Or ext = "webp" Or ext = "bmp" Or ext = "avif" Then
                safeName = MakeTimestamp() & "_" & SafeFileNamePart(fso.GetBaseName(origName)) & "." & ext
                savePath = imgDir & safeName
                uploadedFile.SaveAs savePath
                
                ' Resize if needed
                On Error Resume Next
                Dim imgObj
                Set imgObj = ASPPY.Image.Image.Open(savePath)
                If Err.Number = 0 Then
                    If imgObj.Width > 1920 Then
                        Dim ratio2, newHeight2
                        ratio2 = 1920.0 / CDbl(imgObj.Width)
                        newHeight2 = CInt(CDbl(imgObj.Height) * ratio2)
                        Dim newSz
                        newSz = Array(1920, newHeight2)
                        Dim resizedImg
                        Set resizedImg = imgObj.Resize(newSz)
                        resizedImg.Save savePath
                    End If
                End If
                On Error Goto 0
                
                urls.Add idx, urlPrefix & safeName
                idx = idx + 1
            End If
        End If
    Next
    
    If idx = 0 Then
        Response.Write JsonError("No valid image files uploaded. Allowed: jpg, jpeg, png, gif, webp, bmp, avif", 400)
        Exit Function
    End If
    
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "url", urls(0)
    result.Add "urls", urls
    result.Add "count", idx
    
    Response.Write JsonOk(result)
End Function

' =================== AI GENERATION (async job based) ===================

Function GetConfigValue(keyName)
    On Error Resume Next
    Dim conn, rs
    GetConfigValue = ""
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT value FROM config WHERE key=" & SqlEscape(keyName))
    If Not rs.EOF Then
        If Not IsNull(rs("value").Value) Then GetConfigValue = rs("value").Value
    End If
    rs.Close
    conn.Close
    On Error Goto 0
End Function

Function RandomToken()
    Dim chars, i, r
    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Randomize
    RandomToken = ""
    For i = 1 To 32
        r = Int(Rnd() * Len(chars)) + 1
        RandomToken = RandomToken & Mid(chars, r, 1)
    Next
End Function

Sub EnsureJobsTable(conn)
    conn.Execute "CREATE TABLE IF NOT EXISTS jobs (id INTEGER PRIMARY KEY AUTOINCREMENT, token TEXT NOT NULL, project_id INTEGER NOT NULL, project_name TEXT NOT NULL, username TEXT NOT NULL, folder TEXT NOT NULL DEFAULT '', prompt TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'pending', result TEXT DEFAULT '', notify INTEGER NOT NULL DEFAULT 0, email TEXT NOT NULL DEFAULT '', cancel INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL DEFAULT (datetime('now')))"
End Sub

' Returns True when the given job has been cancelled by the user
Function JobIsCancelled(jobId)
    On Error Resume Next
    Dim conn, rs
    JobIsCancelled = False
    If CLng(jobId) = 0 Then Exit Function
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT cancel, status FROM jobs WHERE id=" & CLng(jobId))
    If Not rs.EOF Then
        If CLng(rs("cancel").Value) = 1 Or rs("status").Value = "cancelled" Then
            JobIsCancelled = True
        End If
    End If
    rs.Close
    conn.Close
    On Error Goto 0
End Function

' Send the "generation finished" email if the user opted in for this job
Sub NotifyJobDone(jobId)
    On Error Resume Next
    Dim conn, rs, notify, email, projName, jStatus, uName
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT notify, email, project_name, username, status FROM jobs WHERE id=" & CLng(jobId))
    If rs.EOF Then
        rs.Close
        conn.Close
        Exit Sub
    End If
    notify = CLng(rs("notify").Value)
    email = rs("email").Value
    projName = rs("project_name").Value
    uName = rs("username").Value
    jStatus = rs("status").Value
    rs.Close
    conn.Close
    
    If notify <> 1 Or Len(email) = 0 Then Exit Sub
    If jStatus <> "done" And jStatus <> "error" Then Exit Sub
    
    Dim subject, bodyTpl, publicUrl, statusText
    publicUrl = GetConfigValue("public_url")
    If Len(publicUrl) = 0 Then publicUrl = "http://localhost:8080"
    If Right(publicUrl, 1) = "/" Then publicUrl = Left(publicUrl, Len(publicUrl) - 1)
    
    If jStatus = "done" Then
        statusText = "completed successfully"
    Else
        statusText = "failed"
    End If
    
    subject = GetConfigValue("notify_email_subject")
    If Len(subject) = 0 Then subject = "Your website generation has finished"
    bodyTpl = GetConfigValue("notify_email_body")
    If Len(bodyTpl) = 0 Then bodyTpl = "Hello {username}," & vbLf & vbLf & "The AI generation for your project ""{project}"" has finished with status: {status}." & vbLf & vbLf & "Open your project: {link}"
    bodyTpl = Replace(bodyTpl, "{username}", uName)
    bodyTpl = Replace(bodyTpl, "{project}", projName)
    bodyTpl = Replace(bodyTpl, "{status}", statusText)
    bodyTpl = Replace(bodyTpl, "{link}", publicUrl)
    
    SendMail email, subject, bodyTpl
    On Error Goto 0
End Sub

' Core generation logic, independent of Session/Response so the worker can run it.
' jobId is used to detect cancellation (pass 0 when not running as a job).
' Returns a Dictionary: ok (bool) + message/backup/files_written/reasoning OR raw/parsed OR error
Function RunGeneration(username, projectId, projectName, folderName, userPrompt, jobId)
    Dim res
    Set res = Server.CreateObject("Scripting.Dictionary")
    
    ' Get LLM config
    Dim conn, rs
    Dim llmEndpoint, llmModel, llmApiKey, systemPrompt, enableWikipedia, llmParams
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT key, value FROM config")
    Do While Not rs.EOF
        If rs("key").Value = "llm_endpoint" Then llmEndpoint = rs("value").Value
        If rs("key").Value = "llm_model" Then llmModel = rs("value").Value
        If rs("key").Value = "llm_api_key" Then llmApiKey = rs("value").Value
        If rs("key").Value = "system_prompt" Then systemPrompt = rs("value").Value
        If rs("key").Value = "enable_wikipedia" Then enableWikipedia = rs("value").Value
        If rs("key").Value = "llm_parameters" Then llmParams = rs("value").Value
        rs.MoveNext
    Loop
    rs.Close
    conn.Close
    
    ' Single-file mode: no section markers / patch mode anymore.
    Dim sectionsEnabled
    sectionsEnabled = False
    
    If Len(llmEndpoint) = 0 Then
        res.Add "ok", False
        res.Add "error", "LLM endpoint not configured. Admin must set it up first."
        Set RunGeneration = res
        Exit Function
    End If
    
    Dim projectDir, fso
    projectDir = GetProjectDir(folderName)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(projectDir) Then
        EnsureDir projectDir
        EnsureDir projectDir & "img"
        EnsureDir projectDir & "backups"
    End If
    
    ' Read existing project file (single-file: only index.html matters)
    Dim existingIndex, existingCss, existingJs, fc
    existingIndex = ""
    existingCss = ""
    existingJs = ""
    
    If fso.FileExists(projectDir & "index.html") Then
        fc = ReadFileContent(fso, projectDir & "index.html")
        If Not IsNull(fc) Then existingIndex = fc
    End If
    
    Dim fullPrompt, wikienabled
    wikienabled = (enableWikipedia = "1")
    
    ' Collect image paths (relative: img/...) from the project's img folder
    Dim imgUrls, imgDir
    imgDir = projectDir & "img\"
    imgUrls = GetImageUrls(imgDir)
    
    ' Load the last 5 prompts for this project (oldest first) as conversation context
    Dim promptHistory, hrs, histConn
    promptHistory = ""
    Set histConn = OpenDB()
    histConn.Execute "CREATE TABLE IF NOT EXISTS prompts (id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, prompt TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT (datetime('now')))"
    Set hrs = histConn.Execute("SELECT prompt FROM prompts WHERE project_id=" & CLng(projectId) & " ORDER BY id DESC LIMIT 5")
    Dim hIdx
    hIdx = 0
    Dim histLines(4)
    Do While Not hrs.EOF
        If hIdx <= 4 Then histLines(hIdx) = hrs("prompt").Value
        hIdx = hIdx + 1
        hrs.MoveNext
    Loop
    hrs.Close
    histConn.Close
    Dim hj
    For hj = hIdx - 1 To 0 Step -1
        If hj <= 4 Then
            promptHistory = promptHistory & "- " & histLines(hj) & vbCrLf
        End If
    Next
    
    fullPrompt = buildAiPrompt(userPrompt, systemPrompt, projectName, existingIndex, existingCss, existingJs, imgUrls, wikienabled, promptHistory, sectionsEnabled)
    
    ' LLM call with up to 2 retries (3 attempts total). Accepts a complete
    ' document first (even when it echoes old ID markers), otherwise merges
    ' echoed partial <!--ID:n--> blocks into the existing tagged document.
    ' Retries are adaptive: attempts 2-3 carry a repair hint describing the
    ' previous failure (truncation, unknown IDs, ...) instead of repeating
    ' the same prompt.
    Dim roundResult, totalStart, genSecs, attempt, gotContent, lastTransportErr
    Dim aiResponse, aiReasoning, newIndexContent, mergedHtml, partialEntries
    Dim legacyFiles, li, le, attemptPrompt, repairHint
    Dim lastFailKind, lastFailDetail, fullBroken, partUnknown, partMergedBroken, anchorsLost
    gotContent = False
    lastTransportErr = ""
    lastFailKind = ""
    lastFailDetail = ""
    newIndexContent = Null
    aiResponse = ""
    aiReasoning = ""
    totalStart = Timer
    For attempt = 1 To 3
        If attempt = 1 Or Len(lastFailKind) = 0 Then
            attemptPrompt = fullPrompt
        Else
            repairHint = BuildRepairHint(lastFailKind, lastFailDetail, existingIndex)
            If Len(repairHint) > 0 Then
                attemptPrompt = fullPrompt & vbCrLf & vbCrLf & repairHint
            Else
                attemptPrompt = fullPrompt
            End If
        End If
        Set roundResult = CallLLMSimple(llmEndpoint, llmModel, llmApiKey, attemptPrompt, wikienabled, llmParams)
        If roundResult("type") = "error" Then
            lastTransportErr = roundResult("error")
            If JobIsCancelled(jobId) Then Exit For
            If attempt = 3 Then Exit For
        Else
            ' Cancelled while the LLM was working? Discard, write nothing.
            If JobIsCancelled(jobId) Then
                lastTransportErr = ""
                Exit For
            End If
            ' Post-process every AI text: long dashes become plain hyphens,
            ' and absolute image paths are rewritten to relative ones (img/...)
            aiResponse = SanitizeAiText(roundResult("content"))
            aiResponse = MakeImagePathsRelative(aiResponse, folderName)
            aiReasoning = SanitizeAiText(roundResult("reasoning"))
            fullBroken = False
            partUnknown = ""
            partMergedBroken = False
            anchorsLost = False
            ' Parse the structured output: single ```html block (with <file> fallback)
            newIndexContent = ParseSingleFileHtml(aiResponse)
            If IsNull(newIndexContent) Then
                ' Legacy fallback: <file name="index.html">...</file>
                legacyFiles = ParseAIResponse(aiResponse)
                If Not IsNull(legacyFiles) Then
                    For li = 0 To legacyFiles.Count - 1
                        Set le = legacyFiles(li)
                        If LCase(le("name")) = "index.html" Then
                            newIndexContent = le("content")
                            Exit For
                        End If
                    Next
                    If IsNull(newIndexContent) And legacyFiles.Count > 0 Then
                        Set le = legacyFiles(0)
                        newIndexContent = le("content")
                    End If
                End If
            End If
            If Not IsNull(newIndexContent) Then
                If Len(Trim(newIndexContent)) > 0 Then
                    If HtmlSkeletonOk(newIndexContent) Then
                        gotContent = True
                        Exit For
                    Else
                        fullBroken = True ' complete doc attempted but cut off
                        newIndexContent = Null
                    End If
                Else
                    newIndexContent = Null
                End If
            End If
            ' Partial path: merge echoed ID blocks into the tagged base
            If HasIdMarkers(existingIndex) Then
                partialEntries = ParsePartialBlocks(aiResponse)
                If IsArray(partialEntries) Then
                    mergedHtml = MergeIdBlocks(existingIndex, partialEntries)
                    If Not IsNull(mergedHtml) Then
                        mergedHtml = MakeImagePathsRelative(mergedHtml, folderName)
                        If HtmlSkeletonOk(mergedHtml) Then
                            newIndexContent = mergedHtml
                            gotContent = True
                            Exit For
                        Else
                            partMergedBroken = True
                        End If
                    Else
                        partUnknown = FindUnknownIds(partialEntries, existingIndex)
                        If Len(partUnknown) = 0 Then anchorsLost = True
                    End If
                End If
            End If
            newIndexContent = Null
            ' Failure classification for the next attempt's repair hint.
            ' A partial-style answer (with ID markers) ALWAYS fails the full
            ' skeleton check, so "truncated" only applies to markerless answers.
            If HasIdMarkers(aiResponse) Then
                If partMergedBroken Then
                    lastFailKind = "merged_broken"
                    lastFailDetail = ""
                ElseIf Len(partUnknown) > 0 Then
                    lastFailKind = "badids"
                    lastFailDetail = partUnknown
                ElseIf anchorsLost Then
                    lastFailKind = "anchors"
                    lastFailDetail = ""
                Else
                    lastFailKind = "noblocks"
                    lastFailDetail = ""
                End If
            ElseIf fullBroken Then
                lastFailKind = "truncated"
                lastFailDetail = ""
            Else
                lastFailKind = "noblocks"
                lastFailDetail = ""
            End If
            If attempt = 3 Then Exit For
        End If
    Next
    genSecs = Timer - totalStart
    If genSecs < 0 Then genSecs = genSecs + 86400 ' crossed midnight
    genSecs = CLng(genSecs)

    ' Cancelled while the LLM was working? Discard the result, write nothing.
    If JobIsCancelled(jobId) Then
        res.Add "ok", False
        res.Add "cancelled", True
        res.Add "error", "Generation cancelled by user"
        Set RunGeneration = res
        Exit Function
    End If

    If Len(lastTransportErr) > 0 And Not gotContent Then
        res.Add "ok", False
        res.Add "error", lastTransportErr
        res.Add "fail_kind", "transport"
        Set RunGeneration = res
        Exit Function
    End If

    If IsNull(newIndexContent) Then
        res.Add "ok", True
        res.Add "raw", aiResponse
        res.Add "reasoning", aiReasoning
        res.Add "parsed", False
        res.Add "fail_kind", lastFailKind
        res.Add "fail_detail", lastFailDetail
        Set RunGeneration = res
        Exit Function
    End If
    If Len(Trim(newIndexContent)) = 0 Then
        res.Add "ok", True
        res.Add "raw", aiResponse
        res.Add "reasoning", aiReasoning
        res.Add "parsed", False
        res.Add "fail_kind", lastFailKind
        res.Add "fail_detail", lastFailDetail
        Set RunGeneration = res
        Exit Function
    End If

    ' Safety net: only accept complete HTML documents (full or merged)
    If Not HtmlSkeletonOk(newIndexContent) Then
        res.Add "ok", False
        res.Add "error", "The AI returned a broken or incomplete index.html (the page structure would be lost). Nothing was changed. Please try again, or rephrase the request."
        res.Add "reasoning", aiReasoning
        res.Add "fail_kind", lastFailKind
        res.Add "raw_excerpt", Left(aiResponse, 1500)
        Set RunGeneration = res
        Exit Function
    End If

    ' Autonome owns all IDs: strip any echoed markers and re-tag the whole
    ' document with fresh <!--ID:n-->...<!--/ID:n--> numbers.
    newIndexContent = TagDocument(StripIdMarkers(newIndexContent))
    If Not HtmlSkeletonOk(newIndexContent) Then
        res.Add "ok", False
        res.Add "error", "The AI returned a broken or incomplete index.html (the page structure would be lost). Nothing was changed. Please try again, or rephrase the request."
        res.Add "reasoning", aiReasoning
        Set RunGeneration = res
        Exit Function
    End If
    
    ' Create backup of the current index.html
    Dim backupDir, ts
    ts = MakeTimestamp()
    backupDir = projectDir & "backups\" & ts & "\"
    EnsureDir backupDir
    
    If fso.FileExists(projectDir & "index.html") Then
        fso.CopyFile projectDir & "index.html", backupDir & "index.html"
    End If
    
    ' Write the single self-contained file
    Dim fileCount, patchCount
    fileCount = 1
    patchCount = 0
    WriteFileUtf8 projectDir & "index.html", newIndexContent
    
    ' Remove legacy external files: single-file output must not reference them
    On Error Resume Next
    If fso.FileExists(projectDir & "style.css") Then fso.DeleteFile projectDir & "style.css", True
    If fso.FileExists(projectDir & "script.js") Then fso.DeleteFile projectDir & "script.js", True
    On Error Goto 0
    
    Dim nic
    nic = newIndexContent
    
    ' Social sharing (Facebook / WhatsApp / LinkedIn): inject Open Graph tags
    ' with ABSOLUTE urls so link previews show the site's photo.
    Dim ogHtml, ogPublic, ogSite, ogImg, ogResult
    ogHtml = ""
    If fso.FileExists(projectDir & "index.html") Then
        nic = ReadFileContent(fso, projectDir & "index.html")
        If Not IsNull(nic) Then ogHtml = nic
    End If
    If Len(ogHtml) > 0 Then
        ogPublic = GetConfigValue("public_url")
        If Len(ogPublic) = 0 Then ogPublic = "http://localhost:8080"
        If Right(ogPublic, 1) = "/" Then ogPublic = Left(ogPublic, Len(ogPublic) - 1)
        ogSite = ogPublic & GetProjectUrlPrefix(folderName)
        ogImg = FirstProjectImage(ogHtml, fso, projectDir)
        ogResult = InjectOpenGraphTags(ogHtml, projectName, ogPublic, ogSite, ogImg)
        If ogResult <> ogHtml Then WriteFileUtf8 projectDir & "index.html", ogResult
    End If

    ' Final guarantee: responsive safety net plus fresh sequential
    ' <!--ID:n--> markers on the file on disk. Guard first (it only touches
    ' <head>/<style>), then strip and re-tag the whole document.
    Dim finalHtml, finalTagged, guardedHtml
    finalHtml = ReadFileContent(fso, projectDir & "index.html")
    If Not IsNull(finalHtml) Then
        If Len(finalHtml) > 0 Then
            guardedHtml = EnsureResponsiveGuard(finalHtml)
            If guardedHtml <> finalHtml Then
                WriteFileUtf8 projectDir & "index.html", guardedHtml
                finalHtml = guardedHtml
            End If
            finalTagged = TagDocument(StripIdMarkers(finalHtml))
            If Len(finalTagged) > 0 Then
                If finalTagged <> finalHtml Then
                    If HtmlSkeletonOk(finalTagged) Then WriteFileUtf8 projectDir & "index.html", finalTagged
                End If
            End If
        End If
    End If
    
    ' Update project timestamp + record this prompt in the history
    Set conn = OpenDB()
    conn.Execute "UPDATE projects SET updated_at=datetime('now') WHERE id=" & CLng(projectId)
    conn.Execute "INSERT INTO prompts (project_id, prompt) VALUES (" & CLng(projectId) & ", " & SqlEscape(userPrompt) & ")"
    On Error Resume Next
    conn.Execute "CREATE TABLE IF NOT EXISTS chat_messages (id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, role TEXT NOT NULL, content TEXT NOT NULL, reasoning TEXT NOT NULL DEFAULT '', tokens_in INTEGER NOT NULL DEFAULT 0, tokens_out INTEGER NOT NULL DEFAULT 0, duration_seconds INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL DEFAULT (datetime('now')))"
    ' Ensure user message exists (HandleGenerateWebsite already inserted it, but do not duplicate)
    Dim chatCheck, hasUserMsg
    hasUserMsg = False
    Set chatCheck = conn.Execute("SELECT COUNT(*) AS cnt FROM chat_messages WHERE project_id=" & CLng(projectId) & " AND role='user' AND content=" & SqlEscape(userPrompt) & " AND datetime(created_at) >= datetime('now','-1 minute')")
    If Not chatCheck.EOF Then
        If CLng(chatCheck("cnt").Value) > 0 Then hasUserMsg = True
    End If
    chatCheck.Close
    If Not hasUserMsg Then
        conn.Execute "INSERT INTO chat_messages (project_id, role, content) VALUES (" & CLng(projectId) & ", 'user', " & SqlEscape(userPrompt) & ")"
    End If
    On Error Goto 0
    conn.Close
    
    ' Build stats suffix: token usage (when reported by the LLM) + duration
    Dim tokIn, tokOut, statsTxt, durTxt, tpsTxt
    tokIn = 0
    tokOut = 0
    If roundResult.Exists("tokens_in") Then tokIn = CLng(roundResult("tokens_in"))
    If roundResult.Exists("tokens_out") Then tokOut = CLng(roundResult("tokens_out"))
    If genSecs >= 60 Then
        durTxt = (genSecs \ 60) & "m " & (genSecs Mod 60) & "s"
    Else
        durTxt = genSecs & "s"
    End If
    tpsTxt = ""
    If genSecs > 0 And tokOut > 0 Then
        Dim tpsVal
        tpsVal = Round(CDbl(tokOut) / CDbl(genSecs) * 10) / 10
        Dim tpsStr
        tpsStr = CStr(tpsVal)
        ' Ensure 1 decimal place
        If InStr(1, tpsStr, ".", 1) = 0 And InStr(1, tpsStr, ",", 1) = 0 Then
            tpsStr = tpsStr & ".0"
        End If
        tpsStr = Replace(tpsStr, ".", ",")
        tpsTxt = tpsStr & " t/s"
    End If
    If tokIn > 0 Or tokOut > 0 Then
        statsTxt = tokIn & " tokens in, " & tokOut & " tokens out, " & durTxt
        If Len(tpsTxt) > 0 Then statsTxt = statsTxt & ", " & tpsTxt
    Else
        statsTxt = durTxt
        If Len(tpsTxt) > 0 Then statsTxt = statsTxt & ", " & tpsTxt
    End If
    
    Dim msgTxt
    msgTxt = "Website generated successfully (" & statsTxt & ")"
    
    ' Persist assistant message to chat history
    Set conn = OpenDB()
    On Error Resume Next
    conn.Execute "INSERT INTO chat_messages (project_id, role, content, reasoning, tokens_in, tokens_out, duration_seconds) VALUES (" & CLng(projectId) & ", 'assistant', " & SqlEscape(msgTxt) & ", " & SqlEscape(aiReasoning) & ", " & CLng(tokIn) & ", " & CLng(tokOut) & ", " & CLng(genSecs) & ")"
    On Error Goto 0
    conn.Close
    
    res.Add "ok", True
    res.Add "message", msgTxt
    res.Add "backup", "backups/" & ts
    res.Add "files_written", fileCount
    res.Add "sections_patched", patchCount
    res.Add "reasoning", aiReasoning
    res.Add "tokens_in", tokIn
    res.Add "tokens_out", tokOut
    res.Add "duration_seconds", genSecs
    Set RunGeneration = res
End Function

' Spawn a detached worker process that calls generate_worker via localhost.
' Requires ASP_PY_ALLOW_PYTHON=1; returns False when unavailable (caller falls back to sync).
Function SpawnWorker(jobId, token)
    On Error Resume Next
    Dim internalUrl, workerUrl, pyCode, ret
    
    internalUrl = GetConfigValue("internal_url")
    If Len(internalUrl) = 0 Then internalUrl = "http://127.0.0.1:8080"
    If Right(internalUrl, 1) = "/" Then internalUrl = Left(internalUrl, Len(internalUrl) - 1)
    workerUrl = internalUrl & "/api.asp?action=generate_worker&job_id=" & jobId & "&token=" & token
    
    pyCode = "import subprocess, sys" & vbLf
    pyCode = pyCode & "child = 'import urllib.request,sys' + chr(10) + 'try:' + chr(10) + '    urllib.request.urlopen(sys.argv[1], timeout=28800).read()' + chr(10) + 'except Exception:' + chr(10) + '    pass'" & vbLf
    pyCode = pyCode & "kw = {}" & vbLf
    pyCode = pyCode & "if sys.platform == 'win32':" & vbLf
    pyCode = pyCode & "    kw['creationflags'] = 0x08000008" & vbLf
    pyCode = pyCode & "else:" & vbLf
    pyCode = pyCode & "    kw['start_new_session'] = True" & vbLf
    pyCode = pyCode & "subprocess.Popen([sys.executable, '-c', child, ASPPY_ARGS], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, **kw)" & vbLf
    pyCode = pyCode & "ASPPY_RETURN('ok')"
    
    ret = ASPPY.ExecutePython(pyCode, workerUrl, 20)
    
    If Err.Number <> 0 Then
        Err.Clear
        SpawnWorker = False
    ElseIf ret = "ok" Then
        SpawnWorker = True
    Else
        SpawnWorker = False
    End If
    On Error Goto 0
End Function

Function HandleGenerateWebsite()
    RequireAuth()
    
    Dim body, projectId, prompt, conn, rs, projectName, folderName, username, notifyFlag, userEmail
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    projectId = body("project_id")
    prompt = body("prompt")
    
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    If Len(prompt) = 0 Then
        Response.Write JsonError("prompt required", 400)
        Exit Function
    End If
    
    notifyFlag = 0
    If body.Exists("notify") Then
        If body("notify") = True Or CStr(body("notify")) = "1" Or CStr(body("notify")) = "True" Then notifyFlag = 1
    End If
    
    projectName = GetOwnedProjectName(projectId)
    If IsNull(projectName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    folderName = GetOwnedProjectFolder(projectId)
    username = Session("username")
    
    Set conn = OpenDB()
    EnsureJobsTable conn
    
    ' Refuse a new prompt while a generation for this project is still running
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM jobs WHERE project_id=" & CLng(projectId) & " AND status IN ('pending','running') AND cancel=0 AND created_at > datetime('now','-3 hours')")
    If CLng(rs("cnt").Value) > 0 Then
        rs.Close
        conn.Close
        Response.Write JsonError("A generation is already running for this project. Wait until it finishes (or cancel it) before submitting a new prompt.", 409)
        Exit Function
    End If
    rs.Close
    
    ' The user's email, for the optional completion notification
    userEmail = ""
    Set rs = conn.Execute("SELECT email FROM users WHERE id=" & Session("user_id"))
    If Not rs.EOF Then
        If Not IsNull(rs("email").Value) Then userEmail = rs("email").Value
    End If
    rs.Close
    
    ' Persist user message to chat history immediately (so reload shows it while generating)
    On Error Resume Next
    conn.Execute "CREATE TABLE IF NOT EXISTS chat_messages (id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, role TEXT NOT NULL, content TEXT NOT NULL, reasoning TEXT NOT NULL DEFAULT '', tokens_in INTEGER NOT NULL DEFAULT 0, tokens_out INTEGER NOT NULL DEFAULT 0, duration_seconds INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL DEFAULT (datetime('now')))"
    conn.Execute "INSERT INTO chat_messages (project_id, role, content) VALUES (" & CLng(projectId) & ", 'user', " & SqlEscape(prompt) & ")"
    On Error Goto 0

    ' Create a job record
    Dim token, jobId
    token = RandomToken()
    conn.Execute "INSERT INTO jobs (token, project_id, project_name, username, folder, prompt, status, notify, email) VALUES (" & SqlEscape(token) & ", " & CLng(projectId) & ", " & SqlEscape(projectName) & ", " & SqlEscape(username) & ", " & SqlEscape(folderName) & ", " & SqlEscape(prompt) & ", 'pending', " & notifyFlag & ", " & SqlEscape(userEmail) & ")"
    Set rs = conn.Execute("SELECT last_insert_rowid() AS id")
    jobId = CLng(rs("id").Value)
    rs.Close
    conn.Close
    Session.Timeout = 60
    
    ' Try to run async via a detached worker; fall back to synchronous
    If SpawnWorker(jobId, token) Then
        Dim jobResult
        Set jobResult = Server.CreateObject("Scripting.Dictionary")
        jobResult.Add "job_id", jobId
        Response.Write JsonOk(jobResult)
    Else
        ' Synchronous fallback (no ASP_PY_ALLOW_PYTHON): blocks until done
        Dim genRes
        Set genRes = RunGeneration(username, projectId, projectName, folderName, prompt, jobId)
        
        Set conn = OpenDB()
        If genRes.Exists("cancelled") Or JobIsCancelled(jobId) Then
            conn.Execute "UPDATE jobs SET status='cancelled', result=" & SqlEscape(ASPPY.JSON.Encode(genRes)) & " WHERE id=" & jobId
        ElseIf genRes("ok") Then
            conn.Execute "UPDATE jobs SET status='done', result=" & SqlEscape(ASPPY.JSON.Encode(genRes)) & " WHERE id=" & jobId
        Else
            conn.Execute "UPDATE jobs SET status='error', result=" & SqlEscape(ASPPY.JSON.Encode(genRes)) & " WHERE id=" & jobId
        End If
        conn.Close
        
        NotifyJobDone jobId
        
        If genRes("ok") Then
            Response.Write JsonOk(genRes)
        Else
            Response.Write JsonError(genRes("error"), 500)
        End If
    End If
End Function

' Worker endpoint: no session, authenticated via the job token (localhost call)
Function HandleGenerateWorker()
    Dim jobId, token, conn, rs
    jobId = Request.QueryString("job_id")
    token = Request.QueryString("token")
    
    If Len(jobId) = 0 Or Len(token) = 0 Then
        Response.Write JsonError("job_id and token required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    EnsureJobsTable conn
    Set rs = conn.Execute("SELECT id, token, project_id, project_name, username, folder, prompt, status FROM jobs WHERE id=" & CLng(jobId))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Job not found", 404)
        Exit Function
    End If
    
    If rs("token").Value <> token Then
        rs.Close
        conn.Close
        Response.Write JsonError("Invalid token", 403)
        Exit Function
    End If
    
    If rs("status").Value <> "pending" Then
        rs.Close
        conn.Close
        Response.Write JsonError("Job already processed", 409)
        Exit Function
    End If
    
    Dim wProjectId, wProjectName, wUsername, wFolder, wPrompt
    wProjectId = CLng(rs("project_id").Value)
    wProjectName = rs("project_name").Value
    wUsername = rs("username").Value
    wFolder = rs("folder").Value
    wPrompt = rs("prompt").Value
    rs.Close
    
    conn.Execute "UPDATE jobs SET status='running' WHERE id=" & CLng(jobId)
    conn.Close
    
    Dim genRes
    Set genRes = RunGeneration(wUsername, wProjectId, wProjectName, wFolder, wPrompt, CLng(jobId))
    
    Set conn = OpenDB()
    If genRes.Exists("cancelled") Or JobIsCancelled(jobId) Then
        conn.Execute "UPDATE jobs SET status='cancelled', result=" & SqlEscape(ASPPY.JSON.Encode(genRes)) & " WHERE id=" & CLng(jobId)
    ElseIf genRes("ok") Then
        conn.Execute "UPDATE jobs SET status='done', result=" & SqlEscape(ASPPY.JSON.Encode(genRes)) & " WHERE id=" & CLng(jobId)
    Else
        conn.Execute "UPDATE jobs SET status='error', result=" & SqlEscape(ASPPY.JSON.Encode(genRes)) & " WHERE id=" & CLng(jobId)
    End If
    conn.Close
    
    NotifyJobDone jobId
    
    Response.Write JsonOk("worker finished")
End Function

' Status polling endpoint for the frontend
Function HandleGenerateStatus()
    RequireAuth()
    
    Dim jobId, conn, rs
    jobId = Request.QueryString("job_id")
    If Len(jobId) = 0 Then
        Response.Write JsonError("job_id required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    EnsureJobsTable conn
    Set rs = conn.Execute("SELECT username, status, result FROM jobs WHERE id=" & CLng(jobId))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Job not found", 404)
        Exit Function
    End If
    
    If rs("username").Value <> Session("username") Then
        rs.Close
        conn.Close
        Response.Write JsonError("Access denied", 403)
        Exit Function
    End If
    
    Dim statusResult, jStatus, jResult
    jStatus = rs("status").Value
    jResult = rs("result").Value
    rs.Close
    conn.Close
    
    Set statusResult = Server.CreateObject("Scripting.Dictionary")
    statusResult.Add "status", jStatus
    If Len(jResult) > 0 Then
        On Error Resume Next
        statusResult.Add "result", ASPPY.JSON.Decode(jResult)
        On Error Goto 0
    End If
    
    Response.Write JsonOk(statusResult)
End Function

' Cancel an in-progress generation job (e.g. after accidentally pressing Enter)
Function HandleCancelGeneration()
    RequireAuth()
    
    Dim body, jobId, conn, rs, jStatus
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    jobId = body("job_id")
    If Len(jobId) = 0 Then
        Response.Write JsonError("job_id required", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    EnsureJobsTable conn
    Set rs = conn.Execute("SELECT username, status FROM jobs WHERE id=" & CLng(jobId))
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonError("Job not found", 404)
        Exit Function
    End If
    
    If rs("username").Value <> Session("username") Then
        rs.Close
        conn.Close
        Response.Write JsonError("Access denied", 403)
        Exit Function
    End If
    
    jStatus = rs("status").Value
    rs.Close
    
    If jStatus <> "pending" And jStatus <> "running" Then
        conn.Close
        Response.Write JsonError("Job is not running (status: " & jStatus & ")", 409)
        Exit Function
    End If
    
    conn.Execute "UPDATE jobs SET cancel=1 WHERE id=" & CLng(jobId)
    If jStatus = "pending" Then
        ' Not picked up yet: mark cancelled right away (the worker will refuse it)
        conn.Execute "UPDATE jobs SET status='cancelled' WHERE id=" & CLng(jobId)
    End If
    conn.Close
    
    Response.Write JsonOk("Generation cancelled. Any result from this job will be discarded.")
End Function

' Return the currently running generation job for a project (if any), so the
' UI can pick up polling again after the user navigated away and back.
Function HandleActiveJob()
    RequireAuth()
    
    Dim projectId, folderName, conn, rs
    projectId = Request.QueryString("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    Set conn = OpenDB()
    EnsureJobsTable conn
    Set rs = conn.Execute("SELECT id, status, prompt, notify FROM jobs WHERE project_id=" & CLng(projectId) & " AND username=" & SqlEscape(Session("username")) & " AND status IN ('pending','running') AND cancel=0 AND created_at > datetime('now','-3 hours') ORDER BY id DESC LIMIT 1")
    
    If rs.EOF Then
        rs.Close
        conn.Close
        Response.Write JsonOk(Null)
        Exit Function
    End If
    
    Dim job
    Set job = Server.CreateObject("Scripting.Dictionary")
    job.Add "job_id", CLng(rs("id").Value)
    job.Add "status", rs("status").Value
    job.Add "prompt", rs("prompt").Value
    job.Add "notify", CLng(rs("notify").Value)
    rs.Close
    conn.Close
    
    Response.Write JsonOk(job)
End Function

Function GetImageUrls(imgDir)
    On Error Resume Next
    Dim fso, folder, file, urls, ext, prefix
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    urls = ""
    If Not fso.FolderExists(imgDir) Then
        GetImageUrls = ""
        Exit Function
    End If
    ' RELATIVE paths (img/...) so generated sites can be hosted anywhere without broken links
    prefix = "img/"
    Set folder = fso.GetFolder(imgDir)
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "gif" Or ext = "webp" Or ext = "bmp" Or ext = "avif" Then
            If Len(urls) > 0 Then urls = urls & vbCrLf
            urls = urls & prefix & file.Name
        End If
    Next
    GetImageUrls = urls
    On Error Goto 0
End Function

Function buildAiPrompt(userPrompt, systemPrompt, projectName, existingIndex, existingCss, existingJs, imgUrls, wikienabled, promptHistory, sectionsEnabled)
    Dim sb
    sb = ""
    
    If Len(systemPrompt) > 0 Then
        sb = sb & systemPrompt & vbCrLf & vbCrLf
    End If
    
    sb = sb & "You build a complete website based on the user's description." & vbCrLf
    
    Dim dtNow, dayName2
    dtNow = Now()
    dayName2 = WeekdayName(Weekday(dtNow))
    sb = sb & "Current date and time: " & dayName2 & " " & Year(dtNow) & "-" & Right("0" & Month(dtNow), 2) & "-" & Right("0" & Day(dtNow), 2) & " " & Right("0" & Hour(dtNow), 2) & ":" & Right("0" & Minute(dtNow), 2) & vbCrLf & vbCrLf
    
    If Len(imgUrls) > 0 Then
        sb = sb & "AVAILABLE IMAGES (use these RELATIVE paths EXACTLY as written, e.g. src=""img/photo.jpg"". Never prefix them with a slash, folder or domain):" & vbCrLf
        sb = sb & imgUrls & vbCrLf & vbCrLf
    End If
    
    If wikienabled Then
        sb = sb & "WIKIPEDIA TOOLS AVAILABLE:" & vbCrLf
        sb = sb & "You have access to Wikipedia search tools. Use them to research topics and get accurate, up-to-date information for the website content. Available functions: search_wikipedia (by query) and get_wikipedia_page (by exact title)." & vbCrLf & vbCrLf
    End If
    
    sb = sb & "SINGLE-FILE HTML BUILDER:" & vbCrLf
    sb = sb & "You build complete, self-contained single-file websites, using Bootstrap 5 (latest version, via CDN)." & vbCrLf
    sb = sb & "ALWAYS output exactly ONE complete index.html document and nothing else." & vbCrLf
    ' FIXED: Explicitly specify CUSTOM css/js and carve out an exception for allowed CDNs
    sb = sb & "Put ALL CUSTOM CSS inside a single <style> tag in the <head>, and ALL CUSTOM JavaScript inside a single <script> tag before </body>. The ONLY external files allowed are the required Bootstrap, Icon, and Font CDNs." & vbCrLf
    sb = sb & "Output ONLY one fenced ```html code block: no explanations, no commentary, no design critique, and no multiple code blocks." & vbCrLf
    sb = sb & "The file must be ready to save and open directly in a browser." & vbCrLf & vbCrLf
    sb = sb & "SINGLE-FILE WEBSITE REFERENCE:" & vbCrLf
    sb = sb & "- Output EXACTLY ONE complete index.html in a single ```html block - no commentary, no extra blocks." & vbCrLf
    ' FIXED: Clarified custom vs external links
    sb = sb & "- ALL CUSTOM CSS in one <style> in <head>; ALL CUSTOM JS in one <script> before </body>. Never link external .css/.js files (except the allowed CDNs)." & vbCrLf
    sb = sb & "- Use Bootstrap 5 via CDN (CSS link in <head>, bundle JS before </body>); add Bootstrap Icons CDN if icons are needed." & vbCrLf
    sb = sb & "- Make it responsive (viewport meta + Bootstrap grid/utilities), accessible (semantic tags, alt text, labels), and ready to save-and-open." & vbCrLf
    sb = sb & "- Use Bootstrap 5 classes and components (navbar, cards, grid, carousel, etc.) as the foundation of the layout." & vbCrLf
    sb = sb & "- For UI elements, icons, and logos use inline SVG or Bootstrap Icons." & vbCrLf
    sb = sb & "- Write ALL website content (headings, paragraphs, buttons, labels) in the SAME LANGUAGE as the USER REQUEST below." & vbCrLf & vbCrLf
    sb = sb & "IMAGE STRATEGY:" & vbCrLf
    sb = sb & "1. In case uploaded imagelinks are provided below, use them (RELATIVE paths EXACTLY as written, e.g. src=""img/photo.jpg"")!" & vbCrLf
    sb = sb & "2. In case NO uploaded images are provided, or only 1 or 2, attempt to use real, working URLs for public domain or hotlink-friendly photos where contextually appropriate." & vbCrLf
    sb = sb & "3. If real images are unavailable, use reliable placeholders like https://picsum.photos/800/600 or https://placehold.co/800x600." & vbCrLf
    sb = sb & "4. ALWAYS include an onerror attribute on every <img> tag as a fallback for broken links (e.g., onerror=""this.onerror=null;this.src='https://placehold.co/800x600?text=Image+Not+Found';"")." & vbCrLf
    sb = sb & "5. For UI elements, icons, and logos, use inline SVGs or Bootstrap Icons." & vbCrLf
    sb = sb & "6. For gallery-sections, use lightbox to display an image or video in a full-screen overlay." & vbCrLf & vbCrLf
    sb = sb & "DESIGN GUIDELINES:" & vbCrLf
    sb = sb & "- Start with a strong hero section; use an uploaded image there if available" & vbCrLf
    sb = sb & "- Consistent spacing: py-5 for sections, generous whitespace, mb-4 between elements" & vbCrLf
    sb = sb & "- One cohesive color palette: a single primary accent color, neutrals for the rest" & vbCrLf
    sb = sb & "- Max 2 fonts; Google Fonts are allowed via a <link> tag in index.html" & vbCrLf
    sb = sb & "- Cards with subtle shadows (shadow-sm) and rounded corners; add gentle hover effects in the inline <style> block" & vbCrLf
    sb = sb & "- Images: use object-fit cover and fixed heights so different photo sizes look uniform" & vbCrLf
    sb = sb & "- Always end the page with a proper footer (site name, copyright with the current year)" & vbCrLf & vbCrLf

    sb = sb & "MOBILE-FIRST (mandatory - mentally test at 360px width):" & vbCrLf
    sb = sb & "- Fluid headings with clamp(), e.g. h1: clamp(2rem, 8vw, 4.5rem); never fixed pixel font-sizes above 40px." & vbCrLf
    sb = sb & "- Hero must fit small screens: min-height with svh units, generous top padding so a fixed/sticky navbar NEVER overlaps the title; buttons wrap underneath each other when narrow." & vbCrLf
    sb = sb & "- Every Bootstrap .row sits inside a .container/.container-fluid (bare rows cause horizontal overflow); no fixed pixel widths on mobile; avoid 100vw full-bleed." & vbCrLf
    sb = sb & "- Buttons/links wrap (flex-wrap), stay >=44px tall and inside the viewport; images always max-width:100% + height:auto (or Bootstrap img-fluid); long headings break with overflow-wrap." & vbCrLf
    sb = sb & "- A small responsive safety net (viewport, overflow-x guard) is injected automatically, but author mobile-first CSS yourself anyway." & vbCrLf & vbCrLf
    
    sb = sb & "SEO & ACCESSIBILITY (mandatory):" & vbCrLf
    sb = sb & "- Every <img> must have a meaningful alt attribute" & vbCrLf
    sb = sb & "- Include a descriptive <title> and <meta name=""description""> in <head>" & vbCrLf
    sb = sb & "- Exactly ONE <h1> per page; headings in logical order (h1 then h2 then h3)" & vbCrLf
    sb = sb & "- Ensure sufficient color contrast between text and background" & vbCrLf
    sb = sb & "- Use aria-label on icon-only buttons and nav toggles" & vbCrLf
    sb = sb & "- Do NOT add Open Graph (og:) or twitter: meta tags - the system injects them automatically with correct absolute URLs" & vbCrLf & vbCrLf

    sb = sb & "LEGAL / COOKIE & PRIVACY (mandatory, always include in EVERY generation, in the SAME LANGUAGE as the website):" & vbCrLf
    sb = sb & "- Cookie notice: add a Bootstrap 5 modal with id=""cookieModal"" (class=""modal fade"", tabindex=""-1""). It MUST show only once: on DOMContentLoaded check if !localStorage.getItem('cookieConsent') then new bootstrap.Modal(document.getElementById('cookieModal')).show(). On Accept/Reject/Close store localStorage.setItem('cookieConsent','accepted'|'rejected'|'dismissed') and never show again. Modal body must list cookies actually used on the site (if any): at minimum functional/localStorage, and if the design embeds analytics, maps, video, fonts etc. mention them explicitly; otherwise explicitly state that no tracking/marketing cookies are used. Include two buttons: Accept (id=""cookieAccept"") and Reject/Close (id=""cookieReject""). All texts in the website language. Put the JS logic in the inline <script> block before </body>." & vbCrLf
    sb = sb & "- Privacy statement: add a second Bootstrap 5 modal with id=""privacyModal"" containing a concise but complete privacy statement (data controller, what is collected, purpose, legal basis, retention, user rights, contact). Also in the website language." & vbCrLf
    ' FIXED: Added the word "custom" before JS here as well for consistency.
    sb = sb & "- Footer MUST contain two links/buttons that open the modals: e.g. <a href=""#"" data-bs-toggle=""modal"" data-bs-target=""#privacyModal"">Privacy</a> and <a href=""#"" data-bs-toggle=""modal"" data-bs-target=""#cookieModal"">Cookies / Cookie settings</a>. All custom CSS goes in the single inline <style> tag in <head>; all custom JS goes in the single inline <script> tag before </body>." & vbCrLf & vbCrLf
    
    If Len(promptHistory) > 0 Then
        sb = sb & "PREVIOUS REQUESTS for this website (chronological, for context - the current request builds on these):" & vbCrLf
        sb = sb & promptHistory & vbCrLf
    End If
    
    ' Single-file mode: always send the current index.html so the AI can
    ' improve upon it. Legacy style.css/script.js are ignored (single-file
    ' output is self-contained).
    If Len(existingIndex) > 0 Then
        sb = sb & "EXISTING WEBSITE CODE (improve upon this, do NOT start from scratch unless asked):" & vbCrLf & vbCrLf
        sb = sb & "--- Current index.html ---" & vbCrLf
        sb = sb & existingIndex & vbCrLf & vbCrLf
    End If
    
    sb = sb & "USER REQUEST: " & userPrompt & vbCrLf & vbCrLf
    If Len(existingIndex) > 0 And HasIdMarkers(existingIndex) Then
        sb = sb & "BLOCK INVENTORY (block ID: content preview - use this to pick the RIGHT id):" & vbCrLf
        sb = sb & BuildBlockInventory(existingIndex) & vbCrLf
        sb = sb & "BLOCK UPDATE PROTOCOL (the current code above contains <!--ID:n-->...<!--/ID:n--> markers owned by the system):" & vbCrLf
        sb = sb & "- Change ONLY what the USER REQUEST asks. Leave all other blocks untouched." & vbCrLf
        sb = sb & "- If the requested element already exists on the page (e.g. a gallery, a contact form), MODIFY that block instead of adding a duplicate." & vbCrLf
        sb = sb & "- PREFER partial blocks: return ONLY the changed blocks, each as its own fenced ```html block that starts with the echoed <!--ID:n--> marker and ends with the matching <!--/ID:n--> marker. Echo the numbers EXACTLY as received; never invent numbers. Short answers are more reliable than long ones." & vbCrLf
        sb = sb & "- Decide per block: change (return the block with new inner HTML), delete (return the block with empty inner HTML or the word DELETE), add (return a larger replacement of the parent block, or a new <!--ID:NEW-->...<!--/ID:NEW--> block with a placement anchor as its first line)." & vbCrLf
        sb = sb & "- Placement anchors for NEW blocks: <!--AFTER:n--> inserts after block n, <!--AFTER:0--> or <!--FIRST--> inserts as the first block in <body>, <!--INTO:n--> appends inside block n (e.g. a new card into a cards row), <!--LAST--> or no anchor appends before </body>. Example: <!--ID:NEW--> <!--AFTER:9--> <section>...</section> <!--/ID:NEW-->." & vbCrLf
        sb = sb & "- Move: to RELOCATE a block, echo it with its ID, the (optionally new) content, and a placement anchor as first line (e.g. gallery under the hero: echo the gallery block starting with <!--AFTER:2-->). Never use NEW for moves and never duplicate the block." & vbCrLf
        sb = sb & "- Copy image paths EXACTLY as listed under AVAILABLE IMAGES above; never invent or shorten filenames." & vbCrLf
        sb = sb & "- Theme, colors, fonts or other styling: change ONLY the head <style> block and return it with its echoed ID (same <!--ID:n--> markers, CSS inside unchanged in structure)." & vbCrLf
        sb = sb & "- Do NOT output CSS/JS comment markers. If you return a larger block that contains other <!--ID:n--> markers, the outer block wins and inner numbers are reassigned by the system. Blocks can nest (a section block may contain card blocks): always echo the SMALLEST block that covers your change. No text outside the code blocks." & vbCrLf
        sb = sb & "If the request affects most of the page, you may instead return exactly one fenced ```html block with the COMPLETE index.html (without any ID markers). Prefer partial blocks whenever possible."
    Else
        sb = sb & "Now write the website. Your answer MUST be exactly one fenced ```html block containing the COMPLETE index.html from <!DOCTYPE html> to </html>. No text outside the code block."
    End If
    
    buildAiPrompt = sb
End Function

' Agentic loop: supports multiple rounds of tool calls + reasoning continuations.
' Messages are built as a VBScript array of Dictionaries and encoded via
' ASPPY.JSON.Encode so a VBScript array becomes a real JSON array.
' ---- Simple single-shot LLM helpers ----

Function LlmHttpPost(url, apiKey, bodyStr)
    On Error Resume Next
    Dim xmlhttp
    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    If InStr(1, url, "anthropic.com", 1) > 0 Then
        If Len(apiKey) > 0 Then
            xmlhttp.setRequestHeader "x-api-key", apiKey
        End If
        xmlhttp.setRequestHeader "anthropic-version", "2023-06-01"
    Else
        If Len(apiKey) > 0 Then
            xmlhttp.setRequestHeader "Authorization", "Bearer " & apiKey
        End If
    End If
    ' Wait up to 2 hours for the complete answer (slow reasoning models)
    xmlhttp.setTimeouts 28800000, 28800000, 28800000, 28800000
    xmlhttp.Send bodyStr
    
    If Err.Number <> 0 Then
        LlmHttpPost = "ERROR: LLM request failed: " & Err.Description & " (URL: " & url & ")"
        Err.Clear
        On Error Goto 0
        Exit Function
    End If
    If xmlhttp.Status < 200 Or xmlhttp.Status >= 300 Then
        LlmHttpPost = "ERROR: HTTP " & xmlhttp.Status & " from LLM: " & Left(xmlhttp.responseText, 300)
        On Error Goto 0
        Exit Function
    End If
    LlmHttpPost = xmlhttp.responseText
    On Error Goto 0
End Function

Function BuildLlmReqBody(model, msgArrJson, enableTools, params)
    Dim sb
    sb = "{""model"": """ & model & """, ""stream"": false, ""messages"": " & msgArrJson
    If enableTools Then
        sb = sb & ", ""tools"": " & wikipediaToolsJson()
    End If
    If Len(params) > 0 Then
        sb = sb & ", " & params
        If InStr(1, params, "max_tokens", 1) = 0 Then
            sb = sb & ", ""max_tokens"": -1"
        End If
    Else
        sb = sb & ", ""max_tokens"": -1"
    End If
    sb = sb & "}"
    BuildLlmReqBody = sb
End Function

Function BuildAnthropicReqBody(model, prompt, params)
    Dim sb
    sb = "{""model"": """ & model & """"
    If Len(params) > 0 Then
        sb = sb & ", " & params
    End If
    sb = sb & ", ""messages"": [{""role"": ""user"", ""content"": """ & EscapeJsonString(prompt) & """}]"
    sb = sb & "}"
    BuildAnthropicReqBody = sb
End Function

Function IsAnthropicEndpoint(endpoint)
    IsAnthropicEndpoint = (InStr(1, endpoint, "anthropic.com", 1) > 0)
End Function

Function AnthropicExtractText(parsed)
    Dim out, arr, i, blk
    out = ""
    If Not IsObject(parsed) Then
        AnthropicExtractText = ""
        Exit Function
    End If
    If Not parsed.Exists("content") Then
        AnthropicExtractText = ""
        Exit Function
    End If
    Set arr = parsed("content")
    ' content is array-like Dictionary with numeric keys (0,1...) or Collection
    On Error Resume Next
    For i = 0 To arr.Count - 1
        Set blk = arr(i)
        If IsObject(blk) Then
            If blk.Exists("type") Then
                If blk("type") = "text" And blk.Exists("text") Then
                    If Not IsNull(blk("text")) Then out = out & CStr(blk("text"))
                End If
            End If
        End If
    Next
    ' Fallback: iterate via For Each in case Count is not reliable
    If Len(out) = 0 Then
        For Each blk In arr
            If IsObject(blk) Then
                If blk.Exists("text") Then
                    If Not IsNull(blk("text")) Then out = out & CStr(blk("text"))
                End If
            End If
        Next
    End If
    On Error Goto 0
    AnthropicExtractText = out
End Function

' Single-shot LLM call: send the prompt, wait for the complete answer.
' If the model requests tools, execute them and do ONE follow-up call.
' Falls back to reasoning_content when content is empty.
' Returns Dictionary: type = "final" | "error", content, reasoning, error
Function CallLLMSimple(endpoint, model, apiKey, prompt, enableTools, params)
    On Error Resume Next
    Dim res
    Set res = Server.CreateObject("Scripting.Dictionary")
    res.Add "reasoning", ""
    res.Add "tokens_in", 0
    res.Add "tokens_out", 0
    
    Dim finalEndpoint, url
    finalEndpoint = endpoint
    If Right(finalEndpoint, 1) = "/" Then
        finalEndpoint = Left(finalEndpoint, Len(finalEndpoint) - 1)
    End If

    ' --- Anthropic Messages API branch ---
    If IsAnthropicEndpoint(endpoint) Then
        ' Build correct URL for Anthropic (handle /v1/messages already in endpoint)
        If Right(LCase(finalEndpoint), Len("/v1/messages")) = "/v1/messages" Then
            url = finalEndpoint
        ElseIf Right(LCase(finalEndpoint), Len("/v1")) = "/v1" Then
            url = finalEndpoint & "/messages"
        Else
            url = finalEndpoint & "/v1/messages"
        End If
        Dim anthResp
        anthResp = LlmHttpPost(url, apiKey, BuildAnthropicReqBody(model, prompt, params))
        If Left(anthResp, 6) = "ERROR:" Then
            res.Add "type", "error"
            res.Add "error", Mid(anthResp, 8)
            Set CallLLMSimple = res
            On Error Goto 0
            Exit Function
        End If
        Dim anthParsed, anthContent
        Set anthParsed = ASPPY.JSON.Decode(anthResp)
        If Not IsObject(anthParsed) Then
            res.Add "type", "error"
            res.Add "error", "Invalid JSON from Anthropic: " & Left(anthResp, 400)
            Set CallLLMSimple = res
            On Error Goto 0
            Exit Function
        End If
        ' Anthropic error shape: {"type":"error","error":{"type":"...","message":"..."}}
        If anthParsed.Exists("type") Then
            If anthParsed("type") = "error" Then
                Dim errMsg
                errMsg = "Anthropic error"
                If anthParsed.Exists("error") Then
                    If IsObject(anthParsed("error")) Then
                        If anthParsed("error").Exists("message") Then errMsg = anthParsed("error")("message")
                    Else
                        errMsg = CStr(anthParsed("error"))
                    End If
                End If
                res.Add "type", "error"
                res.Add "error", errMsg & " (model=" & model & ", url=" & url & ")"
                Set CallLLMSimple = res
                On Error Goto 0
                Exit Function
            End If
        End If
        anthContent = AnthropicExtractText(anthParsed)
        If Len(anthContent) = 0 Then
            ' Fallback: try OpenAI-style choices for compatibility proxies
            If anthParsed.Exists("choices") Then
                On Error Resume Next
                anthContent = anthParsed("choices")(0)("message")("content")
                On Error Goto 0
            End If
        End If
        If Len(anthContent) = 0 Then
            res.Add "type", "error"
            res.Add "error", "Empty response from Anthropic (no content). Raw: " & Left(anthResp, 500)
            Set CallLLMSimple = res
            On Error Goto 0
            Exit Function
        End If
        ' Truncated by the output token limit? Fail clearly instead of writing half a website.
        If anthParsed.Exists("stop_reason") Then
            If Not IsNull(anthParsed("stop_reason")) Then
                If anthParsed("stop_reason") = "max_tokens" Then
                    res.Add "type", "error"
                    res.Add "error", "The response was cut off by the max_tokens limit (stop_reason=max_tokens). Nothing was changed. Increase max_tokens in the admin panel (Extra LLM Parameters) and try again."
                    Set CallLLMSimple = res
                    On Error Goto 0
                    Exit Function
                End If
            End If
        End If
        ' Token usage (Anthropic: usage.input_tokens / usage.output_tokens)
        If anthParsed.Exists("usage") Then
            If IsObject(anthParsed("usage")) Then
                If anthParsed("usage").Exists("input_tokens") Then
                    If Not IsNull(anthParsed("usage")("input_tokens")) Then res("tokens_in") = CLng(anthParsed("usage")("input_tokens"))
                End If
                If anthParsed("usage").Exists("output_tokens") Then
                    If Not IsNull(anthParsed("usage")("output_tokens")) Then res("tokens_out") = CLng(anthParsed("usage")("output_tokens"))
                End If
            End If
        End If
        res.Add "type", "final"
        res.Add "content", anthContent
        Set CallLLMSimple = res
        On Error Goto 0
        Exit Function
    End If

    url = finalEndpoint & "/v1/chat/completions"
    
    Dim q, msgArrJson
    q = Chr(34)
    msgArrJson = "[{" & q & "role" & q & ":" & q & "user" & q & "," & q & "content" & q & ":" & q & EscapeJsonString(prompt) & q & "}]"
    
    Dim resp
    resp = LlmHttpPost(url, apiKey, BuildLlmReqBody(model, msgArrJson, enableTools, params))
    If Left(resp, 6) = "ERROR:" Then
        res.Add "type", "error"
        res.Add "error", Mid(resp, 8)
        Set CallLLMSimple = res
        On Error Goto 0
        Exit Function
    End If
    
    Dim parsed, msg, content, reasoning, hasToolCalls
    Set parsed = ASPPY.JSON.Decode(resp)
    If Not IsObject(parsed) Then
        res.Add "type", "error"
        res.Add "error", "Invalid JSON from LLM: " & Left(resp, 300)
        Set CallLLMSimple = res
        On Error Goto 0
        Exit Function
    End If
    If Not parsed.Exists("choices") Then
        res.Add "type", "error"
        res.Add "error", "No choices in LLM response: " & Left(resp, 300)
        Set CallLLMSimple = res
        On Error Goto 0
        Exit Function
    End If
    
    Set msg = parsed("choices")(0)("message")
    content = ""
    If msg.Exists("content") Then
        If Not IsNull(msg("content")) Then content = msg("content")
    End If
    reasoning = ""
    If msg.Exists("reasoning_content") Then
        If Not IsNull(msg("reasoning_content")) Then reasoning = msg("reasoning_content")
    End If
    
    hasToolCalls = False
    If msg.Exists("tool_calls") Then
        If Not IsNull(msg("tool_calls")) Then
            Dim tcChk
            For Each tcChk In msg("tool_calls")
                hasToolCalls = True
                Exit For
            Next
        End If
    End If
    
    ' One follow-up round when the model wants tools
    If hasToolCalls Then
        Dim oldArr, n, i, tcCount, tc
        oldArr = ASPPY.JSON.Decode(msgArrJson)
        n = UBound(oldArr)
        
        tcCount = 0
        For Each tc In msg("tool_calls")
            tcCount = tcCount + 1
        Next
        
        Dim newArr()
        ReDim newArr(n + 1 + tcCount)
        For i = 0 To n
            Set newArr(i) = oldArr(i)
        Next
        
        Dim asstMsg
        Set asstMsg = Server.CreateObject("Scripting.Dictionary")
        asstMsg.Add "role", "assistant"
        asstMsg.Add "content", content
        asstMsg.Add "tool_calls", msg("tool_calls")
        Set newArr(n + 1) = asstMsg
        
        Dim k, tcId, tcName, tcArgs, execResult, toolMsg
        k = n + 2
        For Each tc In msg("tool_calls")
            tcId = tc("id")
            tcName = tc("function")("name")
            tcArgs = tc("function")("arguments")
            execResult = ExecuteWikipediaTool(tcName, tcArgs)
            reasoning = reasoning & vbCrLf & "[Tool: " & tcName & " " & Left(tcArgs, 150) & "]"
            
            Set toolMsg = Server.CreateObject("Scripting.Dictionary")
            toolMsg.Add "role", "tool"
            toolMsg.Add "tool_call_id", tcId
            toolMsg.Add "content", execResult
            Set newArr(k) = toolMsg
            k = k + 1
        Next
        
        ' Add an explicit instruction so the model writes the files instead of planning further
        ReDim Preserve newArr(k)
        Dim goMsg
        Set goMsg = Server.CreateObject("Scripting.Dictionary")
        goMsg.Add "role", "user"
        goMsg.Add "content", "You now have all the information you need. Write the complete website immediately. Your answer must be exactly one fenced ```html block containing the COMPLETE single-file index.html. No plan, no explanation."
        Set newArr(k) = goMsg
        
        ' Follow-up: no tools this time, force the final answer
        resp = LlmHttpPost(url, apiKey, BuildLlmReqBody(model, ASPPY.JSON.Encode(newArr), False, params))
        If Left(resp, 6) = "ERROR:" Then
            res.Add "type", "error"
            res.Add "error", Mid(resp, 8)
            Set CallLLMSimple = res
            On Error Goto 0
            Exit Function
        End If
        
        Set parsed = ASPPY.JSON.Decode(resp)
        If IsObject(parsed) Then
            If parsed.Exists("choices") Then
                Set msg = parsed("choices")(0)("message")
                content = ""
                If msg.Exists("content") Then
                    If Not IsNull(msg("content")) Then content = msg("content")
                End If
                If msg.Exists("reasoning_content") Then
                    If Not IsNull(msg("reasoning_content")) Then
                        reasoning = reasoning & vbCrLf & msg("reasoning_content")
                        ' Fallback: use reasoning as content when content stays empty
                        If Len(content) = 0 Then content = msg("reasoning_content")
                    End If
                End If
            End If
        End If
    End If
    
    ' Token usage (OpenAI-compatible: usage.prompt_tokens / usage.completion_tokens)
    If IsObject(parsed) Then
        If parsed.Exists("usage") Then
            If IsObject(parsed("usage")) Then
                If parsed("usage").Exists("prompt_tokens") Then
                    If Not IsNull(parsed("usage")("prompt_tokens")) Then res("tokens_in") = CLng(parsed("usage")("prompt_tokens"))
                End If
                If parsed("usage").Exists("completion_tokens") Then
                    If Not IsNull(parsed("usage")("completion_tokens")) Then res("tokens_out") = CLng(parsed("usage")("completion_tokens"))
                End If
            End If
        End If
    End If
    
    ' Truncated by the output token limit? Fail clearly instead of writing half a website.
    Dim simpleFinish
    simpleFinish = ""
    If IsObject(parsed) Then
        If parsed.Exists("choices") Then
            If parsed("choices")(0).Exists("finish_reason") Then
                If Not IsNull(parsed("choices")(0)("finish_reason")) Then simpleFinish = parsed("choices")(0)("finish_reason")
            End If
        End If
    End If
    If simpleFinish = "length" Then
        res.Add "type", "error"
        res.Add "error", "The response was cut off by the token limit (finish_reason=length). Nothing was changed. Increase max_tokens (admin panel, Extra LLM Parameters) and/or the context length in LM Studio, then try again."
        Set CallLLMSimple = res
        On Error Goto 0
        Exit Function
    End If
    
    ' Fallback: reasoning as the answer (Gemma puts output there)
    If Len(content) = 0 And Len(reasoning) > 0 Then
        content = reasoning
    End If
    
    If Len(content) = 0 Then
        res.Add "type", "error"
        res.Add "error", "Empty response from LLM (no content, no reasoning)."
        Set CallLLMSimple = res
        On Error Goto 0
        Exit Function
    End If
    
    res.Add "type", "final"
    res.Add "content", content
    res("reasoning") = reasoning
    Set CallLLMSimple = res
    On Error Goto 0
End Function

' One LLM round. Returns a Dictionary:
'   type      = "final" | "continue" | "error"
'   content   = final answer text            (type=final)
'   reasoning = thinking text from this round (all types)
'   state     = updated messages JSON array   (type=continue)
'   error     = error description             (type=error)
Function CallLLMRound(endpoint, model, apiKey, msgArrJson, enableTools, params)
    On Error Resume Next
    Dim res
    Set res = Server.CreateObject("Scripting.Dictionary")
    res.Add "reasoning", ""
    
    Dim finalEndpoint, url
    finalEndpoint = endpoint
    If Right(finalEndpoint, 1) = "/" Then
        finalEndpoint = Left(finalEndpoint, Len(finalEndpoint) - 1)
    End If
    url = finalEndpoint & "/v1/chat/completions"
    
    ' Build request body (messages is already a JSON array string)
    Dim sb
    sb = "{""model"": """ & model & """, ""stream"": false, ""messages"": " & msgArrJson
    If enableTools Then
        sb = sb & ", ""tools"": " & wikipediaToolsJson()
    End If
    If Len(params) > 0 Then
        sb = sb & ", " & params
        If InStr(1, params, "max_tokens", 1) = 0 Then
            sb = sb & ", ""max_tokens"": -1"
        End If
    Else
        sb = sb & ", ""max_tokens"": -1"
    End If
    sb = sb & "}"
    
    Dim xmlhttp
    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    If Len(apiKey) > 0 Then
        xmlhttp.setRequestHeader "Authorization", "Bearer " & apiKey
    End If
    xmlhttp.setTimeouts 300000, 300000, 300000, 300000
    xmlhttp.Send sb
    
    If Err.Number <> 0 Then
        res.Add "type", "error"
        res.Add "error", "LLM request failed: " & Err.Description & " (URL: " & url & ")"
        Err.Clear
        Set CallLLMRound = res
        On Error Goto 0
        Exit Function
    End If
    If xmlhttp.Status < 200 Or xmlhttp.Status >= 300 Then
        res.Add "type", "error"
        res.Add "error", "HTTP " & xmlhttp.Status & " from LLM: " & Left(xmlhttp.responseText, 300)
        Set CallLLMRound = res
        On Error Goto 0
        Exit Function
    End If
    
    Dim resp, parsed, choice, msg, finishReason
    resp = xmlhttp.responseText
    Set parsed = ASPPY.JSON.Decode(resp)
    If Not IsObject(parsed) Then
        res.Add "type", "error"
        res.Add "error", "Invalid JSON from LLM: " & Left(resp, 300)
        Set CallLLMRound = res
        On Error Goto 0
        Exit Function
    End If
    If Not parsed.Exists("choices") Then
        res.Add "type", "error"
        res.Add "error", "No choices in LLM response: " & Left(resp, 300)
        Set CallLLMRound = res
        On Error Goto 0
        Exit Function
    End If
    
    Set choice = parsed("choices")(0)
    Set msg = choice("message")
    finishReason = ""
    If choice.Exists("finish_reason") Then
        If Not IsNull(choice("finish_reason")) Then finishReason = choice("finish_reason")
    End If
    
    Dim content, reasoning, hasToolCalls
    content = ""
    If msg.Exists("content") Then
        If Not IsNull(msg("content")) Then content = msg("content")
    End If
    reasoning = ""
    If msg.Exists("reasoning_content") Then
        If Not IsNull(msg("reasoning_content")) Then reasoning = msg("reasoning_content")
    End If
    res("reasoning") = reasoning
    
    hasToolCalls = False
    If msg.Exists("tool_calls") Then
        If Not IsNull(msg("tool_calls")) Then hasToolCalls = True
    End If
    
    ' --- Tool calls: execute and append results to conversation state ---
    If hasToolCalls Then
        Dim oldArr, n, i, tcCount, tc
        oldArr = ASPPY.JSON.Decode(msgArrJson)
        n = UBound(oldArr)
        
        tcCount = 0
        For Each tc In msg("tool_calls")
            tcCount = tcCount + 1
        Next
        
        Dim newArr()
        ReDim newArr(n + 1 + tcCount)
        For i = 0 To n
            Set newArr(i) = oldArr(i)
        Next
        
        Dim asstMsg
        Set asstMsg = Server.CreateObject("Scripting.Dictionary")
        asstMsg.Add "role", "assistant"
        asstMsg.Add "content", content
        asstMsg.Add "tool_calls", msg("tool_calls")
        Set newArr(n + 1) = asstMsg
        
        Dim k, tcId, tcName, tcArgs, execResult, toolMsg, toolNotes
        toolNotes = ""
        k = n + 2
        For Each tc In msg("tool_calls")
            tcId = tc("id")
            tcName = tc("function")("name")
            tcArgs = tc("function")("arguments")
            execResult = ExecuteWikipediaTool(tcName, tcArgs)
            toolNotes = toolNotes & "[Tool: " & tcName & " " & Left(tcArgs, 150) & "]" & vbCrLf
            
            Set toolMsg = Server.CreateObject("Scripting.Dictionary")
            toolMsg.Add "role", "tool"
            toolMsg.Add "tool_call_id", tcId
            toolMsg.Add "content", execResult
            Set newArr(k) = toolMsg
            k = k + 1
        Next
        
        res.Add "type", "continue"
        res("reasoning") = reasoning & vbCrLf & toolNotes
        res.Add "state", ASPPY.JSON.Encode(newArr)
        Set CallLLMRound = res
        On Error Goto 0
        Exit Function
    End If
    
    ' --- Final content ---
    If Len(content) > 0 Then
        res.Add "type", "final"
        res.Add "content", content
        Set CallLLMRound = res
        On Error Goto 0
        Exit Function
    End If
    
    ' --- Reasoning only ---
    If Len(reasoning) > 0 Then
        ' If reasoning already contains the single-file block, treat it as final output
        If InStr(1, reasoning, "```html", 1) > 0 Or InStr(1, reasoning, "<file name=", 1) > 0 Then
            res.Add "type", "final"
            res.Add "content", reasoning
            Set CallLLMRound = res
            On Error Goto 0
            Exit Function
        End If
        ' Truncated by token limit? Give a clear error instead of looping
        If finishReason = "length" Then
            res.Add "type", "error"
            res.Add "error", "The model ran out of tokens while thinking (finish_reason=length). Increase max_tokens/context length in LM Studio, or use a model with less verbose reasoning."
            Set CallLLMRound = res
            On Error Goto 0
            Exit Function
        End If
        
        ' Append (trimmed) reasoning + forceful continue instruction
        Dim oldArr2, n2, i2
        oldArr2 = ASPPY.JSON.Decode(msgArrJson)
        n2 = UBound(oldArr2)
        
        Dim newArr2()
        ReDim newArr2(n2 + 2)
        For i2 = 0 To n2
            Set newArr2(i2) = oldArr2(i2)
        Next
        
        Dim trimmedReasoning
        trimmedReasoning = reasoning
        If Len(trimmedReasoning) > 3000 Then
            trimmedReasoning = "..." & Right(trimmedReasoning, 3000)
        End If
        
        Dim thinkMsg
        Set thinkMsg = Server.CreateObject("Scripting.Dictionary")
        thinkMsg.Add "role", "assistant"
        thinkMsg.Add "content", trimmedReasoning
        Set newArr2(n2 + 1) = thinkMsg
        
        Dim contMsg
        Set contMsg = Server.CreateObject("Scripting.Dictionary")
        contMsg.Add "role", "user"
        contMsg.Add "content", "Do NOT think or plan any further. Write the final answer immediately: exactly one fenced ```html block with the COMPLETE single-file index.html."
        Set newArr2(n2 + 2) = contMsg
        
        res.Add "type", "continue"
        res.Add "state", ASPPY.JSON.Encode(newArr2)
        Set CallLLMRound = res
        On Error Goto 0
        Exit Function
    End If
    
    res.Add "type", "error"
    res.Add "error", "Empty response from LLM (no content, no reasoning, no tool calls). finish_reason=" & finishReason
    Set CallLLMRound = res
    On Error Goto 0
End Function

Function buildChatCompletionRequest(model, messagesJson, toolsDef, params)
    Dim sb
    sb = "{""model"": """ & model & """, ""stream"": false, ""messages"": " & messagesJson
    
    If Not IsNull(toolsDef) Then
        sb = sb & ", ""tools"": " & toolsDef
    End If
    
    If Len(params) > 0 Then
        sb = sb & ", " & params
        ' Add unlimited tokens if params do not specify max_tokens
        If InStr(1, params, "max_tokens", 1) = 0 Then
            sb = sb & ", ""max_tokens"": -1"
        End If
    Else
        sb = sb & ", ""max_tokens"": -1"
    End If
    
    sb = sb & "}"
    buildChatCompletionRequest = sb
End Function

' Build a request with the tool result appended - uses proper JSON objects
Function buildToolResultRequest(model, userPrompt, assistantMsg, toolCallId, toolName, toolResult, toolsDef, params)
    On Error Resume Next
    Dim msgs, tcContent, tcObj, funcObj, outObj
    Set outObj = Server.CreateObject("Scripting.Dictionary")
    outObj.Add "model", model
    
    Set msgs = Server.CreateObject("Scripting.Dictionary")
    
    ' Message 0: user prompt
    Dim userMsg
    Set userMsg = Server.CreateObject("Scripting.Dictionary")
    userMsg.Add "role", "user"
    userMsg.Add "content", userPrompt
    msgs.Add 0, userMsg
    
    ' Message 1: assistant with tool call
    tcContent = assistantMsg("content")
    If IsNull(tcContent) Or IsEmpty(tcContent) Then tcContent = "" End If
    Dim asstMsg, innerTc, innerFunc
    Set asstMsg = Server.CreateObject("Scripting.Dictionary")
    asstMsg.Add "role", "assistant"
    asstMsg.Add "content", tcContent
    
    Dim tcArr, tcItem
    Set tcArr = Server.CreateObject("Scripting.Dictionary")
    Set tcItem = Server.CreateObject("Scripting.Dictionary")
    tcItem.Add "id", toolCallId
    tcItem.Add "type", "function"
    Set innerFunc = Server.CreateObject("Scripting.Dictionary")
    innerFunc.Add "name", toolName
    innerFunc.Add "arguments", "{}"
    tcItem.Add "function", innerFunc
    tcArr.Add 0, tcItem
    asstMsg.Add "tool_calls", tcArr
    msgs.Add 1, asstMsg
    
    ' Message 2: tool result
    Dim toolMsg
    Set toolMsg = Server.CreateObject("Scripting.Dictionary")
    toolMsg.Add "role", "tool"
    toolMsg.Add "tool_call_id", toolCallId
    toolMsg.Add "content", toolResult
    msgs.Add 2, toolMsg
    
    outObj.Add "messages", msgs
    outObj.Add "stream", False
    
    If Not IsNull(toolsDef) Then
        outObj.Add "tools", ASPPY.JSON.Decode(toolsDef)
    End If
    If Len(params) > 0 Then
        Dim extraParams
        Set extraParams = ASPPY.JSON.Decode("{" & params & "}")
        If IsObject(extraParams) Then
            Dim pKey
            For Each pKey In extraParams
                outObj.Add pKey, extraParams(pKey)
            Next
        End If
    End If
    
    buildToolResultRequest = ASPPY.JSON.Encode(outObj)
    On Error Goto 0
End Function

Function wikipediaToolsJson()
    Dim sb
    sb = "["
    sb = sb & "{""type"": ""function"", ""function"": {"
    sb = sb & """name"": ""search_wikipedia"", "
    sb = sb & """description"": ""Search Wikipedia for articles matching a query. Returns titles and snippets."", "
    sb = sb & """parameters"": {""type"": ""object"", ""properties"": {"
    sb = sb & """query"": {""type"": ""string"", ""description"": ""The search query""}"
    sb = sb & "}, ""required"": [""query""]}"
    sb = sb & "}},"
    sb = sb & "{""type"": ""function"", ""function"": {"
    sb = sb & """name"": ""get_wikipedia_page"", "
    sb = sb & """description"": ""Get the full content of a Wikipedia page by its exact title."", "
    sb = sb & """parameters"": {""type"": ""object"", ""properties"": {"
    sb = sb & """title"": {""type"": ""string"", ""description"": ""The exact Wikipedia page title""}"
    sb = sb & "}, ""required"": [""title""]}"
    sb = sb & "}}"
    sb = sb & "]"
    wikipediaToolsJson = sb
End Function

Function ExecuteWikipediaTool(name, argsJson)
    On Error Resume Next
    Dim parsed, query, title, escaped
    Set parsed = ASPPY.JSON.Decode(argsJson)
    If Not IsObject(parsed) Then
        ExecuteWikipediaTool = """Error: invalid arguments"""
        On Error Goto 0
        Exit Function
    End If
    
    If name = "search_wikipedia" Then
        If Not parsed.Exists("query") Then
            ExecuteWikipediaTool = """Error: missing query parameter"""
            On Error Goto 0
            Exit Function
        End If
        query = parsed("query")
        ExecuteWikipediaTool = WikipediaSearch(query)
    ElseIf name = "get_wikipedia_page" Then
        If Not parsed.Exists("title") Then
            ExecuteWikipediaTool = """Error: missing title parameter"""
            On Error Goto 0
            Exit Function
        End If
        title = parsed("title")
        ExecuteWikipediaTool = WikipediaGetPage(title)
    Else
        ExecuteWikipediaTool = """Error: unknown tool: " & EscapeJsonString(name) & """"
    End If
    
    On Error Goto 0
End Function

Function WikipediaSearch(query)
    On Error Resume Next
    Dim xmlhttp, url, resp, parsed, pages, idx, result, i, key
    Dim urlEnc
    urlEnc = Server.URLEncode(query)
    url = "https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch=" & urlEnc & "&format=json&srlimit=5"
    
    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "User-Agent", "ASPPY_AI_Builder/1.0"
    xmlhttp.Send
    
    If Err.Number <> 0 Or xmlhttp.Status <> 200 Then
        WikipediaSearch = """Error: Wikipedia API unavailable"""
        On Error Goto 0
        Exit Function
    End If
    
    resp = xmlhttp.responseText
    Set parsed = ASPPY.JSON.Decode(resp)
    If IsObject(parsed) Then
        If parsed.Exists("query") Then
            If parsed("query").Exists("search") Then
                Set pages = parsed("query")("search")
                result = "["
                For i = 0 To pages.Count - 1
                    If i > 0 Then result = result & ","
                    result = result & "{""title"":""" & EscapeJsonString(CStr(pages(i)("title"))) & """,""snippet"":""" & EscapeJsonString(StripHtml(CStr(pages(i)("snippet")))) & """}"
                Next
                result = result & "]"
                WikipediaSearch = result
                On Error Goto 0
                Exit Function
            End If
        End If
    End If
    
    WikipediaSearch = """No results found"""
    On Error Goto 0
End Function

Function WikipediaGetPage(title)
    On Error Resume Next
    Dim xmlhttp, url, resp, parsed, pages, key, extract
    Dim urlEnc
    urlEnc = Server.URLEncode(title)
    url = "https://en.wikipedia.org/w/api.php?action=query&prop=extracts&exintro=1&explaintext=1&titles=" & urlEnc & "&format=json"
    
    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "User-Agent", "ASPPY_AI_Builder/1.0"
    xmlhttp.Send
    
    If Err.Number <> 0 Or xmlhttp.Status <> 200 Then
        WikipediaGetPage = """Error: Wikipedia API unavailable"""
        On Error Goto 0
        Exit Function
    End If
    
    resp = xmlhttp.responseText
    Set parsed = ASPPY.JSON.Decode(resp)
    If IsObject(parsed) Then
        If parsed.Exists("query") Then
            If parsed("query").Exists("pages") Then
                Set pages = parsed("query")("pages")
                ' Dictionary keys are page IDs (numbers as strings)
                Dim keysColl
                For Each key In pages
                    Set pageDict = pages(key)
                    If pageDict.Exists("extract") Then
                        extract = pageDict("extract")
                        ' Truncate extract to ~8000 chars to keep the response manageable and valid JSON
                        If Len(extract) > 8000 Then
                            extract = Left(extract, 8000) & "..."
                        End If
                        WikipediaGetPage = "{""title"":""" & EscapeJsonString(title) & """,""content"":""" & EscapeJsonString(extract) & """}"
                        On Error Goto 0
                        Exit Function
                    End If
                Next
            End If
        End If
    End If
    
    WikipediaGetPage = """Page not found"""
    On Error Goto 0
End Function

Function StripHtml(s)
    Dim result, i, inTag
    result = ""
    inTag = False
    For i = 1 To Len(s)
        Dim ch
        ch = Mid(s, i, 1)
        If ch = "<" Then
            inTag = True
        ElseIf ch = ">" Then
            inTag = False
        ElseIf Not inTag Then
            result = result & ch
        End If
    Next
    StripHtml = result
End Function

Function escapeAppendToolResult(messagesJson, assistantMsg, toolCallId, toolName, toolResult)
    On Error Resume Next
    
    ' Manually build the extended messages JSON array
    ' This avoids complex JSON manipulation in VBScript
    Dim sb, tcContent
    tcContent = assistantMsg("content")
    If IsNull(tcContent) Or IsEmpty(tcContent) Then
        tcContent = ""
    End If
    
    ' Remove the closing "]" from messagesJson, add assistant message, tool result, and close
    sb = messagesJson
    
    ' Append the assistant message with tool_calls
    sb = sb & ",{""role"": ""assistant"", ""content"": """ & EscapeJsonString(tcContent) & """, ""tool_calls"": ["
    sb = sb & "{""id"": """ & EscapeJsonString(toolCallId) & """, ""type"": ""function"", ""function"": {""name"": """ & EscapeJsonString(toolName) & """, ""arguments"": ""{}""}}"
    sb = sb & "]}]"
    
    ' Now append the tool result message
    ' We need to re-wrap: messagesJson had the tool message, now add the tool response
    ' Rebuild properly:
    Dim beforeClose
    ' Find the last '}]' and replace it
    Dim pos2
    pos2 = InStrRev(sb, "}]")
    If pos2 > 0 Then
        sb = Left(sb, pos2 - 1) & ",{""role"": ""tool"", ""tool_call_id"": """ & EscapeJsonString(toolCallId) & """, ""content"": " & toolResult & "}]}]"
    End If
    
    escapeAppendToolResult = sb
    On Error Goto 0
End Function

Function extractOpenAIResponse(jsonStr)
    On Error Resume Next
    Dim parsed, choices, message, content
    Set parsed = ASPPY.JSON.Decode(jsonStr)
    If IsObject(parsed) Then
        If parsed.Exists("choices") Then
            Set choices = parsed("choices")
            If choices.Count > 0 Then
                Set message = choices(0)("message")
                If message.Exists("content") Then
                    content = message("content")
                    If Not IsNull(content) And Len(content) > 0 Then
                        extractOpenAIResponse = content
                        Exit Function
                    End If
                End If
                If message.Exists("reasoning_content") Then
                    content = message("reasoning_content")
                    If Not IsNull(content) And Len(content) > 0 Then
                        extractOpenAIResponse = content
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    extractOpenAIResponse = Null
    On Error Goto 0
End Function

Function extractOllamaResponse(jsonStr)
    On Error Resume Next
    Dim parsed, content
    Set parsed = ASPPY.JSON.Decode(jsonStr)
    If IsObject(parsed) Then
        If parsed.Exists("response") Then
            content = parsed("response")
            extractOllamaResponse = content
            Exit Function
        End If
        If parsed.Exists("message") Then
            If IsObject(parsed("message")) Then
                If parsed("message").Exists("content") Then
                    extractOllamaResponse = parsed("message")("content")
                    Exit Function
                End If
            End If
        End If
    End If
    extractOllamaResponse = Null
    On Error Goto 0
End Function

' Extract the single-file website from a fenced ```html ... ``` block.
' Returns the trimmed HTML string, or Null when no usable block is found.
' Falls back to raw HTML (response starting with <!DOCTYPE or <html>).
Function ParseSingleFileHtml(responseText)
    On Error Resume Next
    Dim p1, p2, fence, html, lfPos, trimmed
    ParseSingleFileHtml = Null
    If IsNull(responseText) Or Len(responseText) = 0 Then Exit Function
    ' Find first ```html (case-insensitive) or plain ``` fence
    p1 = InStr(1, responseText, "```html", 1)
    If p1 > 0 Then
        p1 = p1 + Len("```html")
    Else
        p1 = InStr(1, responseText, "```", 1)
        If p1 > 0 Then p1 = p1 + 3
    End If
    If p1 > 0 Then
        p2 = InStr(p1, responseText, "```", 0)
        If p2 > p1 Then
            html = Trim(Mid(responseText, p1, p2 - p1))
            If Len(html) > 0 Then
                ParseSingleFileHtml = html
                On Error Goto 0
                Exit Function
            End If
        End If
    End If
    ' Fallback: raw HTML without fences
    trimmed = Trim(responseText)
    If InStr(1, trimmed, "<!DOCTYPE", 1) = 1 Or InStr(1, trimmed, "<html", 1) = 1 Then
        ParseSingleFileHtml = trimmed
    End If
    On Error Goto 0
End Function

Function ParseAIResponse(responseText)
    On Error Resume Next
    
    Dim files, idx, i, tagStart, tagEnd, tagContent, nameStart, nameEnd, fileName, fileContent, fileEntry
    Set files = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    
    ' Find all <file name="...">...</file> blocks
    i = 1
    Do While i <= Len(responseText)
        tagStart = InStr(i, responseText, "<file name=", 1)
        If tagStart = 0 Then Exit Do
        
        ' Find end of opening tag
        tagEnd = InStr(tagStart, responseText, ">", 1)
        If tagEnd = 0 Then Exit Do
        
        ' Extract name attribute
        Dim tagStr
        tagStr = Mid(responseText, tagStart, tagEnd - tagStart + 1)
        nameStart = InStr(1, tagStr, """", 1)
        If nameStart = 0 Then
            i = tagEnd + 1
            ' Continue loop
        Else
            nameEnd = InStr(nameStart + 1, tagStr, """", 1)
            If nameEnd = 0 Then
                i = tagEnd + 1
                ' Continue loop
            Else
                fileName = Mid(tagStr, nameStart + 1, nameEnd - nameStart - 1)
                
                ' Find closing tag
                Dim closeTag
                closeTag = "</file>"
                Dim closeStart
                closeStart = InStr(tagEnd + 1, responseText, closeTag, 1)
                If closeStart = 0 Then Exit Do
                
                fileContent = Mid(responseText, tagEnd + 1, closeStart - tagEnd - 1)
                
                ' Trim whitespace
                fileContent = Trim(fileContent)
                
                Set fileEntry = Server.CreateObject("Scripting.Dictionary")
                fileEntry.Add "name", fileName
                fileEntry.Add "content", fileContent
                
                files.Add idx, fileEntry
                idx = idx + 1
                
                i = closeStart + Len(closeTag)
            End If
        End If
    Loop
    
    If idx = 0 Then
        Set ParseAIResponse = Null
    Else
        Set ParseAIResponse = files
    End If
    
    On Error Goto 0
End Function

' Extract a double-quoted attribute value from an opening tag string
Function ExtractTagAttr(tagStr, attrName)
    Dim p, q1, q2
    ExtractTagAttr = ""
    p = InStr(1, tagStr, attrName & "=""", 1)
    If p = 0 Then Exit Function
    q1 = p + Len(attrName) + 2
    q2 = InStr(q1, tagStr, """", 1)
    If q2 = 0 Then Exit Function
    ExtractTagAttr = Mid(tagStr, q1, q2 - q1)
End Function

' Find all <patch file="..." section="...">...</patch> blocks in the AI output.
' Returns Null when none found, otherwise a Dictionary of {file, section, content}.
Function ParseAIPatches(responseText)
    On Error Resume Next
    
    Dim patches, idx, i, tagStart, tagEnd, tagStr, closeStart, entry, pFile, pSection
    Set patches = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    
    i = 1
    Do While i <= Len(responseText)
        tagStart = InStr(i, responseText, "<patch ", 1)
        If tagStart = 0 Then Exit Do
        tagEnd = InStr(tagStart, responseText, ">", 1)
        If tagEnd = 0 Then Exit Do
        
        tagStr = Mid(responseText, tagStart, tagEnd - tagStart + 1)
        pFile = ExtractTagAttr(tagStr, "file")
        pSection = ExtractTagAttr(tagStr, "section")
        
        closeStart = InStr(tagEnd + 1, responseText, "</patch>", 1)
        If closeStart = 0 Then Exit Do
        
        If Len(pFile) > 0 And Len(pSection) > 0 Then
            Set entry = Server.CreateObject("Scripting.Dictionary")
            entry.Add "file", pFile
            entry.Add "section", pSection
            entry.Add "before", ExtractTagAttr(tagStr, "before")
            entry.Add "after", ExtractTagAttr(tagStr, "after")
            entry.Add "content", Trim(Mid(responseText, tagEnd + 1, closeStart - tagEnd - 1))
            patches.Add idx, entry
            idx = idx + 1
        End If
        
        i = closeStart + Len("</patch>")
    Loop
    
    If idx = 0 Then
        Set ParseAIPatches = Null
    Else
        Set ParseAIPatches = patches
    End If
    
    On Error Goto 0
End Function

' Find a section marker token (e.g. "SECTION:hero" or "/SECTION:hero") on a
' word boundary. Start markers must not actually be end markers ("/SECTION:").
' Returns 0 when not found.
Function FindMarkerPos(content, token)
    Dim p, nextCh, tokenLen, okBoundary
    tokenLen = Len(token)
    FindMarkerPos = 0
    p = InStr(1, content, token, 1)
    Do While p > 0
        okBoundary = True
        ' The character after the token may not extend the section name
        If p + tokenLen <= Len(content) Then
            nextCh = LCase(Mid(content, p + tokenLen, 1))
            If (nextCh >= "a" And nextCh <= "z") Or (nextCh >= "0" And nextCh <= "9") Or nextCh = "-" Or nextCh = "_" Then okBoundary = False
        End If
        ' A start token must not be the tail of an end token
        If Left(token, 1) <> "/" And p > 1 Then
            If Mid(content, p - 1, 1) = "/" Then okBoundary = False
        End If
        If okBoundary Then
            FindMarkerPos = p
            Exit Function
        End If
        p = InStr(p + 1, content, token, 1)
    Loop
End Function

' Make sure patch content is wrapped in the correct start/end markers for the file type
Function EnsureSectionMarkers(fileName, sectionName, content)
    Dim t, startMk, endMk, lname
    t = Trim(content)
    If FindMarkerPos(t, "SECTION:" & sectionName) > 0 And FindMarkerPos(t, "/SECTION:" & sectionName) > 0 Then
        EnsureSectionMarkers = t & vbCrLf
        Exit Function
    End If
    lname = LCase(fileName)
    If Right(lname, 4) = ".css" Then
        startMk = "/* SECTION:" & sectionName & " */"
        endMk = "/* /SECTION:" & sectionName & " */"
    ElseIf Right(lname, 3) = ".js" Then
        startMk = "// SECTION:" & sectionName
        endMk = "// /SECTION:" & sectionName
    Else
        startMk = "<!-- SECTION:" & sectionName & " -->"
        endMk = "<!-- /SECTION:" & sectionName & " -->"
    End If
    EnsureSectionMarkers = startMk & vbCrLf & t & vbCrLf & endMk & vbCrLf
End Function

' Position where the comment containing the marker starts. pos points at the
' marker token; walk back over whitespace and include "<!--", "/*" or "//".
Function MarkerCommentStart(content, pos)
    Dim k, ch
    k = pos - 1
    Do While k >= 1
        ch = Mid(content, k, 1)
        If ch = " " Or ch = vbTab Then
            k = k - 1
        Else
            Exit Do
        End If
    Loop
    If k >= 4 Then
        If Mid(content, k - 3, 4) = "<!--" Then
            MarkerCommentStart = k - 3
            Exit Function
        End If
    End If
    If k >= 2 Then
        If Mid(content, k - 1, 2) = "/*" Or Mid(content, k - 1, 2) = "//" Then
            MarkerCommentStart = k - 1
            Exit Function
        End If
    End If
    MarkerCommentStart = pos
End Function

' Position where the comment containing the marker ends. tokenStart points at
' the marker token of length tokenLen; include the trailing "-->" or "*/",
' or (for // comments) run to the end of the line.
Function MarkerCommentEnd(content, tokenStart, tokenLen)
    Dim k, lineEnd
    k = tokenStart + tokenLen
    Do While k <= Len(content)
        If Mid(content, k, 1) = " " Or Mid(content, k, 1) = vbTab Then
            k = k + 1
        Else
            Exit Do
        End If
    Loop
    If k + 2 <= Len(content) Then
        If Mid(content, k, 3) = "-->" Then
            MarkerCommentEnd = k + 2
            Exit Function
        End If
    End If
    If k + 1 <= Len(content) Then
        If Mid(content, k, 2) = "*/" Then
            MarkerCommentEnd = k + 1
            Exit Function
        End If
    End If
    ' Line comment (//) or malformed: the marker runs to the end of its line
    lineEnd = InStr(tokenStart, content, vbLf, 0)
    If lineEnd = 0 Then
        MarkerCommentEnd = Len(content)
    Else
        MarkerCommentEnd = lineEnd - 1
        If MarkerCommentEnd >= 1 Then
            If Mid(content, MarkerCommentEnd, 1) = vbCr Then MarkerCommentEnd = MarkerCommentEnd - 1
        End If
        If MarkerCommentEnd < tokenStart + tokenLen - 1 Then MarkerCommentEnd = tokenStart + tokenLen - 1
    End If
End Function

' Replace one named section in a file's content with new content.
' Replacement is surgical: from the start of the start-marker comment through
' the end of the end-marker comment. Content sharing a line with a marker is
' preserved (small models often put markers and code on one line).
' Unknown sections are INSERTED. Placement, in order of preference:
'   1. after the section named in afterName (when given and found)
'   2. before the section named in beforeName (when given and found)
'   3. HTML: just above the footer (SECTION:footer marker or <footer> tag)
'   4. HTML: before </body>; other files: appended at the end
Function ApplySectionPatch(content, fileName, sectionName, patchContent, beforeName, afterName)
    Dim sPos, ePos, repStart, repEnd, newBlock, lname, insPos, refPos
    newBlock = EnsureSectionMarkers(fileName, sectionName, patchContent)
    
    sPos = FindMarkerPos(content, "SECTION:" & sectionName)
    ePos = FindMarkerPos(content, "/SECTION:" & sectionName)
    
    If sPos > 0 And ePos > sPos Then
        repStart = MarkerCommentStart(content, sPos)
        repEnd = MarkerCommentEnd(content, ePos, Len("/SECTION:" & sectionName))
        ApplySectionPatch = Left(content, repStart - 1) & Trim(newBlock) & Mid(content, repEnd + 1)
        Exit Function
    End If
    
    ' Section not found: insert as a new section
    lname = LCase(fileName)
    insPos = 0
    
    ' 1. Explicit anchor: after an existing section
    If Len(afterName) > 0 Then
        refPos = FindMarkerPos(content, "/SECTION:" & afterName)
        If refPos > 0 Then
            insPos = MarkerCommentEnd(content, refPos, Len("/SECTION:" & afterName)) + 1
            ApplySectionPatch = Left(content, insPos - 1) & vbCrLf & newBlock & Mid(content, insPos)
            Exit Function
        End If
    End If
    
    ' 2. Explicit anchor: before an existing section
    If Len(beforeName) > 0 Then
        refPos = FindMarkerPos(content, "SECTION:" & beforeName)
        If refPos > 0 Then insPos = MarkerCommentStart(content, refPos)
    End If
    
    ' 3. Default for HTML: just above the footer
    If insPos = 0 And (Right(lname, 5) = ".html" Or Right(lname, 4) = ".htm") Then
        refPos = FindMarkerPos(content, "SECTION:footer")
        If refPos > 0 Then
            insPos = MarkerCommentStart(content, refPos)
        Else
            refPos = InStr(1, content, "<footer", 1)
            If refPos > 0 Then
                insPos = refPos
                ' When the footer tag is wrapped in a section with another
                ' name, insert before that section's START marker so the new
                ' markers do not end up nested inside it.
                Dim wrapPos, wnm, wk, wch, wEndPos
                wrapPos = InStrRev(content, "SECTION:", refPos, 1)
                If wrapPos > 1 Then
                    If Mid(content, wrapPos - 1, 1) <> "/" Then
                        wnm = ""
                        wk = wrapPos + Len("SECTION:")
                        Do While wk <= Len(content)
                            wch = LCase(Mid(content, wk, 1))
                            If (wch >= "a" And wch <= "z") Or (wch >= "0" And wch <= "9") Or wch = "-" Or wch = "_" Then
                                wnm = wnm & wch
                                wk = wk + 1
                            Else
                                Exit Do
                            End If
                        Loop
                        If Len(wnm) > 0 Then
                            wEndPos = FindMarkerPos(content, "/SECTION:" & wnm)
                            If wEndPos > refPos Then insPos = MarkerCommentStart(content, wrapPos)
                        End If
                    End If
                End If
            End If
        End If
        ' 4. No footer: before </body>
        If insPos = 0 Then
            refPos = InStr(1, content, "</body>", 1)
            If refPos > 0 Then insPos = refPos
        End If
    End If
    
    If insPos > 0 Then
        ApplySectionPatch = Left(content, insPos - 1) & newBlock & Mid(content, insPos)
        Exit Function
    End If
    ApplySectionPatch = content & vbCrLf & newBlock
End Function

' Sanity check on SECTION markers: properly paired in order, never nested,
' and never after </html>. Small models get this wrong; broken markers make
' patching dangerous, so callers fall back to full-file mode when this fails.
Function SectionMarkersHealthy(content)
    Dim p, isEnd, nm, k, ch, openName, htmlClose
    SectionMarkersHealthy = False
    openName = ""
    htmlClose = InStr(1, content, "</html>", 1)
    p = InStr(1, content, "SECTION:", 1)
    Do While p > 0
        If htmlClose > 0 And p > htmlClose Then Exit Function ' marker after </html>
        isEnd = False
        If p > 1 Then
            If Mid(content, p - 1, 1) = "/" Then isEnd = True
        End If
        ' Extract the section name
        nm = ""
        k = p + Len("SECTION:")
        Do While k <= Len(content)
            ch = LCase(Mid(content, k, 1))
            If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "_" Then
                nm = nm & ch
                k = k + 1
            Else
                Exit Do
            End If
        Loop
        If Len(nm) = 0 Then Exit Function
        If isEnd Then
            If openName <> nm Then Exit Function ' end without matching open
            openName = ""
        Else
            If openName <> "" Then Exit Function ' nested section
            openName = nm
        End If
        p = InStr(k, content, "SECTION:", 1)
    Loop
    If openName <> "" Then Exit Function ' unclosed section
    SectionMarkersHealthy = True
End Function

' Remove all SECTION marker comments from content (the content itself stays).
' Used to self-heal files whose markers are broken.
Function StripSectionMarkers(content)
    Dim s, p, tokStart, k, ch, cStart, cEnd
    s = content
    p = InStr(1, s, "SECTION:", 1)
    Do While p > 0
        tokStart = p
        If p > 1 Then
            If Mid(s, p - 1, 1) = "/" Then tokStart = p - 1
        End If
        ' Token runs through the section name
        k = p + Len("SECTION:")
        Do While k <= Len(s)
            ch = LCase(Mid(s, k, 1))
            If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "_" Then
                k = k + 1
            Else
                Exit Do
            End If
        Loop
        cStart = MarkerCommentStart(s, tokStart)
        cEnd = MarkerCommentEnd(s, tokStart, k - tokStart)
        s = Left(s, cStart - 1) & Mid(s, cEnd + 1)
        p = InStr(1, s, "SECTION:", 1)
    Loop
    StripSectionMarkers = s
End Function

' Find "<tag" at a word boundary (so "nav" never matches "<navbar").
' Returns 0 when not found.
Function FindOpeningTag(s, tag, startAt)
    Dim p, nx
    FindOpeningTag = 0
    p = InStr(startAt, s, "<" & tag, 1)
    Do While p > 0
        nx = ""
        If p + Len(tag) + 1 <= Len(s) Then nx = Mid(s, p + Len(tag) + 1, 1)
        If nx = " " Or nx = ">" Or nx = vbTab Or nx = vbCr Or nx = vbLf Then
            FindOpeningTag = p
            Exit Function
        End If
        p = InStr(p + 1, s, "<" & tag, 1)
    Loop
End Function

' Find the "</tag>" matching the opening tag at openPos, accounting for
' nested tags of the same name. Returns 0 when unbalanced.
Function FindMatchingClose(s, tag, openPos)
    Dim depth, cur, pO, pC
    depth = 1
    cur = openPos + 1
    FindMatchingClose = 0
    Do While depth > 0
        pO = FindOpeningTag(s, tag, cur)
        pC = InStr(cur, s, "</" & tag & ">", 1)
        If pC = 0 Then Exit Function
        If pO > 0 And pO < pC Then
            depth = depth + 1
            cur = pO + 1
        Else
            depth = depth - 1
            cur = pC + 1
            If depth = 0 Then FindMatchingClose = pC
        End If
    Loop
End Function

' Section name for an auto-wrapped tag: its sanitized id attribute,
' falling back to "tag-N".
Function SectionNameForTag(s, tag, openPos, counter)
    Dim tagEnd, tagStr, idVal, nm, i, ch
    nm = ""
    tagEnd = InStr(openPos, s, ">", 0)
    If tagEnd > 0 Then
        tagStr = Mid(s, openPos, tagEnd - openPos + 1)
        idVal = ExtractTagAttr(tagStr, "id")
        For i = 1 To Len(idVal)
            ch = LCase(Mid(idVal, i, 1))
            If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "_" Then nm = nm & ch
        Next
    End If
    If Len(nm) = 0 Then nm = tag & "-" & (counter + 1)
    SectionNameForTag = nm
End Function

' Deterministic fallback when the AI ignores the marker instructions (common
' with small models): wrap the <head> contents and every top-level body block
' (<nav>, <header>, <section>, <article>, <aside>, <footer>) in SECTION
' markers so partial updates keep working. Markers produced here are always
' balanced and never nested.
Function AutoInjectSectionMarkers(html)
    Dim s, pHeadOpen, pHeadClose
    s = html
    
    ' Wrap the <head> contents
    pHeadOpen = FindOpeningTag(s, "head", 1)
    If pHeadOpen > 0 Then
        pHeadOpen = InStr(pHeadOpen, s, ">", 0)
        pHeadClose = InStr(1, s, "</head>", 1)
        If pHeadOpen > 0 And pHeadClose > pHeadOpen Then
            s = Left(s, pHeadOpen) & vbCrLf & "<!-- SECTION:head -->" & Mid(s, pHeadOpen + 1, pHeadClose - pHeadOpen - 1) & "<!-- /SECTION:head -->" & vbCrLf & Mid(s, pHeadClose)
        End If
    End If
    
    ' Wrap sequential blocks inside <body>
    Dim tags, cursor, bodyOpen, bodyClose, bestPos, bestTag, ti, p, secCount
    tags = Array("nav", "header", "section", "article", "aside", "footer")
    bodyOpen = FindOpeningTag(s, "body", 1)
    If bodyOpen = 0 Then
        AutoInjectSectionMarkers = s
        Exit Function
    End If
    cursor = InStr(bodyOpen, s, ">", 0) + 1
    secCount = 0
    
    Do While True
        bodyClose = InStr(1, s, "</body>", 1)
        If bodyClose = 0 Then bodyClose = Len(s) + 1
        
        ' The earliest next block tag from the cursor
        bestPos = 0
        bestTag = ""
        For ti = 0 To UBound(tags)
            p = FindOpeningTag(s, tags(ti), cursor)
            If p > 0 And p < bodyClose Then
                If bestPos = 0 Or p < bestPos Then
                    bestPos = p
                    bestTag = tags(ti)
                End If
            End If
        Next
        If bestPos = 0 Then Exit Do
        
        Dim closePos, nm, blockEnd, startMk, endMk
        closePos = FindMatchingClose(s, bestTag, bestPos)
        If closePos = 0 Then
            cursor = bestPos + 1 ' malformed block: skip it
        Else
            nm = SectionNameForTag(s, bestTag, bestPos, secCount)
            If FindMarkerPos(s, "SECTION:" & nm) > 0 Then nm = nm & "-" & (secCount + 1)
            secCount = secCount + 1
            blockEnd = closePos + Len("</" & bestTag & ">")
            startMk = "<!-- SECTION:" & nm & " -->" & vbCrLf
            endMk = vbCrLf & "<!-- /SECTION:" & nm & " -->"
            s = Left(s, bestPos - 1) & startMk & Mid(s, bestPos, blockEnd - bestPos) & endMk & Mid(s, blockEnd)
            cursor = blockEnd + Len(startMk) + Len(endMk)
        End If
    Loop
    
    AutoInjectSectionMarkers = s
End Function

' First image of a project: the first "img/..." referenced in the HTML, or
' (when none is referenced) the first image file in the img folder.
' Returns a relative path like "img/photo.jpg", or "" when there is none.
Function FirstProjectImage(html, fso, projectDir)
    Dim p, q, imgDir, f, ext
    FirstProjectImage = ""
    
    p = InStr(1, html, "src=""img/", 1)
    If p > 0 Then
        q = InStr(p + 5, html, """", 0)
        If q > p + 5 Then
            FirstProjectImage = Mid(html, p + 5, q - p - 5)
            Exit Function
        End If
    End If
    p = InStr(1, html, "src='img/", 1)
    If p > 0 Then
        q = InStr(p + 5, html, "'", 0)
        If q > p + 5 Then
            FirstProjectImage = Mid(html, p + 5, q - p - 5)
            Exit Function
        End If
    End If
    
    imgDir = projectDir & "img\"
    If fso.FolderExists(imgDir) Then
        For Each f In fso.GetFolder(imgDir).Files
            ext = LCase(fso.GetExtensionName(f.Name))
            If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "gif" Or ext = "webp" Or ext = "avif" Then
                FirstProjectImage = "img/" & f.Name
                Exit Function
            End If
        Next
    End If
End Function

' Ensure Open Graph / Twitter meta tags in <head> so shared links (Facebook,
' WhatsApp, LinkedIn, ...) show a proper title, description and photo.
' og:image requires an ABSOLUTE url, which the AI cannot know - so the server
' injects it. Existing tags are kept; a relative og:image is made absolute.
Function InjectOpenGraphTags(html, projectName, publicUrl, siteUrl, imgRel)
    Dim s, headClose, block, title, desc, tPos, tEnd, tClose, tagStart, tagEnd, tagStr, imgUrl
    s = html
    headClose = InStr(1, s, "</head>", 1)
    If headClose = 0 Then
        InjectOpenGraphTags = s
        Exit Function
    End If
    
    ' Title and description come from the document itself
    title = projectName
    tPos = InStr(1, s, "<title", 1)
    If tPos > 0 Then
        tEnd = InStr(tPos, s, ">", 0)
        tClose = InStr(tPos, s, "</title>", 1)
        If tEnd > 0 And tClose > tEnd Then title = Trim(Mid(s, tEnd + 1, tClose - tEnd - 1))
    End If
    
    desc = ""
    tPos = InStr(1, s, "name=""description""", 1)
    If tPos > 0 Then
        tagStart = InStrRev(s, "<", tPos)
        tagEnd = InStr(tPos, s, ">", 0)
        If tagStart > 0 And tagEnd > tagStart Then
            desc = ExtractTagAttr(Mid(s, tagStart, tagEnd - tagStart + 1), "content")
        End If
    End If
    
    imgUrl = ""
    If Len(imgRel) > 0 Then imgUrl = siteUrl & imgRel
    
    ' Fix an existing og:image that uses a relative URL (previews break on those)
    tPos = InStr(1, s, "property=""og:image""", 1)
    If tPos > 0 Then
        tagStart = InStrRev(s, "<", tPos)
        tagEnd = InStr(tPos, s, ">", 0)
        If tagStart > 0 And tagEnd > tagStart Then
            tagStr = Mid(s, tagStart, tagEnd - tagStart + 1)
            Dim curImg, newImg, newTagStr
            curImg = ExtractTagAttr(tagStr, "content")
            If Len(curImg) > 0 And LCase(Left(curImg, 4)) <> "http" Then
                If Left(curImg, 1) = "/" Then
                    newImg = publicUrl & curImg
                Else
                    newImg = siteUrl & curImg
                End If
                newTagStr = Replace(tagStr, "content=""" & curImg & """", "content=""" & newImg & """")
                s = Replace(s, tagStr, newTagStr, 1, 1)
            End If
        End If
    End If
    
    ' Add whatever is still missing
    block = ""
    If InStr(1, s, "property=""og:type""", 1) = 0 Then
        block = block & "    <meta property=""og:type"" content=""website"">" & vbCrLf
    End If
    If InStr(1, s, "property=""og:title""", 1) = 0 Then
        block = block & "    <meta property=""og:title"" content=""" & Replace(title, """", "&quot;") & """>" & vbCrLf
    End If
    If Len(desc) > 0 And InStr(1, s, "property=""og:description""", 1) = 0 Then
        block = block & "    <meta property=""og:description"" content=""" & Replace(desc, """", "&quot;") & """>" & vbCrLf
    End If
    If InStr(1, s, "property=""og:url""", 1) = 0 Then
        block = block & "    <meta property=""og:url"" content=""" & siteUrl & """>" & vbCrLf
    End If
    If Len(imgUrl) > 0 Then
        If InStr(1, s, "property=""og:image""", 1) = 0 Then
            block = block & "    <meta property=""og:image"" content=""" & imgUrl & """>" & vbCrLf
        End If
        If InStr(1, s, "name=""twitter:card""", 1) = 0 Then
            block = block & "    <meta name=""twitter:card"" content=""summary_large_image"">" & vbCrLf
        End If
        If InStr(1, s, "name=""twitter:image""", 1) = 0 Then
            block = block & "    <meta name=""twitter:image"" content=""" & imgUrl & """>" & vbCrLf
        End If
    End If
    
    If Len(block) > 0 Then
        headClose = InStr(1, s, "</head>", 1)
        s = Left(s, headClose - 1) & block & Mid(s, headClose)
    End If
    
    InjectOpenGraphTags = s
End Function

' Responsive safety net: viewport meta + overflow guards against mobile
' horizontal overflow (bare .row, 100vw, fixed widths, long words).
' Idempotent via the /*autonome-guard*/ marker (not an ID marker, so it
' survives StripIdMarkers/retag cycles). Runs on every generation, right
' before the final re-tag, so AI rewrites can never drop it.
Function EnsureResponsiveGuard(html)
    Dim s, headClose, styleOpen, styleEnd, guardCss, insPos, cPos
    EnsureResponsiveGuard = html
    If IsNull(html) Then Exit Function
    If Len(html) = 0 Then Exit Function
    s = html
    ' 1. viewport meta (AI should emit it, but never assume)
    If InStr(1, s, "name=""viewport""", 1) = 0 Then
        cPos = InStr(1, s, "<meta", 1)
        If cPos = 0 Then cPos = InStr(1, s, "<head", 1)
        If cPos > 0 Then
            insPos = InStr(cPos, s, ">", 0)
            If insPos > 0 Then s = Left(s, insPos) & vbCrLf & "    <meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & Mid(s, insPos + 1)
        End If
    End If
    ' 2. guard CSS, exactly once
    If InStr(1, s, "/*autonome-guard*/", 1) = 0 Then
        guardCss = "/*autonome-guard*/" & vbCrLf & _
            "/* responsive safety net, keep - prevents mobile horizontal overflow */" & vbCrLf & _
            "html{scroll-padding-top:84px;}" & vbCrLf & _
            "html,body{overflow-x:hidden;overflow-x:clip;max-width:100%;}" & vbCrLf & _
            "img,video,iframe{max-width:100%;height:auto;}" & vbCrLf & _
            "h1,h2,h3,h4,.display-1,.display-2,.display-3,.display-4,.display-5,.display-6{overflow-wrap:break-word;}"
        styleOpen = InStr(1, s, "<style", 1)
        If styleOpen > 0 Then
            styleEnd = InStr(styleOpen, s, "</style>", 1)
            If styleEnd > 0 Then
                s = Left(s, styleEnd - 1) & guardCss & vbCrLf & Mid(s, styleEnd)
            End If
        Else
            headClose = InStr(1, s, "</head>", 1)
            If headClose > 0 Then
                s = Left(s, headClose - 1) & "    <style>" & vbCrLf & guardCss & vbCrLf & "    </style>" & vbCrLf & Mid(s, headClose)
            End If
        End If
    End If
    EnsureResponsiveGuard = s
End Function

' True when content still looks like a complete HTML document
' (</head>, <body, </body>, </html> present and in the right order).
Function HtmlSkeletonOk(content)
    Dim pHead, pBodyO, pBodyC, pHtmlC
    HtmlSkeletonOk = False
    pHead = InStr(1, content, "</head>", 1)
    pBodyO = InStr(1, content, "<body", 1)
    pBodyC = InStr(1, content, "</body>", 1)
    pHtmlC = InStr(1, content, "</html>", 1)
    If pHead = 0 Or pBodyO = 0 Or pBodyC = 0 Or pHtmlC = 0 Then Exit Function
    If pHead < pBodyO And pBodyO < pBodyC And pBodyC < pHtmlC Then HtmlSkeletonOk = True
End Function

' =================== BLOCK IDS (<!--ID:XX-->...<!--/ID:XX-->) ===================
' Autonome owns all block IDs. The AI never invents them: it only echoes the
' received IDs as locators (or <!--ID:NEW--> for a new block). After every
' successful content return the server strips all old markers and re-tags the
' whole document with fresh numbers, so IDs are always unique and sequential.
' Phase 1 granularity (basic): top-level body blocks (header/nav/main/section/
' footer/div/article/aside) plus whole <script> elements in body and whole
' <style>/<script> elements in <head>. Title/meta/CDN/OG tags stay untagged.

Function HasIdMarkers(s)
    HasIdMarkers = False
    If IsNull(s) Then Exit Function
    If InStr(1, s, "<!--ID:", 1) > 0 Then HasIdMarkers = True
End Function

' Remove every own marker so re-tagging is idempotent. Handles HTML markers
' (open <!--ID:..--> and close <!--/ID:..-->) and native CSS/JS markers
' (open /*ID:..*/ and close /*/ID:..*/) for future finer granularity.
Function StripIdMarkers(s)
    Dim out, p, q
    If IsNull(s) Then
        StripIdMarkers = ""
        Exit Function
    End If
    out = s
    Do While True
        p = InStr(1, out, "<!--ID:", 1)
        If p = 0 Then Exit Do
        q = InStr(p, out, "-->", 0)
        If q = 0 Then Exit Do
        out = Left(out, p - 1) & Mid(out, q + 3)
    Loop
    Do While True
        p = InStr(1, out, "<!--/ID:", 1)
        If p = 0 Then Exit Do
        q = InStr(p, out, "-->", 0)
        If q = 0 Then Exit Do
        out = Left(out, p - 1) & Mid(out, q + 3)
    Loop
    Do While True
        p = InStr(1, out, "/*ID:", 1)
        If p = 0 Then Exit Do
        q = InStr(p, out, "*/", 0)
        If q = 0 Then Exit Do
        out = Left(out, p - 1) & Mid(out, q + 2)
    Loop
    Do While True
        p = InStr(1, out, "/*/ID:", 1)
        If p = 0 Then Exit Do
        q = InStr(p, out, "*/", 0)
        If q = 0 Then Exit Do
        out = Left(out, p - 1) & Mid(out, q + 2)
    Loop
    StripIdMarkers = out
End Function

Function IdIsVoidElement(tagName)
    Dim t
    t = LCase(tagName)
    IdIsVoidElement = (t = "img" Or t = "br" Or t = "hr" Or t = "meta" Or _
        t = "link" Or t = "input" Or t = "source" Or t = "track" Or _
        t = "wbr" Or t = "embed" Or t = "param" Or t = "col" Or t = "base")
End Function

' Tag name at a "<" position (skips </ and whitespace). Returns "" when none.
Function IdTagNameAt(s, ltPos)
    Dim i, ch, name
    IdTagNameAt = ""
    If ltPos + 1 > Len(s) Then Exit Function
    i = ltPos + 1
    ch = Mid(s, i, 1)
    If ch = "/" Then i = i + 1
    If ch = "!" Or ch = "?" Then Exit Function
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = ">" Or ch = "/" Then Exit Do
        i = i + 1
    Loop
    name = Mid(s, ltPos + 1, i - ltPos - 1)
    If Left(name, 1) = "/" Then name = Mid(name, 2)
    If Right(name, 1) = "/" Then name = Left(name, Len(name) - 1)
    If Right(name, 1) = ">" Then name = Left(name, Len(name) - 1)
    IdTagNameAt = LCase(name)
End Function

' True when "<" starts a plausible tag: optional "/", a letter-led name of
' letters/digits/-/:, then whitespace, "/" or ">". Rejects typos like <br(
' or JS comparisons like a<b, which must be treated as plain text so one
' malformed "<" can never swallow the next ">" (and with it, real closers).
Function IdTagOpenAt(s, lt)
    Dim i, ch
    IdTagOpenAt = False
    i = lt + 1
    If i > Len(s) Then Exit Function
    ch = Mid(s, i, 1)
    If ch = "/" Then
        i = i + 1
        If i > Len(s) Then Exit Function
        ch = Mid(s, i, 1)
    End If
    If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z")) Then Exit Function
    i = i + 1
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "-" Or ch = ":" Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    If i > Len(s) Then Exit Function
    ch = Mid(s, i, 1)
    If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = "/" Or ch = ">" Then IdTagOpenAt = True
End Function

' Position of the ">" closing the tag opened at lt, skipping ">" inside
' quoted attribute values (e.g. alt="a>b"). Returns 0 when lt starts no
' valid tag or the tag never closes.
Function IdTagEnd(s, lt)
    Dim i, ch, q
    IdTagEnd = 0
    If Not IdTagOpenAt(s, lt) Then Exit Function
    i = lt + 1
    q = ""
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If q = "" Then
            If ch = """" Or ch = "'" Then
                q = ch
            ElseIf ch = ">" Then
                IdTagEnd = i
                Exit Function
            End If
        Else
            If ch = q Then q = ""
        End If
        i = i + 1
    Loop
End Function

Function IdIsBlockContainer(tagName)
    Dim t
    t = LCase(tagName)
    IdIsBlockContainer = (t = "header" Or t = "nav" Or t = "main" Or _
        t = "section" Or t = "footer" Or t = "div" Or t = "article" Or _
        t = "aside" Or t = "script" Or t = "style" Or _
        t = "ul" Or t = "ol" Or t = "li")
End Function

' Count direct-child (depth 0) tags in s. Returns Dictionary tag->count.
' Skips comments and raw <script>/<style> bodies; void elements count
' without changing depth.
Function DirectChildCounts(s)
    Dim d, pos, lt, gt, tn, depth, isClose, selfClose, cEnd
    Set d = Server.CreateObject("Scripting.Dictionary")
    pos = 1
    depth = 0
    Do While pos <= Len(s)
        lt = InStr(pos, s, "<", 0)
        If lt = 0 Then Exit Do
        If Mid(s, lt + 1, 3) = "!--" Then
            gt = InStr(lt + 4, s, "-->", 0)
            If gt = 0 Then Exit Do
            pos = gt + 3
        ElseIf Not IdTagOpenAt(s, lt) Then
            ' Typo/text "<" (e.g. "<br(" or "a<b"): skip one char, depth untouched
            pos = lt + 1
        Else
            gt = IdTagEnd(s, lt)
            If gt = 0 Then Exit Do
            tn = IdTagNameAt(s, lt)
            isClose = (Mid(s, lt + 1, 1) = "/")
            selfClose = (Mid(s, gt - 1, 1) = "/")
            If tn = "script" Or tn = "style" Then
                If Not isClose And Not selfClose Then
                    If depth = 0 Then
                        If d.Exists(tn) Then d(tn) = d(tn) + 1 Else d.Add tn, 1
                    End If
                    cEnd = InStr(gt + 1, s, "</" & tn & ">", 1)
                    If cEnd = 0 Then Exit Do
                    pos = cEnd + Len(tn) + 3
                Else
                    pos = gt + 1
                End If
            ElseIf tn = "" Then
                pos = lt + 1
            ElseIf isClose Then
                If depth > 0 Then depth = depth - 1
                pos = gt + 1
            ElseIf selfClose Or IdIsVoidElement(tn) Then
                If depth = 0 Then
                    If d.Exists(tn) Then d(tn) = d(tn) + 1 Else d.Add tn, 1
                End If
                pos = gt + 1
            Else
                If depth = 0 Then
                    If d.Exists(tn) Then d(tn) = d(tn) + 1 Else d.Add tn, 1
                End If
                depth = depth + 1
                pos = gt + 1
            End If
        End If
    Loop
    Set DirectChildCounts = d
End Function

' True when a container stays an untagged shell so its children keep their
' own IDs: a <section> anywhere inside, or 2+ same-tag direct children that
' are repeatable items (cards/cols, articles, list items).
Function ShouldSplitShell(innerRange)
    Dim d, k
    ShouldSplitShell = False
    If InStr(1, innerRange, "<section", 1) > 0 Then
        ShouldSplitShell = True
        Exit Function
    End If
    Set d = DirectChildCounts(innerRange)
    For Each k In d.Keys
        If k = "div" Or k = "article" Or k = "li" Then
            If d(k) >= 2 Then
                ShouldSplitShell = True
                Exit Function
            End If
        End If
    Next
End Function

' Position of the first depth-0 opening tag with the given name, or 0.
Function FindDepthZeroOpen(s, tagName)
    Dim pos, lt, gt, tn, depth, isClose, selfClose, cEnd
    FindDepthZeroOpen = 0
    pos = 1
    depth = 0
    Do While pos <= Len(s)
        lt = InStr(pos, s, "<", 0)
        If lt = 0 Then Exit Function
        If Mid(s, lt + 1, 3) = "!--" Then
            gt = InStr(lt + 4, s, "-->", 0)
            If gt = 0 Then Exit Function
            pos = gt + 3
        ElseIf Not IdTagOpenAt(s, lt) Then
            pos = lt + 1
        Else
            gt = IdTagEnd(s, lt)
            If gt = 0 Then Exit Function
            tn = IdTagNameAt(s, lt)
            isClose = (Mid(s, lt + 1, 1) = "/")
            selfClose = (Mid(s, gt - 1, 1) = "/")
            If tn = "script" Or tn = "style" Then
                If Not isClose And Not selfClose Then
                    cEnd = InStr(gt + 1, s, "</" & tn & ">", 1)
                    If cEnd = 0 Then Exit Function
                    pos = cEnd + Len(tn) + 3
                Else
                    pos = gt + 1
                End If
            ElseIf tn = "" Then
                pos = lt + 1
            ElseIf isClose Then
                If depth > 0 Then depth = depth - 1
                pos = gt + 1
            ElseIf selfClose Or IdIsVoidElement(tn) Then
                pos = gt + 1
            Else
                If depth = 0 And tn = LCase(tagName) Then
                    FindDepthZeroOpen = lt
                    Exit Function
                End If
                depth = depth + 1
                pos = gt + 1
            End If
        End If
    Loop
End Function

' Wrap repeating item-children nested inside an element that itself stays one
' block (e.g. cards inside a section: section keeps its ID, each card gets a
' nested ID). Looks through single structural wrappers (div/article/aside).
' One extra level per call; wrappers recurse with level+1 (level>4 stops).
Function NestRepeatItems(s, nextId, level)
    Dim d, k, repDiv, repSec, repArt, repLi, hasRep
    Dim out, pos, lt, gt, tn, depth, isClose, selfClose, chunkEnd, cEnd
    Dim singleTag, singleCount, wLt, wGt, wFin, wCls, wInner
    NestRepeatItems = s
    If level > 4 Then Exit Function
    If nextId > 40 Then Exit Function
    If Len(s) = 0 Then Exit Function
    Set d = DirectChildCounts(s)
    repDiv = False
    repSec = False
    repArt = False
    repLi = False
    If d.Exists("div") Then
        If d("div") >= 2 Then repDiv = True
    End If
    If d.Exists("section") Then
        If d("section") >= 2 Then repSec = True
    End If
    If d.Exists("article") Then
        If d("article") >= 2 Then repArt = True
    End If
    If d.Exists("li") Then
        If d("li") >= 2 Then repLi = True
    End If
    hasRep = (repDiv Or repSec Or repArt Or repLi)
    If Not hasRep Then
        ' No direct repeats: descend once through a single structural
        ' wrapper (e.g. section > h2 + div.row: the row holds the cards).
        singleTag = ""
        singleCount = 0
        For Each k In d.Keys
            If k = "div" Or k = "article" Or k = "aside" Then
                If d(k) = 1 Then
                    singleCount = singleCount + 1
                    singleTag = k
                Else
                    singleCount = 99
                End If
            ElseIf k = "section" Or k = "header" Or k = "footer" Or k = "nav" Or k = "main" Or k = "ul" Or k = "ol" Or k = "li" Or k = "script" Or k = "style" Then
                singleCount = 99
            End If
        Next
        If singleCount = 1 Then
            wLt = FindDepthZeroOpen(s, singleTag)
            If wLt > 0 Then
                wGt = IdTagEnd(s, wLt)
                wFin = IdFindMatchingClose(s, wLt, singleTag)
                If wFin > 0 And wGt > 0 Then
                    wCls = InStrRev(s, "<", wFin - 1)
                    If wCls > wGt Then
                        wInner = Mid(s, wGt + 1, wCls - wGt - 1)
                        wInner = NestRepeatItems(wInner, nextId, level + 1)
                        NestRepeatItems = Left(s, wGt) & wInner & Mid(s, wCls)
                        Exit Function
                    End If
                End If
            End If
        End If
        Exit Function
    End If
    out = ""
    pos = 1
    depth = 0
    Do While pos <= Len(s)
        lt = InStr(pos, s, "<", 0)
        If lt = 0 Then
            out = out & Mid(s, pos)
            Exit Do
        End If
        If Mid(s, lt + 1, 3) = "!--" Then
            gt = InStr(lt + 4, s, "-->", 0)
            If gt = 0 Then
                out = out & Mid(s, pos)
                Exit Do
            End If
            out = out & Mid(s, pos, gt + 3 - pos)
            pos = gt + 3
        ElseIf Not IdTagOpenAt(s, lt) Then
            out = out & Mid(s, pos, lt - pos + 1)
            pos = lt + 1
        Else
            gt = IdTagEnd(s, lt)
            If gt = 0 Then
                out = out & Mid(s, pos)
                Exit Do
            End If
            tn = IdTagNameAt(s, lt)
            isClose = (Mid(s, lt + 1, 1) = "/")
            selfClose = (Mid(s, gt - 1, 1) = "/")
            If tn = "script" Or tn = "style" Then
                If Not isClose And Not selfClose Then
                    cEnd = InStr(gt + 1, s, "</" & tn & ">", 1)
                    If cEnd = 0 Then
                        out = out & Mid(s, pos)
                        Exit Do
                    End If
                    out = out & Mid(s, pos, cEnd + Len(tn) + 3 - pos)
                    pos = cEnd + Len(tn) + 3
                Else
                    out = out & Mid(s, pos, gt + 1 - pos)
                    pos = gt + 1
                End If
            ElseIf tn = "" Then
                out = out & Mid(s, pos, lt - pos + 1)
                pos = lt + 1
            ElseIf isClose Then
                If depth > 0 Then depth = depth - 1
                out = out & Mid(s, pos, gt + 1 - pos)
                pos = gt + 1
            ElseIf selfClose Or IdIsVoidElement(tn) Then
                out = out & Mid(s, pos, gt + 1 - pos)
                pos = gt + 1
            ElseIf depth = 0 Then
                If (tn = "div" And repDiv) Or (tn = "section" And repSec) Or (tn = "article" And repArt) Or (tn = "li" And repLi) Then
                    chunkEnd = IdFindMatchingClose(s, lt, tn)
                    If chunkEnd = 0 Then
                        out = out & Mid(s, pos)
                        Exit Do
                    End If
                    out = out & Mid(s, pos, lt - pos) & "<!--ID:" & nextId & "-->" & Mid(s, lt, chunkEnd - lt) & "<!--/ID:" & nextId & "-->"
                    nextId = nextId + 1
                    pos = chunkEnd
                Else
                    depth = depth + 1
                    out = out & Mid(s, pos, gt + 1 - pos)
                    pos = gt + 1
                End If
            Else
                depth = depth + 1
                out = out & Mid(s, pos, gt + 1 - pos)
                pos = gt + 1
            End If
        End If
        If nextId > 40 Then
            out = out & Mid(s, pos)
            Exit Do
        End If
    Loop
    NestRepeatItems = out
End Function

' Find the position just after the matching close tag for the element opened
' at openPos (position of "<"). Returns 0 when not found. Counts nested
' elements with the same tag name; skips comments and raw <script>/<style>.
Function IdFindMatchingClose(s, openPos, tagName)
    Dim openEnd, depth, p, lt, gt, tn, isClose, selfClose
    IdFindMatchingClose = 0
    openEnd = IdTagEnd(s, openPos)
    If openEnd = 0 Then Exit Function
    If Mid(s, openEnd - 1, 1) = "/" Then Exit Function ' self-closing opener
    depth = 1
    p = openEnd + 1
    Do While p <= Len(s)
        lt = InStr(p, s, "<", 0)
        If lt = 0 Then Exit Function
        ' Skip HTML comments
        If Mid(s, lt + 1, 3) = "!--" Then
            gt = InStr(lt + 4, s, "-->", 0)
            If gt = 0 Then Exit Function
            p = gt + 3
        ElseIf Not IdTagOpenAt(s, lt) Then
            ' Typo/text "<" (e.g. "<br(" or JS "a<b"): skip one char
            p = lt + 1
        Else
            gt = IdTagEnd(s, lt)
            If gt = 0 Then Exit Function
            tn = IdTagNameAt(s, lt)
            If tn = LCase(tagName) Then
                isClose = (Mid(s, lt + 1, 1) = "/")
                selfClose = (Mid(s, gt - 1, 1) = "/")
                If Not selfClose And Not IdIsVoidElement(tn) Then
                    If isClose Then
                        depth = depth - 1
                        If depth = 0 Then
                            IdFindMatchingClose = gt + 1
                            Exit Function
                        End If
                    Else
                        depth = depth + 1
                    End If
                End If
            End If
            ' Skip raw script/style bodies when scanning other containers
            If tn = "script" Or tn = "style" Then
                If Mid(s, lt + 1, 1) <> "/" Then
                    If LCase(tagName) <> tn Then
                        Dim cPos
                        cPos = InStr(gt + 1, s, "</" & tn & ">", 1)
                        If cPos = 0 Then Exit Function
                        p = cPos + Len(tn) + 3
                    Else
                        p = gt + 1
                    End If
                Else
                    p = gt + 1
                End If
            Else
                p = gt + 1
            End If
        End If
    Loop
End Function

' Tag the whole document with fresh <!--ID:n-->...<!--/ID:n--> markers.
' Always strips old markers first; <head> stays untagged.
Function TagDocument(html)
    Dim clean, bodyTag, bodyOpenEnd, bodyClose, headPart, inner, tailPart, nextId
    If IsNull(html) Then
        TagDocument = ""
        Exit Function
    End If
    clean = StripIdMarkers(html)
    bodyTag = InStr(1, clean, "<body", 1)
    If bodyTag = 0 Then
        TagDocument = clean
        Exit Function
    End If
    bodyOpenEnd = IdTagEnd(clean, bodyTag)
    If bodyOpenEnd = 0 Then bodyOpenEnd = InStr(bodyTag, clean, ">", 0)
    bodyClose = InStr(bodyOpenEnd, clean, "</body>", 1)
    If bodyOpenEnd = 0 Or bodyClose = 0 Then
        TagDocument = clean
        Exit Function
    End If
    headPart = Left(clean, bodyOpenEnd)
    inner = Mid(clean, bodyOpenEnd + 1, bodyClose - bodyOpenEnd - 1)
    tailPart = Mid(clean, bodyClose)
    nextId = 1
    ' Document order: head <style>/<script> first, so theme and color
    ' changes are addressable via a block ID instead of a full return.
    headPart = TagHeadStyles(headPart, nextId)
    TagDocument = headPart & TagInnerBody(inner, nextId) & tailPart
End Function

' Wrap <style> and <script> elements inside <head> with ID markers.
' Everything else in <head> (title/meta/CDN/OG tags) stays untagged.
Function TagHeadStyles(head, nextId)
    Dim out, pos, lt, gt, tn, chunkEnd
    out = ""
    pos = 1
    Do While pos <= Len(head)
        lt = InStr(pos, head, "<", 0)
        If lt = 0 Then
            out = out & Mid(head, pos)
            Exit Do
        End If
        If Mid(head, lt + 1, 3) = "!--" Then
            gt = InStr(lt + 4, head, "-->", 0)
            If gt = 0 Then
                out = out & Mid(head, pos)
                Exit Do
            End If
            out = out & Mid(head, pos, gt + 3 - pos)
            pos = gt + 3
        ElseIf Not IdTagOpenAt(head, lt) Then
            ' Typo/text "<": copy through untouched
            out = out & Mid(head, pos, lt - pos + 1)
            pos = lt + 1
        Else
            tn = IdTagNameAt(head, lt)
            If (tn = "style" Or tn = "script") And Mid(head, lt + 1, 1) <> "/" Then
                gt = IdTagEnd(head, lt)
                If gt = 0 Then
                    out = out & Mid(head, pos)
                    Exit Do
                End If
                If Mid(head, gt - 1, 1) = "/" Then
                    out = out & Mid(head, pos, lt - pos) & "<!--ID:" & nextId & "-->" & Mid(head, lt, gt - lt + 1) & "<!--/ID:" & nextId & "-->"
                    nextId = nextId + 1
                    pos = gt + 1
                Else
                    chunkEnd = IdFindMatchingClose(head, lt, tn)
                    If chunkEnd = 0 Then
                        out = out & Mid(head, pos)
                        Exit Do
                    Else
                        out = out & Mid(head, pos, lt - pos) & "<!--ID:" & nextId & "-->" & Mid(head, lt, chunkEnd - lt) & "<!--/ID:" & nextId & "-->"
                        nextId = nextId + 1
                        pos = chunkEnd
                    End If
                End If
            Else
                out = out & Mid(head, pos, lt - pos + 1)
                pos = lt + 1
            End If
        End If
        If nextId > 40 Then
            If pos <= Len(head) Then
                out = out & "<!--ID:" & nextId & "-->" & Mid(head, pos) & "<!--/ID:" & nextId & "-->"
                nextId = nextId + 1
            End If
            Exit Do
        End If
    Loop
    TagHeadStyles = out
End Function

' Tag one body-level string; the running number in nextId is shared with the
' caller (VBScript passes ByRef by default). <section>/<script>/<style> are
' always atomic blocks. Other containers holding sections or repeatable items
' (cards/cols, articles, list items) stay untagged shells so every child keeps
' its own ID; containers without those become one block each. Numbering is
' always fresh per document (TagDocument strips first); IDs are never stored.
Function TagInnerBody(s, nextId)
    Dim out, pos, lt, gt, tn, chunkEnd, openEnd, closeStart, innerRange, trimmed
    out = ""
    pos = 1
    Do While pos <= Len(s)
        lt = InStr(pos, s, "<", 0)
        If lt = 0 Then
            trimmed = Trim(Mid(s, pos))
            If Len(trimmed) > 30 Then
                out = out & "<!--ID:" & nextId & "-->" & Mid(s, pos) & "<!--/ID:" & nextId & "-->"
                nextId = nextId + 1
            Else
                out = out & Mid(s, pos)
            End If
            Exit Do
        End If
        ' Keep comments / whitespace between blocks untagged
        If Mid(s, lt + 1, 3) = "!--" Then
            gt = InStr(lt + 4, s, "-->", 0)
            If gt = 0 Then
                out = out & Mid(s, pos)
                Exit Do
            End If
            out = out & Mid(s, pos, gt + 3 - pos)
            pos = gt + 3
        ElseIf Not IdTagOpenAt(s, lt) Then
            ' Typo/text "<" (e.g. "<br(" or JS "a<b"): copy through untouched
            out = out & Mid(s, pos, lt - pos + 1)
            pos = lt + 1
        Else
            tn = IdTagNameAt(s, lt)
            If tn = "" Or Mid(s, lt + 1, 1) = "/" Then
                out = out & Mid(s, pos, lt - pos + 1)
                pos = lt + 1
            ElseIf tn = "section" Or tn = "script" Or tn = "style" Then
                If lt > pos Then out = out & Mid(s, pos, lt - pos)
                gt = IdTagEnd(s, lt)
                If gt = 0 Then
                    out = out & Mid(s, lt)
                    Exit Do
                End If
                If Mid(s, gt - 1, 1) = "/" Or IdIsVoidElement(tn) Then
                    out = out & "<!--ID:" & nextId & "-->" & Mid(s, lt, gt - lt + 1) & "<!--/ID:" & nextId & "-->"
                    nextId = nextId + 1
                    pos = gt + 1
                ElseIf tn = "section" Then
                    chunkEnd = IdFindMatchingClose(s, lt, tn)
                    If chunkEnd = 0 Then
                        ' Unbalanced: tag rest as one block so nothing is lost
                        out = out & "<!--ID:" & nextId & "-->" & Mid(s, lt) & "<!--/ID:" & nextId & "-->"
                        nextId = nextId + 1
                        Exit Do
                    Else
                        openEnd = gt
                        closeStart = InStrRev(s, "<", chunkEnd - 1)
                        If closeStart <= openEnd Then closeStart = chunkEnd
                        innerRange = Mid(s, openEnd + 1, closeStart - openEnd - 1)
                        If InStr(1, innerRange, "<section", 1) > 0 Then
                            ' Nested sections surface individually
                            out = out & Mid(s, lt, openEnd - lt + 1) & TagInnerBody(innerRange, nextId) & Mid(s, closeStart, chunkEnd - closeStart)
                        Else
                            ' Section keeps its ID; repeatable items inside
                            ' (cards, list entries) get nested IDs
                            If nextId <= 40 Then innerRange = NestRepeatItems(innerRange, nextId, 0)
                            out = out & "<!--ID:" & nextId & "-->" & Mid(s, lt, openEnd - lt + 1) & innerRange & Mid(s, closeStart, chunkEnd - closeStart) & "<!--/ID:" & nextId & "-->"
                            nextId = nextId + 1
                        End If
                        pos = chunkEnd
                    End If
                Else
                    chunkEnd = IdFindMatchingClose(s, lt, tn)
                    If chunkEnd = 0 Then
                        ' Unbalanced: tag rest as one block so nothing is lost
                        out = out & "<!--ID:" & nextId & "-->" & Mid(s, lt) & "<!--/ID:" & nextId & "-->"
                        nextId = nextId + 1
                        Exit Do
                    Else
                        out = out & "<!--ID:" & nextId & "-->" & Mid(s, lt, chunkEnd - lt) & "<!--/ID:" & nextId & "-->"
                        nextId = nextId + 1
                        pos = chunkEnd
                    End If
                End If
            ElseIf IdIsBlockContainer(tn) Then
                If lt > pos Then out = out & Mid(s, pos, lt - pos)
                gt = IdTagEnd(s, lt)
                If gt = 0 Then
                    out = out & Mid(s, lt)
                    Exit Do
                End If
                If Mid(s, gt - 1, 1) = "/" Or IdIsVoidElement(tn) Then
                    out = out & "<!--ID:" & nextId & "-->" & Mid(s, lt, gt - lt + 1) & "<!--/ID:" & nextId & "-->"
                    nextId = nextId + 1
                    pos = gt + 1
                Else
                    chunkEnd = IdFindMatchingClose(s, lt, tn)
                    If chunkEnd = 0 Then
                        ' Unbalanced: tag rest as one block so nothing is lost
                        out = out & "<!--ID:" & nextId & "-->" & Mid(s, lt) & "<!--/ID:" & nextId & "-->"
                        nextId = nextId + 1
                        Exit Do
                    Else
                        openEnd = gt
                        closeStart = InStrRev(s, "<", chunkEnd - 1)
                        If closeStart <= openEnd Then closeStart = chunkEnd
                        innerRange = Mid(s, openEnd + 1, closeStart - openEnd - 1)
                        If nextId <= 40 And ShouldSplitShell(innerRange) Then
                            ' Shell with sections/items: recurse untagged so
                            ' each child keeps its own ID. Past ~40 blocks the
                            ' rest is wrapped whole (cap against ID overload).
                            out = out & Mid(s, lt, openEnd - lt + 1) & TagInnerBody(innerRange, nextId) & Mid(s, closeStart, chunkEnd - closeStart)
                        Else
                            ' Whole block, but repeatable items inside still
                            ' get nested IDs (e.g. cards in a wrapped div)
                            If nextId <= 40 Then innerRange = NestRepeatItems(innerRange, nextId, 0)
                            out = out & "<!--ID:" & nextId & "-->" & Mid(s, lt, openEnd - lt + 1) & innerRange & Mid(s, closeStart, chunkEnd - closeStart) & "<!--/ID:" & nextId & "-->"
                            nextId = nextId + 1
                        End If
                        pos = chunkEnd
                    End If
                End If
            Else
                ' Inline content between blocks: keep as-is until next container
                out = out & Mid(s, pos, lt - pos + 1)
                pos = lt + 1
            End If
        End If
        If nextId > 40 Then
            ' Cap against ID overload: wrap the rest whole as one final
            ' block instead of splitting further (max ~41 IDs per document)
            If pos <= Len(s) Then
                out = out & "<!--ID:" & nextId & "-->" & Mid(s, pos) & "<!--/ID:" & nextId & "-->"
                nextId = nextId + 1
            End If
            Exit Do
        End If
    Loop
    TagInnerBody = out
End Function

' Collect every fenced ```html block from AI output (joined). Null if none.
Function CollectFencedHtml(responseText)
    Dim combined, p1, p2, fence
    CollectFencedHtml = Null
    If IsNull(responseText) Then Exit Function
    If Len(responseText) = 0 Then Exit Function
    combined = ""
    p1 = 1
    Do While True
        p1 = InStr(p1, responseText, "```html", 1)
        If p1 = 0 Then Exit Do
        p1 = p1 + Len("```html")
        p2 = InStr(p1, responseText, "```", 0)
        If p2 = 0 Then Exit Do
        If Len(combined) > 0 Then combined = combined & vbCrLf
        combined = combined & Trim(Mid(responseText, p1, p2 - p1))
        p1 = p2 + 3
    Loop
    If Len(combined) = 0 Then
        p1 = InStr(1, responseText, "```", 1)
        If p1 > 0 Then
            p1 = p1 + 3
            p2 = InStr(p1, responseText, "```", 0)
            If p2 > p1 Then combined = Trim(Mid(responseText, p1, p2 - p1))
        End If
    End If
    If Len(combined) = 0 Then Exit Function
    CollectFencedHtml = combined
End Function

' Parse echoed partial blocks <!--ID:..-->...<!--/ID:..--> from AI output.
' Returns a VBScript array of Dictionaries {id, content} or Null when unusable.
' Outer ID wins: nested markers inside a returned block are stripped, so a
' larger replacement block never breaks the flat numbering (fresh IDs are
' assigned by TagDocument after the merge anyway).
Function ParsePartialBlocks(responseText)
    Dim combined, entries(), n, p, q, idStr, inner, closeP
    ParsePartialBlocks = Null
    combined = CollectFencedHtml(responseText)
    If IsNull(combined) Then
        If InStr(1, responseText, "<!--ID:", 1) = 0 Then Exit Function
        combined = responseText
    End If
    n = -1
    p = 1
    Do While True
        p = InStr(p, combined, "<!--ID:", 1)
        If p = 0 Then Exit Do
        q = InStr(p, combined, "-->", 0)
        If q = 0 Then Exit Do
        idStr = Trim(Mid(combined, p + 7, q - p - 7))
        If Len(idStr) = 0 Then Exit Do
        closeP = InStr(q + 3, combined, "<!--/ID:" & idStr & "-->", 1)
        If closeP = 0 Then Exit Do
        inner = Mid(combined, q + 3, closeP - q - 3)
        inner = StripIdMarkers(inner) ' outer wins: drop nested markers
        n = n + 1
        ReDim Preserve entries(n)
        Dim e
        Set e = Server.CreateObject("Scripting.Dictionary")
        e.Add "id", idStr
        e.Add "content", inner
        Set entries(n) = e
        p = closeP + Len("<!--/ID:" & idStr & "-->")
    Loop
    If n < 0 Then Exit Function
    ParsePartialBlocks = entries
End Function

' Merge echoed partial blocks into the tagged base document.
' Replace by ID; empty content (or DELETE) clears the block; ID:NEW blocks
' are inserted at their placement anchor (AFTER/FIRST/LAST/INTO).
' Returns merged HTML or Null on any mismatch.
Function MergeIdBlocks(baseHtml, entries)
    Dim merged, i, idStr, content, t, openM, openEnd, closeM, newArr, newCount, moved
    MergeIdBlocks = Null
    If IsNull(baseHtml) Then Exit Function
    If Not IsArray(entries) Then Exit Function
    merged = baseHtml
    newCount = -1
    For i = 0 To UBound(entries)
        idStr = Trim(entries(i)("id"))
        content = entries(i)("content")
        If UCase(idStr) = "NEW" Then
            If Len(Trim(content)) > 0 Then
                If UCase(Trim(content)) <> "DELETE" Then
                    newCount = newCount + 1
                    ReDim Preserve newArr(newCount)
                    newArr(newCount) = content
                End If
            End If
        Else
            openM = InStr(1, merged, "<!--ID:" & idStr & "-->", 1)
            closeM = InStr(1, merged, "<!--/ID:" & idStr & "-->", 1)
            If openM = 0 Or closeM = 0 Then Exit Function ' unknown ID
            If closeM < openM Then Exit Function
            openEnd = openM + Len("<!--ID:" & idStr & "-->")
            t = Trim(content)
            If UCase(t) = "DELETE" Then t = ""
            If Len(t) > 0 Then
                ' Move when the echoed block carries a placement anchor,
                ' otherwise a plain in-place replace (anchors never leak)
                moved = MoveBlock(merged, idStr, content)
                If IsNull(moved) Then
                    merged = Left(merged, openEnd - 1) & StripAnchorComments(t) & Mid(merged, closeM)
                Else
                    merged = moved
                End If
            Else
                merged = Left(merged, openEnd - 1) & t & Mid(merged, closeM)
            End If
        End If
    Next
    If newCount >= 0 Then
        ' Reverse order: each anchor resolves against the current document,
        ' so applying last-first keeps the listed order for shared anchors
        For i = newCount To 0 Step -1
            merged = InsertNewBlock(merged, newArr(i))
            If IsNull(merged) Then Exit Function
        Next
    End If
    ' Silent-corruption guard: consent anchors present in the base must
    ' survive the merge (a wrong echoed ID must never drop them)
    If Not MergeKeepsAnchors(baseHtml, merged) Then Exit Function
    MergeIdBlocks = merged
End Function

' Remove placement anchor comments so they never reach the file
Function StripAnchorComments(s)
    Dim out, a, b
    If IsNull(s) Then
        StripAnchorComments = ""
        Exit Function
    End If
    out = s
    Do While True
        a = InStr(1, out, "<!--AFTER:", 1)
        If a = 0 Then Exit Do
        b = InStr(a, out, "-->", 0)
        If b = 0 Then Exit Do
        out = Left(out, a - 1) & Mid(out, b + 3)
    Loop
    Do While True
        a = InStr(1, out, "<!--INTO:", 1)
        If a = 0 Then Exit Do
        b = InStr(a, out, "-->", 0)
        If b = 0 Then Exit Do
        out = Left(out, a - 1) & Mid(out, b + 3)
    Loop
    Do While True
        a = InStr(1, out, "<!--FIRST-->", 1)
        If a = 0 Then Exit Do
        out = Left(out, a - 1) & Mid(out, a + 12)
    Loop
    Do While True
        a = InStr(1, out, "<!--LAST-->", 1)
        If a = 0 Then Exit Do
        out = Left(out, a - 1) & Mid(out, a + 11)
    Loop
    StripAnchorComments = out
End Function

' Parse a placement anchor from block content. Sets kind (AFTER/FIRST/LAST/
' INTO) and id, returns True when an anchor comment was found.
Function ParseAnchor(content, anchorKind, anchorId)
    Dim a, b
    ParseAnchor = False
    anchorKind = ""
    anchorId = ""
    If IsNull(content) Then Exit Function
    a = InStr(1, content, "<!--AFTER:", 1)
    If a > 0 Then
        b = InStr(a, content, "-->", 0)
        If b > 0 Then
            anchorKind = "AFTER"
            anchorId = Trim(Mid(content, a + 10, b - a - 10))
            ParseAnchor = True
            Exit Function
        End If
    End If
    If InStr(1, content, "<!--FIRST-->", 1) > 0 Then
        anchorKind = "FIRST"
        anchorId = ""
        ParseAnchor = True
        Exit Function
    End If
    a = InStr(1, content, "<!--INTO:", 1)
    If a > 0 Then
        b = InStr(a, content, "-->", 0)
        If b > 0 Then
            anchorKind = "INTO"
            anchorId = Trim(Mid(content, a + 9, b - a - 9))
            ParseAnchor = True
            Exit Function
        End If
    End If
    If InStr(1, content, "<!--LAST-->", 1) > 0 Then
        anchorKind = "LAST"
        anchorId = ""
        ParseAnchor = True
        Exit Function
    End If
End Function

' Insert blockText at an anchor. selfId is the moving block's own id ("" for
' NEW). Returns Null when the anchor cannot be resolved (caller falls back).
Function InsertAt(merged, blockText, anchorKind, anchorId, selfId)
    Dim insAt, bodyTag, bodyEnd, closeM
    InsertAt = Null
    If IsNull(merged) Or IsNull(blockText) Then Exit Function
    If anchorKind = "FIRST" Or (anchorKind = "AFTER" And anchorId = "0") Then
        bodyTag = InStr(1, merged, "<body", 1)
        If bodyTag = 0 Then Exit Function
        bodyEnd = IdTagEnd(merged, bodyTag)
        If bodyEnd = 0 Then bodyEnd = InStr(bodyTag, merged, ">", 0)
        If bodyEnd = 0 Then Exit Function
        InsertAt = Left(merged, bodyEnd) & vbCrLf & blockText & vbCrLf & Mid(merged, bodyEnd + 1)
        Exit Function
    End If
    If anchorKind = "INTO" Or anchorKind = "AFTER" Then
        If Len(anchorId) = 0 Then Exit Function
        If Len(selfId) > 0 Then
            If UCase(anchorId) = UCase(selfId) Then Exit Function ' self-target
        End If
        closeM = InStr(1, merged, "<!--/ID:" & anchorId & "-->", 1)
        If closeM = 0 Then Exit Function ' unknown target
        If anchorKind = "INTO" Then
            InsertAt = Left(merged, closeM - 1) & blockText & vbCrLf & Mid(merged, closeM)
        Else
            insAt = closeM + Len("<!--/ID:" & anchorId & "-->")
            InsertAt = Left(merged, insAt - 1) & vbCrLf & blockText & Mid(merged, insAt)
        End If
        Exit Function
    End If
    If anchorKind = "LAST" Then
        insAt = InStr(1, merged, "</body>", 1)
        If insAt = 0 Then Exit Function
        InsertAt = Left(merged, insAt - 1) & blockText & vbCrLf & Mid(merged, insAt)
        Exit Function
    End If
End Function

' Insert one NEW block at its placement anchor: <!--AFTER:n--> (0 = first in
' <body>), <!--FIRST-->, <!--LAST-->, <!--INTO:n--> (last child inside n).
' No anchor (or unknown target) falls back to before </body>.
' Returns Null on failure.
Function InsertNewBlock(merged, content)
    Dim body, anchorKind, anchorId, insAt
    InsertNewBlock = Null
    If IsNull(merged) Or IsNull(content) Then Exit Function
    body = StripAnchorComments(content)
    If Len(Trim(body)) = 0 Then
        InsertNewBlock = merged
        Exit Function
    End If
    anchorKind = ""
    anchorId = ""
    If Not ParseAnchor(content, anchorKind, anchorId) Then anchorKind = "LAST"
    insAt = InsertAt(merged, body, anchorKind, anchorId, "")
    If IsNull(insAt) Then
        InsertNewBlock = InsertAt(merged, body, "LAST", "", "")
    Else
        InsertNewBlock = insAt
    End If
End Function

' Relocate an existing block: remove it with its markers and re-insert the
' same block (with new inner content) at the anchor. Self/unknown anchors
' and anchor-less content return Null so the caller keeps a plain replace
' (content is never lost on placement).
Function MoveBlock(merged, idStr, newInner)
    Dim openM, openEnd, closeM, closeEnd, anchorKind, anchorId, blockText, moved
    MoveBlock = Null
    If IsNull(merged) Or IsNull(newInner) Then Exit Function
    openM = InStr(1, merged, "<!--ID:" & idStr & "-->", 1)
    closeM = InStr(1, merged, "<!--/ID:" & idStr & "-->", 1)
    If openM = 0 Or closeM = 0 Then Exit Function
    If closeM < openM Then Exit Function
    anchorKind = ""
    anchorId = ""
    If Not ParseAnchor(newInner, anchorKind, anchorId) Then Exit Function
    newInner = StripAnchorComments(newInner)
    openEnd = openM + Len("<!--ID:" & idStr & "-->")
    closeEnd = closeM + Len("<!--/ID:" & idStr & "-->")
    blockText = "<!--ID:" & idStr & "-->" & newInner & "<!--/ID:" & idStr & "-->"
    moved = Left(merged, openM - 1) & Mid(merged, closeEnd)
    moved = InsertAt(moved, blockText, anchorKind, anchorId, idStr)
    If IsNull(moved) Then Exit Function
    MoveBlock = moved
End Function

' True when every consent anchor of the base still exists after the merge
Function MergeKeepsAnchors(baseHtml, mergedHtml)
    Dim anchors, i, a
    MergeKeepsAnchors = True
    anchors = Array("id=""cookieModal""", "id=""privacyModal""", "id=""cookieAccept""")
    For i = 0 To UBound(anchors)
        a = anchors(i)
        If InStr(1, baseHtml, a, 1) > 0 Then
            If InStr(1, mergedHtml, a, 1) = 0 Then
                MergeKeepsAnchors = False
                Exit Function
            End If
        End If
    Next
End Function

' Comma-separated list of block IDs in a tagged document (max maxN)
Function ListBlockIds(html, maxN)
    Dim out, seen, p, q, idStr, n
    ListBlockIds = ""
    If IsNull(html) Then Exit Function
    out = ""
    n = 0
    Set seen = Server.CreateObject("Scripting.Dictionary")
    p = 1
    Do While p <= Len(html)
        p = InStr(p, html, "<!--ID:", 1)
        If p = 0 Then Exit Do
        q = InStr(p, html, "-->", 0)
        If q = 0 Then Exit Do
        idStr = Trim(Mid(html, p + 7, q - p - 7))
        If Len(idStr) > 0 And UCase(idStr) <> "NEW" Then
            If Not seen.Exists(idStr) Then
                seen.Add idStr, True
                If Len(out) > 0 Then out = out & ", "
                out = out & idStr
                n = n + 1
                If n >= maxN Then Exit Do
            End If
        End If
        p = q + 3
    Loop
    ListBlockIds = out
End Function

' Comma-separated echoed IDs missing from the base ("" when all known)
Function FindUnknownIds(entries, baseHtml)
    Dim out, i, idStr
    FindUnknownIds = ""
    If Not IsArray(entries) Then Exit Function
    out = ""
    For i = 0 To UBound(entries)
        idStr = Trim(entries(i)("id"))
        If UCase(idStr) <> "NEW" And Len(idStr) > 0 Then
            If InStr(1, baseHtml, "<!--ID:" & idStr & "-->", 1) = 0 Then
                If Len(out) > 0 Then out = out & ", "
                out = out & idStr
            End If
        End If
    Next
    FindUnknownIds = out
End Function

' Strip all <...> tags (crude preview helper for the block inventory)
Function StripTags(s)
    Dim out, p, q
    If IsNull(s) Then
        StripTags = ""
        Exit Function
    End If
    out = s
    Do While True
        p = InStr(1, out, "<", 0)
        If p = 0 Then Exit Do
        q = InStr(p, out, ">", 0)
        If q = 0 Then Exit Do
        out = Left(out, p - 1) & " " & Mid(out, q + 1)
    Loop
    StripTags = Trim(out)
End Function

' Collapse whitespace for previews
Function CollapseSpaces(s)
    Dim out
    out = Replace(Replace(Replace(s, vbCrLf, " "), vbCr, " "), vbLf, " ")
    Do While InStr(1, out, "  ", 0) > 0
        out = Replace(out, "  ", " ")
    Loop
    CollapseSpaces = Trim(out)
End Function

' One line per block: ID + tag/content preview, so the AI picks the RIGHT id
Function BuildBlockInventory(html)
    Dim out, p, q, idStr, closeM, inner, txt, n
    BuildBlockInventory = ""
    If IsNull(html) Then Exit Function
    out = ""
    n = 0
    p = 1
    Do While p <= Len(html)
        p = InStr(p, html, "<!--ID:", 1)
        If p = 0 Then Exit Do
        q = InStr(p, html, "-->", 0)
        If q = 0 Then Exit Do
        idStr = Trim(Mid(html, p + 7, q - p - 7))
        If Len(idStr) > 0 And UCase(idStr) <> "NEW" Then
            closeM = InStr(q + 3, html, "<!--/ID:" & idStr & "-->", 1)
            If closeM > 0 Then
                inner = Mid(html, q + 3, closeM - q - 3)
                inner = StripIdMarkers(inner)
                txt = CollapseSpaces(StripTags(inner))
                If Len(txt) > 70 Then txt = Left(txt, 70) & "..."
                out = out & "ID " & idStr & ": " & txt & vbCrLf
                n = n + 1
                If n >= 40 Then Exit Do
            End If
        End If
        p = q + 3
    Loop
    BuildBlockInventory = out
End Function

' Repair hint appended to retries (attempts 2-3) based on the previous failure
Function BuildRepairHint(failKind, failDetail, existingIndex)
    Dim hint, validIds
    BuildRepairHint = ""
    hint = ""
    validIds = ListBlockIds(existingIndex, 60)
    If failKind = "truncated" Then
        hint = "REPAIR INSTRUCTION (your previous answer was INCOMPLETE - it stopped mid-page): do NOT return the complete page again. Return ONLY the blocks changed by the USER REQUEST, each as its own fenced ```html block starting with the echoed <!--ID:n--> marker and ending with <!--/ID:n-->. Keep the answer short."
    ElseIf failKind = "badids" Then
        hint = "REPAIR INSTRUCTION (your previous answer used block IDs that DO NOT EXIST: " & failDetail & ". Valid IDs are: " & validIds & "). Echo them EXACTLY as received and never invent numbers. For brand-new content use <!--ID:NEW-->...<!--/ID:NEW-->. Return ONLY the changed blocks as fenced ```html blocks."
    ElseIf failKind = "noblocks" Then
        hint = "REPAIR INSTRUCTION (your previous answer contained NO usable blocks): return one fenced ```html block per changed block, each starting with the echoed <!--ID:n--> marker and ending with <!--/ID:n-->. Valid IDs: " & validIds & ". Return ONLY the changed blocks."
    ElseIf failKind = "anchors" Then
        hint = "REPAIR INSTRUCTION (your blocks removed required page elements such as the cookie/privacy consent modals): only change what the USER REQUEST asks and leave all other blocks untouched. Return ONLY the changed blocks as fenced ```html blocks with echoed IDs."
    ElseIf failKind = "merged_broken" Then
        hint = "REPAIR INSTRUCTION (your merged blocks broke the page structure): return smaller valid HTML fragments only, each wrapped in its echoed <!--ID:n-->...<!--/ID:n--> markers. Valid IDs: " & validIds & "."
    End If
    BuildRepairHint = hint
End Function

Function EscapeJsonString(s)
    Dim result, i, ch, code
    result = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "\" Then
            result = result & "\\"
        ElseIf ch = """" Then
            result = result & "\"""
        ElseIf ch = vbCr Then
            result = result & "\r"
        ElseIf ch = vbLf Then
            result = result & "\n"
        ElseIf ch = vbTab Then
            result = result & "\t"
        Else
            ' AscW: full Unicode code point. Only strip real control chars (0-31).
            ' Negative values = high Unicode (e.g. emoji surrogates) - keep them.
            code = AscW(ch)
            If code >= 0 And code < 32 Then
                ' skip control character
            Else
                result = result & ch
            End If
        End If
    Next
    EscapeJsonString = result
End Function

' =================== ADMIN USER MANAGEMENT ===================

Function HandleListImages()
    RequireAuth()
    
    Dim projectId, folderName, projectDir, fso, imgDir
    projectId = Request.QueryString("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    projectDir = GetProjectDir(folderName)
    imgDir = projectDir & "img\"
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    ' URL prefix for displaying thumbnails in the app (random folder, no username)
    Dim prefix
    prefix = GetProjectUrlPrefix(folderName) & "img/"
    
    Dim images, idx, file, ext
    Set images = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    
    If fso.FolderExists(imgDir) Then
        Dim folder
        Set folder = fso.GetFolder(imgDir)
        For Each file In folder.Files
            ext = LCase(fso.GetExtensionName(file.Name))
            If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "gif" Or ext = "webp" Or ext = "bmp" Or ext = "avif" Then
                Dim imgData
                Set imgData = Server.CreateObject("Scripting.Dictionary")
                imgData.Add "filename", file.Name
                imgData.Add "url", prefix & file.Name
                imgData.Add "size", file.Size
                images.Add idx, imgData
                idx = idx + 1
            End If
        Next
    End If
    
    Response.Write JsonOk(images)
End Function

Function HandleDeleteImage()
    RequireAuth()
    
    Dim body, projectId, filename, folderName, projectDir, fso, imgPath
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    projectId = body("project_id")
    filename = body("filename")
    
    If Len(projectId) = 0 Or Len(filename) = 0 Then
        Response.Write JsonError("project_id and filename required", 400)
        Exit Function
    End If
    
    If ContainsPathTraversal(filename) Then
        Response.Write JsonError("Invalid filename", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    projectDir = GetProjectDir(folderName)
    imgPath = projectDir & "img\" & filename
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(imgPath) Then
        fso.DeleteFile imgPath, True
        Response.Write JsonOk("Image deleted")
    Else
        Response.Write JsonError("Image not found", 404)
    End If
End Function

' Look up a project owned by the current user; returns its name or Null
Function GetOwnedProjectName(projectId)
    Dim conn, rs
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT project_name FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id"))
    If rs.EOF Then
        rs.Close
        conn.Close
        GetOwnedProjectName = Null
        Exit Function
    End If
    GetOwnedProjectName = rs("project_name").Value
    rs.Close
    conn.Close
End Function

' Look up a project owned by the current user; returns its random folder name or Null
Function GetOwnedProjectFolder(projectId)
    Dim conn, rs
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT folder FROM projects WHERE id=" & CLng(projectId) & " AND user_id=" & Session("user_id"))
    If rs.EOF Then
        rs.Close
        conn.Close
        GetOwnedProjectFolder = Null
        Exit Function
    End If
    GetOwnedProjectFolder = rs("folder").Value
    rs.Close
    conn.Close
End Function

Function HandleListBackups()
    RequireAuth()
    
    Dim projectId, folderName, projectDir, backupsDir, fso
    projectId = Request.QueryString("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    projectDir = GetProjectDir(folderName)
    backupsDir = projectDir & "backups\"
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    
    Dim backups, idx, folder, sub_
    Set backups = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    
    If fso.FolderExists(backupsDir) Then
        Set folder = fso.GetFolder(backupsDir)
        For Each sub_ In folder.SubFolders
            Dim bData, fCount, bf
            Set bData = Server.CreateObject("Scripting.Dictionary")
            bData.Add "name", sub_.Name
            fCount = 0
            For Each bf In sub_.Files
                fCount = fCount + 1
            Next
            bData.Add "files", fCount
            backups.Add idx, bData
            idx = idx + 1
        Next
    End If
    
    Response.Write JsonOk(backups)
End Function

Function HandleRestoreBackup()
    RequireAuth()
    
    Dim body, projectId, backupName, folderName, projectDir, backupDir, fso
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    projectId = "" & body("project_id")
    backupName = "" & body("backup")
    
    If Len(projectId) = 0 Or Len(backupName) = 0 Then
        Response.Write JsonError("project_id and backup required", 400)
        Exit Function
    End If
    
    If ContainsPathTraversal(backupName) Then
        Response.Write JsonError("Invalid backup name", 400)
        Exit Function
    End If
    If InStr(1, backupName, "\", 1) > 0 Or InStr(1, backupName, "/", 1) > 0 Then
        Response.Write JsonError("Invalid backup name", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    projectDir = GetProjectDir(folderName)
    backupDir = projectDir & "backups\" & backupName & "\"
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(backupDir) Then
        Response.Write JsonError("Backup not found", 404)
        Exit Function
    End If
    
    ' First back up the CURRENT version, so a restore is also undoable
    Dim safeDir, ts
    ts = MakeTimestamp()
    safeDir = projectDir & "backups\" & ts & "\"
    EnsureDir safeDir
    If fso.FileExists(projectDir & "index.html") Then
        fso.CopyFile projectDir & "index.html", safeDir & "index.html"
    End If
    If fso.FileExists(projectDir & "style.css") Then
        fso.CopyFile projectDir & "style.css", safeDir & "style.css"
    End If
    If fso.FileExists(projectDir & "script.js") Then
        fso.CopyFile projectDir & "script.js", safeDir & "script.js"
    End If
    
    ' Restore files from the chosen backup
    Dim restored, bf2, ext2
    restored = 0
    Dim bFolder
    Set bFolder = fso.GetFolder(backupDir)
    For Each bf2 In bFolder.Files
        ext2 = LCase(fso.GetExtensionName(bf2.Name))
        If ext2 = "html" Or ext2 = "css" Or ext2 = "js" Then
            fso.CopyFile bf2.Path, projectDir & bf2.Name, True
            restored = restored + 1
        End If
    Next
    
    ' Update project timestamp
    Dim conn2
    Set conn2 = OpenDB()
    conn2.Execute "UPDATE projects SET updated_at=datetime('now') WHERE id=" & CLng(projectId)
    conn2.Close
    
    Dim result
    Set result = Server.CreateObject("Scripting.Dictionary")
    result.Add "message", "Restored " & restored & " file(s) from backup " & backupName
    result.Add "safety_backup", ts
    Response.Write JsonOk(result)
End Function

' Download the project (site files + images) as a single autonome.zip
Function HandleDownloadProject()
    RequireAuth()
    
    Dim projectId, folderName, projectDir, fso, tmpBase, stageDir, zipPath
    projectId = Request.QueryString("project_id")
    If Len(projectId) = 0 Then
        Response.Write JsonError("project_id required", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    projectDir = GetProjectDir(folderName)
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(projectDir) Then
        Response.Write JsonError("Project folder not found", 404)
        Exit Function
    End If
    
    ' Stage the deliverable files (no backups, no temp files) in autonome\ and zip that
    tmpBase = projectDir & "_ziptmp\"
    If fso.FolderExists(tmpBase) Then DeleteFolderRecursive fso, tmpBase
    stageDir = tmpBase & "autonome\"
    EnsureDir stageDir
    
    Dim names, i
    names = Array("index.html", "style.css", "script.js", "favicon.ico")
    For i = 0 To UBound(names)
        If fso.FileExists(projectDir & names(i)) Then
            fso.CopyFile projectDir & names(i), stageDir & names(i), True
        End If
    Next
    
    If fso.FolderExists(projectDir & "img") Then
        EnsureDir stageDir & "img"
        Dim imgFolder, f
        Set imgFolder = fso.GetFolder(projectDir & "img")
        For Each f In imgFolder.Files
            fso.CopyFile f.Path, stageDir & "img\" & f.Name, True
        Next
    End If
    
    zipPath = tmpBase & "autonome.zip"
    ASPPY.Zip.Zip Left(stageDir, Len(stageDir) - 1), zipPath
    
    ' Remove the staging folder; the zip itself is deleted after sending
    DeleteFolderRecursive fso, Left(stageDir, Len(stageDir) - 1)
    
    Response.BinaryFile zipPath, False, True
End Function

' Serve a backed-up index.html so it can be previewed in an iframe/modal.
' Injects a <base> tag pointing at the backup folder and rewrites relative
' image paths so they resolve to the live project's img folder.
Function HandlePreviewBackup()
    RequireAuth()
    
    Dim projectId, backupName, folderName, projectDir, backupDir, fso, content
    projectId = Request.QueryString("project_id")
    backupName = Request.QueryString("backup")
    
    If Len(projectId) = 0 Or Len(backupName) = 0 Then
        Response.Write JsonError("project_id and backup required", 400)
        Exit Function
    End If
    
    If ContainsPathTraversal(backupName) Then
        Response.Write JsonError("Invalid backup name", 400)
        Exit Function
    End If
    If InStr(1, backupName, "\", 1) > 0 Or InStr(1, backupName, "/", 1) > 0 Then
        Response.Write JsonError("Invalid backup name", 400)
        Exit Function
    End If
    
    folderName = GetOwnedProjectFolder(projectId)
    If IsNull(folderName) Then
        Response.Write JsonError("Project not found", 404)
        Exit Function
    End If
    
    projectDir = GetProjectDir(folderName)
    backupDir = projectDir & "backups\" & backupName & "\"
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(backupDir & "index.html") Then
        Response.ContentType = "text/html"
        Response.Write "<!DOCTYPE html><html><body style=""font-family:sans-serif;padding:2rem;color:#555""><p>This backup has no index.html to preview.</p></body></html>"
        Exit Function
    End If
    
    content = ReadFileContent(fso, backupDir & "index.html")
    If IsNull(content) Then content = ""
    
    ' Relative img/ references point at the live project's images (backups store no images)
    content = Replace(content, "src=""img/", "src=""../../img/")
    content = Replace(content, "src='img/", "src='../../img/")
    content = Replace(content, "url('img/", "url('../../img/")
    content = Replace(content, "url(""img/", "url(""../../img/")
    
    ' <base> makes css/js resolve inside the backup folder itself
    Dim baseHref, headPos
    baseHref = GetProjectUrlPrefix(folderName) & "backups/" & backupName & "/"
    headPos = InStr(1, content, "<head>", 1)
    If headPos > 0 Then
        content = Left(content, headPos + 5) & vbCrLf & "<base href=""" & baseHref & """>" & Mid(content, headPos + 6)
    Else
        content = "<base href=""" & baseHref & """>" & content
    End If
    
    Response.ContentType = "text/html"
    Response.Write content
End Function

Function HandleAdminGetUsers()
    RequireAdmin()
    
    Dim conn, rs, users, idx
    Set conn = OpenDB()
    Set rs = conn.Execute("SELECT id, username, email, role, created_at FROM users ORDER BY created_at ASC")
    
    Set users = Server.CreateObject("Scripting.Dictionary")
    idx = 0
    Do While Not rs.EOF
        Dim u
        Set u = Server.CreateObject("Scripting.Dictionary")
        u.Add "id", CLng(rs("id").Value)
        u.Add "username", rs("username").Value
        u.Add "email", rs("email").Value
        u.Add "role", rs("role").Value
        u.Add "created_at", rs("created_at").Value
        
        users.Add idx, u
        idx = idx + 1
        rs.MoveNext
    Loop
    rs.Close
    conn.Close
    
    Response.Write JsonOk(users)
End Function

Function HandleAdminUpdateUser()
    RequireAdmin()
    
    Dim body, userId, newRole, conn
    Set body = ParseJsonBody()
    If IsNull(body) Then
        Response.Write JsonError("Invalid JSON body", 400)
        Exit Function
    End If
    
    userId = body("user_id")
    newRole = body("role")
    
    If Len(userId) = 0 Or Len(newRole) = 0 Then
        Response.Write JsonError("user_id and role required", 400)
        Exit Function
    End If
    
    If newRole <> "admin" And newRole <> "user" Then
        Response.Write JsonError("Role must be 'admin' or 'user'", 400)
        Exit Function
    End If
    
    Set conn = OpenDB()
    conn.Execute "UPDATE users SET role=" & SqlEscape(newRole) & " WHERE id=" & CLng(userId)
    conn.Close
    
    Response.Write JsonOk("User role updated")
End Function

' =================== MAIN ROUTER ===================

Dim action
action = LCase(Request.QueryString("action"))

' First-run: create the database automatically when it does not exist yet
EnsureDatabase

If action = "setup" Then
    HandleSetup
ElseIf action = "login" Then
    HandleLogin
ElseIf action = "logout" Then
    HandleLogout
ElseIf action = "session" Then
    HandleSession
ElseIf action = "register" Then
    HandleRegister
ElseIf action = "change_password" Then
    HandleChangePassword
ElseIf action = "update_email" Then
    HandleUpdateEmail
ElseIf action = "forgot_password" Then
    HandleForgotPassword
ElseIf action = "reset_password" Then
    HandleResetPassword
ElseIf action = "get_config" Then
    HandleGetConfig
ElseIf action = "save_config" Then
    HandleSaveConfig
ElseIf action = "list_models" Then
    HandleListModels
ElseIf action = "get_projects" Then
    HandleGetProjects
ElseIf action = "create_project" Then
    HandleCreateProject
ElseIf action = "reset_project" Then
    HandleResetProject()
ElseIf action = "copy_project" Then
    HandleCopyProject()
ElseIf action = "delete_project" Then
    HandleDeleteProject
ElseIf action = "get_project_files" Then
    HandleGetProjectFiles
ElseIf action = "upload_image" Then
    HandleUploadImage
ElseIf action = "generate_website" Then
    HandleGenerateWebsite
ElseIf action = "generate_worker" Then
    HandleGenerateWorker
ElseIf action = "generate_status" Then
    HandleGenerateStatus
ElseIf action = "cancel_generation" Then
    HandleCancelGeneration
ElseIf action = "active_job" Then
    HandleActiveJob
ElseIf action = "download_project" Then
    HandleDownloadProject
ElseIf action = "preview_backup" Then
    HandlePreviewBackup
ElseIf action = "list_images" Then
    HandleListImages
ElseIf action = "delete_image" Then
    HandleDeleteImage
ElseIf action = "list_backups" Then
    HandleListBackups
ElseIf action = "restore_backup" Then
    HandleRestoreBackup
ElseIf action = "admin_get_users" Then
    HandleAdminGetUsers
ElseIf action = "admin_update_user" Then
    HandleAdminUpdateUser
ElseIf action = "rename_project" Then
    HandleRenameProject
ElseIf action = "update_notes" Then
    HandleUpdateNotes
ElseIf action = "get_notes" Then
    HandleGetNotes
ElseIf action = "get_chat_history" Then
    HandleGetChatHistory
ElseIf action = "keepalive" Then
    HandleKeepAlive
ElseIf action = "ping" Then
    HandleKeepAlive
ElseIf action = "version" Then
    Dim v
    Set v = Server.CreateObject("Scripting.Dictionary")
    v.Add "version", "2026-09-10-blocks-guard"
    v.Add "parse", "single-file index.html + 1920px resize + ID blocks"
    Response.Write JsonOk(v)
Else
    Response.Status = "404 Not Found"
    Response.Write JsonError("Unknown action: " & action, 404)
End If
%>
