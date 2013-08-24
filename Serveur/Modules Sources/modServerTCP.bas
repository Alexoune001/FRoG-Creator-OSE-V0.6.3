Attribute VB_Name = "modServerTCP"
Option Explicit
'Affichage de petits d�tails
Sub UpdateCaption()

Dim ReturnedHTML As String, sIp As Long, sStop As Long, affIP As String
    ReturnedHTML = SendAPIRequest("http://checkip.dyndns.org/")
    sIp = InStr(ReturnedHTML, "Address: ") + 9
    If InStr(sIp, ReturnedHTML, ",") > 0 Then
        sStop = InStr(sIp, ReturnedHTML, ",")
    Else
        sStop = InStr(sIp, ReturnedHTML, "<")
    End If
    If sStop > 0 Then
        affIP = Mid(ReturnedHTML, sIp, sStop - sIp)
    Else
        affIP = "0.0.0.0"
    End If
    
With frmServer
    .Caption = GAME_NAME & " - FRoG Serveur 0.6.3"
    .lblIP.Caption = "Adresse IP : " & affIP
    '.lblIP.Caption = frmServer.lblIP & " (Local : " & frmServer.Socket(0).LocalIP & ")"
    .lblPort.Caption = "Port : " & STR$(frmServer.Socket(0).LocalPort)
    .TPO.Caption = "Nombre de joueurs en ligne : " & TotalOnlinePlayers
    .lblgamename.Caption = "Jeu : " & GAME_NAME
    .lbliplocal.Caption = "Local : " & frmServer.Socket(0).LocalIP
    If Scripting = 1 Then
        .lblscript.Caption = "Etat : activ�"
        .lblscript.ForeColor = RGB(51, 204, 0)
    Else
        .lblscript.Caption = "Etat : d�sactiv�"
        .lblscript.ForeColor = RGB(255, 102, 102)
    End If
End With
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    On Error GoTo er:
    If frmServer.Socket(Index).State = sckConnected Then IsConnected = True Else IsConnected = False
    Exit Function
er:
    IsConnected = False
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If Index <= 0 Or Index > MAX_PLAYERS Then IsPlaying = False: Exit Function
    If IsConnected(Index) And Player(Index).InGame = True Then IsPlaying = True Else IsPlaying = False
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If IsConnected(Index) And Trim$(Player(Index).Login) <> vbNullString Then IsLoggedIn = True Else IsLoggedIn = False
End Function
'Multicompte
Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase$(Trim$(Player(i).Login)) = LCase$(Trim$(Login)) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next i
End Function
'IP multiples
Function IsMultiIPOnline(ByVal IP As String) As Boolean
Dim i As Long
Dim n As Long

    n = 0
    IsMultiIPOnline = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And Trim$(GetPlayerIP(i)) = Trim$(IP) And GetPlayerAccess(i) < 4 Then
            n = n + 1
            
            If (n > 5) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next i
End Function
'Banni
Function IsBanned(ByVal IP As String) As Boolean
Dim FileName As String, fIP As String, fName As String
Dim f As Long
On Error Resume Next
    IsBanned = False
    
    FileName = App.Path & "\banlist.txt"
    
    ' Verification de l'existance de banlist
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open FileName For Input As #f
    
    Do While Not EOF(f)
        Input #f, fIP
        Input #f, fName
    
        ' Est-il banni
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If
    Loop
    
    Close #f
End Function

Sub SendDataTo(ByVal Index As Long, ByVal data As String)
Dim i As Long, n As Long, startc As Long
On Error Resume Next

    If IsConnected(Index) Then frmServer.Socket(Index).SendData data: NewDoEvents
End Sub

Sub SendDataToAll(ByVal data As String)
Dim i As Long
On Error Resume Next

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal data As String)
Dim i As Long
On Error Resume Next

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> Index Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal data As String)
Dim i As Long
On Error Resume Next

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then If GetPlayerMap(i) = MapNum Then Call SendDataTo(i, data)
    Next i
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal data As String)
Dim i As Long
On Error Resume Next

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then If GetPlayerMap(i) = MapNum And i <> Index Then Call SendDataTo(i, data)
    Next i
End Sub
'Message global
Sub GlobalMsg(ByVal Msg As String, ByVal Color As Long)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub
'Message admin
Sub AdminMsg(ByVal Msg As String, ByVal Color As Long)
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then Call SendDataTo(i, Packet)
    Next i
End Sub
'Message priv�
Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Long)
Dim Packet As String

    If Not IsPlaying(Index) Then Exit Sub

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub
'Message de guilde
Sub GuildeMsg(ByVal Index As Long, ByVal Msg As String)
    Dim i As Long
    Dim s As String
    
    If Player(Index).Mute Then Exit Sub
       
    If GetPlayerGuild(Index) = vbNullString Then Call PlayerMsg(Index, "Tu n'es pas dans une guilde.", AlertColor): Exit Sub
    
    s = GetPlayerName(Index) & " (" & GetPlayerGuild(Index) & ") : " & Msg
    Call AddLog(s, PLAYER_LOG)
       
    For i = 1 To MAX_PLAYERS
        If GetPlayerGuild(Index) = GetPlayerGuild(i) Then Call PlayerMsg(i, s, CouleurDesGuilde)
    Next i
End Sub

Public Sub QueteMsg(ByVal Index As Long, ByVal Msg As String)
Dim Packet As String

If Mid(Msg, 1, 2) = "**" Then Msg = Mid(Msg, InStr(1, Msg, ":"))
Packet = "QMSG" & SEP_CHAR & Msg & END_CHAR

Call SendDataTo(Index, Packet)
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Long)
Dim Packet As String
Dim text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub
'Message d'alertes
Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & END_CHAR
    
    Call SendDataTo(Index, Packet)
    Call CloseSocket(Index)
    If Index > 0 And Index < MAX_PLAYERS Then
    If IsPlaying(Index) Then
    If IBErr Then Call IBMsg("Un joueur a re�u un message d'alerte. (Login : " & GetPlayerLogin(Index) & " perso : " & GetPlayerName(Index) & " Message : " & Msg & ").", BrightRed)
    End If
    End If
End Sub

Sub PlainMsg(ByVal Index As Long, ByVal Msg As String, ByVal Num As Long)
Dim Packet As String

    Packet = "PLAINMSG" & SEP_CHAR & Msg & SEP_CHAR & Num & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub
' En cas de hack
Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    On Error Resume Next
    If Index > 0 And Index < MAX_PLAYERS Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " a �t� d�connect� pour (" & Reason & ")", White)
        End If
    
        Call AlertMsg(Index, "Tu as perdu ta connexion avec " & GAME_NAME & "." & Reason)
        If IBErr Then Call IBMsg("Tentative de hack d�tect�, raison : " & Reason & " Login : " & GetPlayerLogin(Index) & " Personnage : " & GetPlayerName(Index) & ").", BrightRed)
    End If
    Exit Sub
End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
Dim i As Long
On Error Resume Next
    If (Index = 0) Then
        i = FindOpenPlayerSlot
        
        If i <> 0 Then
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    On Error Resume Next
    If Index <> 0 Then
        ' Tentative de connexion multiple ?
        If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(Index)) Then
                Call TextAdd(frmServer.txtText(0), "Connexion re�ue depuis l'IP " & GetPlayerIP(Index) & ".", True)
            Else
                Call AlertMsg(Index, "Tu as �t� banni de " & GAME_NAME & ".")
            End If
        Else
           ' Tentative d'IP multiple
            Call AlertMsg(Index, GAME_NAME & " n'autorise pas les IP multiples.")
        End If
    End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim Top As String * 3
Dim Start As Long

    If Index > 0 Then
        frmServer.Socket(Index).GetData Buffer, vbString, DataLength
        
        If Buffer = "top" Then
            Top = STR$(TotalOnlinePlayers)
            Call SendDataTo(Index, Top)
            Call CloseSocket(Index)
        End If
        
        Player(Index).Buffer = Player(Index).Buffer & Buffer
        
        Start = InStr(Player(Index).Buffer, END_CHAR)
        Do While Start > 0
            Packet = Mid$(Player(Index).Buffer, 1, Start - 1)
            Player(Index).Buffer = Mid$(Player(Index).Buffer, Start + 1, Len(Player(Index).Buffer))
            Player(Index).DataPackets = Player(Index).DataPackets + 1
            Start = InStr(Player(Index).Buffer, END_CHAR)
            If Len(Packet) > 0 Then
                If Not IsPlaying(Index) Then
                    ' Parse's Without Being Online
                    Call HandleLoginData(Index, Packet)
                Else
                    ' Parse's With Being Online And Playing
                    Call HandleData(Index, Packet)
                End If
            End If
        Loop
                
        ' Not useful
        ' Check if elapsed time has passed
        Player(Index).DataBytes = Player(Index).DataBytes + DataLength
        If GetTickCount >= Player(Index).DataTimer + 1000 Then
            Player(Index).DataTimer = GetTickCount
            Player(Index).DataBytes = 0
            Player(Index).DataPackets = 0
            Exit Sub
        End If
        
        ' Check for data flooding
        If Player(Index).DataBytes > 1000 And GetPlayerAccess(Index) <= 0 Then
            Call HackingAttempt(Index, "Data Flooding")
            Exit Sub
        End If
        
        ' Check for packet flooding
        If Player(Index).DataPackets > 25 And GetPlayerAccess(Index) <= 0 Then
            Call HackingAttempt(Index, "Packet Flooding")
            Exit Sub
        End If
    End If
End Sub

Sub HandleLoginData(ByVal Index As Long, ByVal data As String)
Dim Parse() As String
Dim Name As String
Dim Password As String
Dim Sex As Long
Dim Class As Long
Dim CharNum As Long
Dim i As Integer, n As Integer, f As Integer
    On Error GoTo er:
    
    Player(Index).sync = True
    ' Handle Data
    Parse = Split(data, SEP_CHAR)

    Select Case LCase$(Parse(0))
        Case "desync"
        Call CloseSocket(Index, True)
        Exit Sub
        Case "logination"
            If Not IsLoggedIn(Index) Then
                Name = Parse(1)
                Password = Parse(2)
                
                If Len(Name) < 3 Then
                Call PlainMsg(Index, "Votre pseudo est trop court, il doit contenir au minimum 3 caract�res.", 3)
                Exit Sub
                End If
                
                For i = 1 To Len(Name)
                    n = Asc(Mid$(Name, i, 1))
                    If (n <= 65 And n >= 90) Or (n <= 97 And n >= 122) Or (n = 95) Or (n = 32) Or (n <= 48 And n >= 57) Then
                        Call PlainMsg(Index, "Le nom est invalide, il ne doit pas contenir des caract�res sp�ciaux.", 3)
                        Exit Sub
                    End If
                Next i
        
                            
                If Not AccountExist(Name) Then Call PlainMsg(Index, "Aucun compte ne poss�de ce nom.", 3): Exit Sub
            
                If Not PasswordOK(Name, Password) Then Call PlainMsg(Index, "Le mot de passe est incorrect.", 3): Exit Sub
            
                If IsMultiAccounts(Name) Then Call PlainMsg(Index, "Le multi-compte est interdit.", 3): Exit Sub
                
                If frmServer.Closed.value = Checked Then Call PlainMsg(Index, "Le serveur va fermer dans quelques minutes, merci de revenir plus tard.", 3): Exit Sub
                    
                If Parse(6) <> SEC_CODE1 And Parse(7) = SEC_CODE2 And Parse(8) = SEC_CODE3 And Val(Parse(9)) = SEC_CODE4 Then
                    Call AlertMsg(Index, "Cl� de s�curit� incorrecte.")
                    Exit Sub
                End If
                            
                Dim Packs As String
                Packs = "MAXINFO" & SEP_CHAR
                Packs = Packs & GAME_NAME & SEP_CHAR
                Packs = Packs & MAX_PLAYERS & SEP_CHAR
                Packs = Packs & MAX_ITEMS & SEP_CHAR
                Packs = Packs & MAX_NPCS & SEP_CHAR
                Packs = Packs & MAX_SHOPS & SEP_CHAR
                Packs = Packs & MAX_SPELLS & SEP_CHAR
                Packs = Packs & MAX_MAPS & SEP_CHAR
                Packs = Packs & MAX_MAP_ITEMS & SEP_CHAR
                Packs = Packs & MAX_MAPX & SEP_CHAR
                Packs = Packs & MAX_MAPY & SEP_CHAR
                Packs = Packs & MAX_EMOTICONS & SEP_CHAR
                Packs = Packs & MAX_LEVEL & SEP_CHAR
                Packs = Packs & MAX_QUETES & SEP_CHAR
                Packs = Packs & MAX_INV & SEP_CHAR
                Packs = Packs & MAX_NPC_SPELLS & SEP_CHAR
                Packs = Packs & MAX_PETS & SEP_CHAR
                Packs = Packs & MAX_METIER & SEP_CHAR
                Packs = Packs & MAX_RECETTE & SEP_CHAR
                Packs = Packs & END_CHAR
                Call SendDataTo(Index, Packs)
        
                Call LoadPlayer(Index, Name)
                Call SendChars(Index)
        
                Call AddLog(GetPlayerLogin(Index) & " s'est connect� depuis " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & " s'est connect� depuis " & GetPlayerIP(Index) & ".", True)
                If IBJoueur Then Call IBMsg(GetPlayerLogin(Index) & " s'est connect� � " & GAME_NAME, IBCJoueur)
            End If
            Exit Sub
    
        Case "usagakarim"
                
                If Not FileExist("\accounts\" & Trim$(Player(Index).Login) & ".ini") Then
                Call HackingAttempt(Index, "Erreur : Vous avez tent� de surcharger le serveur")
                Exit Sub
                End If
                
                CharNum = Val(Parse(1))
    
                If CharNum < 1 Or CharNum > MAX_CHARS Then Call HackingAttempt(Index, "Num�ro de personnage invalide."): Exit Sub
            
                If CharExist(Index, CharNum) Then
                    Player(Index).CharNum = CharNum
                    
                    If frmServer.GMOnly.value = Checked And GetPlayerAccess(Index) <= 0 Then
                        Call PlainMsg(Index, "Le serveur est seulement accessible aux membres de l'�quipe pour le moment.", 5)
                        Exit Sub
                    End If
                                        
                    Call JoinGame(Index)
                
                    CharNum = Player(Index).CharNum
                
                    If Val(Scripting) = 1 Then
                        MyScript.ExecuteStatement "Scripts\Main.txt", "UseChar " & Index & "," & CharNum
                    End If
                    
                    Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " est en train de jouer � " & GAME_NAME & ".", PLAYER_LOG)
                    Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " est en train de jouer � " & GAME_NAME & ".", True)
                    If IBJoueur Then Call IBMsg(GetPlayerName(Index) & " vient de se connecter � " & GAME_NAME & ".", IBCJoueur)
                    Call UpdateCaption
                    If Not FindChar(GetPlayerName(Index)) Then
                        f = FreeFile
                        Open App.Path & "\accounts\charlist.txt" For Append As #f
                            Print #f, GetPlayerName(Index)
                        Close #f
                    End If
                Else
                    Call PlainMsg(Index, "Le personnage n'existe pas.", 5)
                End If
            Exit Sub
        Case "addachara"
                Name = Parse(1)
                Sex = Val(Parse(2))
                Class = Val(Parse(3))
                CharNum = Val(Parse(4))
                
                For i = 1 To Len(Name)
                    n = Asc(Mid$(Name, i, 1))
                    
                    If (n <= 65 And n >= 90) Or (n <= 97 And n >= 122) Or (n = 95) Or (n = 32) Or (n <= 48 And n >= 57) Then
                        Call PlainMsg(Index, "Nom invalide, il ne doit pas contenir des caract�res sp�ciaux.", 4)
                        Exit Sub
                    End If
                Next i
                
                If Not FileExist("\accounts\" & Trim$(Player(Index).Login) & ".ini") Then
                Call HackingAttempt(Index, "Erreur : Vous avez tent� de surcharger le serveur")
                Exit Sub
                End If
                
                
                If CharNum < 1 Or CharNum > MAX_CHARS Then Call HackingAttempt(Index, "Num�ros de personnage invalide."): Exit Sub
            
                If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then Call HackingAttempt(Index, "Le sexe est invalide."): Exit Sub
                
                If Class < 0 Or Class > Max_Classes Then Call HackingAttempt(Index, "La classe est invalide."): Exit Sub
    
                If CharExist(Index, CharNum) Then Call PlainMsg(Index, "Le personnage existe d�j�.", 4): Exit Sub
    
                If FindChar(Name) Then Call PlainMsg(Index, "Ce pseudo est d�j� utilis�.", 4): Exit Sub
                
                Call AddChar(Index, Name, Sex, Class, CharNum)
                Call SavePlayer(Index)
                If Val(Scripting) = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "NewChar " & Index & "," & CharNum
                End If
                Call AddLog("Le personnage " & Name & " a �t� ajout� au compte de " & GetPlayerLogin(Index) & ".", PLAYER_LOG)
                Call SendChars(Index)
                Call PlainMsg(Index, "Le personnage a �t� cr��.", 5)
                If IBJoueur Then Call IBMsg("Le personnage " & Name & " a �t� ajout� au compte de " & GetPlayerLogin(Index) & ".", IBCJoueur)
            Exit Sub
    
        Case "serverresults"
            Call SendDataTo(Index, "serverresults" & SEP_CHAR & Val(Parse(1)) & SEP_CHAR & TotalOnlinePlayers & SEP_CHAR & MAX_PLAYERS & END_CHAR)
            Exit Sub
            
        Case "gatglasses"
            Call SendNewCharClasses(Index)
            Exit Sub
            
        Case "newfaccountied"
            If Not IsLoggedIn(Index) Then
                Name = Parse(1)
                Password = Parse(2)
                        
                For i = 1 To Len(Name)
                    n = Asc(Mid$(Name, i, 1))
                    
                    If (n <= 65 And n >= 90) Or (n <= 97 And n >= 122) Or (n = 95) Or (n = 32) Or (n <= 48 And n >= 57) Then
                        Call PlainMsg(Index, "Le nom est invalide, il ne doit pas contenir des caract�res sp�ciaux.", 1)
                        Exit Sub
                    End If
                Next i
                If Not AccountExist(Name) Then
                    Call AddAccount(Index, Name, Password)
                    Call ClearPlayer(Index)
                    Call TextAdd(frmServer.txtText(0), "Compte " & Name & " a �t� cr��.", True)
                    Call AddLog("Compte " & Name & " a �t� cr��.", PLAYER_LOG)
                    Call PlainMsg(Index, "Votre compte a �t� cr�e.", 1)
                    If IBJoueur Then Call IBMsg("Un joueur a cr�e un compte nomm� " & Name, IBCJoueur)
                Else
                    Call PlainMsg(Index, "Un compte du m�me nom existe d�j�.", 1)
                End If
            End If
            Exit Sub
   
        Case "delimaccounted"
            If Not IsLoggedIn(Index) Then
                Name = Parse(1)
                Password = Parse(2)
                            
                If Not AccountExist(Name) Then Call PlainMsg(Index, "Ce compte n'existe pas.", 2): Exit Sub
                
                If Not PasswordOK(Name, Password) Then Call PlainMsg(Index, "Le mot de passe est incorrect.", 2): Exit Sub
                            
                Call LoadPlayer(Index, Name)
                
                For i = 1 To MAX_CHARS
                    If Trim$(Player(Index).Char(i).Name) <> vbNullString Then Call DeleteName(Player(Index).Char(i).Name)
                Next i
                Call ClearPlayer(Index)
                
                Call Kill(App.Path & "\accounts\" & Trim$(Name) & ".ini")
                Call AddLog("Account " & Trim$(Name) & " a �t� effac�.", PLAYER_LOG)
                Call PlainMsg(Index, "Votre compte a �t� effac�.", 2)
                If IBJoueur Then Call IBMsg("Un joueur a �ffac� son compte nomm� " & Name, IBCJoueur)
            End If
            Exit Sub
    
        Case "picvalue"
            Packs = "PICVALUE" & SEP_CHAR & PIC_PL & SEP_CHAR & PIC_NPC1 & SEP_CHAR & PIC_NPC2 & SEP_CHAR & AccModo & SEP_CHAR & AccMapeur & SEP_CHAR & AccDevelopeur & SEP_CHAR & AccAdmin & END_CHAR
            Call SendDataTo(Index, Packs)
            Exit Sub

        Case "delimbocharu"
                CharNum = Val(Parse(1))
    
                If CharNum < 1 Or CharNum > MAX_CHARS Then Call HackingAttempt(Index, "Num�ro de personnage invalide."): Exit Sub
                
                Call DelChar(Index, CharNum)
                Call AddLog("Un personnage a �t� suprim� du compte " & GetPlayerLogin(Index) & ".", PLAYER_LOG)
                Call SendChars(Index)
                Call PlainMsg(Index, "Le personnage a �t� effac�.", 5)
                If IBJoueur Then Call IBMsg("Le personnage " & CharNum & " a �t� suprim� du compte de " & GetPlayerLogin(Index) & ".", IBCJoueur)
            Exit Sub
         
         Case "sync"
                Player(Index).sync = True
         Exit Sub
    End Select
    
    Call HackingAttempt(Index, "Erreur dans l'envoi d'un packet (" & Parse(0) & ")")
    Exit Sub
    
er:
Call AddLog("Le " & Date & " � " & time & " : Erreur dans la r�ception du serveur. D�tails - Num :" & Err.Number & " Description : " & Err.Description & " Source : " & Err.Source & "...", "logs\Err.txt")
On Error Resume Next
If IBErr Then Call IBMsg("Un erreur s'est produite dans la r�ception du serveur.", BrightRed)
If Not IsPlaying(Index) Then Call PlainMsg(Index, "Erreur d'envoi, merci de relancer le jeu.", 3)
End Sub

Sub HandleData(ByVal Index As Long, ByVal data As String)
Dim Parse() As String
Dim Packs As String
Dim Name As String
Dim CharNum As Integer
Dim Msg As String
Dim IPMask As String
Dim BanSlot As Long
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Amount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Movement As Long
Dim i As Long, n As Long, x As Long, y As Long, f As Long
Dim MapNum As Long
Dim s As String
Dim tMapStart As Long, tMapEnd As Long
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long
Dim z As Long
Dim Packet As String
Dim BX As Long, BY As Long
Dim SlotC As Long
Dim Islot As Long
Dim Cnum As Long
Dim Cval As Long
Dim Cdur As Long
Dim SlotI As Long
Dim Cslot As Long
Dim INum As Long
Dim IVal As Long
Dim IDur As Long
Dim PChar As String * 1

On Error GoTo er:
Player(Index).sync = True
    ' Handle Data
    Parse = Split(data, SEP_CHAR)
    Parse(0) = LCase$(Parse(0))
    PChar = Left$(Parse(0), 1)
        
    If Not IsPlaying(Index) Then Exit Sub
    If Not IsConnected(Index) Then Exit Sub

    Select Case PChar
        Case "c"
            Select Case Parse(0)
                Case "changechar"
                    Call SavePlayerOptim(Index)
                    Player(Index).InGame = False
                    Exit Sub
                    
                Case "cast"
                    Call CastSpell(Index, Val(Parse(1)))
                    Exit Sub
                Case "crafter"
                    'Call craft(Index, Val(Parse(1)))
                    Call craftv2(Index, Val(Parse(1)))
                    Exit Sub
                Case "coffreitem"
                    Dim cof As Long
                    
                    If IsPlaying(Index) = False Then Call HackingAttempt(Index, "Vous n'�tes pas en train de jouer (packet coffreitem)."): Exit Sub
                            
                    Packet = "DATACOFR"
                    
                    For cof = 1 To 30
                        Packet = Packet & SEP_CHAR & GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemnum" & cof)
                        Packet = Packet & SEP_CHAR & GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemval" & cof)
                        Packet = Packet & SEP_CHAR & GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemdur" & cof)
                    Next cof
                    
                    Packet = Packet & END_CHAR
                    
                    Call SendDataTo(Index, Packet)
                    Exit Sub
                
                Case "checkcommands"
                    s = Parse(1)
                    If LCase$(Mid$(s, 1, 5)) = "/hdvs" Then HdvCmd Index, s: Exit Sub
                    If Scripting = 1 Then
                        PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & Index, Trim$(s)
                        MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & Index
                    Else
                        Call PlayerMsg(Index, "Ce n'est pas une commande valide.", 12)
                    End If
                    Exit Sub
                    
                                Case "checkarrows"
                    n = Arrows(Val(Parse(1))).Pic
                    Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & n & END_CHAR)
                    Exit Sub
                    
                Case "checkemoticons"
                    n = Emoticons(Val(Parse(1))).Pic
                    Call SendDataToMap(GetPlayerMap(Index), "checkemoticons" & SEP_CHAR & Index & SEP_CHAR & n & END_CHAR)
                    Exit Sub
                    
                Case "chgclasses"
                    If GetPlayerAccess(Index) > ADMIN_MAPPER And CClasses Then Call LoadClasses Else Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Exit Sub
            End Select
        Case "p"
            Select Case Parse(0)

                ' d�placement du personnage
                Case "playermove"
                    If Player(Index).GettingMap = YES Then Exit Sub
                    
                    Dir = Val(Parse(1))
                    Movement = Val(Parse(2))
                    
                    ' Prevent hacking
                    If Dir < DIR_DOWN Or Dir > DIR_UP Then Call HackingAttempt(Index, "Direction invalide."): Exit Sub
                            
                    ' Prevent hacking
                    If Movement < 1 Or Movement > 2 Then Call HackingAttempt(Index, "Mouvement invalide."): Exit Sub
                    
                    ' Prevent player from moving if they have casted a spell
                    If Player(Index).CastedSpell = YES Then
                        ' Check if they have already casted a spell, and if so we can't let them move
                        If GetTickCount > Player(Index).AttackTimer + 1000 Then Player(Index).CastedSpell = NO Else Call SendPlayerXY(Index): Exit Sub
                    End If
                    
                    Call PlayerMove(Index, Dir, Movement)
                    Exit Sub
                
                ' :: Metier ::
                Case "playermetier"
                    If Player(Index).Char(Player(Index).CharNum).metier > 0 Then
                        Call SendPlayerMetier(Index)
                        Call SendDataTo(Index, "METIER" & END_CHAR)
                    Else
                        Call PlayerMsg(Index, "Pas de m�tier", White)
                    End If
                    Exit Sub
                    
                Case "playermetieroublie"
                    If Player(Index).Char(Player(Index).CharNum).metier > 0 Then
                        Player(Index).Char(Player(Index).CharNum).metier = 0
                        Player(Index).Char(Player(Index).CharNum).MetierLvl = 1
                        Player(Index).Char(Player(Index).CharNum).MetierExp = 0
                        Call PlayerMsg(Index, "Le m�tier a �t� oubli�.", White)
                        Call SendPlayerMetier(Index)
                    Else
                        Call PlayerMsg(Index, "Pas de m�tier", White)
                    End If
                    Exit Sub
                    
                ' :: Moving character packet ::
                Case "playerdir"
                    If Player(Index).GettingMap = YES Then Exit Sub
                    Dir = Val(Parse(1))
                    ' Prevent hacking
                    If Dir < DIR_DOWN Or Dir > DIR_UP Then Call HackingAttempt(Index, "Direction Invalide"): Exit Sub
                    
                    Call SetPlayerDir(Index, Dir)
                    Call SendDataToMapBut(Index, GetPlayerMap(Index), "PLAYERDIR" & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
                    Exit Sub
                    
                Case "playermsg"
                    MsgTo = FindPlayer(Parse(1))
                    Msg = Parse(2)
                    ' Prevent hacking
                    If MMsg(Msg) Then Call HackingAttempt(Index, "Caract�re incorrect dans ses paroles (joueurs)"): Exit Sub
            
                    If frmServer.chkP.value = Unchecked Then If GetPlayerAccess(Index) <= 0 Then Call PlayerMsg(Index, "Les messages priv�s ont �t� d�sactiv�s sur le serveur.", BrightRed): Exit Sub
            
                    If FindPlayer(Parse(1)) = 0 Then Call PlayerMsg(Index, "Le joueur est hors-ligne.", White): Exit Sub
                            
                    ' Check if they are trying to talk to themselves
                    If MsgTo <> Index Then
                        If MsgTo > 0 And MsgTo < MAX_PLAYERS Then
                            Call AddLog(GetPlayerName(Index) & " dit � " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                            Call PlayerMsg(MsgTo, GetPlayerName(Index) & " vous dit : '" & Msg & "'", TellColor)
                            Call PlayerMsg(Index, "Vous dite a " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
                        Else
                            Call PlayerMsg(Index, "Le joueur n'est pas en ligne.", White)
                        End If
                    Else
                        Call AddLog("Carte #" & GetPlayerMap(Index) & " : " & GetPlayerName(Index) & " se parle � lui m�me...", PLAYER_LOG)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " murmure quelque chose � lui m�me.", Green)
                    End If
                    TextAdd frmServer.txtText(4), "� : " & GetPlayerName(MsgTo) & " De : " & GetPlayerName(Index) & " : " & Msg, True
                    Exit Sub
                    
                Case "partychat"
                    If Player(Index).InParty > 0 Then
                        For i = 1 To Party.MemberCount(Player(Index).InParty)
                            Call PlayerMsg(Party.PlayerIndex(Player(Index).InParty, i), Parse(1), Blue)
                        Next i
                    Else
                        PlayerMsg Index, "Vous n'�tes pas dans un groupe.", BrightRed
                    End If
                    Exit Sub
                    
                Case "pptrade"
                    n = FindPlayer(Parse(1))
                    ' Check if player is online
                    If n < 1 Then Call PlayerMsg(Index, "Le joueur n'est pas en ligne.", White): Exit Sub
                    ' Prevent trading with self
                    If n = Index Then Exit Sub
                    ' Check if the player is in another trade
                    If Player(Index).InTrade = 1 Then Call PlayerMsg(Index, "Tu �changes d�j� avec quelqu'un.", Pink): Exit Sub
                    ' Check where both players are
                    Dim CanTrade As Boolean
                    CanTrade = False
                    
                    If GetPlayerX(Index) = GetPlayerX(n) And GetPlayerY(Index) + 1 = GetPlayerY(n) Then CanTrade = True
                    If GetPlayerX(Index) = GetPlayerX(n) And GetPlayerY(Index) - 1 = GetPlayerY(n) Then CanTrade = True
                    If GetPlayerX(Index) + 1 = GetPlayerX(n) And GetPlayerY(Index) = GetPlayerY(n) Then CanTrade = True
                    If GetPlayerX(Index) - 1 = GetPlayerX(n) And GetPlayerY(Index) = GetPlayerY(n) Then CanTrade = True
                        
                    If CanTrade = True Then
                        ' Check to see if player is already in a trade
                        If Player(n).InTrade = 1 Then Call PlayerMsg(Index, "Le joueur echange d�j� avec quelq'un.", Pink): Exit Sub
                        Call PlayerMsg(Index, "Requ�te d'�change envoy� � " & GetPlayerName(n) & ".", Pink)
                        Call PlayerMsg(n, GetPlayerName(Index) & " veut faire un �change avec vous.  Entrez /accept pour accepter, ou /refu pour refuser.", Pink)
                        Player(n).TradePlayer = Index
                        Player(Index).TradePlayer = n
                    Else
                        Call PlayerMsg(Index, "Vous avez besoin d'�tre devant le joueur pour �changer.", Pink)
                        Call PlayerMsg(n, "Le joueur doit �tre devant vous pour �changer.", Pink)
                    End If
                    Exit Sub
                    
                Case "party"
                    n = FindPlayer(Parse(1))
                    ' Prevent partying with self
                    If n = Index Then Exit Sub
                    ' Check for a full party and if so drop it
                    Dim g As Integer
                    i = Player(Index).InParty
                    If i > 0 Then
                        g = Party.MemberCount(i)
                        If g = MAX_PARTY_MEMBERS Then
                            Call PlayerMsg(Index, "Le groupe est complet.", Pink)
                            Exit Sub
                        End If
                    End If
                    If n > 0 Then
                        ' Verification : le joueur est il admin ? Si vous voulez que les admins puissent faire des groupes, effacez les DEUX lignes suivantes
                        If GetPlayerAccess(Index) > ADMIN_MONITER Then Call PlayerMsg(Index, "Vous ne pouvez joindre un groupe, vous �tes un admin.", BrightBlue): Exit Sub
                        If GetPlayerAccess(n) > ADMIN_MONITER Then Call PlayerMsg(Index, "Un admin ne peut rejoindre un groupe.", BrightBlue): Exit Sub
                        
                        ' V�rification : le joueur est d�ja dans un groupe
                        If Player(n).InParty = 0 Then
                            Call PlayerMsg(Index, GetPlayerName(n) & " a �t� invit� � joindre votre groupe.", Pink)
                            Call PlayerMsg(n, GetPlayerName(Index) & " t'invite � joindre son groupe. /join pour joindre, ou /leave pour refuser.", Pink)
                            Player(n).InvitedBy = Index
                        Else
                            Call PlayerMsg(Index, "Le joueur est d�j� dans le groupe.", Pink)
                        End If
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
                    
                Case "playerchat"
                    n = FindPlayer(Parse(1))
                    If n < 1 Then Call PlayerMsg(Index, "Personnage hors-ligne.", White): Exit Sub
                    If n = Index Then Exit Sub
                    If Player(Index).InChat = 1 Then Call PlayerMsg(Index, "Vous discutez d�j� avec quelqu'un d'autre.", Pink): Exit Sub
                    If Player(n).InChat = 1 Then Call PlayerMsg(Index, "Le joueur est d�j� en discution avec quelqu'un d'autre.", Pink): Exit Sub
                            
                    Call PlayerMsg(Index, "Requ�te de discutions envoy� a " & GetPlayerName(n) & ".", Pink)
                    Call PlayerMsg(n, GetPlayerName(Index) & " veut discuter avec vous.  taper /chat pour accepter, ou /chatrefu pour refuser.", Pink)
                
                    Player(n).ChatPlayer = Index
                    Player(Index).ChatPlayer = n
                    Exit Sub
                    
                Case "prompt"
                    If Scripting = 1 Then MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & Index & "," & Val(Parse(1)) & "," & Val(Parse(2))
                    Exit Sub
                    
                Case "playerinforequest"
                    Name = Parse(1)
                    
                    i = FindPlayer(Name)
                    If i > 0 Then
                        Call PlayerMsg(Index, "Compte : " & Trim$(Player(i).Login) & ", Nom : " & GetPlayerName(i), BrightGreen)
                        If GetPlayerAccess(Index) > ADMIN_MONITER Then
                            Call PlayerMsg(Index, "-=- Statistique pour " & GetPlayerName(i) & " -=-", BrightGreen)
                            Call PlayerMsg(Index, "Niveau : " & GetPlayerLevel(i) & "  Exp : " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
                            Call PlayerMsg(Index, "PV : " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  PM : " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP : " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
                            Call PlayerMsg(Index, "FOR : " & GetPlayerStr(i) & "  DEF : " & GetPlayerDEF(i) & "  MAGIE : " & GetPlayerMAGI(i) & "  VIT : " & GetPlayerSPEED(i), BrightGreen)
                            n = (GetPlayerStr(i) \ 2) + (GetPlayerLevel(i) \ 2)
                            i = (GetPlayerDEF(i) \ 2) + (GetPlayerLevel(i) \ 2)
                            z = Int(GetPlayerSPEED(Index) * 0.576)
                            If n > 100 Then n = 100
                            If i > 100 Then i = 100
                            If z > 100 Then z = 100
                            Call PlayerMsg(Index, "Chance de coups critique : " & n & "%, Chance de bloquer : " & i & "%, Chance d'esquive : " & z & "%", BrightGreen)
                        End If
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            End Select
        Case "a"
            Select Case Parse(0)
                ' :: Player attack packet ::
                Case "attack"
                
                    If GetPlayerWeaponSlot(Index) > 0 Then
                        If item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).data3 > 0 Then
                            Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).data3 & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
                            Exit Sub
                        End If
                    End If
                    
                    ' Essaye d'attaquer un joueur
                    For i = 1 To MAX_PLAYERS
                        ' Etre sur qu'on s'attaque pas sois m�me en mode gros boulet .
                        If i <> Index Then
                            Randomize
                            
                            ' Peut on attaquer un joueur
                            If CanAttackPlayer(Index, i) Then
                                If Not CanPlayerBlockHit(i) And Not CanPlayerEsquiveHit(i) Then
                                    ' Optention du domage qu'on fera
                                    If Not CanPlayerCriticalHit(Index) Then
                                        Damage = GetPlayerDamage(Index) - GetPlayerProtection(i)
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & END_CHAR)
                                    Else
                                        n = GetPlayerDamage(Index)
                                        Damage = n + Int(Rnd * (n \ 2)) + 1 - GetPlayerProtection(i)
                                        Call BattleMsg(Index, "Vous fa�tes un coup critique .", BrightCyan, 0)
                                        Call BattleMsg(i, GetPlayerName(Index) & " a fait un coup critique .", BrightCyan, 1)
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & END_CHAR)
                                    End If
                                    
                                    If Damage > 0 Then
                                        Call AttackPlayer(Index, i, Damage)
                                    Else
                                        Call PlayerMsg(Index, "Votre attaque n'a aucune effet.", BrightRed)
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                                    End If
                                Else
                                    Call BattleMsg(Index, GetPlayerName(i) & " Bloque/Esquive votre attaque.", BrightCyan, 0)
                                    Call BattleMsg(i, "Vous bloquer/esquiver " & GetPlayerName(Index) & " par chance.", BrightCyan, 1)
                                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                                End If
                                
                                Exit Sub
                            End If
                        End If
                    Next i
                    
                    ' Try to attack a npc
                    For i = 1 To MAX_MAP_NPCS
                        ' Can we attack the npc?
                        If CanAttackNpc(Index, i) Then
                            ' Get the damage we can do
                            If Not CanPlayerCriticalHit(Index) Then
                                Damage = GetPlayerDamage(Index) - (Npc(MapNpc(GetPlayerMap(Index), i).Num).def \ 2)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & END_CHAR)
                            Else
                                n = GetPlayerDamage(Index)
                                Damage = n + Int(Rnd * (n \ 2)) + 1 - (Npc(MapNpc(GetPlayerMap(Index), i).Num).def \ 2)
                                Call BattleMsg(Index, "Vous fa�tes un coup critique .", BrightCyan, 0)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & END_CHAR)
                            End If
                            
                            If Damage > 0 Then
                                Call AttackNpc(Index, i, Damage)
                                If CLng(Npc(MapNpc(GetPlayerMap(Index), i).Num).Inv) = 0 Then Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & 0 & END_CHAR)
                            Else
                                Call BattleMsg(Index, "Votre attaque n'occasionne aucun d�g�t.", BrightRed, 0)
                                Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & 0 & END_CHAR)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                            End If
                            Exit Sub
                        End If
                    Next i
                    Exit Sub
                    
                Case "arrowhit"
                    ' Merci � Xamus (Fontor), Tom13 et Revorn qui ont trouv� ce hack possible
                    ' The player was able, from a 3rd party program, to send a packet arrowhit and kill peaple without a bow
                    n = Val(Parse(1))
                    z = Val(Parse(2))
                    x = Val(Parse(3))
                    y = Val(Parse(4))
                    
                    If GetPlayerWeaponSlot(Index) <= 0 Then Call PlayerMsg(Index, "Vous devez avoir un arc d'�quip�", BrightRed): Exit Sub
                    If item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).data3 <= 0 Then Call PlayerMsg(Index, "Vous devez avoir un arc d'�quip�", BrightRed): Exit Sub
                    
                    If n = TARGET_TYPE_PLAYER Then
                        ' Etre vraiment sur qu'on ne s'attaque pas sois m�me en mode encore plus gros boulet .
                        If z <> Index Then
                            ' Peut on attaquer ?
                            If CanAttackPlayerWithArrow(Index, z) Then
                                If Not CanPlayerBlockHit(z) And Not CanPlayerEsquiveHit(z) Then
                                    ' Quels dommages peut on faire ?
                                    If Not CanPlayerCriticalHit(Index) Then
                                        Damage = GetPlayerDamage(Index) - GetPlayerProtection(z)
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & END_CHAR)
                                    Else
                                        n = GetPlayerDamage(Index)
                                        Damage = n + Int(Rnd * (n \ 2)) + 1 - GetPlayerProtection(z)
                                        Call BattleMsg(Index, "Vous fa�tes un coup critique .", BrightCyan, 0)
                                        Call BattleMsg(z, GetPlayerName(Index) & " vous touche en faisant un coup critique .", BrightCyan, 1)
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & END_CHAR)
                                    End If
                                    
                                    If Damage > 0 Then
                                        Call AttackPlayer(Index, z, Damage)
                                    Else
                                        Call BattleMsg(Index, "Vous n'occasionnez aucun dommage.", BrightRed, 0)
                                        Call BattleMsg(z, GetPlayerName(z) & " ne vous occasionne aucun dommage.", BrightRed, 1)
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                                    End If
                                Else
                                    Call BattleMsg(Index, GetPlayerName(z) & " Bloque/Esquive votre attaque.", BrightCyan, 0)
                                    Call BattleMsg(z, "Vous bloquer/esquiver l'attaque de " & GetPlayerName(Index) & ".", BrightCyan, 1)
                                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                                End If
                                Exit Sub
                            End If
                        End If
                    ElseIf n = TARGET_TYPE_NPC Then
                        ' Peut on attaquer le PNJ
                        If CanAttackNpcWithArrow(Index, z) Then
                            ' Quels dommages peut on faire ?
                            If Not CanPlayerCriticalHit(Index) Then
                                Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), z).Num).def / 2)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & END_CHAR)
                            Else
                                n = GetPlayerDamage(Index)
                                Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), z).Num).def / 2)
                                Call BattleMsg(Index, "Vous sentez une grande �nergie quand vous tirez.", BrightCyan, 0)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "critical" & END_CHAR)
                            End If
                            
                            If Damage > 0 Then
                                Call AttackNpc(Index, z, Damage)
                                Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & 0 & END_CHAR)
                            Else
                                Call BattleMsg(Index, "Votre attaque n'occasione aucun dommage.", BrightRed, 0)
                                Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & 0 & END_CHAR)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                            End If
                            Exit Sub
                        End If
                    End If
                    Exit Sub
            
                Case "acoffre"
                    SlotI = Val(Parse(1))
                    INum = Val(Parse(2))
                    IVal = Val(Parse(3))
                    IDur = Val(Parse(4))
                    'prevent hacking
                    If IsPlaying(Index) = False Then Call HackingAttempt(Index, "Le joueur n'est pas en train de jouer"): Exit Sub
                    If (SlotI > 24 Or SlotI < 1) And INum <> 0 Then Call HackingAttempt(Index, "Invalidee Inv Slot"): Exit Sub
                    If INum < 0 Or INum > MAX_ITEMS Then Call HackingAttempt(Index, "Invalidee Item Num"): Exit Sub
                    On Error Resume Next
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type <> TILE_TYPE_COFFRE And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type <> TILE_TYPE_COFFRE And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type <> TILE_TYPE_COFFRE And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type <> TILE_TYPE_COFFRE Then Call HackingAttempt(Index, "Essaye de hacker l'Atributs Coffre,� peut �tre modifier le client"): Exit Sub
                            
                    Dim AY As Long
                    Dim AX As Long
                    AY = 0
                    AX = 0
                    
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).type = TILE_TYPE_COFFRE Then AY = -1: AX = 0
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).type = TILE_TYPE_COFFRE Then AY = 1: AX = 0
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).type = TILE_TYPE_COFFRE Then AY = 0: AX = 1
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).type = TILE_TYPE_COFFRE Then AY = 0: AX = -1
                            
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + AX, GetPlayerY(Index) + AY).data3 > 0 And Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + AX, GetPlayerY(Index) + AY).data3 < MAX_ITEMS Then
                        i = FindOpenInvSlot(Index, Val(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + AX, GetPlayerY(Index) + AY).data3))
                        
                        If i > 0 Then
                            Call SetPlayerInvItemNum(Index, i, Val(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + AX, GetPlayerY(Index) + AY).data3))
                            Call SetPlayerInvItemValue(Index, i, 1)
                            Call SetPlayerInvItemDur(Index, i, item(Val(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + AX, GetPlayerY(Index) + AY).data3)).data1)
                            Call PlayerMsg(Index, "Vous prenez un " & item(Val(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + AX, GetPlayerY(Index) + AY).data3)).Name, Green)
                            Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + AX, GetPlayerY(Index) + AY).data3 = 0
                        End If
                    Else
                        Call PlayerMsg(Index, "Le coffre est vide", Green)
                    End If
                    Exit Sub
                
                Case "adminmsg"
                    Msg = Parse(1)
                    ' Prevent hacking
                    If MMsg(Msg) Then Call HackingAttempt(Index, "Caract�re incorrect dans ses paroles(admin)"): Exit Sub
                    
                    If frmServer.chkA.value = Unchecked Then Call PlayerMsg(Index, "Les messages aux admins ont �t� d�sactiv�s.", BrightRed): Exit Sub
                            
                    If GetPlayerAccess(Index) > 0 Then
                        Call AddLog("(Admin : " & GetPlayerName(Index) & ") " & Msg, ADMIN_LOG)
                        Call AdminMsg("(Admin : " & GetPlayerName(Index) & ") " & Msg, AdminColor)
                    End If
                    TextAdd frmServer.txtText(5), GetPlayerName(Index) & " : " & Msg, True
                    Exit Sub
                    
                Case "atrade"
                    n = Player(Index).TradePlayer
                    ' Check if anyone requested a trade
                    If n < 1 Then Call PlayerMsg(Index, "Aucune requ�te d'�change avec vous.", Pink): Exit Sub
                    ' Check if its the right player
                    If Player(n).TradePlayer <> Index Then Call PlayerMsg(Index, "L'echange a echou�...", Pink): Exit Sub
                    ' Check where both players are
                    CanTrade = False
                    
                    If GetPlayerX(Index) = GetPlayerX(n) And GetPlayerY(Index) + 1 = GetPlayerY(n) Then CanTrade = True
                    If GetPlayerX(Index) = GetPlayerX(n) And GetPlayerY(Index) - 1 = GetPlayerY(n) Then CanTrade = True
                    If GetPlayerX(Index) + 1 = GetPlayerX(n) And GetPlayerY(Index) = GetPlayerY(n) Then CanTrade = True
                    If GetPlayerX(Index) - 1 = GetPlayerX(n) And GetPlayerY(Index) = GetPlayerY(n) Then CanTrade = True
                        
                    If CanTrade = True Then
                        Call PlayerMsg(Index, "Tu commerces avec " & GetPlayerName(n) & ".", Pink)
                        Call PlayerMsg(n, GetPlayerName(Index) & " accepte ta demande d'�change.", Pink)
                        Call SendDataTo(Index, "PPTRADING" & END_CHAR)
                        Call SendDataTo(n, "PPTRADING" & END_CHAR)
                        For i = 1 To MAX_PLAYER_TRADES
                            Player(Index).Trading(i).InvNum = 0
                            Player(Index).Trading(i).InvName = vbNullString
                            Player(n).Trading(i).InvNum = 0
                            Player(n).Trading(i).InvName = vbNullString
                        Next i
                        Player(Index).InTrade = 1
                        Player(Index).TradeItemMax = 0
                        Player(Index).TradeItemMax2 = 0
                        Player(n).InTrade = 1
                        Player(n).TradeItemMax = 0
                        Player(n).TradeItemMax2 = 0
                    Else
                        Call PlayerMsg(Index, "Le joueur doit �tre devant vous pour �changer.", Pink)
                        Call PlayerMsg(n, "Tu as besoin d'�tre devant le joueur pour �changer.", Pink)
                    End If
                    Exit Sub
                    
                ' Chat Packet
                Case "achat"
                    n = Player(Index).ChatPlayer
                    If n < 1 Then Call PlayerMsg(Index, "Aucune requ�te pour discuter avec vous.", Pink): Exit Sub
                    If Player(n).ChatPlayer <> Index Then
                    Call SendDataTo(Index, "PPCHATTING" & SEP_CHAR & n & END_CHAR)
                    Call SendDataTo(n, "PPCHATTING" & SEP_CHAR & Index & END_CHAR)
                    Else
                    Call PlayerMsg(Index, "La discution a �chou�..", Pink)
                    End If
                    Exit Sub
            End Select
        Case "m"
            Select Case Parse(0)
                Case "mapgetitem"
                    Call PlayerMapGetItem(Index)
                    Exit Sub
                    
                Case "mapdropitem"
                    InvNum = Val(Parse(1))
                    Amount = Val(Parse(2))
                    ' Prevent hacking
                    If InvNum < 1 Or InvNum > MAX_INV Then Call HackingAttempt(Index, "Invalide InvNum"): Exit Sub
                    If item(GetPlayerInvItemNum(Index, InvNum)).type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(Index, InvNum)).Empilable <> 0 Then
                        ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
                        If Amount <= 0 Then Call PlayerMsg(Index, "Tu dois jetter plus que 0.", BrightRed): Exit Sub
                        If Amount > GetPlayerInvItemValue(Index, InvNum) Then Call PlayerMsg(Index, "Tu n'as pas suffisement d'objets � jetter.", BrightRed): Exit Sub
                    Else
                        If Amount > GetPlayerInvItemValue(Index, InvNum) Then Call HackingAttempt(Index, "Modification du nombre d'objets"): Exit Sub
                    End If
                    
                    Call PlayerMapDropItem(Index, InvNum, Amount)
                    Call SendStats(Index)
                    Call SendHP(Index)
                    Call SendMP(Index)
                    Call SendSP(Index)
                    Exit Sub
                    
                Case "maprespawn"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' Clear out it all
                    For i = 1 To MAX_MAP_ITEMS
                        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
                        Call ClearMapItem(i, GetPlayerMap(Index))
                    Next i
                    ' Respawn
                    Call SpawnMapItems(GetPlayerMap(Index))
                    ' Respawn NPCS
                    For i = 1 To MAX_MAP_NPCS
                        Call SpawnNpc(i, GetPlayerMap(Index))
                    Next i
                    Call PlayerMsg(Index, "Carte r�initialis�e", Blue)
                    Call AddLog(GetPlayerName(Index) & " a r�initialis�(e) la carte #" & GetPlayerMap(Index), ADMIN_LOG)
                    Exit Sub
                
                ' Faire une nouvelle guilde
                Case "makeguild"
                    ' Check if the Owner is Online
                    If FindPlayer(Parse(1)) = 0 Then Call PlayerMsg(Index, "Personnage hors-ligne", White): Exit Sub
                   
                    ' Verifie si il est pas d�ja dans une guilde
                    If GetPlayerGuild(FindPlayer(Parse(1))) <> vbNullString Then Call PlayerMsg(Index, "Le joueur est d�j� dans une guilde", Red): Exit Sub
            
                    Dim Level As Integer
                    Dim Level_mini As Integer
            
                    Level = GetPlayerLevel(Index)
                    Level_mini = Val(GetVar(App.Path & "\Data.ini", "GUILDE", "LEVEL_MINI"))
            
            
                    If Level <= Level_mini Then
                        Call PlayerMsg(Index, "Tu dois �tre au minimum au niveau " & Level_mini & " pour cr�e ta guilde .", Red)
                    Else
                        ' If everything is ok then lets make the guild
                        Call SetPlayerGuild(FindPlayer(Parse(1)), (Parse(2)))
                        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 3)
                        Call SendPlayerData(FindPlayer(Parse(1)))
                        Call AddLog(GetPlayerName(Index) & " a cr�er sa guilde nommer : " & Parse(2) & ".", GUILDE_LOG)
                    End If
            
                    Exit Sub
                    
                Case "mapdata"
                     ' Prevent hacking
                     If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                     n = 1
                     MapNum = GetPlayerMap(Index)
                     Map(MapNum).Name = Parse(n + 1)
                     Map(MapNum).Revision = Map(MapNum).Revision + 1
                     Map(MapNum).Moral = Val(Parse(n + 3))
                     Map(MapNum).Up = Val(Parse(n + 4))
                     Map(MapNum).Down = Val(Parse(n + 5))
                     Map(MapNum).Left = Val(Parse(n + 6))
                     Map(MapNum).Right = Val(Parse(n + 7))
                     Map(MapNum).Music = Parse(n + 8)
                     Map(MapNum).BootMap = Val(Parse(n + 9))
                     Map(MapNum).BootX = Val(Parse(n + 10))
                     Map(MapNum).BootY = Val(Parse(n + 11))
                     Map(MapNum).Indoors = Val(Parse(n + 12))
                     n = n + 13
                     
                     For y = 0 To MAX_MAPY
                         For x = 0 To MAX_MAPX
                             Map(MapNum).Tile(x, y).Ground = Val(Parse(n))
                             Map(MapNum).Tile(x, y).Mask = Val(Parse(n + 1))
                             Map(MapNum).Tile(x, y).Anim = Val(Parse(n + 2))
                             Map(MapNum).Tile(x, y).Mask2 = Val(Parse(n + 3))
                             Map(MapNum).Tile(x, y).M2Anim = Val(Parse(n + 4))
                             Map(MapNum).Tile(x, y).Mask3 = Val(Parse(n + 32)) '<--
                             Map(MapNum).Tile(x, y).M3Anim = Val(Parse(n + 30)) '<--
                             Map(MapNum).Tile(x, y).Fringe = Val(Parse(n + 5))
                             Map(MapNum).Tile(x, y).FAnim = Val(Parse(n + 6))
                             Map(MapNum).Tile(x, y).Fringe2 = Val(Parse(n + 7))
                             Map(MapNum).Tile(x, y).F2Anim = Val(Parse(n + 8))
                             Map(MapNum).Tile(x, y).Fringe3 = Val(Parse(n + 26)) '<--
                             Map(MapNum).Tile(x, y).F3Anim = Val(Parse(n + 27)) '<--
                             Map(MapNum).Tile(x, y).type = Val(Parse(n + 9))
                             Map(MapNum).Tile(x, y).data1 = Val(Parse(n + 10))
                             Map(MapNum).Tile(x, y).data2 = Val(Parse(n + 11))
                             Map(MapNum).Tile(x, y).data3 = Val(Parse(n + 12))
                             Map(MapNum).Tile(x, y).String1 = Parse(n + 13)
                             Map(MapNum).Tile(x, y).String2 = Parse(n + 14)
                             Map(MapNum).Tile(x, y).String3 = Parse(n + 15)
                             Map(MapNum).Tile(x, y).Light = Val(Parse(n + 16))
                             Map(MapNum).Tile(x, y).GroundSet = Val(Parse(n + 17))
                             Map(MapNum).Tile(x, y).MaskSet = Val(Parse(n + 18))
                             Map(MapNum).Tile(x, y).AnimSet = Val(Parse(n + 19))
                             Map(MapNum).Tile(x, y).Mask2Set = Val(Parse(n + 20))
                             Map(MapNum).Tile(x, y).M2AnimSet = Val(Parse(n + 21))
                             Map(MapNum).Tile(x, y).Mask3Set = Val(Parse(n + 33)) '<--
                             Map(MapNum).Tile(x, y).M3AnimSet = Val(Parse(n + 31)) '<--
                             Map(MapNum).Tile(x, y).FringeSet = Val(Parse(n + 22))
                             Map(MapNum).Tile(x, y).FAnimSet = Val(Parse(n + 23))
                             Map(MapNum).Tile(x, y).Fringe2Set = Val(Parse(n + 24))
                             Map(MapNum).Tile(x, y).F2AnimSet = Val(Parse(n + 25))
                             Map(MapNum).Tile(x, y).Fringe3Set = Val(Parse(n + 28)) '<--
                             Map(MapNum).Tile(x, y).F3AnimSet = Val(Parse(n + 29)) '<--
                             n = n + 34
                         Next x
                     Next y
                    
                     For x = 1 To MAX_MAP_NPCS
                         Map(MapNum).Npc(x) = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).x = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).y = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).x1 = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).y1 = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).x2 = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).y2 = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).Hasardm = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).Hasardp = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).boucle = Val(Parse(n))
                         n = n + 1
                         Map(MapNum).Npcs(x).Imobile = Val(Parse(n))
                         n = n + 1
                         Call ClearMapNpc(x, MapNum)
                     Next x
                     Map(MapNum).PanoInf = Parse(n)
                     Map(MapNum).TranInf = Val(Parse(n + 1))
                     Map(MapNum).PanoSup = Parse(n + 2)
                     Map(MapNum).TranSup = Val(Parse(n + 3))
                     Map(MapNum).Fog = Val(Parse(n + 4))
                     Map(MapNum).FogAlpha = Val(Parse(n + 5))
                     Map(MapNum).guildSoloView = Parse(n + 6)
                     Map(MapNum).petView = Parse(n + 7)
                     Map(MapNum).traversable = Parse(n + 8)
                     Map(MapNum).meteo = Val(Parse(n + 9))
                     
                     ' Clear out it all
                     For i = 1 To MAX_MAP_ITEMS
                         Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
                         Call ClearMapItem(i, GetPlayerMap(Index))
                     Next i
                     
                     ' Sauvegarde la map
                     Call SaveMap(MapNum)
                     
                     ' Respawn
                     Call SpawnMapItems(GetPlayerMap(Index))
                     
                     ' Respawn NPCS
                     For i = 1 To MAX_MAP_NPCS
                         Call SpawnNpc(i, GetPlayerMap(Index))
                     Next i
                     
                     ' Rafraichir la map pour tous les connect�s
                     For i = 1 To MAX_PLAYERS
                         If IsPlaying(i) And GetPlayerMap(i) = MapNum And i <> Index Then Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & END_CHAR)
                     Next i
                             
                     'V�rifier si les bords sont li�s a une autre map et la modifier en cons�quence
                     If Map(MapNum).Up > 0 And Map(MapNum).Up < MAX_MAPS Then Map(Map(MapNum).Up).Down = MapNum: Map(Map(MapNum).Up).Revision = Map(Map(MapNum).Up).Revision + 1: Call SaveMap(Map(MapNum).Up)
                     If Map(MapNum).Down > 0 And Map(MapNum).Down < MAX_MAPS Then Map(Map(MapNum).Down).Up = MapNum: Map(Map(MapNum).Down).Revision = Map(Map(MapNum).Down).Revision + 1: Call SaveMap(Map(MapNum).Down)
                     If Map(MapNum).Left > 0 And Map(MapNum).Left < MAX_MAPS Then Map(Map(MapNum).Left).Right = MapNum: Map(Map(MapNum).Left).Revision = Map(Map(MapNum).Left).Revision + 1: Call SaveMap(Map(MapNum).Left)
                     If Map(MapNum).Right > 0 And Map(MapNum).Right < MAX_MAPS Then Map(Map(MapNum).Right).Left = MapNum: Map(Map(MapNum).Right).Revision = Map(Map(MapNum).Right).Revision + 1: Call SaveMap(Map(MapNum).Right)
                     
                     Call AddLog(GetPlayerName(Index) & " a modifi�(e) la carte #" & GetPlayerMap(Index), ADMIN_LOG)
                     Call SendDataTo(Index, "CARTESAVE" & END_CHAR)
                     Exit Sub
                     
                Case "mapdown"
                    Dim url As String
                    Dim rep As String
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    
                    z = GetPlayerMap(Index)
                    rep = GetVar(App.Path & "\Data.ini", "FTP", "REP")
                    url = GetVar(App.Path & "\Data.ini", "FTP", "URL")
                    
                    If z <= 0 Or z > MAX_MAPS Then Exit Sub
                    If Mid(url, Len(url)) <> "/" And Mid(rep, 1, 1) <> "/" Then url = url & "/"
                    If rep <> "/" Then If Mid(rep, Len(rep)) <> "/" Then rep = rep & "/"
                    If Mid(url, Len(url)) = "/" And Mid(rep, 1, 1) = "/" Then rep = Mid(rep, 2)
                    
                    Call MapDo(z, url, rep)
                    Call LoadMap(z)
                    
                    ' Clear out it all
                    For i = 1 To MAX_MAP_ITEMS
                        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
                        Call ClearMapItem(i, GetPlayerMap(Index))
                    Next i
                    
                    Call SendMapNpcsToMap(GetPlayerMap(Index))
                    
                    ' Respawn
                    Call SpawnMapItems(GetPlayerMap(Index))
                    
                    ' Respawn NPCS
                    For i = 1 To MAX_MAP_NPCS
                        Call ClearMapNpc(i, GetPlayerMap(Index))
                        Call SpawnNpc(i, GetPlayerMap(Index))
                    Next i
                    
                    ' Rafraichir la map pour tous les joueurs en ligne
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) And GetPlayerMap(i) = MapNum And i <> Index Then Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & END_CHAR)
                    Next i
                    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                    Call AddLog(GetPlayerName(Index) & " a modifi�(e) la carte #" & GetPlayerMap(Index), ADMIN_LOG)
                    Call SendDataTo(Index, "CARTESAVE" & END_CHAR)
                    Exit Sub
                    
                Case "modifinv"
                    SlotI = Val(Parse(1))
                    INum = Val(Parse(2))
                    IVal = Val(Parse(3))
                    IDur = Val(Parse(4))
                    Cslot = Val(Parse(5))
                            
                    'Prevent Hacking
                    If IsPlaying(Index) = False Then Call HackingAttempt(Index, "Le joueur n'est pas en trian de jouer"): Exit Sub
                    If (SlotI > 24 Or SlotI < 1) And INum <> 0 Then Call HackingAttempt(Index, "Inv Slot Invalide"): Exit Sub
                    If (Cslot > 30 Or Cslot < 1) And INum <> 0 Then Call HackingAttempt(Index, "Slot de Coffre Invalide"): Exit Sub
                    If INum < 0 Or INum > MAX_ITEMS Then Call HackingAttempt(Index, "Num�ros d'objet Invalide"): Exit Sub
                    If GetPlayerInvItemNum(Index, SlotI) <> INum And INum <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont inventaire"): Exit Sub
                    If GetPlayerInvItemValue(Index, SlotI) < IVal And IVal <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont inventaire"): Exit Sub
                    If GetPlayerInvItemDur(Index, SlotI) <> IDur And IDur <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont inventaire"): Exit Sub
                            
                    Call SetPlayerInvItemNum(Index, SlotI, INum)
                    Call SetPlayerInvItemValue(Index, SlotI, IVal)
                    Call SetPlayerInvItemDur(Index, SlotI, IDur)
                    Call SendInventoryUpdate(Index, SlotI)
                    Exit Sub
                    
                Case "modifcoffre"
                    SlotC = Val(Parse(1))
                    Cnum = Val(Parse(2))
                    Cval = Val(Parse(3))
                    Cdur = Val(Parse(4))
                    Islot = Val(Parse(5))
            
                    'Prevent Hacking
                    If IsPlaying(Index) = False Then Call HackingAttempt(Index, "Le joueur n'est pas en train de jouer"): Exit Sub
                    If (Islot > 24 Or Islot < 1) And Cnum <> 0 Then Call HackingAttempt(Index, "Inv Slot Invalide"): Exit Sub
                    If (SlotC > 30 Or SlotC < 1) And Cnum <> 0 Then Call HackingAttempt(Index, "Slot de Coffre Invalide"): Exit Sub
                    If Cnum < 0 Or Cnum > MAX_ITEMS Then Call HackingAttempt(Index, "Num�ros d'objet Invalide"): Exit Sub
                    If Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemnum" & SlotC)) <> Cnum And Cnum <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont coffre"): Exit Sub
                    If Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemval" & SlotC)) < Cval And Cval <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont coffre"): Exit Sub
                    If Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemdur" & SlotC)) <> Cdur And Cdur <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont coffre"): Exit Sub
                            
                    Call PutVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemnum" & SlotC, " " & Val(Cnum))
                    Call PutVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemval" & SlotC, " " & Val(Cval))
                    Call PutVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemdur" & SlotC, " " & Val(Cdur))
                    Exit Sub
                    
                Case "mapreport"
                    Packs = "mapreport" & SEP_CHAR
                    For i = 1 To MAX_MAPS
                        Packs = Packs & Map(i).Name & SEP_CHAR
                    Next i
                    Packs = Packs & END_CHAR
                    Call SendDataTo(Index, Packs)
                    Exit Sub
            End Select
        Case "n"
            Select Case Parse(0)
                Case "needmap"
                    ' Get yes/no value
                    s = LCase$(Parse(1))
                    If s = "yes" Then Call SendMap(Index, GetPlayerMap(Index))
                    Call SendMapItemsTo(Index, GetPlayerMap(Index))
                    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
                    Call SendJoinMap(Index)
                    Player(Index).GettingMap = NO
                    Call SendDataTo(Index, "MAPDONE" & SEP_CHAR & CStr(CarteFTP) & END_CHAR)
                    
                    Call SendDataTo(Index, "CARTESAVE" & END_CHAR)
                    If AvMonture(Index) And Map(GetPlayerMap(Index)).Indoors >= 1 Then
                        Call SetPlayerArmorSlot(Index, 0)
                        s = Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "monture"))
                        Call SetPlayerSprite(Index, s)
                        Call SendPlayerData(Index)
                        Call SendInventory(Index)
                        Call SendWornEquipment(Index)
                    End If
                    Exit Sub
                Case "newmetier"
                    Player(Index).Char(Player(Index).CharNum).metier = Val(Parse(1))
                    Player(Index).Char(Player(Index).CharNum).MetierLvl = 1
                    Player(Index).Char(Player(Index).CharNum).MetierExp = 0
                    Call PlayerMsg(Index, "Vous avez appris un m�tier", Green)
                    Exit Sub
                    
                Case "needsmap"
                    Dim Heur As Long
                    Dim Jour As Long
                    Dim Mois As Long
                    Dim Anne As Long
                    Dim Dmod As String
                    Dmod = Parse(1)
                    Anne = Val(Year(Dmod))
                    Mois = Val(Month(Dmod))
                    Jour = Val(Day(Dmod))
                    Heur = Val(Hour(Dmod))
                    s = vbNullString
                                    
                    If Val(Year(FileDateTime(App.Path & "\maps\map" & GetPlayerMap(Index) & ".fcc"))) > Anne Then
                        s = "yes"
                    ElseIf Val(Year(FileDateTime(App.Path & "\maps\map" & GetPlayerMap(Index) & ".fcc"))) = Anne Then
                        If Val(Month(FileDateTime(App.Path & "\maps\map" & GetPlayerMap(Index) & ".fcc"))) > Mois Then
                            s = "yes"
                        ElseIf Val(Month(FileDateTime(App.Path & "\maps\map" & GetPlayerMap(Index) & ".fcc"))) = Mois Then
                            If Val(Day(FileDateTime(App.Path & "\maps\map" & GetPlayerMap(Index) & ".fcc"))) > Jour Then
                                s = "yes"
                            ElseIf Val(Day(FileDateTime(App.Path & "\maps\map" & GetPlayerMap(Index) & ".fcc"))) = Jour Then
                                If Val(Hour(FileDateTime(App.Path & "\maps\map" & GetPlayerMap(Index) & ".fcc"))) > Heur Then
                                    s = "yes"
                                Else
                                    s = vbNullString
                                End If
                            End If
                        End If
                    End If
                    If s = "yes" Then Call SendMap(Index, GetPlayerMap(Index))
                    Call SendMapItemsTo(Index, GetPlayerMap(Index))
                    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
                    Call SendJoinMap(Index)
                    Player(Index).GettingMap = NO
                    Call SendDataTo(Index, "MAPDONE" & SEP_CHAR & CStr(CarteFTP) & END_CHAR)
                    
                    Call SendDataTo(Index, "CARTESAVE" & END_CHAR)
                    If AvMonture(Index) And Map(GetPlayerMap(Index)).Indoors >= 1 Then
                        Call SetPlayerArmorSlot(Index, 0)
                        s = Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "monture"))
                        Call SetPlayerSprite(Index, s)
                        Call SendPlayerData(Index)
                        Call SendInventory(Index)
                        Call SendWornEquipment(Index)
                    End If
                    Exit Sub
            End Select
        Case Else
            Select Case Parse(0)
                Case "useitem"
                    InvNum = Val(Parse(1))
                    CharNum = Player(Index).CharNum
                    
                    ' Prevent hacking
                    If InvNum < 1 Or InvNum > MAX_INV Then Call HackingAttempt(Index, "Invalide InvNum"): Exit Sub
                            
                    ' Prevent hacking
                    If CharNum < 1 Or CharNum > MAX_CHARS Then Call HackingAttempt(Index, "Num�ros de personnage invalide."): Exit Sub
                            
                    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
                        n = item(GetPlayerInvItemNum(Index, InvNum)).data2
                        
                        Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, n5 As Long, mi As Long, niveauRequis As Integer
                        n1 = item(GetPlayerInvItemNum(Index, InvNum)).StrReq
                        n2 = item(GetPlayerInvItemNum(Index, InvNum)).DefReq
                        n3 = item(GetPlayerInvItemNum(Index, InvNum)).SpeedReq
                        n4 = item(GetPlayerInvItemNum(Index, InvNum)).ClassReq
                        n5 = item(GetPlayerInvItemNum(Index, InvNum)).AccessReq
                        niveauRequis = item(GetPlayerInvItemNum(Index, InvNum)).LevelReq
                        
                        If item(GetPlayerInvItemNum(Index, InvNum)).Empilable <> 0 Then
                            mi = 1
                        Else
                            mi = 0
                        End If
                                    
                        Select Case item(GetPlayerInvItemNum(Index, InvNum)).type
                            Case ITEM_TYPE_ARMOR
                                If InvNum <> GetPlayerArmorSlot(Index) Then
                                    If GetPlayerArmorSlot(Index) > 0 Then If item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).type = ITEM_TYPE_MONTURE Then Call EnMonture(Index)
                                    If n4 > -1 Then If GetPlayerClass(Index) <> n4 Then Call PlayerMsg(Index, "Tu as besoin d'�tre un " & GetClassName(n4) & " pour utiliser ceci.", BrightRed): Exit Sub
                                    If GetPlayerAccess(Index) < n5 Then Call PlayerMsg(Index, "Votre acc�s doit �tre sup�rieur � " & n5 & ".", BrightRed): Exit Sub
                                    If item(GetPlayerInvItemNum(Index, InvNum)).Sex <> GetPlayerSex(Index) And item(GetPlayerInvItemNum(Index, InvNum)).Sex <> 2 Then Call PlayerMsg(Index, "Tu n'es pas du bon sexe pour utiliser ceci.", BrightRed): Exit Sub
                                    If Int(GetPlayerStr(Index)) < n1 Then
                                        Call PlayerMsg(Index, "Votre force n'est pas suffisante pour ceci.  Force requise : " & n1, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                        Call PlayerMsg(Index, "Votre d�fense n'est pas suffisante pour ceci.  D�fense requise : " & n2, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                        Call PlayerMsg(Index, "Votre vitesse n'est pas suffisante pour ceci.  Vitesse requise : " & n3, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                    End If
                                    Call MyScript.ExecuteStatement("Scripts\Main.txt", "OnArmorUse " & Index & "," & GetPlayerInvItemNum(Index, InvNum))
                                    Call SetPlayerArmorSlot(Index, InvNum)
                                Else
                                    Call SetPlayerArmorSlot(Index, 0)
                                End If
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                                
                            Case ITEM_TYPE_MONTURE
                                If InvNum <> GetPlayerArmorSlot(Index) Then
                                    If Map(GetPlayerMap(Index)).Indoors >= 1 Then Call PlayerMsg(Index, "Vous �tes en int�rieur vous ne pouvez pas avoir de monture.", BrightRed): Exit Sub
                                    If n4 > -1 Then If GetPlayerClass(Index) <> n4 Then Call PlayerMsg(Index, "Tu as besoin d'�tre un " & GetClassName(n4) & " pour utiliser ceci.", BrightRed): Exit Sub
                                    If GetPlayerAccess(Index) < n5 Then Call PlayerMsg(Index, "Votre acc�s doit �tre sup�rieur � " & n5 & ".", BrightRed): Exit Sub
                                    If Int(GetPlayerStr(Index)) < n1 Then
                                        Call PlayerMsg(Index, "Votre force n'est pas suffisante pour ceci.  Force requise : " & n1, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                        Call PlayerMsg(Index, "Votre d�fense n'est pas suffisante pour ceci.  D�fense requise : " & n2, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                        Call PlayerMsg(Index, "Votre vitesse n'est pas suffisante pour ceci.  Vitesse requise : " & n3, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                    End If
                                    Call SetPlayerArmorSlot(Index, InvNum)
                                    Call PutVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "monture", GetPlayerSprite(Index))
                                    n = item(GetPlayerInvItemNum(Index, InvNum)).data1
                                    Call SetPlayerSprite(Index, n)
                                    Call MyScript.ExecuteStatement("Scripts\Main.txt", "OnMontureUse " & Index & "," & GetPlayerInvItemNum(Index, InvNum))
                                    Call SendPlayerData(Index)
                                Else
                                    Call SetPlayerArmorSlot(Index, 0)
                                    Call EnMonture(Index)
                                End If
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
            
                            Case ITEM_TYPE_WEAPON
                                If InvNum <> GetPlayerWeaponSlot(Index) Then
                                    If n4 > -1 Then If GetPlayerClass(Index) <> n4 Then Call PlayerMsg(Index, "Tu as besoin d'�tre un " & GetClassName(n4) & " pour utiliser ceci.", BrightRed): Exit Sub
                                    If GetPlayerAccess(Index) < n5 Then Call PlayerMsg(Index, "Votre acc�s doit �tre sup�rieur � " & n5 & ".", BrightRed): Exit Sub
                                    If item(GetPlayerInvItemNum(Index, InvNum)).Sex <> GetPlayerSex(Index) And item(GetPlayerInvItemNum(Index, InvNum)).Sex <> 2 Then Call PlayerMsg(Index, "Tu n'es pas du bon sexe pour utiliser ceci.", BrightRed): Exit Sub
                                    If Int(GetPlayerStr(Index)) < n1 Then
                                        Call PlayerMsg(Index, "Votre force n'est pas suffisante pour ceci.  Force requise : " & n1, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                        Call PlayerMsg(Index, "Votre d�fense n'est pas suffisant pour ceci.  D�fense requise : " & n2, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                        Call PlayerMsg(Index, "Votre vitesse n'est pas suffisant pour ceci.  Vitesse requise : " & n3, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                    End If
                                    Call SetPlayerWeaponSlot(Index, InvNum)
                                    Call MyScript.ExecuteStatement("Scripts\Main.txt", "OnWeaponUse " & Index & "," & GetPlayerInvItemNum(Index, InvNum))
                                Else
                                    Call SetPlayerWeaponSlot(Index, 0)
                                End If
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                                    
                            Case ITEM_TYPE_HELMET
                                If InvNum <> GetPlayerHelmetSlot(Index) Then
                                    If n4 > -1 Then If GetPlayerClass(Index) <> n4 Then Call PlayerMsg(Index, "Tu as besoin d'�tre un " & GetClassName(n4) & " pour utiliser ceci.", BrightRed): Exit Sub
                                    If GetPlayerAccess(Index) < n5 Then Call PlayerMsg(Index, "Votre acc�s doit �tre sup�rieur � " & n5 & ".", BrightRed): Exit Sub
                                    If item(GetPlayerInvItemNum(Index, InvNum)).Sex <> GetPlayerSex(Index) And item(GetPlayerInvItemNum(Index, InvNum)).Sex <> 2 Then Call PlayerMsg(Index, "Tu n'es pas du bon sexe pour utiliser ceci.", BrightRed): Exit Sub
                                    If Int(GetPlayerStr(Index)) < n1 Then
                                        Call PlayerMsg(Index, "Votre force n'est pas suffisante pour ceci.  Force requise : " & n1, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                        Call PlayerMsg(Index, "Votre d�fense n'est pas suffisante pour ceci.  D�fense requise : " & n2, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                        Call PlayerMsg(Index, "Votre vitesse n'est pas suffisante pour ceci.  Vitesse requise : " & n3, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                    End If
                                    Call SetPlayerHelmetSlot(Index, InvNum)
                                    Call MyScript.ExecuteStatement("Scripts\Main.txt", "OnHelmetUse " & Index & "," & GetPlayerInvItemNum(Index, InvNum))
                                Else
                                    Call SetPlayerHelmetSlot(Index, 0)
                                End If
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                        
                            Case ITEM_TYPE_SHIELD
                                If InvNum <> GetPlayerShieldSlot(Index) Then
                                    If n4 > -1 Then If GetPlayerClass(Index) <> n4 Then Call PlayerMsg(Index, "Tu as besoin d'�tre un " & GetClassName(n4) & " pour utiliser ceci.", BrightRed): Exit Sub
                                    If GetPlayerAccess(Index) < n5 Then Call PlayerMsg(Index, "Votre acc�s doit �tre sup�rieur � " & n5 & ".", BrightRed): Exit Sub
                                    If item(GetPlayerInvItemNum(Index, InvNum)).Sex <> GetPlayerSex(Index) And item(GetPlayerInvItemNum(Index, InvNum)).Sex <> 2 Then Call PlayerMsg(Index, "Tu n'es pas du bon sexe pour utiliser ceci.", BrightRed): Exit Sub
                                    If Int(GetPlayerStr(Index)) < n1 Then
                                        Call PlayerMsg(Index, "Votre force n'est pas suffisante pour ceci.  Force requise : " & n1, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                        Call PlayerMsg(Index, "Votre d�fense n'est pas suffisante pour ceci.  D�fense requise : " & n2, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                        Call PlayerMsg(Index, "Votre vitesse n'est pas suffisante pour ceci.  Vitesse requise : " & n3, BrightRed)
                                        Exit Sub
                                    ElseIf Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                    End If
                                    Call SetPlayerShieldSlot(Index, InvNum)
                                    Call MyScript.ExecuteStatement("Scripts\Main.txt", "OnShieldUse " & Index & "," & GetPlayerInvItemNum(Index, InvNum))
                                Else
                                    Call SetPlayerShieldSlot(Index, 0)
                                End If
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                        
                            Case ITEM_TYPE_SCRIPT
                                If Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                End If
                                n = item(GetPlayerInvItemNum(Index, InvNum)).data1
                                If item(Player(Index).Char(CharNum).Inv(InvNum).Num).data2 = 1 Then Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, mi)
                                MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Val(n)
                            
                            Case ITEM_TYPE_PET
                                If Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis (" & niveauRequis & ")", BrightRed)
                                        Exit Sub
                                End If
                                If InvNum <> GetPlayerPetSlot(Index) Then
                                    Call SetPlayerPetSlot(Index, InvNum)
                                    'Call PlayerPet(Index, 0, GetPlayerDir(Index))
                                    Call PetMove(Index)
                                Else
                                    Call SetPlayerPetSlot(Index, 0)
                                End If
                                Call SendInventory(Index)
                                Call SendWornEquipment(Index)
                                
                            Case ITEM_TYPE_POTIONADDHP
                                If Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                End If
                                Call SetPlayerHP(Index, GetPlayerHP(Index) + item(Player(Index).Char(CharNum).Inv(InvNum).Num).data1)
                                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, mi)
                                Call SendHP(Index)
                            
                            Case ITEM_TYPE_POTIONADDMP
                                If Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                End If
                                Call SetPlayerMP(Index, GetPlayerMP(Index) + item(Player(Index).Char(CharNum).Inv(InvNum).Num).data1)
                                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, mi)
                                Call SendMP(Index)
                    
                            Case ITEM_TYPE_POTIONADDSP
                                If Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                End If
                                Call SetPlayerSP(Index, GetPlayerSP(Index) + item(Player(Index).Char(CharNum).Inv(InvNum).Num).data1)
                                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, mi)
                                Call SendSP(Index)
            
                            Case ITEM_TYPE_POTIONSUBHP
                                If Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                End If
                                Call SetPlayerHP(Index, GetPlayerHP(Index) - item(Player(Index).Char(CharNum).Inv(InvNum).Num).data1)
                                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, mi)
                                Call SendHP(Index)
                            
                            Case ITEM_TYPE_POTIONSUBMP
                                If Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                End If
                                Call SetPlayerMP(Index, GetPlayerMP(Index) - item(Player(Index).Char(CharNum).Inv(InvNum).Num).data1)
                                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, mi)
                                Call SendMP(Index)
                    
                            Case ITEM_TYPE_POTIONSUBSP
                                If Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                End If
                                Call SetPlayerSP(Index, GetPlayerSP(Index) - item(Player(Index).Char(CharNum).Inv(InvNum).Num).data1)
                                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, mi)
                                Call SendSP(Index)
                                
                            Case ITEM_TYPE_KEY
                                If Int(GetPlayerLevel(Index)) < niveauRequis Then
                                        Call PlayerMsg(Index, "Votre niveau n'est pas suffisant pour ceci.  Niveau requis : " & niveauRequis, BrightRed)
                                        Exit Sub
                                End If
                                Select Case GetPlayerDir(Index)
                                    Case DIR_UP
                                        If GetPlayerY(Index) > 0 Then x = GetPlayerX(Index): y = GetPlayerY(Index) - 1 Else Exit Sub
                                    Case DIR_DOWN
                                        If GetPlayerY(Index) < MAX_MAPY Then x = GetPlayerX(Index): y = GetPlayerY(Index) + 1 Else Exit Sub
                                    Case DIR_LEFT
                                        If GetPlayerX(Index) > 0 Then x = GetPlayerX(Index) - 1: y = GetPlayerY(Index) Else Exit Sub
                                    Case DIR_RIGHT
                                        If GetPlayerX(Index) < MAX_MAPY Then x = GetPlayerX(Index) + 1: y = GetPlayerY(Index) Else Exit Sub
                                End Select
                                
                                ' Check if a key exists
                                If Map(GetPlayerMap(Index)).Tile(x, y).type = TILE_TYPE_KEY Then
                                    ' Check if the key they are using matches the map key
                                    If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(x, y).data1 Then
                                        TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                        
                                        Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                                        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                                            Call MapMsg(GetPlayerMap(Index), "La porte a �t� ouverte.", White)
                                        Else
                                            Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
                                        End If
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & END_CHAR)
                                        
                                        ' Check if we are supposed to take away the item
                                        If Map(GetPlayerMap(Index)).Tile(x, y).data2 = 1 Then
                                            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), mi)
                                            Call PlayerMsg(Index, "La cl� se dissous.", Yellow)
                                        End If
                                    End If
                                End If
                                
                                ' Check if a key exists
                                If Map(GetPlayerMap(Index)).Tile(x, y).type = TILE_TYPE_COFFRE Then
                                    ' Check if the key they are using matches the map key
                                    If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(x, y).data1 Then
                                        TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                        
                                        Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                                        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                                            Call MapMsg(GetPlayerMap(Index), "Le coffre a �t� ouvert.", White)
                                        Else
                                            Call MapMsg(GetPlayerMap(Index), "Il faut un code pour ouvrir se coffre.", White)
                                        End If
                                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & END_CHAR)
                                        
                                        ' Check if we are supposed to take away the item
                                        If Map(GetPlayerMap(Index)).Tile(x, y).data2 = 1 Or Map(GetPlayerMap(Index)).Tile(x, y).data2 = "1" Then
                                            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), mi)
                                            Call PlayerMsg(Index, "La cl� se dissous.", Yellow)
                                        End If
                                        
                                        Call GiveItem(Index, Val(Map(GetPlayerMap(Index)).Tile(x, y).data3), 1)
                                    End If
                                End If
                                
                                If Map(GetPlayerMap(Index)).Tile(x, y).type = TILE_TYPE_DOOR Then
                                    TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                                    TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                    
                                    Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & END_CHAR)
                                End If
                                
                            Case ITEM_TYPE_SPELL
                                ' Optention du num�ro du sort
                                n = item(GetPlayerInvItemNum(Index, InvNum)).data1
                                
                                If n > 0 Then
                                    ' Etre sur que on est dans la bonne classe
                                    If Spell(n).ClassReq - 1 = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                                        If Spell(n).LevelReq = 0 And Player(Index).Char(Player(Index).CharNum).Access < 1 Then Call PlayerMsg(Index, "Ce sort peut uniquement �tre utilis� par un admin.", BrightRed): Exit Sub
                                                                    
                                        ' Etre sur qu'on a un level suffisant
                                        i = GetSpellReqLevel(Index, n)
                                        If i <= GetPlayerLevel(Index) Then
                                            i = FindOpenSpellSlot(Index)
                                            
                                            ' Etre sur que le slot est libre
                                            If i > 0 Then
                                                ' Etre sur on a pas d�ja le sort
                                                If Not HasSpell(Index, n) Then
                                                    Call SetPlayerSpell(Index, i, n)
                                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), mi)
                                                    Call PlayerMsg(Index, "Tu �tudies le sort avec concentration...", Yellow)
                                                    Call PlayerMsg(Index, "Tu as appris un nouveau sort.", White)
                                                Else
                                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), mi)
                                                    Call PlayerMsg(Index, "Tu as d�j� appris ce sort. Le sort disparait.", BrightRed)
                                                End If
                                            Else
                                                Call PlayerMsg(Index, "Tu as appris tous ce que tu peux.", BrightRed)
                                            End If
                                        Else
                                            Call PlayerMsg(Index, "Tu dois �tre au niveau " & i & " pour apprendre ce sort.", White)
                                        End If
                                    Else
                                        Call PlayerMsg(Index, "Ce sort peut �tre appris uniquement par un " & GetClassName(Spell(n).ClassReq - 1) & ".", White)
                                    End If
                                Else
                                    Call PlayerMsg(Index, "Ce parchemin n'est pas li� � un sort, contactez l'admin.", White)
                                End If
                        End Select
                        
                        Call SendStats(Index)
                        Call SendHP(Index)
                        Call SendMP(Index)
                        Call SendSP(Index)
                    End If
                    Exit Sub
                    
                Case "usestatpoint"
                    PointType = Val(Parse(1))
                    
                    ' Prevent hacking
                    If (PointType < 0) Or (PointType > 3) Then Call HackingAttempt(Index, "Type de points Invalide"): Exit Sub
                                    
                    ' Make sure they have points
                    If GetPlayerPOINTS(Index) > 0 Then
                        If Scripting = 1 Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & Index & "," & PointType
                        Else
                            Select Case PointType
                                Case 0
                                    Call SetPlayerStr(Index, Player(Index).Char(Player(Index).CharNum).STR + 1)
                                    Call BattleMsg(Index, "Vous avez gagn� plus de force.", 15, 0)
                                Case 1
                                    Call SetPlayerDEF(Index, Player(Index).Char(Player(Index).CharNum).def + 1)
                                    Call BattleMsg(Index, "Vous avez gagn� plus de d�fense.", 15, 0)
                                Case 2
                                    Call SetPlayerMAGI(Index, Player(Index).Char(Player(Index).CharNum).magi + 1)
                                    Call BattleMsg(Index, "Vous avez gagn� plus de magie.", 15, 0)
                                Case 3
                                    Call SetPlayerSPEED(Index, Player(Index).Char(Player(Index).CharNum).Speed + 1)
                                    Call BattleMsg(Index, "Vous avez gagn� plus de vitesse.", 15, 0)
                            End Select
                            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
                        End If
                    Else
                        Call BattleMsg(Index, "Tu n'as pas de points � distribuer.", BrightRed, 0)
                    End If
                    
                    Call SendHP(Index)
                    Call SendMP(Index)
                    Call SendSP(Index)
                    Call SendStats(Index)
                    
                    Call SendDataTo(Index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR)
                    Exit Sub
            'End Select
        'Case Else
            'Select Case Parse(0)
           
                ' :: Guilds Packet ::
                ' Access
                Case "guildchangeaccess"
                    ' Check the requirements.
                    If GetPlayerGuildAccess(Index) < 1 Then Call PlayerMsg(Index, "Vous n'avez pas un acc�s suffisant.", Red): Exit Sub
                    
                    If FindPlayer(Parse(1)) = 0 Then Call PlayerMsg(Index, "Personnage hors-ligne", White): Exit Sub
                
                    If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then Call PlayerMsg(Index, "Le joueur n'est pas dans votre Guilde", Red): Exit Sub
                    
                    If GetPlayerGuild(Index) = vbNullString Then Call PlayerMsg(Index, "Vous n'�tes pas dans une guilde", Red): Exit Sub
                    
                    'Set the player's new access level
                    Call AddLog("Changement de l'acc�s de " & GetPlayerName(FindPlayer(Parse(1))) & " de " & GetPlayerGuildAccess(FindPlayer(Parse(1))) & " � " & Parse(2) & ".", GUILDE_LOG)
                    Call SetPlayerGuildAccess(FindPlayer(Parse(1)), Parse(2))
                    Call SendPlayerData(FindPlayer(Parse(1)))
                    Exit Sub
                
                ' Disown
                Case "guilddisown"
                    ' Check if all the requirements
                    If GetPlayerGuildAccess(Index) < 1 Then Call PlayerMsg(Index, "Vous n'avez pas un acc�s suffisant.", Red): Exit Sub
                    
                    If FindPlayer(Parse(1)) = 0 Then Call PlayerMsg(Index, "Personnage hors-ligne", White): Exit Sub
                            
                    If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then Call PlayerMsg(Index, "Le joueur n'est pas dans votre guilde", Red): Exit Sub
                    
                    If GetPlayerGuildAccess(FindPlayer(Parse(1))) > GetPlayerGuildAccess(Index) Then Call PlayerMsg(Index, "Le joueur a un privil�ge plus �lev� dans la guilde.", Red): Exit Sub
                            
                    'Player checks out, take him out of the guild
                    Call SetPlayerGuild(FindPlayer(Parse(1)), "")
                    Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
                    Call SendPlayerData(FindPlayer(Parse(1)))
                    Call AddLog(GetPlayerName(FindPlayer(Parse(1))) & " a �t� vir� de la guilde : " & GetPlayerGuild(Index) & ".", GUILDE_LOG)
                    Exit Sub
            
                ' Leave Guild
                Case "guildleave"
                    ' Check if they can leave
                    If GetPlayerGuild(Index) = vbNullString Then Call PlayerMsg(Index, "Tu n'es pas dans une guilde.", Red): Exit Sub
                    
                    Call SetPlayerGuild(Index, "")
                    Call SetPlayerGuildAccess(Index, 0)
                    Call SendPlayerData(Index)
                    Call AddLog(GetPlayerName(Index) & " a quitt� sa guilde.", GUILDE_LOG)
                    Exit Sub
                
                ' Make A Member
                Case "guildmember"
                    If GetPlayerGuildAccess(Index) < 1 Then Call PlayerMsg(Index, "Vous n'avez pas un acc�s suffisant.", Red): Exit Sub
                    
                    ' Check if its possible to admit the member
                    If FindPlayer(Parse(1)) = 0 Then Call PlayerMsg(Index, "Personnage hors-ligne", White): Exit Sub
                            
                    If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then Call PlayerMsg(Index, "Ce joueur n'est pas dans votre guilde", Red): Exit Sub
                    
                    If GetPlayerGuildAccess(FindPlayer(Parse(1))) > 1 Then Call PlayerMsg(Index, "Ce joueur est d�j� membre de votre guilde", Red): Exit Sub
                    
                    'All has gone well, set the guild access to 1
                    Call SetPlayerGuild(FindPlayer(Parse(1)), GetPlayerGuild(Index))
                    Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 1)
                    Call SendPlayerData(FindPlayer(Parse(1)))
                    Call AddLog("Recrutement de " & GetPlayerName(FindPlayer(Parse(1))) & " comme recruteur dans la guilde : " & GetPlayerGuild(Index) & ".", GUILDE_LOG)
                    Exit Sub
                
                ' Make A Trainie
                Case "guildtrainee"
                    'It is possible, so set the guild to index's guild, and the access level to 0
                    Call SetPlayerGuild(Val(Parse(1)), GetPlayerGuild(Val(Parse(2))))
                    Call SetPlayerGuildAccess(Val(Parse(1)), 0)
                    Call SendPlayerData(Val(Parse(1)))
                    Call AddLog("Recrutement de " & GetPlayerName(Val(Parse(1))) & " dans la guilde : " & GetPlayerGuild(Index) & ".", GUILDE_LOG)
                    Exit Sub
                                        
                Case "guildtraineevbyesno"
                    If GetPlayerGuildAccess(Index) < 1 Then Call PlayerMsg(Index, "Vous n'avez pas un acc�s suffisant.", Red): Exit Sub
                    ' Check if its possible to induct member
                    If FindPlayer(Parse(1)) = 0 Then Call PlayerMsg(Index, "Le joueur est Hors-ligne", White): Exit Sub
                    If GetPlayerGuild(FindPlayer(Parse(1))) <> vbNullString Then Call PlayerMsg(Index, "Le joueur est d�j� dans une guilde.", Red): Exit Sub
                    
                    Call SendDataTo(FindPlayer(Parse(1)), "guildtraineevbyesno" & SEP_CHAR & FindPlayer(Parse(1)) & SEP_CHAR & Index & END_CHAR)
                    Exit Sub
                    
                ' :: Social packets ::
                Case "saymsg"
                    Msg = Parse(1)
                    
                    ' Prevent hacking
                    If MMsg(Msg) Then Call HackingAttempt(Index, "Caract�re incorrecte dans ses paroles"): Exit Sub
                    
                    If frmServer.chkM.value = Unchecked And GetPlayerAccess(Index) <= 0 Then Call PlayerMsg(Index, "Les discutions ne sont pas autoris�s sur les cartes.", BrightRed): Exit Sub
                    
                    Call AddLog("Carte #" & GetPlayerMap(Index) & " : " & GetPlayerName(Index) & " : " & Msg & "", PLAYER_LOG)
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " : " & Msg & "", SayColor)
                    Call MapMsg2(GetPlayerMap(Index), Msg, Index)
                    TextAdd frmServer.txtText(3), GetPlayerName(Index) & " Sur la carte " & GetPlayerMap(Index) & ": " & Msg, True
                    Exit Sub
                    
                Case "guildemsg"
                   Msg = Parse(1)
                   
                   If Player(Index).Mute = True Then Exit Sub
                   
                   If GetPlayerGuild(Index) = vbNullString Then Call PlayerMsg(Index, "Tu n'es pas dans une guilde.", AlertColor): Exit Sub
                   
                   s = GetPlayerName(Index) & " (" & GetPlayerGuild(Index) & ") : " & Msg
                   Call AddLog(s, PLAYER_LOG)
                   
                   For i = 1 To MAX_PLAYERS
                       If GetPlayerGuild(Index) = GetPlayerGuild(i) Then Call PlayerMsg(i, s, CouleurDesGuilde)
                   Next i
                   Exit Sub
            
                Case "emotemsg"
                    Msg = Parse(1)
                    If MMsg(Msg) Then Call HackingAttempt(Index, "Caract�re incorrecte dans ses paroles(�moticons)"): Exit Sub
                    
                    If frmServer.chkE.value = Unchecked Then If GetPlayerAccess(Index) <= 0 Then Call PlayerMsg(Index, "Les �motes ont �t� d�sactiv�s.", BrightRed): Exit Sub
                    
                    Call AddLog("Carte #" & GetPlayerMap(Index) & " : " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Msg, EmoteColor)
                    TextAdd frmServer.txtText(6), GetPlayerName(Index) & " " & Msg, True
                    Exit Sub
             
                Case "broadcastmsg"
                    Msg = Parse(1)
                    ' Prevent hacking
                    If MMsg(Msg) Then Call HackingAttempt(Index, "Caract�re incorrecte dans ses paroles(global)"): Exit Sub
                    
                    If frmServer.chkBC.value = Unchecked Then If GetPlayerAccess(Index) <= 0 Then Call PlayerMsg(Index, "Les hurlement ont �t� d�sactiv�s.", BrightRed): Exit Sub
                    
                    If Player(Index).Mute = True Then Exit Sub
                    
                    s = GetPlayerName(Index) & " : " & Msg
                    Call AddLog(s, PLAYER_LOG)
                    Call GlobalMsg(s, BroadcastColor)
                    Call TextAdd(frmServer.txtText(0), s, True)
                    TextAdd frmServer.txtText(1), GetPlayerName(Index) & " : " & Msg, True
                    Exit Sub
                
                Case "globalmsg"
                    Msg = Parse(1)
                    ' Prevent hacking
                    If MMsg(Msg) Then Call HackingAttempt(Index, "Caract�re incorrects dans ses paroles(global)"): Exit Sub
                    
                    If frmServer.chkG.value = Unchecked Then If GetPlayerAccess(Index) <= 0 Then Call PlayerMsg(Index, "Les messages globaux ont �t� d�sactiv�s.", BrightRed): Exit Sub
                        
                    If Player(Index).Mute = True Then Exit Sub
                    
                    If GetPlayerAccess(Index) > 0 Then
                        s = "(global) " & GetPlayerName(Index) & ": " & Msg
                        Call AddLog(s, ADMIN_LOG)
                        Call GlobalMsg(s, GlobalColor)
                        Call TextAdd(frmServer.txtText(0), s, True)
                    End If
                    TextAdd frmServer.txtText(2), GetPlayerName(Index) & ": " & Msg, True
                    Exit Sub
                    
                Case "ouvrire"
                    If Val(Parse(1)) < 0 And Val(Parse(1)) > MAX_MAPX Then Call HackingAttempt(Index, "Position hors cartes"): Exit Sub
                    If Val(Parse(2)) < 0 And Val(Parse(2)) > MAX_MAPX Then Call HackingAttempt(Index, "Position hors cartes"): Exit Sub
                    
                    TempTile(GetPlayerMap(Index)).DoorOpen(Val(Parse(1)), Val(Parse(2))) = YES
                    TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                    Exit Sub
                    
                Case "dansinv"
                    SlotI = Val(Parse(1))
                    INum = Val(Parse(2))
                    IVal = Val(Parse(3))
                    IDur = Val(Parse(4))
                    Cslot = Val(Parse(5))
                            
                    'Prevent Hacking
                    If IsPlaying(Index) = False Then Call HackingAttempt(Index, "Le joueur n'est pas en train de jouer"): Exit Sub
                    If (SlotI > 24 Or SlotI < 1) And INum <> 0 Then Call HackingAttempt(Index, "Inv Slot Invalide"): Exit Sub
                    If (Cslot > 30 Or Cslot < 1) And INum <> 0 Then Call HackingAttempt(Index, "Slot de Coffre Invalide"): Exit Sub
                    If INum < 1 Or INum > MAX_ITEMS Then Call HackingAttempt(Index, "Num�ros d'objet Invalide"): Exit Sub
                    If Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemnum" & Cslot)) <> INum And INum <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont inventaire"): Exit Sub
                    If Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemval" & Cslot)) < (IVal - GetPlayerInvItemValue(Index, SlotI)) And IVal <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont inventaire"): Exit Sub
                    If Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemdur" & Cslot)) <> IDur And IDur <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont inventaire"): Exit Sub
                            
                    Call SetPlayerInvItemNum(Index, SlotI, INum)
                    Call SetPlayerInvItemValue(Index, SlotI, IVal)
                    Call SetPlayerInvItemDur(Index, SlotI, IDur)
                    Call SendInventoryUpdate(Index, SlotI)
                    Exit Sub
                
                Case "danscoffre"
                    SlotC = Val(Parse(1))
                    Cnum = Val(Parse(2))
                    Cval = Val(Parse(3))
                    Cdur = Val(Parse(4))
                    Islot = Val(Parse(5))
                           
                    'Prevent Hacking
                    If IsPlaying(Index) = False Then Call HackingAttempt(Index, "Le joueur n'ait pas entrin de jouer"): Exit Sub
                    If (Islot > 24 Or Islot < 1) And Cnum <> 0 Then Call HackingAttempt(Index, "Inv Slot Invalide"): Exit Sub
                    If (SlotC > 30 Or SlotC < 1) And Cnum <> 0 Then Call HackingAttempt(Index, "Slot de Coffre Invalide"): Exit Sub
                    If Cnum < 1 Or Cnum > MAX_ITEMS Then Call HackingAttempt(Index, "Num�ros d'objet Invalide"): Exit Sub
                    If GetPlayerInvItemNum(Index, Islot) <> Cnum And Cnum <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont coffre"): Exit Sub
                    If GetPlayerInvItemValue(Index, Islot) < (Cval - Val(GetVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemval" & SlotC))) And Cval <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont coffre"): Exit Sub
                    If GetPlayerInvItemDur(Index, Islot) <> Cdur And Cdur <> 0 Then Call HackingAttempt(Index, "Essaye de hacker sont coffre"): Exit Sub
                    
                    ' Make sure if the item we transfer to the bank is unequiped
                    ' Thanks to Xamus (Fontor), Tom13 and Revorn who found and told me about this possible hack
                    If GetPlayerWeaponSlot(Index) = Islot Then SetPlayerWeaponSlot Index, 0
                    If GetPlayerArmorSlot(Index) = Islot Then SetPlayerArmorSlot Index, 0
                    If GetPlayerHelmetSlot(Index) = Islot Then SetPlayerHelmetSlot Index, 0
                    If GetPlayerShieldSlot(Index) = Islot Then SetPlayerShieldSlot Index, 0
                    If GetPlayerPetSlot(Index) = Islot Then SetPlayerPetSlot Index, 0
                    
                    Call PutVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemnum" & SlotC, " " & Val(Cnum))
                    Call PutVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemval" & SlotC, " " & Val(Cval))
                    Call PutVar(App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini", "CHAR" & Player(Index).CharNum, "cofitemdur" & SlotC, " " & Val(Cdur))
                    Exit Sub
            
            
                Case "setsprite"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The sprite
                    n = Val(Parse(1))
                    Call SetPlayerSprite(Index, n)
                    Call SendPlayerData(Index)
                    Exit Sub
            
                Case "setplayersprite"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The sprite
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerSprite(i, n)
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� le sprite de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setname"
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                            
                    Call SetPlayerName(Index, Parse(1))
                    Call SendPlayerData(Index)
                    Exit Sub
            
                Case "setplayername"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerName(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� le nom de " & GetPlayerName(i) & "(nouveau nom)", ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setplayerstr"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerStr(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� la force de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
                    
                Case "setplayerdef"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerDEF(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� la d�fense de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setplayervit"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerSPEED(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� la vitesse de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setplayermagi"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerMAGI(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� la magie de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setplayerpk"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerPK(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� les PK de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setplayerexp"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerExp(i, Parse(2))
                        Call CheckPlayerLevelUp(i)
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� l'exp�rience de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setplayerniveau"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerLevel(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� le niveau de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setplayerpoint"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerPOINTS(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� les points de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setplayermaxpv"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerHP(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� les PV de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "setplayermaxpm"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    i = FindPlayer(Parse(1))
                    n = Val(Parse(2))
                    If i > 0 Then
                        Call SetPlayerMP(i, Parse(2))
                        Call SendPlayerData(i)
                        Call AddLog(GetPlayerName(Index) & " a chang� les PM de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "getstats"
                    Call PlayerMsg(Index, "-=- Statistiques de " & GetPlayerName(Index) & " -=-", White)
                    Call PlayerMsg(Index, "Niveau : " & GetPlayerLevel(Index) & "  Exp : " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
                    Call PlayerMsg(Index, "PV : " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index) & "  PM : " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index) & "  SP : " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index), White)
                    Call PlayerMsg(Index, "FOR : " & GetPlayerStr(Index) & "  DEF : " & GetPlayerDEF(Index) & "  MAGIE : " & GetPlayerMAGI(Index) & "  VIT : " & GetPlayerSPEED(Index), BrightGreen)
                    n = Int(GetPlayerStr(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
                    i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
                    z = Int(GetPlayerSPEED(Index) * 0.576)
                    If n > 100 Then n = 100
                    If i > 100 Then i = 100
                    If z > 100 Then z = 100
                    Call PlayerMsg(Index, "Chance de coup critique : " & n & "%, Chance de bloquer : " & i & "%, Chance d'esquive : " & z & "%", White)
                    Exit Sub
            
                Case "warpmeto"
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Name = Parse(1)
                    i = FindPlayer(Name)
                    If i > 0 Then
                        Call PlayerWarp(Index, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i))
                        Call PlayerMsg(i, GetPlayerName(Index) & " a �t� t�l�port� � cot� de toi.", White)
                        Call PlayerMsg(Index, "Vous avez �t� t�l�port� � cot� de " & GetPlayerName(i) & ".", White)
                        Call AddLog(GetPlayerName(Index) & " s'est t�l�port� � cot� de " & GetPlayerName(i), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "warptome"
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Name = Parse(1)
                    i = FindPlayer(Name)
                    If i > 0 Then
                        Call PlayerWarp(i, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                        Call PlayerMsg(i, "Vous avez �t� t�l�port� par " & GetPlayerName(Index) & ".", White)
                        Call PlayerMsg(Index, GetPlayerName(i) & " a �t� t�l�port�.", White)
                        Call AddLog(GetPlayerName(i) & " a �t� t�l�port� par " & GetPlayerName(Index), ADMIN_LOG)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "getotherstats"
                    Name = Parse(1)
                    i = FindPlayer(Name)
                    If i > 0 Then
                        Call PlayerMsg(Index, "-=- Statistique pour " & GetPlayerName(i) & " -=-", BrightGreen)
                        Call PlayerMsg(Index, "Niveau : " & GetPlayerLevel(i), BrightGreen)
                        n = Int(GetPlayerStr(i) / 2) + Int(GetPlayerLevel(i) / 2)
                        i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)
                        z = Int(GetPlayerSPEED(Index) * 0.576)
                        If n > 100 Then n = 100
                        If i > 100 Then i = 100
                        If z < 100 Then z = 100
                        Call PlayerMsg(Index, "Chance de coups critique : " & n & "%, Chance de bloquer : " & i & "%, Chance d'esquive : " & z & "%", BrightGreen)
                    Else
                        Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                    End If
                    Exit Sub
            
                Case "getadminhelp"
                    Call GlobalMsg(GetPlayerName(Index) & " a besoin d'un admin.", White)
                    If IBAdmin Then Call IBMsg(GetPlayerName(Index) & " a besoin d'un admin.", IBCJoueur)
                    Exit Sub
            
                Case "requestnewmap"
                    Dir = Val(Parse(1))
                    ' Prevent hacking
                    If Dir < DIR_DOWN Or Dir > DIR_UP Then Call HackingAttempt(Index, "Direction Invalide"): Exit Sub
                    Call PlayerMove(Index, Dir, 1)
                    Exit Sub
                Case "remplacemetier"
                    Player(Index).Char(Player(Index).CharNum).metier = Val(Parse(1))
                    Player(Index).Char(Player(Index).CharNum).MetierLvl = 1
                    Player(Index).Char(Player(Index).CharNum).MetierExp = 0
                    Call PlayerMsg(Index, "Vous avez oubli� votre m�tier puis appris un nouveau m�tier.", Green)
                    Exit Sub
                    
                Case "envmap"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "MAPENV" & SEP_CHAR & GetVar(App.Path & "\Data.ini", "FTP", "HOTE") & SEP_CHAR & GetVar(App.Path & "\Data.ini", "FTP", "REP") & END_CHAR)
                    Exit Sub
            
                Case "kickplayer"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) <= 0 Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The player index
                    n = FindPlayer(Parse(1))
                    If n <> Index Then
                        If n > 0 Then
                            If GetPlayerAccess(n) <= GetPlayerAccess(Index) Then
                                Call GlobalMsg(GetPlayerName(n) & " a �t� deconnect� de " & GAME_NAME & " par " & GetPlayerName(Index) & ".", White)
                                Call AddLog(GetPlayerName(Index) & " a d�connect�(kicker) " & GetPlayerName(n) & ".", ADMIN_LOG)
                                Call AlertMsg(n, "Vous avez �t� banni par " & GetPlayerName(Index) & ".")
                                If IBAdmin Then Call IBMsg(GetPlayerName(Index) & " a d�connect�(kicker) " & GetPlayerName(n) & ".", IBCAdmin)
                            Else
                                Call PlayerMsg(Index, "Cette personne poss�de un acc�s sup�rieur au votre.", White)
                            End If
                        Else
                            Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                        End If
                    Else
                        Call PlayerMsg(Index, "Tu ne peux te kicker toi m�me, imbecile.", White)
                    End If
                    Exit Sub
            
                Case "banlist"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = 1
                    f = FreeFile
                    Open App.Path & "\banlist.txt" For Input As #f
                    Do While Not EOF(f)
                        Input #f, s
                        Input #f, Name
                        Call PlayerMsg(Index, n & " : IP banni " & s & " par " & Name, White)
                        n = n + 1
                    Loop
                    Close #f
                    Exit Sub
            
                Case "bandestroy"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call Kill(App.Path & "\banlist.txt")
                    Call PlayerMsg(Index, "Liste des bannis effac�e.", White)
                    If IBAdmin Then Call IBMsg(GetPlayerName(Index) & " a d�truit la liste des bannis.", IBCAdmin)
                    Exit Sub
            
                Case "banplayer"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The player index
                    n = FindPlayer(Parse(1))
                    If n <> Index Then
                        If n > 0 Then
                            If GetPlayerAccess(n) <= GetPlayerAccess(Index) Then Call BanIndex(n, Index) Else Call PlayerMsg(Index, "Cette utilisateur a un acc�s sup�rieur au votre.", White)
                        Else
                            Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                        End If
                    Else
                        Call PlayerMsg(Index, "Tu ne peux te bannir toi m�me.", White)
                    End If
                    Exit Sub
            
                Case "requesteditmap"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "EDITMAP" & END_CHAR)
                    Exit Sub
            
                Case "requestedititem"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "ITEMEDITOR" & END_CHAR)
                    Exit Sub
            
                Case "edititem"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The item #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_ITEMS Then Call HackingAttempt(Index, "Index d'objet Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite l'objet #" & n & ".", ADMIN_LOG)
                    Call SendEditItemTo(Index, n)
                    Exit Sub
            
                Case "saveitem"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = Val(Parse(1))
                    If n < 0 Or n > MAX_ITEMS Then Call HackingAttempt(Index, "Index de l'objet Invalide"): Exit Sub
                    ' Update the item
                    item(n).Name = Parse(2)
                    item(n).Pic = Val(Parse(3))
                    item(n).type = Val(Parse(4))
                    item(n).data1 = Val(Parse(5))
                    item(n).data2 = Val(Parse(6))
                    item(n).data3 = Val(Parse(7))
                    item(n).StrReq = Val(Parse(8))
                    item(n).DefReq = Val(Parse(9))
                    item(n).SpeedReq = Val(Parse(10))
                    item(n).ClassReq = Val(Parse(11))
                    item(n).AccessReq = Val(Parse(12))
                    
                    item(n).AddHP = Val(Parse(13))
                    item(n).AddMP = Val(Parse(14))
                    item(n).AddSP = Val(Parse(15))
                    item(n).AddStr = Val(Parse(16))
                    item(n).AddDef = Val(Parse(17))
                    item(n).AddMagi = Val(Parse(18))
                    item(n).AddSpeed = Val(Parse(19))
                    item(n).AddEXP = Val(Parse(20))
                    item(n).desc = Parse(21)
                    item(n).AttackSpeed = Val(Parse(22))
                    item(n).NCoul = Val(Parse(23))
                    
                    item(n).paperdoll = Val(Parse(24))
                    item(n).paperdollPic = Val(Parse(25))
                    
                    item(n).Empilable = Val(Parse(26))
                    
                    item(n).Sex = Val(Parse(27))
                    item(n).tArme = Val(Parse(28))
                    item(n).LevelReq = Val(Parse(29))
                    
                    ' Sauvegarder l'objet
                    Call SendUpdateItemToAll(n)
                    Call SaveItem(n)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde l'objet #" & n & ".", ADMIN_LOG)
                    Exit Sub
                
                Case "requesteditpet"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "PETEDITOR" & END_CHAR)
                    Exit Sub
                
                Case "editpet"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The item #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_PETS Then Call HackingAttempt(Index, "Index d'objet Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite le famillier #" & n & ".", ADMIN_LOG)
                    Call SendEditPetTo(Index, n)
                    Exit Sub
                
                Case "savepet"
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = Val(Parse(1))
                    If n < 0 Or n > MAX_PETS Then Call HackingAttempt(Index, "Index du famillier Invalide"): Exit Sub
                    Pets(n).nom = Parse(2)
                    Pets(n).sprite = Val(Parse(3))
                    Pets(n).addForce = Val(Parse(4))
                    Pets(n).addDefence = Val(Parse(5))
                    
                    Call SendUpdatePetToAll(n)
                    Call SavePet(n)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde le famillier #" & n & ".", ADMIN_LOG)
                    Exit Sub
                    
                Case "requesteditmetier"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "metierEDITOR" & END_CHAR)
                    Exit Sub
                
                Case "editmetier"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The item #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_METIER Then Call HackingAttempt(Index, "Index d'objet Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite le metier #" & n & ".", ADMIN_LOG)
                    Call SendEditmetierTo(Index, n)
                    Exit Sub
                
                Case "savemetier"
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = Val(Parse(1))
                    If n < 0 Or n > MAX_METIER Then Call HackingAttempt(Index, "Index du metier Invalide"): Exit Sub
                    metier(n).nom = Parse(2)
                    metier(n).type = Val(Parse(3))
                    metier(n).desc = Parse(4)
                    x = 5
                    For i = 0 To MAX_DATA_METIER
                        For z = 0 To 1
                            metier(n).data(i, z) = Val(Parse(x))
                            x = x + 1
                        Next z
                    Next i
                    
                    Call SendUpdatemetierToAll(n)
                    Call SaveMetier(n)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde le metier #" & n & ".", ADMIN_LOG)
                    Exit Sub
                Case "requesteditrecette"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "recetteEDITOR" & END_CHAR)
                    Exit Sub
                
                Case "editrecette"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The item #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_RECETTE Then Call HackingAttempt(Index, "Index d'objet Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite le recette #" & n & ".", ADMIN_LOG)
                    Call SendEditrecetteTo(Index, n)
                    Exit Sub
                
                Case "saverecette"
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = Val(Parse(1))
                    If n < 0 Or n > MAX_RECETTE Then Call HackingAttempt(Index, "Index du recette Invalide"): Exit Sub
                    recette(n).nom = Parse(2)
                    x = 3
                    For i = 0 To 9
                        For z = 0 To 1
                            recette(n).InCraft(i, z) = Val(Parse(x))
                            x = x + 1
                        Next z
                    Next i
                    For z = 0 To 1
                        recette(n).craft(z) = Val(Parse(x))
                        x = x + 1
                    Next z
                    
                    Call SendUpdaterecetteToAll(n)
                    Call Saverecette(n)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde le recette #" & n & ".", ADMIN_LOG)
                    Exit Sub
                    
                Case "requesteditnpc"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "NPCEDITOR" & END_CHAR)
                    Exit Sub
            
                Case "editnpc"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The npc #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_NPCS Then Call HackingAttempt(Index, "Index du PNJ Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite le PNJ #" & n & ".", ADMIN_LOG)
                    Call SendEditNpcTo(Index, n)
                    Exit Sub
            
                Case "savenpc"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_NPCS Then Call HackingAttempt(Index, "Index du PNJ Invalide"): Exit Sub
                    ' Update the npc
                    Npc(n).Name = Parse(2)
                    Npc(n).AttackSay = Parse(3)
                    Npc(n).sprite = Val(Parse(4))
                    Npc(n).SpawnSecs = Val(Parse(5))
                    Npc(n).Behavior = Val(Parse(6))
                    Npc(n).Range = Val(Parse(7))
                    Npc(n).STR = Val(Parse(8))
                    Npc(n).def = Val(Parse(9))
                    Npc(n).Speed = Val(Parse(10))
                    Npc(n).magi = Val(Parse(11))
                    Npc(n).MaxHp = Val(Parse(12))
                    Npc(n).Exp = Val(Parse(13))
                    Npc(n).SpawnTime = Val(Parse(14))
                    Npc(n).QueteNum = Val(Parse(15))
                    Npc(n).Inv = Val(Parse(16))
                    Npc(n).Vol = Val(Parse(17))
                    
                    z = 18
                    For i = 1 To MAX_NPC_DROPS
                        Npc(n).ItemNPC(i).chance = Val(Parse(z))
                        Npc(n).ItemNPC(i).ItemNum = Val(Parse(z + 1))
                        Npc(n).ItemNPC(i).ItemValue = Val(Parse(z + 2))
                        z = z + 3
                    Next i
                    
                    For i = 1 To MAX_NPC_SPELLS
                        Npc(n).Spell(i) = Val(Parse(z))
                        z = z + 1
                    Next
                    
                    ' Sauvegarde du pnj
                    Call SendUpdateNpcToAll(n)
                    Call SaveNpc(n)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde le PNJ #" & n & ".", ADMIN_LOG)
                    Exit Sub
            
                Case "requesteditquetes"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "QUETEEDITOR" & END_CHAR)
                    Exit Sub
            
                Case "editquetes"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The spell #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_QUETES Then Call HackingAttempt(Index, "Indes de qete Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite la quete #" & n & ".", ADMIN_LOG)
                    Call SendEditQuetesTo(Index, n)
                    Exit Sub
            
                Case "demarequete"
                    
                    n = Val(Parse(1))
                    If n < 0 Or n > MAX_QUETES Then Call HackingAttempt(Index, "Index de qu�te invalide."): Exit Sub
                    Player(Index).Char(Player(Index).CharNum).QueteEnCour = Val(Parse(1))
                    If n = 0 Then Exit Sub
                    
                    If quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).type = QUETE_TYPE_APORT Then
                        i = FindOpenInvSlot(Index, quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).data1)
                        If i = 0 Then
                            Call PlayerMsg(Index, "Ton inventaire est plein, tu ne peux pas faire cette qu�te.", Red)
                            Player(Index).Char(Player(Index).CharNum).QueteEnCour = 0
                            Call SendDataTo(Index, "QUETECOUR" & SEP_CHAR & 0 & END_CHAR)
                            Exit Sub
                        End If
                        Call GiveItem(Index, quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).data1, 1)
                    End If
                    If quete(Val(Parse(1))).temps > 0 Then Call SendDataTo(Index, "TEMPSQUETE" & SEP_CHAR & quete(Val(Parse(1))).temps & END_CHAR)
                    Exit Sub
            
                Case "savequete"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' Spell #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_QUETES Then Call HackingAttempt(Index, "Indes de quete Invalide"): Exit Sub
                    ' Update the quete
                    quete(n).nom = Parse(2)
                    quete(n).data1 = Val(Parse(3))
                    quete(n).data2 = Val(Parse(4))
                    quete(n).data3 = Val(Parse(5))
                    quete(n).Description = Parse(6)
                    quete(n).reponse = Parse(7)
                    quete(n).String1 = Parse(8)
                    quete(n).temps = Val(Parse(9))
                    quete(n).type = Val(Parse(10))
                    
                    Dim l As Long
                    i = 10
                    For l = 1 To 15
                        i = i + 1
                        quete(n).indexe(l).data1 = Val(Parse(i))
                        i = i + 1
                        quete(n).indexe(l).data2 = Val(Parse(i))
                        i = i + 1
                        quete(n).indexe(l).data3 = Val(Parse(i))
                        i = i + 1
                        quete(n).indexe(l).String1 = Parse(i)
                    Next l
                    quete(n).Recompence.Exp = Val(Parse(i + 1))
                    quete(n).Recompence.objn1 = Val(Parse(i + 2))
                    quete(n).Recompence.objn2 = Val(Parse(i + 3))
                    quete(n).Recompence.objn3 = Val(Parse(i + 4))
                    quete(n).Recompence.objq1 = Val(Parse(i + 5))
                    quete(n).Recompence.objq2 = Val(Parse(i + 6))
                    quete(n).Recompence.objq3 = Val(Parse(i + 7))
                    quete(n).Case = Val(Parse(i + 8))
                    
                    ' Sauvegarde de la quete
                    Call SendUpdateQueteToAll(n)
                    Call SaveQuete(n)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde la quete #" & n & ".", ADMIN_LOG)
                    Exit Sub
            
                Case "requesteditshop"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "SHOPEDITOR" & END_CHAR)
                    Exit Sub
            
                Case "editshop"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The shop #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_SHOPS Then Call HackingAttempt(Index, "Index du magasin Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite le magasin #" & n & ".", ADMIN_LOG)
                    Call SendEditShopTo(Index, n)
                    Exit Sub
            
                Case "saveshop"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ShopNum = Val(Parse(1))
                    ' Prevent hacking
                    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then Call HackingAttempt(Index, "Index de magasin Invalide"): Exit Sub
                    ' Update the shop
                    Shop(ShopNum).Name = Parse(2)
                    Shop(ShopNum).JoinSay = Parse(3)
                    Shop(ShopNum).LeaveSay = Parse(4)
                    Shop(ShopNum).FixesItems = Val(Parse(5))
                    Shop(ShopNum).FixObjet = Val(Parse(6))
                    n = 7
                    For z = 1 To 6
                        For i = 1 To MAX_TRADES
                            Shop(ShopNum).TradeItem(z).value(i).GiveItem = Val(Parse(n))
                            Shop(ShopNum).TradeItem(z).value(i).GiveValue = Val(Parse(n + 1))
                            Shop(ShopNum).TradeItem(z).value(i).GetItem = Val(Parse(n + 2))
                            Shop(ShopNum).TradeItem(z).value(i).GetValue = Val(Parse(n + 3))
                            n = n + 4
                        Next i
                    Next z
                    
                    ' Save it
                    Call SendUpdateShopToAll(ShopNum)
                    Call SaveShop(ShopNum)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde le magasin #" & ShopNum & ".", ADMIN_LOG)
                    Exit Sub
            
                Case "requesteditspell"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "SPELLEDITOR" & END_CHAR)
                    Exit Sub
            
                Case "editspell"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' The spell #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_SPELLS Then Call HackingAttempt(Index, "Indes de sort Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite le sorrt #" & n & ".", ADMIN_LOG)
                    Call SendEditSpellTo(Index, n)
                    Exit Sub
            
                Case "savespell"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    ' Spell #
                    n = Val(Parse(1))
                    ' Prevent hacking
                    If n < 0 Or n > MAX_SPELLS Then Call HackingAttempt(Index, "Invalide Spell Index"): Exit Sub
                    ' Update the spell
                    Spell(n).Name = Parse(2)
                    Spell(n).ClassReq = Val(Parse(3))
                    Spell(n).LevelReq = Val(Parse(4))
                    Spell(n).type = Val(Parse(5))
                    Spell(n).data1 = Val(Parse(6))
                    Spell(n).data2 = Val(Parse(7))
                    Spell(n).data3 = Val(Parse(8))
                    Spell(n).MPCost = Val(Parse(9))
                    Spell(n).Sound = Val(Parse(10))
                    Spell(n).Range = Val(Parse(11))
                    Spell(n).SpellAnim = Val(Parse(12))
                    Spell(n).SpellTime = Val(Parse(13))
                    Spell(n).SpellDone = Val(Parse(14))
                    Spell(n).AE = Val(Parse(15))
                    Spell(n).Big = Val(Parse(16))
                    Spell(n).SpellIco = Val(Parse(17))
                    
                    ' Sauvegarde su sort
                    Call SendUpdateSpellToAll(n)
                    Call SaveSpell(n)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde le sort #" & n & ".", ADMIN_LOG)
                    Exit Sub
            
                Case "setaccess"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Call HackingAttempt(Index, "Essaye d'utiliser des pouvoirs qu'il n'a pas"): Exit Sub
                    ' The index
                    n = FindPlayer(Parse(1))
                    If n < 1 Or n > MAX_PLAYERS Then Exit Sub
                    ' The access
                    i = Val(Parse(2))
                    ' Check for invalid access level
                    If i >= 0 Or i <= 3 Then
                        If GetPlayerName(Index) <> GetPlayerName(n) Then
                            If GetPlayerAccess(Index) > GetPlayerAccess(n) Then
                                ' Check if player is on
                                If n > 0 Then
                                    If GetPlayerAccess(n) <= 0 Then Call GlobalMsg(GetPlayerName(n) & " est devenu mod�rateur.", BrightBlue)
                                    Call SetPlayerAccess(n, i)
                                    Call SendPlayerData(n)
                                    Call AddLog(GetPlayerName(Index) & " a modifi�(e) l'acc�s de " & GetPlayerName(n) & ".", ADMIN_LOG)
                                Else
                                    Call PlayerMsg(Index, "Personnage hors-ligne.", White)
                                End If
                            Else
                                Call PlayerMsg(Index, "Votre acc�s est plus bas que " & GetPlayerName(n) & ".", Red)
                            End If
                        Else
                            Call PlayerMsg(Index, "Tu ne peux changer ton acc�s.", Red)
                        End If
                    Else
                        Call PlayerMsg(Index, "Niveau d'acc�s invalide.", Red)
                    End If
                    Exit Sub
                
                Case "whosonline"
                    Call SendWhosOnline(Index)
                    Exit Sub
            
                Case "onlinelist"
                    Call SendOnlineList
                    Exit Sub
            
                Case "setmotd"
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse(1))
                    Call GlobalMsg("Mot de bienvenue remplac� par : " & Parse(1), BrightCyan)
                    Call AddLog(GetPlayerName(Index) & " a chang� le mot de bienvenue par : " & Parse(1), ADMIN_LOG)
                    If IBAdmin Then Call IBMsg(GetPlayerName(Index) & " a chang� le mot de bienvenue.", IBCAdmin)
                    Exit Sub
                
                Case "leaveshop"
                    Call QueteMsg(Index, Shop(Parse(1)).LeaveSay)
                    Exit Sub
                
                Case "traderequest"
                    ' Trade num
                    n = Val(Parse(1))
                    z = Val(Parse(2))
                    ' Prevent hacking
                    If (n < 1) Or (n > 6) Then Call HackingAttempt(Index, "Modification d'une requet d'�change"): Exit Sub
                    If (z <= 0) Or (z > (MAX_TRADES * 6)) Then Call HackingAttempt(Index, "Modification d'une requet d'�change"): Exit Sub
                    ' Index for shop
                    i = Player(Index).Char(Player(Index).CharNum).vendeur
                    ' Check if inv full
                    If i <= 0 Then Exit Sub
                    x = FindOpenInvSlot(Index, Shop(i).TradeItem(n).value(z).GetItem)
                    If x = 0 Then Call PlayerMsg(Index, "L'�change a �chou�, Inventaire pleins.", BrightRed): Exit Sub
                    ' Check if they have the item
                    If HasItem(Index, Shop(i).TradeItem(n).value(z).GiveItem) >= Shop(i).TradeItem(n).value(z).GiveValue Then
                        Call TakeItem(Index, Shop(i).TradeItem(n).value(z).GiveItem, Shop(i).TradeItem(n).value(z).GiveValue)
                        Call GiveItem(Index, Shop(i).TradeItem(n).value(z).GetItem, Shop(i).TradeItem(n).value(z).GetValue)
                        Call PlayerMsg(Index, "Echange r�ussit.", Yellow)
                        If Player(Index).Char(Player(Index).CharNum).QueteEnCour > 0 Then
                            If quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).type = QUETE_TYPE_RECUP Then
                                Call PlayerQueteTypeRecup(Index, Player(Index).Char(Player(Index).CharNum).QueteEnCour, Shop(i).TradeItem(n).value(z).GetItem, Shop(i).TradeItem(n).value(z).GetValue)
                            End If
                        End If
                    Else
                        Call PlayerMsg(Index, "Vous n'avez pas l'objet demand�.", BrightRed)
                    End If
                    Exit Sub
                Case "vendrerequest"
                    ' Trade num
                    n = Val(Parse(1))
                    z = Val(Parse(2))
                    ' Prevent hacking
                    If (n < 1) Or (n > 6) Then Call HackingAttempt(Index, "Modification d'une requet d'�change"): Exit Sub
                    If (z <= 0) Or (z > (MAX_TRADES * 6)) Then Call HackingAttempt(Index, "Modification d'une requet d'�change"): Exit Sub
                    ' Index for shop
                    i = Player(Index).Char(Player(Index).CharNum).vendeur
                    ' Check if inv full
                    If i <= 0 Then Exit Sub
                    x = FindOpenInvSlot(Index, Shop(i).TradeItem(n).value(z).GiveItem) 'Shop(i).TradeItem(N).value(z).GetItem)
                    If x = 0 Then Call PlayerMsg(Index, "L'�change a �chou�, Inventaire pleins.", BrightRed): Exit Sub
                    ' Check if they have the item
                    If HasItem(Index, Shop(i).TradeItem(n).value(z).GetItem) >= Shop(i).TradeItem(n).value(z).GetValue Then
                        Call GiveItem(Index, Shop(i).TradeItem(n).value(z).GiveItem, Math.Round(Shop(i).TradeItem(n).value(z).GiveValue / 2))
                        Call TakeItem(Index, Shop(i).TradeItem(n).value(z).GetItem, Shop(i).TradeItem(n).value(z).GetValue)
                        Call PlayerMsg(Index, "Echange r�ussit.", Yellow)
                        If Player(Index).Char(Player(Index).CharNum).QueteEnCour > 0 Then
                            If quete(Player(Index).Char(Player(Index).CharNum).QueteEnCour).type = QUETE_TYPE_RECUP Then
                                Call PlayerQueteTypeRecup(Index, Player(Index).Char(Player(Index).CharNum).QueteEnCour, Shop(i).TradeItem(n).value(z).GetItem, Shop(i).TradeItem(n).value(z).GetValue)
                            End If
                        End If
                    Else
                        Call PlayerMsg(Index, "Vous n'avez pas l'objet demand�.", BrightRed)
                    End If
                    Exit Sub
                Case "fixitem"
                    Dim d As Currency
                    ' Inv num
                    n = Val(Parse(1))
                    ' Make sure its a equipable item
                    If item(GetPlayerInvItemNum(Index, n)).type < ITEM_TYPE_WEAPON Or item(GetPlayerInvItemNum(Index, n)).type > ITEM_TYPE_SHIELD Then
                        Call PlainMsg(Index, "Tu peux seulement r�parer les armes, armure, casque et bouclier.", 6)
                        Exit Sub
                    End If
                    ' Now check the rate of pay
                    ItemNum = GetPlayerInvItemNum(Index, n)
                    d = item(GetPlayerInvItemNum(Index, n)).data2 / 5
                    DurNeeded = item(ItemNum).data1 - GetPlayerInvItemDur(Index, n)
                    GoldNeeded = (DurNeeded * d \ 2)
                    If GoldNeeded <= 0 Then GoldNeeded = 1
                    
                    ' Check if they even need it repaired
                    If DurNeeded <= 0 Then Call PlainMsg(Index, "Cette objets est en parfait �tat.", 6): Exit Sub
                            
                    ' Check if they have enough for at least one point
                    If HasItem(Index, Val(Parse(2))) >= d Then
                        ' Check if they have enough for a total restoration
                        If HasItem(Index, Val(Parse(2))) >= GoldNeeded Then
                            Call TakeItem(Index, Val(Parse(2)), GoldNeeded)
                            Call SetPlayerInvItemDur(Index, n, item(ItemNum).data1)
                            Call PlainMsg(Index, "Cette objet a totalement �t� r�par� pour " & GoldNeeded & Trim$(item(Val(Parse(2))).Name), 6)
                        Else
                            ' They dont so restore as much as we can
                            DurNeeded = (HasItem(Index, Val(Parse(2))) \ d)
                            GoldNeeded = Int(DurNeeded * d \ 2)
                            If GoldNeeded <= 0 Then GoldNeeded = 1
                            Call TakeItem(Index, Val(Parse(2)), GoldNeeded)
                            Call SetPlayerInvItemDur(Index, n, GetPlayerInvItemDur(Index, n) + DurNeeded)
                            Call PlainMsg(Index, "Cette objet a �t� r�par� pour " & GoldNeeded & Trim$(item(Val(Parse(2))).Name), 6)
                        End If
                    Else
                        Call PlainMsg(Index, "Pas assez de " & Trim$(item(Val(Parse(2))).Name) & " pour r�parer cet objet.", 6)
                    End If
                    Call SendInventory(Index)
                    Exit Sub
            
                Case "search"
                    x = Val(Parse(1))
                    y = Val(Parse(2))
                    ' Prevent subscript out of range
                    If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then Exit Sub
                    ' Check for a player
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) And GetPlayerX(i) = x And GetPlayerY(i) = y Then
                            ' Consider the player
                            If GetPlayerLevel(i) >= GetPlayerLevel(Index) + 5 Then
                                Call PlayerMsg(Index, "Vos chance semble minime...", BrightRed)
                            Else
                                If GetPlayerLevel(i) > GetPlayerLevel(Index) Then
                                    Call PlayerMsg(Index, "Ce joueur semble avoir une force que vous ne poss�dez pas.", Yellow)
                                Else
                                    If GetPlayerLevel(i) = GetPlayerLevel(Index) Then
                                        Call PlayerMsg(Index, "Cela risque d'�tre un combat m�morable.", White)
                                    Else
                                        If GetPlayerLevel(Index) >= GetPlayerLevel(i) + 5 Then
                                            Call PlayerMsg(Index, "Tu peux facilement tuer ce joueur.", BrightBlue)
                                        Else
                                            If GetPlayerLevel(Index) > GetPlayerLevel(i) Then Call PlayerMsg(Index, "Vous avez un avantage sur ce joueur.", Yellow)
                                        End If
                                    End If
                                End If
                            End If
                            ' Change target
                            Player(Index).Target = i
                            Player(Index).TargetType = TARGET_TYPE_PLAYER
                            Call PlayerMsg(Index, "Votre cible est maintenant " & GetPlayerName(i) & ".", Yellow)
                            Exit Sub
                        End If
                    Next i
                    
                    ' Check for an npc
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(GetPlayerMap(Index), i).Num > 0 Then
                            If MapNpc(GetPlayerMap(Index), i).x = x And MapNpc(GetPlayerMap(Index), i).y = y Then
                                ' Change target
                                Player(Index).Target = i
                                Player(Index).TargetType = TARGET_TYPE_NPC
                                Call PlayerMsg(Index, "Votre cible est maintenant " & Trim$(Npc(MapNpc(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                                Exit Sub
                            End If
                        End If
                    Next i
                    
                    ' Check for an item
                    For i = 1 To MAX_MAP_ITEMS
                        If MapItem(GetPlayerMap(Index), i).Num > 0 Then
                            If MapItem(GetPlayerMap(Index), i).x = x And MapItem(GetPlayerMap(Index), i).y = y Then
                                Call PlayerMsg(Index, "Vous voyez un " & Trim$(item(MapItem(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                                Exit Sub
                            End If
                        End If
                    Next i
                    Exit Sub
                
                Case "dchat"
                    n = Player(Index).ChatPlayer
                    If n < 1 Then Call PlayerMsg(Index, "Aucune requ�te pour discuter avec vous.", Pink): Exit Sub
                    
                    Call PlayerMsg(Index, "Vous declinez la requ�te de chat.", Pink)
                    Call PlayerMsg(n, GetPlayerName(Index) & " refuse votre demande.", Pink)
                    
                    Player(Index).ChatPlayer = 0
                    Player(Index).InChat = 0
                    Player(n).ChatPlayer = 0
                    Player(n).InChat = 0
                    Exit Sub
            
                Case "qchat"
                    n = Player(Index).ChatPlayer
                    If n < 1 Then Call PlayerMsg(Index, "Aucune requ�te pour discuter avec vous.", Pink): Exit Sub
                    
                    Call SendDataTo(Index, "qchat" & END_CHAR)
                    Call SendDataTo(n, "qchat" & END_CHAR)
                    
                    Player(Index).ChatPlayer = 0
                    Player(Index).InChat = 0
                    Player(n).ChatPlayer = 0
                    Player(n).InChat = 0
                    Exit Sub
                
                Case "sendchat"
                    n = Player(Index).ChatPlayer
                    If n < 1 Then Call PlayerMsg(Index, "Aucune requ�te pour discuter avec vous.", Pink): Exit Sub
            
                    Call SendDataTo(n, "sendchat" & SEP_CHAR & Parse(1) & SEP_CHAR & Index & END_CHAR)
                    Exit Sub
            
                Case "qtrade"
                    n = Player(Index).TradePlayer
                    ' Check if anyone trade with player
                    If n < 1 Then Call PlayerMsg(Index, "Aucune requ�te pour �changer avec vous.", Pink): Exit Sub
                    Call PlayerMsg(Index, "Arr�t de l'�change.", Pink)
                    Call PlayerMsg(n, GetPlayerName(Index) & " a arr�t� d'�changer avec vous.", Pink)
                    Player(Index).TradeOk = 0
                    Player(n).TradeOk = 0
                    Player(Index).TradePlayer = 0
                    Player(Index).InTrade = 0
                    Player(n).TradePlayer = 0
                    Player(n).InTrade = 0
                    Call SendDataTo(Index, "qtrade" & END_CHAR)
                    Call SendDataTo(n, "qtrade" & END_CHAR)
                    Exit Sub
            
                Case "dtrade"
                    n = Player(Index).TradePlayer
                    ' Check if anyone trade with player
                    If n < 1 Then Call PlayerMsg(Index, "Personne ne veut �changer avec vous.", Pink): Exit Sub
                    Call PlayerMsg(Index, "Refus de la requ�te.", Pink)
                    Call PlayerMsg(n, GetPlayerName(Index) & " refuse ta requ�te.", Pink)
                    Player(Index).TradePlayer = 0
                    Player(Index).InTrade = 0
                    Player(n).TradePlayer = 0
                    Player(n).InTrade = 0
                    Exit Sub
            
                Case "updatetradeinv"
                    n = Val(Parse(1))
                    Player(Index).Trading(n).InvNum = Val(Parse(2))
                    Player(Index).Trading(n).InvName = Trim$(Parse(3))
                    Player(Index).Trading(n).InvVal = Val(Parse(4))
                    If Player(Index).Trading(n).InvNum = 0 Then
                        Player(Index).TradeItemMax = Player(Index).TradeItemMax - 1
                        Player(Index).TradeOk = 0
                        Player(n).TradeOk = 0
                        Call SendDataTo(Index, "trading" & SEP_CHAR & 0 & END_CHAR)
                        Call SendDataTo(n, "trading" & SEP_CHAR & 0 & END_CHAR)
                    Else
                        Player(Index).TradeItemMax = Player(Index).TradeItemMax + 1
                    End If
                    Call SendDataTo(Player(Index).TradePlayer, "updatetradeitem" & SEP_CHAR & n & SEP_CHAR & Player(Index).Trading(n).InvNum & SEP_CHAR & Player(Index).Trading(n).InvName & SEP_CHAR & Player(Index).Trading(n).InvVal & END_CHAR)
                    Exit Sub
                
                Case "swapitems"
                    n = Player(Index).TradePlayer
                    If Player(Index).TradeOk = 0 Then
                        Player(Index).TradeOk = 1
                        Call SendDataTo(n, "trading" & SEP_CHAR & 1 & END_CHAR)
                    ElseIf Player(Index).TradeOk = 1 Then
                        Player(Index).TradeOk = 0
                        Call SendDataTo(n, "trading" & SEP_CHAR & 0 & END_CHAR)
                    End If
                    If Player(Index).TradeOk = 1 And Player(n).TradeOk = 1 Then
                        Player(Index).TradeItemMax2 = 0
                        Player(n).TradeItemMax2 = 0
                        For i = 1 To MAX_INV
                            If Player(Index).TradeItemMax = Player(Index).TradeItemMax2 Then Exit For
                            If GetPlayerInvItemNum(n, i) < 1 Then Player(Index).TradeItemMax2 = Player(Index).TradeItemMax2 + 1
                        Next i
            
                        For i = 1 To MAX_INV
                            If Player(n).TradeItemMax = Player(n).TradeItemMax2 Then Exit For
                            If GetPlayerInvItemNum(Index, i) < 1 Then Player(n).TradeItemMax2 = Player(n).TradeItemMax2 + 1
                        Next i
                        
                        If Player(Index).TradeItemMax2 = Player(Index).TradeItemMax And Player(n).TradeItemMax2 = Player(n).TradeItemMax Then
                            For i = 1 To MAX_PLAYER_TRADES
                                For x = 1 To MAX_INV
                                    If GetPlayerInvItemNum(n, x) < 1 Then
                                        If Player(Index).Trading(i).InvNum > 0 Then
                                            Call GiveItem(n, GetPlayerInvItemNum(Index, Player(Index).Trading(i).InvNum), Player(Index).Trading(i).InvVal)
                                            Call TakeItem(Index, GetPlayerInvItemNum(Index, Player(Index).Trading(i).InvNum), Player(Index).Trading(i).InvVal)
                                            Exit For
                                        End If
                                    End If
                                Next x
                            Next i
            
                            For i = 1 To MAX_PLAYER_TRADES
                                For x = 1 To MAX_INV
                                    If GetPlayerInvItemNum(Index, x) < 1 Then
                                        If Player(n).Trading(i).InvNum > 0 Then
                                            Call GiveItem(Index, GetPlayerInvItemNum(n, Player(n).Trading(i).InvNum), Player(n).Trading(i).InvVal)
                                            Call TakeItem(n, GetPlayerInvItemNum(n, Player(n).Trading(i).InvNum), Player(n).Trading(i).InvVal)
                                            Exit For
                                        End If
                                    End If
                                Next x
                            Next i
                            Call PlayerMsg(n, "Echange r�ussit.", BrightGreen)
                            Call PlayerMsg(Index, "Echange r�ussit.", BrightGreen)
                            Call SendInventory(n)
                            Call SendInventory(Index)
                        Else
                            If Player(Index).TradeItemMax2 < Player(Index).TradeItemMax Then
                                Call PlayerMsg(Index, "Votre inventaire est plein.", BrightRed)
                                Call PlayerMsg(n, "L'inventaire de " & GetPlayerName(n) & " est plein.", BrightRed)
                            ElseIf Player(n).TradeItemMax2 < Player(n).TradeItemMax Then
                                Call PlayerMsg(n, "Votre inventaire est pleins.", BrightRed)
                                Call PlayerMsg(Index, "L'inventaire de " & GetPlayerName(n) & " est plein.", BrightRed)
                            End If
                        End If
                        
                        Player(Index).TradePlayer = 0
                        Player(Index).InTrade = 0
                        Player(Index).TradeOk = 0
                        Player(n).TradePlayer = 0
                        Player(n).InTrade = 0
                        Player(n).TradeOk = 0
                        Call SendDataTo(Index, "qtrade" & END_CHAR)
                        Call SendDataTo(n, "qtrade" & END_CHAR)
                    End If
                    Exit Sub
            
                Case "joinparty"
                    n = Player(Index).InvitedBy
                   
                    If n > 0 Then
                        ' Check to make sure they aren't the starter
                            ' Check to make sure that each of there party players match
                        Call PlayerMsg(Index, "Tu as rejoins le groupe de " & GetPlayerName(n) & " .", Pink)
                        If Player(n).InParty = 0 Then ' Set the party leader up
                            Party.CreateParty n, Index
                        Else
                            Party.AddMember Player(n).InParty, Index
                        End If
                        
                        For i = 1 To Party.MemberCount(Player(n).InParty) - 1
                            Call PlayerMsg(Party.PlayerIndex(Player(n).InParty, i), GetPlayerName(Index) & " a rejoint votre groupe.", Pink)
                        Next i
                    Else
                        Call PlayerMsg(Index, "Tu n'as pas �t� invit� dans un groupe.", Pink)
                    End If
                    Exit Sub
            
                Case "leaveparty"
                    n = Player(Index).InvitedBy
                   
                    If Player(Index).InParty > 0 Then
                        'If Party.PlayerIndex(Player(Index).InParty, Party.Leader(Player(Index).InParty)) = Index Then Exit Sub
                        Call PlayerMsg(Index, "Tu as quitter le groupe.", Pink)
                        For i = 1 To Party.MemberCount(Player(Index).InParty)
                            If i <> Player(Index).PartyPlayer Then Call PlayerMsg(Party.PlayerIndex(Player(Index).InParty, i), GetPlayerName(Index) & " a quitt� le groupe.", Pink)
                        Next i
                        Party.RemoveMember Player(Index).InParty, Player(Index).PartyPlayer
                    ElseIf n > 0 Then
                        Call PlayerMsg(Index, "Tu refuse la demande de groupe.", Pink)
                        Call PlayerMsg(n, GetPlayerName(Index) & " refuse la demande de groupe.", Pink)
                        Player(Index).InParty = 0
                        Player(Index).InvitedBy = 0
                    Else
                        Call PlayerMsg(Index, "Vous n'�tes pas dans un groupe.", Pink)
                    End If
                    Exit Sub
                
                Case "spells"
                    Call SendPlayerSpells(Index)
                    Exit Sub
                            
                Case "requestlocation"
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call PlayerMsg(Index, "Carte : " & GetPlayerMap(Index) & ", X : " & GetPlayerX(Index) & ", Y : " & GetPlayerY(Index), Pink)
                    Exit Sub
                
                Case "refresh"
                    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                    Call ContrOnOff(Index)
                    Exit Sub
                
                Case "buysprite"
                    ' Check if player stepped on sprite changing tile
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type <> TILE_TYPE_SPRITE_CHANGE Then Call PlayerMsg(Index, "Tu as besoin d'�tre sur la case de sprite pour faire ceci.", BrightRed): Exit Sub
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2 = 0 Then
                        Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1)
                        Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                        Exit Sub
                    End If
                    
                    For i = 1 To MAX_INV
                        If GetPlayerInvItemNum(Index, i) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data2 Then
                            If item(GetPlayerInvItemNum(Index, i)).type = ITEM_TYPE_CURRENCY Then
                                If GetPlayerInvItemValue(Index, i) >= Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data3 Then
                                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data3)
                                    If GetPlayerInvItemValue(Index, i) <= 0 Then Call SetPlayerInvItemNum(Index, i, 0)
                                    Call PlayerMsg(Index, "Tu as un nouveau sprite.", BrightGreen)
                                    Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1)
                                    Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                                    Call SendInventory(Index)
                                End If
                            Else
                                If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i And GetPlayerPetSlot(Index) Then
                                    Call SetPlayerInvItemNum(Index, i, 0)
                                    Call SetPlayerInvItemValue(Index, i, 0)
                                    Call PlayerMsg(Index, "Tu as un nouveau sprite.", BrightGreen)
                                    Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).data1)
                                    Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                                    Call SendInventory(Index)
                                End If
                            End If
                            If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i And GetPlayerPetSlot(Index) <> i Then Exit Sub
                        End If
                    Next i
                    
                    Call PlayerMsg(Index, "Tu ne poss�des pas le n�cessaire.", BrightRed)
                    Exit Sub
                            
                Case "requesteditarrow"
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "arrowEDITOR" & END_CHAR)
                    Exit Sub
            
                Case "editarrow"
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = Val(Parse(1))
                    If n < 0 Or n > MAX_ARROWS Then Call HackingAttempt(Index, "Index de fl�che Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite la fl�che #" & n & ".", ADMIN_LOG)
                    Call SendEditArrowTo(Index, n)
                    Exit Sub
            
                Case "savearrow"
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = Val(Parse(1))
                    If n < 0 Or n > MAX_ITEMS Then Call HackingAttempt(Index, "Index de fl�che Invalide"): Exit Sub
                    Arrows(n).Name = Parse(2)
                    Arrows(n).Pic = Val(Parse(3))
                    Arrows(n).Range = Val(Parse(4))
            
                    Call SendUpdateArrowToAll(n)
                    Call SaveArrow(n)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde la fl�che #" & n & ".", ADMIN_LOG)
                    Exit Sub
                    
                Case "requesteditemoticon"
                    ' Prevent hacking
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call SendDataTo(Index, "EMOTICONEDITOR" & END_CHAR)
                    Exit Sub
            
                Case "editemoticon"
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = Val(Parse(1))
                    If n < 0 Or n > MAX_EMOTICONS Then Call HackingAttempt(Index, "Index d'�motic�ne Invalide"): Exit Sub
                    Call AddLog(GetPlayerName(Index) & " edite l'�motic�ne #" & n & ".", ADMIN_LOG)
                    Call SendEditEmoticonTo(Index, n)
                    Exit Sub
                Case "exscript"
                    If Val(Scripting) = 1 Then
                        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Parse(2)
                    End If
                    Exit Sub
                Case "saveemoticon"
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    n = Val(Parse(1))
                    If n < 0 Or n > MAX_ITEMS Then Call HackingAttempt(Index, "Index d'�motic�ne Invalide"): Exit Sub
                    Emoticons(n).Command = Parse(2)
                    Emoticons(n).Pic = Val(Parse(3))
                    Call SendUpdateEmoticonToAll(n)
                    Call SaveEmoticon(n)
                    Call AddLog(GetPlayerName(Index) & " sauvegarde l'�motic�ne #" & n & ".", ADMIN_LOG)
                    Exit Sub
                    
                Case "gmtime"
                    ' Merci � Xamus (Fontor), Tom13 et Revorn qui m'ont inform� de ce possible hack
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    GameTime = Val(Parse(1))
                    Call SendTimeToAll
                    Exit Sub
                    
                Case "weather"
                    ' MErci � Xamus (Fontor), Tom13 et Revorn qui m'ont inform� de ce possible hack
                    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    GameWeather = Val(Parse(1))
                    Call SendWeatherToAll
                    Exit Sub
                    
                Case "warpto"
                    ' Merci � Xamus (Fontor), Tom13 and Revorn qui m'ont inform� de ce possible hack
                    If GetPlayerAccess(Index) < ADMIN_MAPPER Then Call HackingAttempt(Index, "Clonage d'Admin"): Exit Sub
                    Call PlayerWarp(Index, Val(Parse(1)), GetPlayerX(Index), GetPlayerY(Index))
                    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_COFFRE Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_SIGN Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).type = TILE_TYPE_BLOCK_TOIT Then Call Debloque(Index)
                    Exit Sub
                    
                Case "sync"
                Player(Index).sync = True
                Exit Sub
            
            End Select
            
    End Select

Call HackingAttempt(Index, "Erreur : Probl�me dans l'envoi d'un packet (" & Parse(0) & ")")
Exit Sub
er:
Call AddLog("le : " & Date & "     � : " & time & "...Erreur dans la r�ception du serveur. D�tails : Num :" & Err.Number & " Description : " & Err.Description & " Source : " & Err.Source & "...", "logs\Err.txt")
On Error Resume Next
If IBErr Then Call IBMsg("Un erreur c'est produite dans la r�ception du serveur", BrightRed)
If Not IsPlaying(Index) Then Call PlainMsg(Index, "Erreur d'envoie, relancez s'il vous plait.", 3)
End Sub

Sub MapDo(ByVal z As Long, ByVal url As String, ByVal rep As String)
If FileExist("\maps\map" & z & ".fcc") Then Call Kill(App.Path & "\maps\map" & z & ".fcc")
Call ClearMap(z)
If Mid(url, Len(url)) = "/" And rep = "/" Then
    Call DeleteUrlCacheEntry(url & "map" & z & ".fcc")
    Call URLDownloadToFile(0, url & "map" & z & ".fcc", App.Path & "\maps\map" & z & ".fcc", 0, 0)
ElseIf Mid(url, Len(url)) <> "/" And Mid(rep, 1, 1) = "/" Then
    Call DeleteUrlCacheEntry(url & rep & "map" & z & ".fcc")
    Call URLDownloadToFile(0, url & rep & "map" & z & ".fcc", App.Path & "\maps\map" & z & ".fcc", 0, 0)
Else
    Call DeleteUrlCacheEntry(url & rep & "map" & z & ".fcc")
    Call URLDownloadToFile(0, url & rep & "map" & z & ".fcc", App.Path & "\maps\map" & z & ".fcc", 0, 0)
End If
End Sub

Sub CloseSocket(ByVal Index As Long, Optional ByVal Bypass As Boolean = False)
On Error Resume Next
    ' Make sure player was/is playing the game, and if so, save'm.
    If Index > 0 Then
        If Player(Index).sync = False Then
        Call LeftGame(Index, Bypass)
        frmServer.Socket(Index).Close
        Player(Index).Login = vbNullString
        Player(Index).InGame = False
        Call UpdateCaption
        Call TextAdd(frmServer.txtText(0), "Connexion de " & GetPlayerIP(Index) & " est termin�e.", True)
        End If
    End If
        
End Sub

Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim n As Long, i As Long

    s = vbNullString
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> Index Then
            s = s & GetPlayerName(i) & ", "
            n = n + 1
        End If
    Next i
            
    If n = 0 Then
        s = "Il n'y a pas d'autres joueurs connect�..."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "Il y a " & n & " joueur(s) en ligne : " & s & "."
    End If
        
    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub SendOnlineList()
Dim Packet As String
Dim i As Long
Dim n As Long
Packet = vbNullString
n = 0
For i = 1 To MAX_PLAYERS
    If IsPlaying(i) Then
        Packet = Packet & SEP_CHAR & GetPlayerName(i) & SEP_CHAR
        n = n + 1
    End If
Next i

Packet = "ONLINELIST" & SEP_CHAR & n & Packet & END_CHAR

Call SendDataToAll(Packet)
End Sub

Sub SendChars(ByVal Index As Long)
Dim Packet As String
Dim i As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim$(Player(Index).Char(i).Name) & SEP_CHAR & Trim$(Classe(Player(Index).Char(i).Class).Name) & SEP_CHAR & Player(Index).Char(i).Level & SEP_CHAR & Player(Index).Char(i).sprite & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendJoinMap(ByVal Index As Long)
Dim Packet As String
Dim i As Long

On Error GoTo er:

    Packet = vbNullString
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> Index And GetPlayerMap(i) = GetPlayerMap(Index) Then
            Packet = "PLAYERDATA" & SEP_CHAR
            Packet = Packet & i & SEP_CHAR
            Packet = Packet & GetPlayerName(i) & SEP_CHAR
            Packet = Packet & GetPlayerSprite(i) & SEP_CHAR
            Packet = Packet & GetPlayerMap(i) & SEP_CHAR
            Packet = Packet & GetPlayerX(i) & SEP_CHAR
            Packet = Packet & GetPlayerY(i) & SEP_CHAR
            Packet = Packet & GetPlayerDir(i) & SEP_CHAR
            Packet = Packet & GetPlayerAccess(i) & SEP_CHAR
            Packet = Packet & GetPlayerPK(i) & SEP_CHAR
            Packet = Packet & GetPlayerGuild(i) & SEP_CHAR
            Packet = Packet & GetPlayerGuildAccess(i) & SEP_CHAR
            Packet = Packet & GetPlayerClass(i) & SEP_CHAR
            Packet = Packet & GetPlayerLevel(i) & SEP_CHAR
            Packet = Packet & Player(i).InParty & SEP_CHAR
            Packet = Packet & END_CHAR
            Call SendDataTo(Index, Packet)
        End If
    Next i
    
    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(Index) & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR
    Packet = Packet & Player(Index).InParty & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)

Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     � : " & time & "...Erreur pendant l'envoi du changement de carte d'un joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ",Carte : " & GetPlayerMap(Index) & ". D�tails : Num :" & Err.Number & " Description : " & Err.Description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant l'envoi du changement de carte d'un joueur(" & GetPlayerName(Index) & ")", BrightRed)
Call PlainMsg(Index, "Erreur du serveur, relancez s'il vous plait.(Pour tous probl�me r�curent visitez " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String

On Error GoTo er:

    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR
    Packet = Packet & Player(Index).InParty
    Packet = Packet & END_CHAR
    Call SendDataToMapBut(Index, MapNum, Packet)

Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     � : " & time & "...Erreur pendant le d�pard du joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ",De la carte : " & MapNum & ". D�tails : Num :" & Err.Number & " Description : " & Err.Description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant le d�pard de " & GetPlayerName(Index) & " d'une la carte", BrightRed)
Call PlainMsg(Index, "Erreur du serveur, relancez s'il vous plait.(Pour tous probl�me r�curent visitez " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Sub SendPlayerData(ByVal Index As Long)
Dim Packet As String

On Error GoTo er:

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(Index) & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR
    Packet = Packet & Player(Index).InParty & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     � : " & time & "...Erreur pendant l'envoi des donn�es du joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ". D�tails : Num :" & Err.Number & " Description : " & Err.Description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant l'envoi des donn�es du joueur : " & GetPlayerName(Index), BrightRed)
Call PlainMsg(Index, "Erreur du serveur, relancez s'il vous plait.(Pour tous probl�me r�curent visitez " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Sub SendPlayerQuete(ByVal Index As Long)
Dim Packet As String
Dim i As Long

On Error GoTo er:

Packet = "PLAYERQUETE" & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).QueteEnCour & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.data1 & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.data2 & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.data3 & SEP_CHAR
Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.String1 & SEP_CHAR

For i = 1 To 15
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data1 & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data2 & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).data3 & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Quetep.indexe(i).String1 & SEP_CHAR
Next i

Packet = Packet & END_CHAR
Call SendDataTo(Index, Packet)
Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     � : " & time & "...Erreur pendant l'envoi des donn�es(qu�te) du joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ". D�tails : Num :" & Err.Number & " Description : " & Err.Description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant l'envoi des donn�es(qu�te) du joueur : " & GetPlayerName(Index), BrightRed)
End Sub

Sub SendPlayerMetier(ByVal Index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERMETIER" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).metier & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).MetierLvl & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).MetierExp & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim x As Long
Dim y As Long
Dim i As Long
Dim s As String
On Error GoTo er:

If CarteFTP Then
    Packet = "MAPDOWN" & SEP_CHAR & MapNum & SEP_CHAR & GetVar(App.Path & "\Data.ini", "FTP", "URL") & SEP_CHAR & GetVar(App.Path & "\Data.ini", "FTP", "REP") & END_CHAR
    Call SendDataTo(Index, Packet)
Else
    Packet = "MAPDATAS" & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR & Map(MapNum).PanoInf & SEP_CHAR & Map(MapNum).TranInf & SEP_CHAR & Map(MapNum).PanoSup & SEP_CHAR & Map(MapNum).TranSup & SEP_CHAR & Map(MapNum).Fog & SEP_CHAR & Map(MapNum).FogAlpha & SEP_CHAR & Map(MapNum).guildSoloView & SEP_CHAR & Map(MapNum).petView & SEP_CHAR & Map(MapNum).traversable & SEP_CHAR & Map(MapNum).meteo & END_CHAR
    
    Call SendDataTo(Index, Packet)
    
    Packet = "MAPTILES" & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
        With Map(MapNum).Tile(x, y)
            Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .type & SEP_CHAR & .data1 & SEP_CHAR & .data2 & SEP_CHAR & .data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .Light & SEP_CHAR
            Packet = Packet & .GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR & .Fringe3 & SEP_CHAR & .F3Anim & SEP_CHAR & .Fringe3Set & SEP_CHAR & .F3AnimSet & SEP_CHAR & .M3Anim & SEP_CHAR & .M3AnimSet & SEP_CHAR & .Mask3 & SEP_CHAR & .Mask3Set & SEP_CHAR  '<--
        End With
        Next x
    Next y
    
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
    
    Packet = "MAPNPCS" & SEP_CHAR
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(GetPlayerMap(Index)).Npc(x) & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).x & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).y & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).x1 & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).y1 & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).x2 & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).y2 & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).Hasardm & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).Hasardp & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).boucle & SEP_CHAR
        Packet = Packet & Map(GetPlayerMap(Index)).Npcs(x).Imobile & SEP_CHAR
    Next x
        
    Packet = Packet & END_CHAR
        
    Call SendDataTo(Index, Packet)
End If

Exit Sub
er:
On Error Resume Next
If Index < 0 Or Index > MAX_PLAYERS Then Exit Sub
Call AddLog("le : " & Date & "     � : " & time & "...Erreur pendant l'envoi de la carte " & MapNum & " au joueur : " & GetPlayerName(Index) & ",Compte : " & GetPlayerLogin(Index) & ". D�tails : Num :" & Err.Number & " Description : " & Err.Description & " Source : " & Err.Source & "...", "logs\Err.txt")
If IBErr Then Call IBMsg("Erreur pendant l'envoi de la carte " & MapNum & " au joueur : " & GetPlayerName(Index), BrightRed)
Call PlainMsg(Index, "Erreur du serveur, relancez s'il vous plait.(Pour tous probl�me r�curent visitez " & Trim$(GetVar(App.Path & "\Config\.ini", "CONFIG", "WebSite")) & ").", 3)
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        If MapNum > 0 Then Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        If MapNum > 0 Then Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendItems(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim$(item(i).Name) <> vbNullString Then Call SendUpdateItemTo(Index, i)
    Next i
End Sub

Sub SendPets(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_PETS
        If Trim$(Pets(i).nom) <> vbNullString Then Call SendUpdatePetTo(Index, i)
    Next i
End Sub

Sub SendMetiers(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_METIER
        If Trim$(metier(i).nom) <> vbNullString Then Call SendUpdatemetierTo(Index, i)
    Next i
End Sub

Sub SendRecettes(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_RECETTE
        Call SendUpdaterecetteTo(Index, i)
    Next i
End Sub

Sub SendEmoticons(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(i).Command) <> vbNullString Then Call SendUpdateEmoticonTo(Index, i)
    Next i
End Sub

Sub SendArrows(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(Index, i)
    Next i
End Sub

Sub SendNpcs(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim$(Npc(i).Name) <> vbNullString Then Call SendUpdateNpcTo(Index, i)
    Next i
End Sub

Sub SendInventory(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR & Index & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemDur(Index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendWornEquipment(ByVal Index As Long)
Dim Packet As String
    If IsPlaying(Index) Then
    'CODE ORIGINAL:
       Packet = "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & GetPlayerPetSlot(Index) & END_CHAR
    'CODE MODIFIE POUR PAPERDOLL:
    'Packet = "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).Casque & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).armure & SEP_CHAR & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).arme & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).bouclier & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If
End Sub

Sub SendHP(ByVal Index As Long)
Dim Packet As String, x As Byte

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
    
    If Player(Index).InParty > 0 Then
        For x = 1 To Party.MemberCount(Player(Index).InParty)
            If Player(Index).PartyPlayer <> x Then Call SendDataTo(Party.PlayerIndex(Player(Index).InParty, x), "partyhp" & SEP_CHAR & Index & SEP_CHAR & Player(Index).InParty & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).HP & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & Player(Index).Char(Player(Index).CharNum).MP & END_CHAR)
        Next x
    End If
    
    Packet = "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendStats(ByVal Index As Long)
Dim Packet As String
    
    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerStr(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Classe(i).STR & SEP_CHAR & Classe(i).def & SEP_CHAR & Classe(i).Speed & SEP_CHAR & Classe(i).magi & SEP_CHAR & Classe(i).Locked & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Classe(i).STR & SEP_CHAR & Classe(i).def & SEP_CHAR & Classe(i).Speed & SEP_CHAR & Classe(i).magi & SEP_CHAR & Classe(i).MaleSprite & SEP_CHAR & Classe(i).FemaleSprite & SEP_CHAR & Classe(i).Locked & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(Index, Packet)
End Sub

Sub SendLeftGame(ByVal Index As Long)

Dim Packet As String
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & "" & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & "" & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToAllBut(Index, Packet)
    
End Sub

Sub SendPlayerXY(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String
    
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(item(ItemNum).Name) & SEP_CHAR & item(ItemNum).Pic & SEP_CHAR & item(ItemNum).type & SEP_CHAR & item(ItemNum).data1 & SEP_CHAR & item(ItemNum).data2 & SEP_CHAR & item(ItemNum).data3 & SEP_CHAR & item(ItemNum).StrReq & SEP_CHAR & item(ItemNum).DefReq & SEP_CHAR & item(ItemNum).SpeedReq & SEP_CHAR & item(ItemNum).ClassReq & SEP_CHAR & item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & item(ItemNum).AddHP & SEP_CHAR & item(ItemNum).AddMP & SEP_CHAR & item(ItemNum).AddSP & SEP_CHAR & item(ItemNum).AddStr & SEP_CHAR & item(ItemNum).AddDef & SEP_CHAR & item(ItemNum).AddMagi & SEP_CHAR & item(ItemNum).AddSpeed & SEP_CHAR & item(ItemNum).AddEXP & SEP_CHAR & item(ItemNum).desc & SEP_CHAR & item(ItemNum).AttackSpeed
    Packet = Packet & SEP_CHAR & item(ItemNum).NCoul & SEP_CHAR & item(ItemNum).paperdoll & SEP_CHAR & item(ItemNum).paperdollPic & SEP_CHAR & item(ItemNum).Empilable & SEP_CHAR & item(ItemNum).tArme & SEP_CHAR & item(ItemNum).LevelReq & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String
    
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(item(ItemNum).Name) & SEP_CHAR & item(ItemNum).Pic & SEP_CHAR & item(ItemNum).type & SEP_CHAR & item(ItemNum).data1 & SEP_CHAR & item(ItemNum).data2 & SEP_CHAR & item(ItemNum).data3 & SEP_CHAR & item(ItemNum).StrReq & SEP_CHAR & item(ItemNum).DefReq & SEP_CHAR & item(ItemNum).SpeedReq & SEP_CHAR & item(ItemNum).ClassReq & SEP_CHAR & item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & item(ItemNum).AddHP & SEP_CHAR & item(ItemNum).AddMP & SEP_CHAR & item(ItemNum).AddSP & SEP_CHAR & item(ItemNum).AddStr & SEP_CHAR & item(ItemNum).AddDef & SEP_CHAR & item(ItemNum).AddMagi & SEP_CHAR & item(ItemNum).AddSpeed & SEP_CHAR & item(ItemNum).AddEXP & SEP_CHAR & item(ItemNum).desc & SEP_CHAR & item(ItemNum).AttackSpeed
    Packet = Packet & SEP_CHAR & item(ItemNum).NCoul & SEP_CHAR & item(ItemNum).paperdoll & SEP_CHAR & item(ItemNum).paperdollPic & SEP_CHAR & item(ItemNum).Empilable & SEP_CHAR & item(ItemNum).tArme & SEP_CHAR & item(ItemNum).LevelReq & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String
    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(item(ItemNum).Name) & SEP_CHAR & item(ItemNum).Pic & SEP_CHAR & item(ItemNum).type & SEP_CHAR & item(ItemNum).data1 & SEP_CHAR & item(ItemNum).data2 & SEP_CHAR & item(ItemNum).data3 & SEP_CHAR & item(ItemNum).StrReq & SEP_CHAR & item(ItemNum).DefReq & SEP_CHAR & item(ItemNum).SpeedReq & SEP_CHAR & item(ItemNum).ClassReq & SEP_CHAR & item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & item(ItemNum).AddHP & SEP_CHAR & item(ItemNum).AddMP & SEP_CHAR & item(ItemNum).AddSP & SEP_CHAR & item(ItemNum).AddStr & SEP_CHAR & item(ItemNum).AddDef & SEP_CHAR & item(ItemNum).AddMagi & SEP_CHAR & item(ItemNum).AddSpeed & SEP_CHAR & item(ItemNum).AddEXP & SEP_CHAR & item(ItemNum).desc & SEP_CHAR & item(ItemNum).AttackSpeed
    Packet = Packet & SEP_CHAR & item(ItemNum).NCoul & SEP_CHAR & item(ItemNum).paperdoll & SEP_CHAR & item(ItemNum).paperdollPic & SEP_CHAR & item(ItemNum).Empilable & SEP_CHAR & item(ItemNum).Sex & SEP_CHAR & item(ItemNum).tArme & SEP_CHAR & item(ItemNum).LevelReq & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdatePetToAll(ByVal PetNum As Long)
Dim Packet As String
    
    Packet = "UPDATEPET" & SEP_CHAR & PetNum & SEP_CHAR & Pets(PetNum).nom & SEP_CHAR & Pets(PetNum).sprite & SEP_CHAR & Pets(PetNum).addForce & SEP_CHAR & Pets(PetNum).addDefence & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdatePetTo(ByVal Index As Long, ByVal PetNum As Long)
Dim Packet As String
    
    Packet = "UPDATEPET" & SEP_CHAR & PetNum & SEP_CHAR & Pets(PetNum).nom & SEP_CHAR & Pets(PetNum).sprite & SEP_CHAR & Pets(PetNum).addForce & SEP_CHAR & Pets(PetNum).addDefence & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditPetTo(ByVal Index As Long, ByVal PetNum As Long)
Dim Packet As String

    Packet = "EDITPET" & SEP_CHAR & PetNum & SEP_CHAR & Pets(PetNum).nom & SEP_CHAR & Pets(PetNum).sprite & SEP_CHAR & Pets(PetNum).addForce & SEP_CHAR & Pets(PetNum).addDefence & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdatemetierToAll(ByVal metiernum As Long)
Dim Packet As String
Dim i As Long, z As Long

    Packet = "UPDATEMETIER" & SEP_CHAR & metiernum & SEP_CHAR & metier(metiernum).nom & SEP_CHAR & metier(metiernum).type & SEP_CHAR & metier(metiernum).desc & SEP_CHAR
    For i = 0 To MAX_DATA_METIER
        For z = 0 To 1
            Packet = Packet & metier(metiernum).data(i, z) & SEP_CHAR
        Next z
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdatemetierTo(ByVal Index As Long, ByVal metiernum As Long)
Dim Packet As String
Dim i As Long, z As Long

    Packet = "UPDATEMETIER" & SEP_CHAR & metiernum & SEP_CHAR & metier(metiernum).nom & SEP_CHAR & metier(metiernum).type & SEP_CHAR & metier(metiernum).desc & SEP_CHAR
    For i = 0 To MAX_DATA_METIER
        For z = 0 To 1
            Packet = Packet & metier(metiernum).data(i, z) & SEP_CHAR
        Next z
    Next i
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditmetierTo(ByVal Index As Long, ByVal metiernum As Long)
Dim Packet As String
Dim i As Long, z As Long

    Packet = "EDITMETIER" & SEP_CHAR & metiernum & SEP_CHAR & metier(metiernum).nom & SEP_CHAR & metier(metiernum).type & SEP_CHAR & metier(metiernum).desc & SEP_CHAR
    For i = 0 To MAX_DATA_METIER
        For z = 0 To 1
            Packet = Packet & metier(metiernum).data(i, z) & SEP_CHAR
        Next z
    Next i
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdaterecetteToAll(ByVal recettenum As Long)
Dim Packet As String
Dim i As Long, z As Long

    Packet = "UPDATErecette" & SEP_CHAR & recettenum & SEP_CHAR & recette(recettenum).nom & SEP_CHAR
    For i = 0 To 9
        For z = 0 To 1
            Packet = Packet & recette(recettenum).InCraft(i, z) & SEP_CHAR
        Next z
    Next i
    For z = 0 To 1
        Packet = Packet & recette(recettenum).craft(z) & SEP_CHAR
    Next z
    Packet = Packet & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdaterecetteTo(ByVal Index As Long, ByVal recettenum As Long)
Dim Packet As String
Dim i As Long, z As Long

    Packet = "UPDATErecette" & SEP_CHAR & recettenum & SEP_CHAR & recette(recettenum).nom & SEP_CHAR
    For i = 0 To 9
        For z = 0 To 1
            Packet = Packet & recette(recettenum).InCraft(i, z) & SEP_CHAR
        Next z
    Next i
    For z = 0 To 1
        Packet = Packet & recette(recettenum).craft(z) & SEP_CHAR
    Next z
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditrecetteTo(ByVal Index As Long, ByVal recettenum As Long)
Dim Packet As String
Dim i As Long, z As Long

    Packet = "EDITrecette" & SEP_CHAR & recettenum & SEP_CHAR & recette(recettenum).nom & SEP_CHAR
    For i = 0 To 9
        For z = 0 To 1
            Packet = Packet & recette(recettenum).InCraft(i, z) & SEP_CHAR
        Next z
    Next i
    For z = 0 To 1
        Packet = Packet & recette(recettenum).craft(z) & SEP_CHAR
    Next z
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditEmoticonTo(ByVal Index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateArrowTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditArrowTo(ByVal Index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITArrow" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Arrows(EmoNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal npcnum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(Npc(npcnum).Name) & SEP_CHAR & Npc(npcnum).sprite & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).QueteNum & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & CLng(Npc(npcnum).Inv) & SEP_CHAR & CLng(Npc(npcnum).Vol) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal npcnum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(Npc(npcnum).Name) & SEP_CHAR & Npc(npcnum).sprite & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).QueteNum & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & CLng(Npc(npcnum).Inv) & SEP_CHAR & CLng(Npc(npcnum).Vol) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal npcnum As Long)
Dim Packet As String
Dim i As Long

    'Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & trim$(Npc(NpcNum).Name) & SEP_CHAR & trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR
    'Packet = Packet & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & END_CHAR
    Packet = "EDITNPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(Npc(npcnum).Name) & SEP_CHAR & Trim$(Npc(npcnum).AttackSay) & SEP_CHAR & Npc(npcnum).sprite & SEP_CHAR & Npc(npcnum).SpawnSecs & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & Npc(npcnum).Range & SEP_CHAR & Npc(npcnum).STR & SEP_CHAR & Npc(npcnum).def & SEP_CHAR & Npc(npcnum).Speed & SEP_CHAR & Npc(npcnum).magi & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).Exp & SEP_CHAR & Npc(npcnum).SpawnTime & SEP_CHAR & Npc(npcnum).QueteNum & SEP_CHAR & CLng(Npc(npcnum).Inv) & SEP_CHAR & CLng(Npc(npcnum).Vol) & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(npcnum).ItemNPC(i).chance
        Packet = Packet & SEP_CHAR & Npc(npcnum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(npcnum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    For i = 1 To MAX_NPC_SPELLS
        Packet = Packet & Npc(npcnum).Spell(i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendShops(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim$(Shop(i).Name) <> vbNullString Then Call SendUpdateShopTo(Index, i)
    Next i
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, z As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).FixObjet & SEP_CHAR
    For i = 1 To 6
        For z = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).value(z).GetValue & SEP_CHAR
        Next z
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpells(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).Name) <> vbNullString Then Call SendUpdateSpellTo(Index, i)
    Next i
End Sub

Sub SendQuetes(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_QUETES
        If Trim$(quete(i).nom) <> vbNullString Or quete(i).type <> 0 Then Call SendUpdateQueteTo(Index, i)
    Next i
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).SpellIco & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateQueteToAll(ByVal QueteNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "UPDATEQUETE" & SEP_CHAR & QueteNum & SEP_CHAR & Trim$(quete(QueteNum).nom) & SEP_CHAR & quete(QueteNum).data1 & SEP_CHAR & quete(QueteNum).data2 & SEP_CHAR & quete(QueteNum).data3 & SEP_CHAR & quete(QueteNum).Description & SEP_CHAR & quete(QueteNum).reponse & SEP_CHAR & quete(QueteNum).String1 & SEP_CHAR & quete(QueteNum).temps & SEP_CHAR & quete(QueteNum).type
    
    For i = 1 To 15
        Packet = Packet & SEP_CHAR & quete(QueteNum).indexe(i).data1 & SEP_CHAR & quete(QueteNum).indexe(i).data2 & SEP_CHAR & quete(QueteNum).indexe(i).data3 & SEP_CHAR & quete(QueteNum).indexe(i).String1
    Next i
    
    Packet = Packet & SEP_CHAR & quete(QueteNum).Recompence.Exp & SEP_CHAR & quete(QueteNum).Recompence.objn1 & SEP_CHAR & quete(QueteNum).Recompence.objn2 & SEP_CHAR & quete(QueteNum).Recompence.objn3 & SEP_CHAR & quete(QueteNum).Recompence.objq1 & SEP_CHAR & quete(QueteNum).Recompence.objq2 & SEP_CHAR & quete(QueteNum).Recompence.objq3 & SEP_CHAR & quete(QueteNum).Case
    Packet = Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).SpellIco & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateQueteTo(ByVal Index As Long, ByVal QueteNum As Long)
Dim Packet As String
Dim i As Long
    Packet = "UPDATEQUETE" & SEP_CHAR & QueteNum & SEP_CHAR & Trim$(quete(QueteNum).nom) & SEP_CHAR & quete(QueteNum).data1 & SEP_CHAR & quete(QueteNum).data2 & SEP_CHAR & quete(QueteNum).data3 & SEP_CHAR & quete(QueteNum).Description & SEP_CHAR & quete(QueteNum).reponse & SEP_CHAR & quete(QueteNum).String1 & SEP_CHAR & quete(QueteNum).temps & SEP_CHAR & quete(QueteNum).type
    
    For i = 1 To 15
        Packet = Packet & SEP_CHAR & quete(QueteNum).indexe(i).data1 & SEP_CHAR & quete(QueteNum).indexe(i).data2 & SEP_CHAR & quete(QueteNum).indexe(i).data3 & SEP_CHAR & quete(QueteNum).indexe(i).String1
    Next i
    
    Packet = Packet & SEP_CHAR & quete(QueteNum).Recompence.Exp & SEP_CHAR & quete(QueteNum).Recompence.objn1 & SEP_CHAR & quete(QueteNum).Recompence.objn2 & SEP_CHAR & quete(QueteNum).Recompence.objn3 & SEP_CHAR & quete(QueteNum).Recompence.objq1 & SEP_CHAR & quete(QueteNum).Recompence.objq2 & SEP_CHAR & quete(QueteNum).Recompence.objq3 & SEP_CHAR & quete(QueteNum).Case
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).type & SEP_CHAR & Spell(SpellNum).data1 & SEP_CHAR & Spell(SpellNum).data2 & SEP_CHAR & Spell(SpellNum).data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).SpellIco & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditQuetesTo(ByVal Index As Long, ByVal QueteNum As Long)
Dim Packet As String

    Packet = "EDITQUETES" & SEP_CHAR & QueteNum & SEP_CHAR & Trim$(quete(QueteNum).nom) & SEP_CHAR & quete(QueteNum).data1 & SEP_CHAR & quete(QueteNum).data2 & SEP_CHAR & quete(QueteNum).data3 & SEP_CHAR & quete(QueteNum).Description & SEP_CHAR & quete(QueteNum).reponse & SEP_CHAR & quete(QueteNum).String1 & SEP_CHAR & quete(QueteNum).temps & SEP_CHAR & quete(QueteNum).type & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, x As Long, y As Long, z As Long, XX As Long
    
    Player(Index).Char(Player(Index).CharNum).vendeur = ShopNum
    
    z = 0
    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).FixObjet & SEP_CHAR
    For i = 1 To 6
        For XX = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).value(XX).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).value(XX).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).value(XX).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).value(XX).GetValue & SEP_CHAR
        Next XX
    Next i
    Packet = Packet & END_CHAR
    
    If z = (MAX_TRADES * 6) Then
        Call PlayerMsg(Index, "Ce magasin ne vend rien.", BrightRed)
    Else
        Call SendDataTo(Index, Packet)
    End If
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(Index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherTo(ByVal Index As Long)
Dim Packet As String
    If RainIntensity <= 0 Then RainIntensity = 1
    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & RainIntensity & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherToAll()
Dim i As Long
Dim Weather As String

    PutVar App.Path & "\Data.ini", "CONFIG", "MeteoParMap", 0
    frmServer.Label12.Caption = "La m�t�o par map est actuellement d�sactiv�e"
    frmServer.CommandMeteoMap.Caption = "activer La m�t�o par map"
    
    Select Case GameWeather
        Case 0
            Weather = "Soleil"
        Case 1
            Weather = "Pluie"
        Case 2
            Weather = "Neige"
        Case 3
            Weather = "Orage"
    End Select
    frmServer.Label5.Caption = "Temps actuel : " & Weather
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then Call SendWeatherTo(i)
    Next i
End Sub

Sub SendTimeTo(ByVal Index As Long)
Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTimeToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then Call SendTimeTo(i)
    Next i
    
    Call SpawnAllMapNpcs
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal Index As Long)
Dim Packet As String

    Packet = "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & Index & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Function MMsg(ByVal Msg As String) As Boolean
Dim i As Long
Dim Asct As String

MMsg = True

For i = 1 To Len(Msg)
    Asct = Asc(Mid$(Msg, i, 1))
    If Asct < 32 Or Asct > 126 Then
        If (Not Asct = 253) And (Not Asct = 252) And (Not Asct = 251) And (Not Asct = 250) And (Not Asct = 249) And (Not Asct = 246) And (Not Asct = 245) And (Not Asct = 244) And (Not Asct = 243) And (Not Asct = 242) And (Not Asct = 238) And (Not Asct = 238) And (Not Asct = 237) And (Not Asct = 236) And (Not Asct = 235) And (Not Asct = 234) And (Not Asct = 233) And (Not Asct = 232) And (Not Asct = 231) And (Not Asct = 230) And (Not Asct = 229) And (Not Asct = 228) And (Not Asct = 227) And (Not Asct = 226) And (Not Asct = 225) And (Not Asct = 224) And (Not Asct = 202) And (Not Asct = 128) And (Not Asct = 199) And (Not Asct = 167) And (Not Asct = 164) Then
           MMsg = True
           Exit Function
        End If
    End If
Next i

MMsg = False
End Function
