VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPop3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ricevi Posta"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsk 
      Left            =   6960
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab sstPosta 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ricezione Posta"
      TabPicture(0)   =   "frmPop3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExit"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtLog"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdRicevi"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Pop3 Parameters"
      TabPicture(1)   =   "frmPop3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(3)=   "Label1"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "txtPopParameters(3)"
      Tab(1).Control(8)=   "txtPopParameters(2)"
      Tab(1).Control(9)=   "txtPopParameters(1)"
      Tab(1).Control(10)=   "cmdOp(1)"
      Tab(1).Control(11)=   "cmdOp(0)"
      Tab(1).Control(12)=   "txtPopParameters(0)"
      Tab(1).Control(13)=   "txtPopParameters(4)"
      Tab(1).Control(14)=   "txtPopParameters(5)"
      Tab(1).Control(15)=   "txtPopParameters(6)"
      Tab(1).Control(16)=   "txtPopParameters(7)"
      Tab(1).Control(17)=   "cmdOp(2)"
      Tab(1).ControlCount=   18
      Begin VB.CommandButton cmdOp 
         Caption         =   "Esci"
         Height          =   495
         Index           =   2
         Left            =   -67560
         TabIndex        =   22
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtPopParameters 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   -70920
         TabIndex        =   20
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtPopParameters 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   -70200
         TabIndex        =   18
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtPopParameters 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   -70920
         TabIndex        =   17
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox txtPopParameters 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   -70920
         TabIndex        =   16
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdRicevi 
         BackColor       =   &H00FF8080&
         Caption         =   "Ricevi Messaggi"
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2640
         Width           =   8415
      End
      Begin VB.TextBox txtLog 
         Height          =   1575
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmPop3.frx":0038
         Top             =   840
         Width           =   8415
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   495
         Left            =   7560
         TabIndex        =   11
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtPopParameters 
         Height          =   285
         Index           =   0
         Left            =   -70920
         TabIndex        =   6
         Top             =   540
         Width           =   2175
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "Leggi"
         Height          =   495
         Index           =   0
         Left            =   -72120
         TabIndex        =   5
         Top             =   3780
         Width           =   1215
      End
      Begin VB.CommandButton cmdOp 
         Caption         =   "Salva"
         Height          =   495
         Index           =   1
         Left            =   -70680
         TabIndex        =   4
         Top             =   3780
         Width           =   1215
      End
      Begin VB.TextBox txtPopParameters 
         Height          =   285
         Index           =   1
         Left            =   -70920
         TabIndex        =   3
         Top             =   1020
         Width           =   2175
      End
      Begin VB.TextBox txtPopParameters 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   -70920
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1500
         Width           =   2175
      End
      Begin VB.TextBox txtPopParameters 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   -70920
         TabIndex        =   1
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Subject"
         Height          =   255
         Left            =   -72480
         TabIndex        =   21
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "String: Start / End"
         Height          =   255
         Left            =   -72480
         TabIndex        =   19
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Timeout"
         Height          =   255
         Left            =   -72480
         TabIndex        =   15
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "ALT + Mouse DX + Config"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   3720
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Pop Server"
         Height          =   255
         Left            =   -72480
         TabIndex        =   10
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Pop User"
         Height          =   255
         Left            =   -72480
         TabIndex        =   9
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Pop Password"
         Height          =   255
         Left            =   -72480
         TabIndex        =   8
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Pop Server Port"
         Height          =   255
         Left            =   -72480
         TabIndex        =   7
         Top             =   1980
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPop3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum OP
   cLeggi = 0
   cSalva
   cEsci
End Enum
Private Enum PopParms
   cPopServer = 0
   cPopUser
   cPopPassword
   cPopServerPort
   cPopTimeout
   cPopStartStr
   cPopEndStr
   cPopSubject
End Enum
' Private Const iPopTimeout As Byte = 10
Private iOK As Boolean
Private sPwConfig As String
Private Const MaxParms As Byte = 8, sFieldSep As String = "\"
Dim sKeys(0 To MaxParms - 1) As String
Dim sKeysDefaults(0 To MaxParms - 1) As String
Private sBuff As String
Private sPopServer As String, sPopServerPort As String
Private sPopUser As String, sPopPassw As String
Private iPopTimeout As Integer, sPopStartStr As String, sPopEndStr As String
Private sPopSubject As String

Private Sub AddLog(sText As String)
   txtLog.Text = txtLog.Text & sText & vbCrLf
   txtLog.SelStart = Len(txtLog.Text)
End Sub

Function GetResponse(sType As String, sDati As String) As Boolean
'------------------------------
'  Aspetta risposta via wsk_DataArrival
'  (prende solo la prima riga con +OK o -ERR)
'  sType vale: R per Reply (1° riga)
'              D per Data (il testo del messaggio)
'
'  sDati = solo prima riga per R
'           tutto il messaggio per D
'
   Dim fOK As Boolean, dTimer As Double
   dTimer = Timer
   If sType = "R" Then
'  Prendo la prima riga del messaggio: +OK o -ERR
      Do While Timer < dTimer + iPopTimeout And _
            InStr(sBuff, vbCr) = 0 And InStr(sBuff, vbLf) = 0
         DoEvents
      Loop
      If InStr(sBuff, vbCr) Then
         sBuff = Left$(sBuff, InStr(sBuff, vbCr) - 1)
         fOK = True
      End If
      If InStr(sBuff, vbLf) Then
         sBuff = Left$(sBuff, InStr(sBuff, vbLf) - 1)
         fOK = True
      End If
   Else
'  Prendo il testo del messaggio che termina con crlf + . + crlf
      fOK = True
      Do While Right$(sBuff, 5) <> vbCrLf & "." & vbCrLf
         DoEvents
         If Timer > dTimer + iPopTimeout Then
'            MsgBox "timeout nella ricezione del messaggio"
            fOK = False
         End If
      Loop
   End If
   If fOK Then
      If Trim$(Left$(sBuff, 3)) = "+OK" Then
         fOK = True
         sDati = sBuff
      Else
         fOK = False
      End If
   End If
   sBuff = ""
   GetResponse = fOK
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOp_Click(Index As Integer)
   Dim sAppName As String, sFileNameIni As String
   Dim sTmp As String, nSize As Long, sKey As String, i As Byte
   
   sAppName = App.EXEName & Chr$(0)
   sFileNameIni = App.Path & "\" & App.EXEName & ".ini" & Chr$(0)
   Select Case Index
      Case cLeggi
         For i = 0 To MaxParms - 1
            sTmp = Space$(255)
            sKey = sKeys(i) & Chr$(0)
            nSize = GetPrivateProfileString(sAppName, sKey, _
                        sKeysDefaults(i), sTmp, Len(sTmp), sFileNameIni)
            sTmp = Left$(sTmp, nSize)
            txtPopParameters(i).Text = sTmp
         Next i
      Case cSalva
         For i = 0 To MaxParms - 1
            sTmp = Trim$(txtPopParameters(i).Text) & Chr$(0)
            sKey = sKeys(i) & Chr$(0)
            nSize = WritePrivateProfileString(sAppName, sKey, sTmp, sFileNameIni)
         Next i
      Case cEsci
         sstPosta.TabsPerRow = 1
         sstPosta.TabVisible(1) = False
   End Select
End Sub

Private Sub cmdRicevi_Click()
   Dim dTimer As Double, i As Integer
   Dim nMsg As Integer, nMsgReclami As Integer
   Dim sReply As String, sReclamoText As String
   Dim iPos As Integer
   
   Screen.MousePointer = vbArrowHourglass
'  Pulisco il buffer dei dati
   sBuff = ""
   iOK = True
   With wsk
      AddLog "Connessione in corso con: " & sPopServer & ":" & sPopServerPort
'     Per prima cosa effettuo la connessione
      .Connect sPopServer, sPopServerPort
      dTimer = Timer + iPopTimeout
      Do While Timer < dTimer And .State <> sckConnected
         DoEvents
      Loop
      If .State <> sckConnected Then
         AddLog "Connessione NON effettuata "
'         MsgBox "Errore nella connessione !!"
         iOK = False
         GoTo ExitRicevi
      End If
'
'  Ricevi messaggio di benvenuto
'
      If Not GetResponse("R", sReply) Then
         AddLog "Connessione non effettuata"
         GoTo ExitRicevi
      End If
      
      AddLog "Connessione effettuata con " & .RemoteHost & " (" & .RemoteHostIP & ")"
'      AddLog "sReply: " & sReply
'  Procedo con l'invio dei comandi pop3
'
'  comando: user
      sBuff = ""
      .SendData "USER " & sPopUser & vbCrLf
      If Not GetResponse("R", sReply) Then
         AddLog "User: " & sPopUser & " NON accettata !"
         iOK = False
         GoTo ExitRicevi
      End If
'      AddLog "User: " & sPopUser & " accettata ! - sreply " & sReply
'
'  Comando: pass
      sBuff = ""
      .SendData "PASS " & sPopPassw & vbCrLf
      If Not GetResponse("R", sReply) Then
         iOK = False
         AddLog "Password NON accettata ! - " & sReply
         GoTo ExitRicevi
      End If
'      AddLog "Password accettata !"
'
'  Comando: stat
'
'  Reply: +OK  n  o  - n=Numero messaggi, o=octets
'
      .SendData "STAT " & vbCrLf
      If Not GetResponse("R", sReply) Then
         iOK = False
         AddLog "Comando STAT NON accettato ! "
         GoTo ExitRicevi
      End If
'      AddLog "Comando STAT accettato ! - sReply: " & sReply
      nMsg = Val(Mid$(sReply, 5, InStr(Mid$(sReply, 5), " ")))
      If nMsg > 0 Then
         AddLog "Trovati " & Str$(nMsg) & " messaggi in totale!"
      Else
         AddLog "NON Trovati messaggi !"
         Beep
         GoTo ExitRicevi
      End If
'
'  Ciclo su tutti i messaggi (TOP n 50 righe)
'
      nMsgReclami = 0
      For i = 1 To nMsg
         .SendData "TOP " & Str$(i) & ", 50" & vbCrLf
         If Not GetResponse("D", sReply) Then
            AddLog "Comando TOP NON accettato ! - " & sReply
            GoTo ExitRicevi
         End If
         If InStr(UCase(sReply), "SUBJECT: " & sPopSubject) <> 0 Then
'
'  messaggio OK  ===> lo leggo
'
'            .SendData "TOP " & Str$(i) & ", 50" & vbCrLf
'            If Not GetResponse("D", sReply) Then
'               AddLog "Comando TOP NON accettato ! - " & sReply
'               GoTo ExitRicevi
'            End If
            nMsgReclami = nMsgReclami + 1
'  leggo la string del reclamo venditore
            iPos = InStr(sReply, sPopStartStr)
            If iPos = 0 Then
               AddLog "Non trovato testo messaggio reclamo (" & Str$(i) & ")"
               GoTo ExitRicevi
            End If
            sReclamoText = ""
            Do While iPos > 0
               iPos = iPos + Len(sPopStartStr)
               sReclamoText = sReclamoText & _
                     Mid$(sReply, iPos, InStr(iPos, sReply, sPopEndStr) - iPos)
               iPos = InStr(iPos, sReply, sPopStartStr)
            Loop
'
'  Ho letto e composto la stringa intera in sReclamoText
'
            AddReclamo sReclamoText
'
'  Messaggio scritto in Db ---> lo cancello
'
            .SendData "DELE " & Str$(i) & vbCrLf
            If Not GetResponse("R", sReply) Then
               AddLog "Comando TOP per msg. " & Str$(i) & _
                           " NON accettato ! - " & sReply
               GoTo ExitRicevi
            End If

         End If
      Next i
      .SendData "QUIT" & vbCrLf
      If Not GetResponse("R", sReply) Then
         AddLog "Comando QUIT NON accettato ! "
      Else
'         AddLog "Fine connessione"
      End If
      
   End With
   
ExitRicevi:
'
'  Comunque chiudo il socket Tcp
'
   wsk.Close
   AddLog "Fine ricezione messaggi - caricati: " & Str$(nMsgReclami)
   Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
'
'  Inizializzazione parametri del file INI
'
   sKeys(cPopServer) = "PopServer"
   sKeys(cPopUser) = "PopUser"
   sKeys(cPopPassword) = "PopPassword"
   sKeys(cPopServerPort) = "PopServerPort"
   sKeys(cPopTimeout) = "PopTimeout"
   sKeys(cPopStartStr) = "PopStartStr"
   sKeys(cPopEndStr) = "PopEndStr"
   sKeys(cPopSubject) = "PopSubject"

   sKeysDefaults(cPopServer) = ""
   sKeysDefaults(cPopUser) = ""
   sKeysDefaults(cPopPassword) = ""
   sKeysDefaults(cPopServerPort) = "110"
   sKeysDefaults(cPopTimeout) = "10"
   sKeysDefaults(cPopStartStr) = "$$**"
   sKeysDefaults(cPopEndStr) = "**$$"
   sKeysDefaults(cPopSubject) = "$RECLAMI$"
  
   cmdOp(cLeggi).Value = True
'
'  Inizializzazione parametri
'
   sPopServer = txtPopParameters(cPopServer).Text
   sPopServerPort = txtPopParameters(cPopServerPort).Text
   sPopUser = txtPopParameters(cPopUser).Text
   sPopPassw = txtPopParameters(cPopPassword).Text
   iPopTimeout = Val(txtPopParameters(cPopTimeout).Text)
   sPopStartStr = txtPopParameters(cPopStartStr).Text
   sPopEndStr = txtPopParameters(cPopEndStr).Text
   sPopSubject = txtPopParameters(cPopSubject).Text
   
   With wsk
      .Protocol = sckTCPProtocol
      .RemoteHost = sPopServer
      .RemotePort = sPopServerPort
   End With
   With sstPosta
      .TabVisible(1) = False
      .TabsPerRow = 1
   End With
   With txtLog
      .Enabled = True
      .Locked = True
      .Text = ""
   End With
   With sstPosta
      .Top = 0
      .Left = 0
      .Width = Me.ScaleWidth
      .Height = Me.ScaleHeight
   End With
End Sub

Private Sub sstPosta_KeyPress(KeyAscii As Integer)
   sPwConfig = sPwConfig & Chr$(KeyAscii)
   If sPwConfig = "Config" Then
      sstPosta.TabsPerRow = 2
      sstPosta.TabVisible(1) = True
   End If
End Sub

Private Sub sstPosta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton And Shift = vbAltMask Then
      sPwConfig = ""
   End If
End Sub

Private Sub wsk_DataArrival(ByVal bytesTotal As Long)
   Dim sInp As String
   sInp = ""
   wsk.GetData sInp, vbString
   sBuff = sBuff + sInp
End Sub

Private Sub wsk_Error(ByVal Number As Integer, Description As String, _
                      ByVal Scode As Long, ByVal Source As String, _
                      ByVal HelpFile As String, ByVal HelpContext As Long, _
                      CancelDisplay As Boolean)
   MsgBox "Errore nel controllo WinSock: " & vbCrLf & vbCrLf & _
            Str$(Number) & " - " & Description
   iOK = False
End Sub

Private Sub AddReclamo(s As String)
   Dim sqlc As String, rsReclamo As Recordset
   Dim iPosEnd As Integer, iPosStart As Integer
   Dim i As Integer, sTmp As String

   sqlc = "Select ODL, DataReclamo, QtaReclamata, " & _
            "CausaReclamo, Venditore, NumeroReclamo, DataEMAIL, Stato " & _
            "from ReclamoVenditore " & _
            "where ODL is null"
   Set rsReclamo = db.OpenRecordset(sqlc, dbOpenDynaset)
   rsReclamo.AddNew
   
   i = 0
   iPosStart = 1
   iPosEnd = InStr(s, sFieldSep)
   If iPosEnd = 0 Then
      MsgBox "Non Trovato separatore nel messaggio !"
   End If
   sTmp = Left(s, iPosEnd - 1)                      '  ODL
   rsReclamo("ODL") = sTmp
   
   AddLog "Inserimento ODL: " & sTmp
   
   iPosStart = iPosEnd + Len(sFieldSep)
   iPosEnd = InStr(iPosStart, s, sFieldSep)
   sTmp = Mid$(s, iPosStart, iPosEnd - iPosStart)  '  DataReclamo
   rsReclamo("DataReclamo") = DataUsa(sTmp)
   
   iPosStart = iPosEnd + Len(sFieldSep)
   iPosEnd = InStr(iPosStart, s, sFieldSep)
   sTmp = Mid$(s, iPosStart, iPosEnd - iPosStart)  '  QtaReclamata
   rsReclamo("QtaReclamata") = Val(sTmp)
   
   iPosStart = iPosEnd + Len(sFieldSep)
   iPosEnd = InStr(iPosStart, s, sFieldSep)
   sTmp = Mid$(s, iPosStart, iPosEnd - iPosStart)  '  CausaReclamo
   rsReclamo("CausaReclamo") = sTmp
   
   iPosStart = iPosEnd + Len(sFieldSep)
   iPosEnd = InStr(iPosStart, s, sFieldSep)
   sTmp = Mid$(s, iPosStart, iPosEnd - iPosStart)  '  Venditore
   rsReclamo("Venditore") = sTmp
   
   iPosStart = iPosEnd + Len(sFieldSep)
   iPosEnd = InStr(iPosStart, s, sFieldSep)
   sTmp = Mid$(s, iPosStart, iPosEnd - iPosStart)  '  NumeroReclamo
   rsReclamo("NumeroReclamo") = sTmp

   iPosStart = iPosEnd + Len(sFieldSep)
   iPosEnd = Len(s)
   sTmp = Mid$(s, iPosStart, iPosEnd - iPosStart)  '  DataEMail
   rsReclamo("DataEMail") = DataUsa(sTmp)          ' Sarebbe TimeStamp !!

   rsReclamo("Stato") = "N"
'
'  N.B. BISOGNA decidere cosa fare nel caso di ODL già presente !!!
'        cioè messaggio rimandato - Err.number = 3022 (duplicate PK)
'
   rsReclamo.Update
   rsReclamo.Close
   Set rsReclamo = Nothing
End Sub
