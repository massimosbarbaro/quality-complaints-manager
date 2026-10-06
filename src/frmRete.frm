VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmRete 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Rete"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   Icon            =   "frmRete.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "Esci"
      Default         =   -1  'True
      Height          =   615
      Left            =   3840
      Picture         =   "frmRete.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdRete 
      Caption         =   "Stampa"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin Crystal.CrystalReport cryRete 
      Left            =   3840
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblDettagli 
      Caption         =   "Risposta"
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Risposta"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label lblDettagli 
      Caption         =   "Risposta"
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Risposta n."
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Risposte per la rete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   630
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   4140
   End
   Begin VB.Image imgLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmRete.frx":0E14
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmRete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String


Public Property Let Key(ByVal vValue As String)
   mKey = Trim$(vValue)
   Dim sqlc As String, rsRisposta As Recordset, i As Integer
   sqlc = "SELECT ATC.NumeroRisposta, ATC.DataRisposta " & _
          "FROM ATC where NumeroRisposta = " & Me.Key
   Set rsRisposta = db.OpenRecordset(sqlc, dbOpenSnapshot)
   For i = 0 To rsRisposta.Fields.Count - 1
      If Not IsNull(rsRisposta(i)) Then
         lblDettagli(i).Caption = rsRisposta(i)
      Else
         lblDettagli(i) = ""
      End If
   Next i
   
 
End Property
Public Property Get Key() As String
    Key = mKey
End Property

Private Sub cmdEsci_Click()
 Unload Me
End Sub

Private Sub cmdRete_Click()
   Dim sSelFormula As String
'   Dim sqlc As String, rsRete As Recordset, i As Byte
'   sqlc = "SELECT ATC.NumeroRisposta, ATC.DataRisposta, ATC.Risposta, " & _
'            "ATC.AzioniCorrettive, ATC.CodFamiglia, ATC.CodDifetto, " & _
'            "ATC.Esterno, ATC.StatoReclamo, ATC.IndiceQualita, " & _
'            "ATC.Soluzione, ATC.QtaAccettata, " & _
'            "ReclamoVenditore.DataReclamo, ReclamoVenditore.QtaReclamata, " & _
'            "ReclamoVenditore.CausaReclamo, ReclamoVenditore.Venditore, " & _
'            "ReclamoVenditore.NumeroReclamo, " & _
'            "MFG.ODL, MFG.Rete, MFG.CodCli, MFG.Cliente, MFG.CodProd, " & _
'            "MFG.Prodotto, MFG.LineaProdotto, MFG.Finitura, MFG.Spessore, " & _
'            "MFG.Altezza, MFG.Larghezza, MFG.Qta " & _
'            "FROM MFG INNER JOIN (ATC INNER JOIN ReclamoVenditore ON ATC.ODL = " & _
'            "ReclamoVenditore.ODL) ON (ATC.ODL = MFG.ODL) AND (MFG.ODL = ReclamoVenditore.ODL)" & _
'            "WHERE ATC.NumeroRisposta=" & Me.Key
'   Set rsRete = db.OpenRecordset(sqlc, dbOpenSnapshot)
'   i = 0
   With cryRete
      On Error Resume Next
      .DataFiles(0) = db.Name
      .Destination = crptToWindow
      .ReportFileName = App.Path & "\rete.rpt"
'      Do While Not rsRete.EOF
       '  .Formulas(i) = "F" & Format(i, "0") & _
       '              " = if ({ATC.QtaAccettata} >= " & rsLimiti("LimInf") & _
       '              " and {ATC.QtaAccettata} <= " & _
       '              Format$(rsLimiti("LimSup"), "000000000") & ")" & _
       '              " then 1 else 0"
'         i = i + 1
'         rsRete.MoveNext
'      Loop
      .WindowState = crptNormal
      sSelFormula = "{ATC.NumeroRisposta} = " & Me.Key
'      .Formulas(1) = "DataOra=" & Chr$(34) & _
'          Format(Now, "dd/mm/yy  -  hh:mm") & Chr$(34)
   .SelectionFormula = sSelFormula
      .Action = 1
      If Err.Number <> 0 Then
          MsgBox "Errore nel produrre il report " & vbCrLf & _
              Err.Number & " - " & Err.Description
          Err.Clear
      End If
   End With
'   rsRete.Close
'   Set rsRete = Nothing
End Sub

Private Sub lblRisposta_Click()

End Sub

