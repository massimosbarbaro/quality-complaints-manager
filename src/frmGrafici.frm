VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGrafici 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Grafici"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   Icon            =   "frmGrafici.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbMesi 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdRicerche 
      Caption         =   "Ricerche rete"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccess 
      Caption         =   "Access"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "Visualizza"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   615
      Left            =   4560
      Picture         =   "frmGrafici.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   615
   End
   Begin VB.Frame fraGrafici 
      Height          =   3015
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   3375
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Esterni"
         Height          =   495
         Index           =   8
         Left            =   360
         TabIndex        =   20
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Rete"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Famiglia"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Difetto"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   345
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Esterni"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Rete"
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo Generale"
         Height          =   495
         Index           =   6
         Left            =   360
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Soluzione"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Famiglia"
         Height          =   375
         Index           =   5
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
   End
   Begin MSMask.MaskEdBox mskDataDA 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataA 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Da"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "A"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Mese / Anno (4 cifre)"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Grafici"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmGrafici.frx":0E14
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmGrafici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iIndice As Integer
'Dim ac As Access.Application


Private Sub cmdAccess_Click()
Dim ac As Access.Application

'Set ac = New Access.Application
Set ac = CreateObject("Access.Application")
' put the path to your database in hereac.
'ac.OpenCurrentDatabase ("C:\Programmi\reclami\reclami.mdb")
ac.OpenCurrentDatabase ("C:\Programmi\reclami\reclami.mdb")
' by default, the OpenReport method of the
' DoCmd object will send the report to the printerac.
ac.DoCmd.OpenReport "rDateFamiglie", acViewPreview

' close the database
ac.CloseCurrentDatabase
End Sub

Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub cmdGrafico_Click()
   Dim sSelFormula As String, sGrafico As String
   Dim ac As Access.Application, sReport As String
   Dim sFile As String, sWhere As String, sTabella As String

'-------------------------------------------------------------------------------
' Preparo le date limite per la stampa dei grafici

   Dim sqlc As String, rsLimiti As Recordset, i As Byte
   Dim dDataInizio, dDataFine
   Dim sTitolo As String, sTest As String, sLinea As String, sGraficoFamiglia As String, sFam As String
   Dim sReteGrafico As String, sQtaGrafico As String, sGraficiReti As String
   
   
   
   
   If mskDataA.Text = "__/____" Then
      mskDataA.Text = mskDataDA.Text
   End If
   dDataFine = Format((DateAdd("d", -1, DateAdd("m", 1, "01/" & mskDataA.Text))), "dd/mm/yyyy")
   
   sTest = Format(DateSerial(Val(Mid$(mskDataDA.Text, 4, 4)), _
                            Val(Mid$(mskDataDA.Text, 1, 2)), 1), "dd/mm/yyyy")
   dDataInizio = Format(sTest, "yyyy/mm/dd")
   'Format(DateSerial(Val(Mid$(mskDataDA.Text, 4, 4)), _
                            Val(Mid$(mskDataDA.Text, 1, 2)), 1), "dd/mm/yyyy")
   If dDataFine <= dDataInizio Then
      MsgBox "Date Errate !!"
      Exit Sub
   End If
'Fine delle date
'---------------------------------------------------------------------------------
   'Apri DB con i report
'   Set ac = New Access.Application
    Set ac = CreateObject("Access.Application")
   sFile = App.Path & "\" & "reclami.mdb"
   ac.OpenCurrentDatabase sFile
   '("C:\Programmi\reclami\reclami.mdb")
'   MsgBox "Stiamo Elaborando il grafico"
     sWhere = " DataReclamo Between #" & dDataInizio & "# And #" & dDataFine & "#"
    Select Case iIndice
     Case 0
        sReport = "rGDifetto"
         sTabella = "SELECT Difetto.DescrizioneDifetto, Count(ATC.NumeroRisposta) " & _
                    "AS ConteggioDiNumeroRisposta, Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata, " & _
                    "Year([DataBolla]) AS [date] INTO [$$Difetto] " & _
                    "FROM ((Difetto INNER JOIN ATC ON Difetto.CodDifetto = ATC.CodDifetto) " & _
                    "INNER JOIN MFG ON ATC.ODL = MFG.ODL) INNER JOIN ReclamoVenditore ON ATC.ODL " & _
                    "= ReclamoVenditore.ODL " & _
                    "WHERE (ATC.Soluzione)='0' " & _
                    "AND ((ATC.Esterno)=No) AND " & sWhere & _
                    " GROUP BY Difetto.DescrizioneDifetto, Year([DataBolla]) " & _
                    "HAVING (((Year([DataBolla]))>1998))"
    
     Case 1
        sReport = "rGFamiglia12"
        
        sqlc = "drop table [$$Linea]"
        db.Execute sqlc
       
        sqlc = "SELECT MFG.LineaProdotto, Sum(MFG.Qta) AS SommaDiQta, " & _
                    "Year([DataBolla]) AS Data INTO [$$Linea] " & _
                    "From MFG " & _
                    "WHERE ((MFG.DataBolla)  Between #" & dDataInizio & "# And #" & dDataFine & "#) " & _
                    "GROUP BY MFG.LineaProdotto, Year([DataBolla])"

        db.Execute sqlc
'        ac.DoCmd.RunSQL sLinea
        
        sqlc = "Drop table  Linea"
        On Error Resume Next
        db.Execute (sqlc)
        On Error Resume Next
        
    
    
'    Set q = db.QueryDefs("qQtaLinea1")
'    q.Execute

        sqlc = "SELECT Famiglia.CodFamiglia, [$$Linea].SommaDiQta, [$$Linea].Data INTO Linea " & _
                "FROM [$$Linea] INNER JOIN Famiglia ON [$$Linea].LineaProdotto = Famiglia.Linea1"
        db.Execute sqlc
    
 '   Set q = db.QueryDefs("qQtaLinea2")
 '   q.Execute
        sqlc = "INSERT INTO Linea ( CodFamiglia, SommaDiQta, Data ) " & _
                "SELECT Famiglia.CodFamiglia, [$$Linea].SommaDiQta, [$$Linea].Data " & _
                "FROM [$$Linea] INNER JOIN Famiglia ON [$$Linea].LineaProdotto = Famiglia.Linea2"
    
        db.Execute sqlc
    
'    Set q = db.QueryDefs("qQtaLinea3")
'    q.Execute
        sqlc = "INSERT INTO Linea ( CodFamiglia, SommaDiQta, Data ) " & _
                "SELECT Famiglia.CodFamiglia, [$$Linea].SommaDiQta, [$$Linea].Data " & _
                "FROM [$$Linea] INNER JOIN Famiglia ON [$$Linea].LineaProdotto = Famiglia.Linea3"
        
        db.Execute sqlc
    
'    Set q = db.QueryDefs("qQtaLinea4")
'    q.Execute
    
        sqlc = "INSERT INTO Linea ( CodFamiglia, SommaDiQta, Data ) " & _
                "SELECT Famiglia.CodFamiglia, [$$Linea].SommaDiQta, [$$Linea].Data " & _
                "FROM [$$Linea] INNER JOIN Famiglia ON [$$Linea].LineaProdotto = Famiglia.Linea4"

        db.Execute sqlc
    
        sqlc = "Drop Table FinaleFamiglia"
        On Error Resume Next
        db.Execute (sqlc)
        On Error Resume Next
    
'    Set q = db.QueryDefs("qFinaleFamiglia")
'    q.Execute
        
        sqlc = "SELECT Linea.CodFamiglia, Sum(Linea.SommaDiQta) AS SommaDiSommaDiQta, " & _
                    "Linea.Data INTO FinaleFamiglia " & _
                "From Linea " & _
                "GROUP BY Linea.CodFamiglia, Linea.Data"
        db.Execute sqlc
        
        
'        EseguiQuery
        sqlc = "drop table [$$Fam]"
        db.Execute sqlc
       
        sqlc = " SELECT Famiglia.CodFamiglia, Famiglia.DescrizioneFamiglia, " & _
                    "Count(ATC.ODL) AS Numero, Sum(ATC.QtaAccettata) AS Accettata, " & _
                    "Year([DataReclamo]) AS [Date], Sum(MKG.Reso) AS SommaDiReso, " & _
                    "Sum(MKG.Accredito) AS SommaDiAccredito INTO [$$Fam] " & _
                    "FROM (((ATC LEFT JOIN MKG ON ATC.ODL = MKG.ODL) INNER JOIN " & _
                    "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) INNER JOIN " & _
                    "MFG ON ATC.ODL = MFG.ODL) INNER JOIN Famiglia ON ATC.CodFamiglia = " & _
                    "Famiglia.CodFamiglia " & _
                    "WHERE " & sWhere & " AND ((ATC.Esterno)=No) AND ((ATC.Soluzione)='0') " & _
                    "GROUP BY Famiglia.CodFamiglia, Famiglia.DescrizioneFamiglia, Year([DataReclamo])"
        db.Execute sqlc
'        ac.DoCmd.RunSQL sFam
        

        sqlc = "drop table [$$Grafico]"
        db.Execute sqlc
        
        sqlc = " SELECT Famiglia.DescrizioneFamiglia, Count(ATC.NumeroRisposta) AS " & _
                    "ConteggioDiNumeroRisposta, Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata, " & _
                    "Year([DataBolla]) AS [date] INTO [$$Grafico]" & _
                    "FROM ((Famiglia INNER JOIN ATC ON Famiglia.CodFamiglia = ATC.CodFamiglia) " & _
                    "INNER JOIN MFG ON ATC.ODL = MFG.ODL) INNER JOIN ReclamoVenditore ON ATC.ODL = " & _
                    "ReclamoVenditore.ODL " & _
                    "WHERE (ATC.Soluzione)='0' AND (ATC.Esterno)=No AND " & sWhere & _
                    " GROUP BY Famiglia.DescrizioneFamiglia, Year([DataBolla]) " & _
                    "HAVING (((Year([DataBolla]))>1998))"
        db.Execute sqlc
'        ac.DoCmd.RunSQL sGraficoFamiglia
        
        
        sTabella = "SELECT [$$Fam].CodFamiglia, [$$Fam].DescrizioneFamiglia, " & _
                    "[$$Fam].Numero, [$$Fam].Accettata, [$$Fam].Date, [$$Fam].SommaDiReso, " & _
                    "[$$Fam].SommaDiAccredito, FinaleFamiglia.SommaDiSommaDiQta AS Spedita " & _
                    "INTO [$$Famiglia] " & _
                    "FROM [$$Fam] INNER JOIN FinaleFamiglia ON " & _
                    "([$$Fam].Date = FinaleFamiglia.Data) AND " & _
                    "([$$Fam].CodFamiglia = FinaleFamiglia.CodFamiglia)"





     Case 2
        sReport = "rGRete"
        sReteGrafico = " SELECT Left([Rete],1) AS ReteR, Year([DataBolla]) AS Anno, " & _
                        "Sum(MFG.Qta) AS SommaDiQta INTO [$$ReteGrafico] " & _
                        "From MFG " & _
                        "WHERE (MFG.DataBolla) Between #" & dDataInizio & "# And #" & dDataFine & "# " & _
                        " GROUP BY Left([Rete],1), Year([DataBolla])"
        
        ac.DoCmd.RunSQL sReteGrafico
          
        sQtaGrafico = "SELECT Year([DataReclamo]) AS AnnoReclamo, Left([Rete],1) AS ReteR, " & _
                        "Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata INTO [$$QtaGrafico] " & _
                        "FROM (ATC INNER JOIN ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                        "INNER JOIN MFG ON ATC.ODL = MFG.ODL " & _
                        "WHERE (ATC.Soluzione)='0' AND ((ATC.Esterno)=No) AND " & sWhere & _
                        " GROUP BY Year([DataReclamo]), Left([Rete],1)"

        ac.DoCmd.RunSQL sQtaGrafico
'in questo caso stabella è recgrafico
        sTabella = "SELECT Year([DataReclamo]) AS AnnoReclamo, Left([Rete],1) AS ReteR, " & _
                    "Count(ATC.NumeroRisposta) AS ConteggioDiNumeroRisposta INTO [$$RecGrafico] " & _
                    "FROM (ATC INNER JOIN ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                    "INNER JOIN MFG ON ATC.ODL = MFG.ODL " & _
                    "WHERE (ATC.Soluzione)='0' AND ((ATC.Esterno)=No) AND " & sWhere & _
                    " GROUP BY Year([DataReclamo]), Left([Rete],1)"
        
        sGraficiReti = "SELECT UCase(Left$([Rete],1)) AS ReteRagg, Count(ATC.NumeroRisposta) " & _
                        "AS ConteggioDiNumeroRisposta, Year([DataBolla]) AS [date], " & _
                        "Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata INTO [$$GraficiReti] " & _
                        "FROM (ATC INNER JOIN ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                        "INNER JOIN MFG ON ATC.ODL = MFG.ODL " & _
                        "WHERE (ATC.Soluzione)='0' AND ((ATC.Esterno)=No) AND " & sWhere & _
                        " GROUP BY UCase(Left$([Rete],1)), Year([DataBolla]) " & _
                        "HAVING (((Year([DataBolla]))>1998))"
        
        ac.DoCmd.RunSQL sGraficiReti
        
     Case 3
        sReport = "rGEsterno"
        sTabella = "SELECT Count(ATC.NumeroRisposta) AS ConteggioDiNumeroRisposta, " & _
                    "Sum(ATC.QtaAccettata) AS SommaDiQtaAccettata, Year([DataBolla]) " & _
                    "AS [date], MFG.Esterni " & _
                    "INTO [$$Esterni] " & _
                    "FROM (ATC INNER JOIN ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                    "LEFT JOIN MFG ON ATC.ODL = MFG.ODL " & _
                    "Where (((ATC.Soluzione) = '0') And ((ATC.Esterno) = Yes)) And " & sWhere & _
                    " GROUP BY Year([DataBolla]), MFG.Esterni " & _
                    "HAVING ((Year([DataBolla]))>1998) "
      Case 4
       sReport = "rTotaleRiepilogo"
       sTabella = "SELECT [DataRisposta]-[DataREclamo] AS Media, " & _
                    "ReclamoVenditore.DataReclamo, Famiglia.DescrizioneFamiglia, " & _
                    "ATC.DataRisposta, ATC.ODL, ATC.NumeroRisposta, MFG.Qta, " & _
                    "ReclamoVenditore.QtaReclamata, ATC.QtaAccettata, MKG.Reso, " & _
                    "MKG.Accredito, MFG.Cliente, MFG.Prodotto, Difetto.DescrizioneDifetto, " & _
                    "ATC.IndiceQualita, ATC.Soluzione, ATC.Esterno INTO [$$RiepilogoTotale] " & _
                    "FROM ((((ATC LEFT JOIN MFG ON ATC.ODL = MFG.ODL) LEFT JOIN " & _
                    "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
                    "INNER JOIN Difetto ON ATC.CodDifetto = Difetto.CodDifetto) " & _
                    "INNER JOIN Famiglia ON ATC.CodFamiglia = Famiglia.CodFamiglia) " & _
                    "LEFT JOIN MKG ON ATC.ODL = MKG.ODL " & _
                    "WHERE " & sWhere

       
       
     Case 5
       sReport = "rTotaleRiepilogoFamiglie"
        sTabella = "SELECT [DataRisposta]-[DataREclamo] AS Media, " & _
              "ReclamoVenditore.DataReclamo, Famiglia.DescrizioneFamiglia, " & _
              "ATC.DataRisposta, ATC.ODL, ATC.NumeroRisposta, MFG.Qta, " & _
              "ReclamoVenditore.QtaReclamata, ATC.QtaAccettata, MKG.Reso, " & _
              "MKG.Accredito, MFG.Cliente, MFG.Prodotto, Difetto.DescrizioneDifetto, " & _
              "ATC.IndiceQualita, ATC.Soluzione, ATC.Esterno INTO [$$RiepilogoTotale] " & _
              "FROM ((((ATC LEFT JOIN MFG ON ATC.ODL = MFG.ODL) LEFT JOIN " & _
              "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
              "INNER JOIN Difetto ON ATC.CodDifetto = Difetto.CodDifetto) " & _
              "INNER JOIN Famiglia ON ATC.CodFamiglia = Famiglia.CodFamiglia) " & _
              "LEFT JOIN MKG ON ATC.ODL = MKG.ODL " & _
              "WHERE " & sWhere

     Case 6
       sReport = "rTotaleRiepilogoRisposta"
   
        sTabella = "SELECT [DataRisposta]-[DataREclamo] AS Media, " & _
              "ReclamoVenditore.DataReclamo, Famiglia.DescrizioneFamiglia, " & _
              "ATC.DataRisposta, ATC.ODL, ATC.NumeroRisposta, MFG.Qta, " & _
              "ReclamoVenditore.QtaReclamata, ATC.QtaAccettata, MKG.Reso, " & _
              "MKG.Accredito, MFG.Cliente, MFG.Prodotto, Difetto.DescrizioneDifetto, " & _
              "ATC.IndiceQualita, ATC.Soluzione, ATC.Esterno INTO [$$RiepilogoTotale] " & _
              "FROM ((((ATC LEFT JOIN MFG ON ATC.ODL = MFG.ODL) LEFT JOIN " & _
              "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) " & _
              "INNER JOIN Difetto ON ATC.CodDifetto = Difetto.CodDifetto) " & _
              "INNER JOIN Famiglia ON ATC.CodFamiglia = Famiglia.CodFamiglia) " & _
              "LEFT JOIN MKG ON ATC.ODL = MKG.ODL " & _
              "WHERE " & sWhere
   
    Case 7
        sReport = "rTotaleRiepilogoRete"
        
        sTabella = "SELECT [DataRisposta]-[DataREclamo] AS Media, " & _
                    "ReclamoVenditore.DataReclamo, Left$([Rete],1) AS ReteRagg, " & _
                    "ATC.DataRisposta, ATC.ODL, ATC.NumeroRisposta, MFG.Qta, " & _
                    "ReclamoVenditore.QtaReclamata, ATC.QtaAccettata, MKG.Reso, " & _
                    "MKG.Accredito, MFG.Cliente, MFG.Prodotto, Difetto.DescrizioneDifetto, " & _
                    "ATC.IndiceQualita, ATC.Soluzione, ATC.Esterno INTO [$$RiepilogoTotale] " & _
                    "FROM ((((ATC LEFT JOIN MFG ON ATC.ODL = MFG.ODL) LEFT JOIN " & _
                    "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) INNER JOIN " & _
                    "Difetto ON ATC.CodDifetto = Difetto.CodDifetto) INNER JOIN Famiglia " & _
                    "ON ATC.CodFamiglia = Famiglia.CodFamiglia) LEFT JOIN MKG ON ATC.ODL = MKG.ODL " & _
                    "WHERE " & sWhere
        
    Case 8
        sReport = "rTotaleRiepilogoEsterni"
        
        sTabella = "SELECT [DataRisposta]-[DataREclamo] AS Media, ReclamoVenditore.DataReclamo, " & _
                    "Famiglia.DescrizioneFamiglia, ATC.DataRisposta, ATC.ODL, " & _
                    "ATC.NumeroRisposta, MFG.Qta, ReclamoVenditore.QtaReclamata, " & _
                    "ATC.QtaAccettata, MKG.Reso, MKG.Accredito, MFG.Cliente, " & _
                    "MFG.Prodotto, Difetto.DescrizioneDifetto, ATC.IndiceQualita, " & _
                    "ATC.Soluzione, ATC.Esterno, MFG.Esterni INTO [$$RiepilogoTotaleEsterni] " & _
                    "FROM ((((ATC LEFT JOIN MFG ON ATC.ODL = MFG.ODL) LEFT JOIN " & _
                    "ReclamoVenditore ON ATC.ODL = ReclamoVenditore.ODL) INNER JOIN " & _
                    "Difetto ON ATC.CodDifetto = Difetto.CodDifetto) INNER JOIN Famiglia ON " & _
                    "ATC.CodFamiglia = Famiglia.CodFamiglia) LEFT JOIN MKG ON ATC.ODL = MKG.ODL " & _
                    "WHERE" & sWhere & " AND (ATC.Esterno)=Yes"
   
   End Select
   
   
      If iIndice < 9 Then
         'Mando in stampa il report selezionato
'         ac.DoCmd.OpenReport sReport, acViewPreview, , sWhere
             
        ac.DoCmd.RunSQL sTabella
        ac.DoCmd.OpenReport sReport, acViewPreview
'        ac.DoCmd.Maximize
'        ac.DoCmd.RunCommand acCmdAppMaximize
        ac.Visible = True
       
'       ac.DoCmd.RunCommand acCmdAppMaximize
      
      
      
      Else
'         With cryGrafici
'         On Error Resume Next
'            .Destination = crptToWindow
'            .ReportFileName = App.Path & "\" & sGrafico
'            .WindowState = crptNormal
'            If iIndice = 6 Or (iIndice = 7) Then
'               sSelFormula = "{qReteRaggruppata.Soluzione}='0' and {qReteRaggruppata.IndiceQualita}=true and {qReteRaggruppata.Esterno}=false "
'            Else
'               sSelFormula = "{ATC.Soluzione}='0' and {ATC.IndiceQualita}=true  and {ATC.Esterno}=false "
'            End If
'            If iIndice < 4 Then
'            .SelectionFormula = sSelFormula
'            End If
'
'            .Action = 1
            If Err.Number <> 0 Then
               MsgBox "Errore nel produrre il report " & vbCrLf & _
                   Err.Number & " - " & Err.Description
               Err.Clear
            End If
 '        End With
      End If
  ' MsgBox "Il grafico è stato completato"
   ' close the database
  ' ac.CloseCurrentDatabase

   
End Sub
'Private Sub EseguiQuery()
'Dim q As QueryDef, db As Database, sqlc As String
'
' On Error Resume Next
' Set db = CurrentDb
'    If Err.Number <> 0 Then
'
'        CurrentDb.Close
''        db.Close
''        ac.CloseCurrentDatabase
'
'        Set db = CurrentDb
'    End If
'
'        sqlc = "Drop table  Linea"
'        On Error Resume Next
'        db.Execute (sqlc)
'        On Error Resume Next
'
'    '    On Error GoTo 0
'        sqlc = "drop table finalefamiglia"
'        On Error Resume Next
'        db.Execute (sqlc)
'        On Error Resume Next
'
''    On Error GoTo 0
''    sqlc = "Drop table $$Famiglia"
''    On Error Resume Next
''    db.Execute (sqlc)
''    On Error GoTo 0
'
'
'
'
'    Set q = db.QueryDefs("qQtaLinea1")
'    q.Execute
'    Set q = db.QueryDefs("qQtaLinea2")
'    q.Execute
'    Set q = db.QueryDefs("qQtaLinea3")
'    q.Execute
'    Set q = db.QueryDefs("qQtaLinea4")
'    q.Execute
'    Set q = db.QueryDefs("qFinaleFamiglia")
'    q.Execute
'
'
'
'End Sub

Private Sub cmdRicerche_Click()
         Load frmRicerca
         frmRicerca.Show

End Sub

Private Sub Form_Load()
iIndice = -99
   Dim i As Byte
   For i = 1 To 12
      cmbMesi.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
   Next i
   cmbMesi.Text = Format(Month(Now), "mmmm")
   mskDataDA.Text = Format(Now, "mm/yyyy")
   Me.Show
'   mskDataDA.SetFocus
    optGrafico(0) = True
End Sub

Private Sub mskDataA_Validate(Cancel As Boolean)
   If mskDataA.Text <> "__/____" Then
      If Not IsDate("01/" & mskDataA.Text) Then
         MsgBox "Data A Non valida"
         Cancel = True
      End If
   Else
      mskDataA.Text = mskDataDA.Text
   End If
End Sub

Private Sub mskDataDA_Validate(Cancel As Boolean)
'  La DATADA NON può essere nulla !!
'   If mskDataDA.Text <> "__/____" Then
   If Not IsDate("01/" & mskDataDA.Text) Then
      MsgBox "Data DA Non valida"
      Cancel = True
   End If
'   End If
End Sub


Private Sub optGrafico_Click(Index As Integer)
iIndice = Index

End Sub

