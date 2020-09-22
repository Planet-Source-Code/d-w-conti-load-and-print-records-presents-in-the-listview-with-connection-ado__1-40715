VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: - >> Carica e stampa DB tramite ADO & SQL"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   5400
      Width           =   4695
      Begin VB.OptionButton optTutti 
         Caption         =   "Tutti"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optSelezione 
         Caption         =   "Selezione"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Stampa Record"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList imglstListImages 
      Left            =   360
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":08CA
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":0A24
            Key             =   "Down"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCaricaDB 
      Caption         =   "Carica DB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2483
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin MSComctlLib.ListView LvSearch 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7435
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imglstListImages"
      SmallIcons      =   "imglstListImages"
      ColHdrIcons     =   "imglstListImages"
      ForeColor       =   -2147483640
      BackColor       =   8454143
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cognome"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sesso"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Programma creato da Conti Davide - Contattatemi tramite mail daw_conti@libero.it"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   100
      Width           =   6735
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    '<< CHIUDE IL PROGRAMMA
    End
End Sub

Private Sub cmdCaricaDB_Click()
Dim itmX As ListItem
Dim DataBases As Database
Dim RSocio As Recordset

Dim StrSQL As String
Dim CondizioniRicerca As Boolean

'<< FORMATTIAMO LA LISTVIEW X VISUALIZZARE I DATI
LvSearch.ListItems.Clear
    
'<< CREIAMO LA STRINGA SQL X CARICARE I DATI NELLA LISTVIEW
'<< Carica tutti i dati presenti nel DataBase
StrSQL = "SELECT * FROM Socio" ' WHERE Cognome Like " & " '" & Text1.Text & "*'" & " order by Cognome"

    
'<< CARICHIAMO IL PERCORSO E APRIAMO IL DATABASE
Set DataBases = OpenDatabase(App.Path & "\DatiSocio.MDB")

'<< CARICHIAMO E APRIAMO I RECORDSET CARICATI CON LA STRINGA SQL
Set RSocio = DataBases.OpenRecordset(StrSQL)

'<< Aggiunge i record del DB nella ListView
Do Until RSocio.EOF

    Set itmX = LvSearch.ListItems.Add()
    
    With RSocio
        itmX.Text = .Fields("Cognome")
        itmX.SubItems(1) = .Fields("Nome")
        itmX.SubItems(2) = .Fields("Sesso")
        CondizioniRicerca = True
        RSocio.MoveNext
    End With
Loop

'<< Se i record sono stati caricati vengono visualizzati nella LV altrimenti
'<< appare una msgbox
If CondizioniRicerca = False Then
    LvSearch.ListItems.Add.Text = "Ricerca completata, nessun record trovato..."
End If

LvSearch.GridLines = True

End Sub

Private Sub cmdprint_Click()
'<< Stampa il record scelto tramite checkbox
If optSelezione = True Then
      Dim i As Integer
      Dim i2 As Integer
      Dim ItemChecked As Boolean

      ItemChecked = False

      For i = 1 To LvSearch.ListItems.Count
         If LvSearch.ListItems(i).Checked = True Then
            ItemChecked = True
            Exit For
         End If
      Next i

      If ItemChecked = True Then
         Printer.Font = "Tahoma"
         Printer.FontBold = False
         Printer.FontUnderline = False
         Printer.FontSize = 10
         Printer.Print vbNewLine
         Printer.Print "Esempio di stampa record selezionati"
         Printer.Print
         Printer.Print "Carica e stampa DB in una LV tramite SQL & ADO"
         Printer.Print "Visualizza i record in una LV con possibilità di ordinarle"
         Printer.Print "Sono abilitate le icone al tipo di ordinamento della colonna"
         Printer.Print
         Printer.Print "Programma creato da Davide Conti il 25.09.2002"
         Printer.Print
         Printer.FontUnderline = False
         Printer.FontBold = False
         Printer.Print vbNewLine

         i2 = 0

         For i = 1 To LvSearch.ListItems.Count
            If LvSearch.ListItems(i).Checked = True Then
               i2 = i2 + 1
               Printer.Print Space(6) & "Cognome : " & Str$(i2)
               Printer.Print Space(6) & "Nome : " & LvSearch.ListItems(i).ListSubItems(1).Text
               Printer.Print Space(6) & "Sesso : " & LvSearch.ListItems(i).ListSubItems(2).Text
               Printer.Print vbNewLine
            End If
         Next i
         Printer.EndDoc
      End If

   Else   '<< Se la selezione è contraria stampa tutti i record presenti nella LV

      If LvSearch.ListItems.Count > 0 Then
         Printer.Font = "Tahoma"
         Printer.FontBold = True
         Printer.FontUnderline = False
         Printer.FontSize = 10
         Printer.Print "Esempio di stampa tutti i record presenti nella LV"
         Printer.Print
         Printer.Print "Carica e stampa DB in una LV tramite SQL & ADO"
         Printer.Print "Visualizza i record in una LV con possibilità di ordinarle"
         Printer.Print "Sono abilitate le icone al tipo di ordinamento della colonna"
         Printer.Print
         Printer.Print "Programma creato da Davide Conti il 25.09.2002"
         Printer.Print
         Printer.FontUnderline = False
         Printer.FontBold = False
         Printer.Print vbNewLine

         For i = 1 To LvSearch.ListItems.Count
            Printer.Print Space(6) & "Cognome : " & Str$(i)
            Printer.Print Space(6) & "Nome : " & LvSearch.ListItems(i).ListSubItems(1).Text
            Printer.Print Space(6) & "Sesso : " & LvSearch.ListItems(i).ListSubItems(2).Text
            Printer.Print vbNewLine
         Next i
         Printer.EndDoc
      End If
   End If
End Sub

Private Sub lvSearch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'<< ORDINA LE COLONNE DEI RECORDS E
'<< AGGIUNGE L'ICONA DEL TIPO DI ORDINAMENTO COLONNA ALLE COLONNE

    '<< ABILITA L'IMPOSTAZIONE ORDINA RECORD
    LvSearch.Sorted = True

    If LvSearch.SortOrder = lvwAscending Then
        '<< SE I RECORD SONO ASCENDENTI DIVENTANO DISCENDENTI
        LvSearch.SortOrder = lvwDescending
    Else
        '<< SE I RECORD SONO DISCENDENTI DIVENTANO ASCENDENTI
        LvSearch.SortOrder = lvwAscending
    End If

    '<< CAMBIANDO IL MODO DI ORDINAMENTO VIENE VISUALIZZATA L'ICONA
    '<< IN BASE ALL'ORDINAMENTO
    If LvSearch.ColumnHeaders(ColumnHeader.Index).Icon = 0 Or _
    LvSearch.ColumnHeaders(ColumnHeader.Index).Icon = "Up" Then
        LvSearch.ColumnHeaders(ColumnHeader.Index).Icon = "Down"
        GoTo ClearAllOthers
    End If

    '<< CAMBIANDO IL MODO DI ORDINAMENTO VIENE VISUALIZZATA L'ICONA
    '<< IN BASE ALL'ORDINAMENTO
    If LvSearch.ColumnHeaders(ColumnHeader.Index).Icon = 0 Or _
    LvSearch.ColumnHeaders(ColumnHeader.Index).Icon = "Down" Then
        LvSearch.ColumnHeaders(ColumnHeader.Index).Icon = "Up"
        GoTo ClearAllOthers
    End If
'<< IMPOSTAZIONI PER CARICARE L'ICONA NELLE COLONNE DELLA LISTVIEW
ClearAllOthers:
    ' setup a counter variable
    Dim lngIndex As Long

    ' loop through all of the column headers
    For lngIndex = 1 To LvSearch.ColumnHeaders.Count - 1
        ' except the current one
        If lngIndex <> ColumnHeader.Index Then
            ' and if it has an 'up' or 'down' then
            If LvSearch.ColumnHeaders(lngIndex).Icon = "Up" Or _
            LvSearch.ColumnHeaders(lngIndex).Icon = "Down" Then
                ' dectroy it's icon
                LvSearch.ColumnHeaders(lngIndex).Icon = 0
            End If
        End If
    Next lngIndex
End Sub


