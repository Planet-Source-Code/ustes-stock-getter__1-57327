VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "eTech Stock Lister"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetAll 
      Caption         =   "Get All Stocks Info"
      Height          =   495
      Left            =   210
      TabIndex        =   3
      Top             =   5400
      Width           =   1560
   End
   Begin MSComctlLib.ListView lvwStocks 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   8916
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Symbol"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Trade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Change"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Volume"
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   510
      Left            =   5460
      TabIndex        =   1
      Top             =   5355
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   900
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton cmdGetStockInfo 
      Caption         =   "Get Stock(s) Info"
      Height          =   495
      Left            =   7950
      TabIndex        =   0
      Top             =   5415
      Width           =   1560
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6300
      Top             =   5325
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject
Dim StockText As TextStream

Private Sub cmdGetAll_Click()

    Dim i As Integer
    Dim clsTemp As New clsStock
    
    Set clsTemp = New clsStock
    
    For i = 1 To lvwStocks.ListItems.Count
        clsTemp.StockSymbol = lvwStocks.ListItems(i).Text
        clsTemp.GetInfo
        lvwStocks.ListItems(i).SubItems(2) = clsTemp.LastTrade
        lvwStocks.ListItems(i).SubItems(3) = clsTemp.Change
        lvwStocks.ListItems(i).SubItems(4) = clsTemp.Volume
        DoEvents
    Next i


End Sub

Private Sub cmdGetStockInfo_Click()
    
    Dim i As Integer
    Dim clsTemp As New clsStock
    
    Set clsTemp = New clsStock
    
    ' Ghost selected ListItem.
    If lvwStocks.SelectedItem Is Nothing Then Exit Sub
    For i = 1 To lvwStocks.ListItems.Count
       If lvwStocks.ListItems(i).Selected = True Then
          clsTemp.StockSymbol = lvwStocks.ListItems(i).Text
          clsTemp.GetInfo
          lvwStocks.ListItems(i).SubItems(2) = clsTemp.LastTrade
          lvwStocks.ListItems(i).SubItems(3) = clsTemp.Change
          lvwStocks.ListItems(i).SubItems(4) = clsTemp.Volume
          
          lvwStocks.ListItems(i).Ghosted = True
       End If
    Next i

    
End Sub
Sub GetStockSymbols()
    
    Dim strTemp As String
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim liStock As ListItem
    Dim clsAddStock As clsStock
    
    Screen.MousePointer = vbHourglass
    
    Set clsAddStock = New clsStock
    
    RichTextBox1.Text = Inet1.OpenURL("http://www.investorguide.com/stocklistticker.html")
    RichTextBox1.SaveFile (App.Path & "\test.txt")
    
    Set StockText = fso.OpenTextFile(App.Path & "\test.txt")

    Do While Not StockText.AtEndOfStream
        strTemp = StockText.ReadLine
        If InStr(strTemp, "InvestorGuide.com's research") <> 0 Then
            intStart = StockText.Line
        End If
       If InStr(strTemp, "<!-- END PAGE CONTENT -->") <> 0 Then
            intEnd = StockText.Line
            Exit Do
        End If
    Loop
        
    StockText.Close
    Set StockText = fso.OpenTextFile(App.Path & "\test.txt")
    
    Do While Not StockText.AtEndOfStream
        strTemp = StockText.ReadLine
        Select Case StockText.Line
            Case intStart + 2 To intEnd - 2
                Set clsAddStock = ParseSymbol(strTemp)
                Set liStock = lvwStocks.ListItems.Add(, , clsAddStock.StockSymbol)
                liStock.SubItems(1) = clsAddStock.StockName
'                liStock.SubItems(2) = clsAddStock.LastTrade
'                liStock.SubItems(3) = clsAddStock.Change
'                liStock.SubItems(4) = clsAddStock.Volume
                colStockList.Add clsAddStock.Key, clsAddStock.StockSymbol, clsAddStock.StockName
                
        End Select
        DoEvents
    Loop
    
    StockText.Close
    Set fso = Nothing
    Screen.MousePointer = vbDefault
End Sub
Function ParseSymbol(strLine As String) As clsStock
    Dim strName As String
    Dim strSymbol As String
    Const intNameStart As Integer = 5
    Dim intNameEnd As Integer
    Dim intSymbolStart As Integer
    Dim intSymbolEnd As Integer
    
    Set ParseSymbol = New clsStock
    
    intNameEnd = InStr(intNameStart, strLine, "<b>")
    strName = Trim(Mid(strLine, intNameStart, intNameEnd - intNameStart - 2))
    
    
    intSymbolStart = InStr(strLine, "<a href='/cgi-bin/research.cgi?name=")
    intSymbolStart = intSymbolStart + Len("<a href='/cgi-bin/research.cgi?name=")
    intSymbolEnd = InStr(intSymbolStart, strLine, "'")
    strSymbol = Trim(Mid(strLine, intSymbolStart, intSymbolEnd - intSymbolStart))
    
    ParseSymbol.Key = colStockList.Count + 1
    ParseSymbol.StockName = strName
    ParseSymbol.StockSymbol = strSymbol

End Function

Private Sub Form_Load()
    Dim strMSG As String
    
    Set colStockList = New colStocks
        
    strMSG = "Double click item to get individual stock info or multi-select and press 'Get Stock Info' buttom"
    lvwStocks.ToolTipText = strMSG
    
    lvwStocks.Visible = False
    GetStockSymbols
    lvwStocks.Visible = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lvwStocks_DblClick()

    Dim clsTemp As New clsStock
    Dim strMSG As String
    
    Set clsTemp = New clsStock

    clsTemp.StockSymbol = lvwStocks.SelectedItem.Text
    clsTemp.GetInfo
    If clsTemp.LastTrade <> 0 Then
        lvwStocks.SelectedItem.SubItems(2) = clsTemp.LastTrade
        lvwStocks.SelectedItem.SubItems(3) = clsTemp.Change
        lvwStocks.SelectedItem.SubItems(4) = clsTemp.Volume
    Else
        strMSG = "Symbol: " & clsTemp.StockSymbol & vbCrLf
        strMSG = strMSG & "Name: " & clsTemp.StockName & vbCrLf
        strMSG = strMSG & "This appears to be an invalid stock symbol.  It will be removed from the list"
        MsgBox strMSG, vbCritical + vbOKOnly, "eStockGetter"
        lvwStocks.ListItems.Remove (lvwStocks.SelectedItem.Index)
    End If
    
End Sub
