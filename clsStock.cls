VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarStockSymbol As String 'local copy
Private mvarStockName As String 'local copy

Private mvarLastTrade As Currency 'local copy
Private mvarChange As String 'local copy
'local variable(s) to hold property value(s)
Private mvarVolume As Double 'local copy
Public Property Get Volume() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Volume
    Volume = mvarVolume
End Property



Public Property Get Change() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Change
    Change = mvarChange
End Property
Public Property Get LastTrade() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastTrade
    LastTrade = mvarLastTrade
End Property



Public Sub GetInfo()

    On Error Resume Next
    
    Dim strTextInfo As String
    Dim intStartTrade As Integer
    Dim intEndTrade As Integer
    Dim strTrade As String
    Dim intStartChange As Integer
    Dim intEndChange As Integer
    Dim strChange As String
    Dim intStartVol As Integer
    Dim intEndVol As Integer
    Dim strVol As String
    
    strTextInfo = Form1.Inet1.OpenURL("http://finance.yahoo.com/q?s=" & mvarStockSymbol)

    intStartTrade = InStr(strTextInfo, "Last Trade:") + 41 + Len("Last Trade:")
    intEndTrade = InStr(intStartTrade, strTextInfo, "<")
    strTrade = Trim(Mid(strTextInfo, intStartTrade, intEndTrade - intStartTrade))
    mvarLastTrade = CCur(strTrade)


    intStartChange = InStr(strTextInfo, "Change:")
    intStartChange = instartchange + InStr(intStartChange, strTextInfo, ";" & """>")
    intStartChange = intStartChange + Len(";" & """>")
    intEndChange = InStr(intStartChange, strTextInfo, "<")
    strChange = Trim(Mid(strTextInfo, intStartChange, intEndChange - intStartChange))
    mvarChange = strChange
    
    intStartVol = InStr(strTextInfo, "Volume:") + Len("volume:")
    intStartVol = instartvol + InStr(intStartVol, strTextInfo, "" & """>")
    intStartVol = intStartVol + Len("" & """>")
    intEndVol = InStr(intStartVol, strTextInfo, "<")
    strVol = Trim(Mid(strTextInfo, intStartVol, intEndVol - intStartVol))
    mvarVolume = Format(CDbl(strVol), "###,###,###")
    
    
End Sub

Public Property Let StockName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StockName = 5
    mvarStockName = vData
End Property


Public Property Get StockName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StockName
    StockName = mvarStockName
End Property



Public Property Let StockSymbol(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StockSymbol = 5
    mvarStockSymbol = vData
End Property


Public Property Get StockSymbol() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StockSymbol
    StockSymbol = mvarStockSymbol
End Property



