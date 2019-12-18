VERSION 5.00
Begin VB.Form frmNeuerArtikel 
   Caption         =   "Neuer Artikel"
   ClientHeight    =   2265
   ClientLeft      =   8205
   ClientTop       =   4200
   ClientWidth     =   3225
   Icon            =   "frmNeuerArtikel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmNeuerArtikel"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3225
   Begin VB.TextBox Kommentar 
      Height          =   285
      Left            =   1320
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Eintragen"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      OLEDropMode     =   1  'Manuell
      TabIndex        =   7
      Tag             =   "6"
      Top             =   1845
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Zusätzliche Eingaben nicht löschen"
      Height          =   315
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   5
      Tag             =   "5"
      ToolTipText     =   "Klick, um das Fenster oben links zu minimieren"
      Top             =   1500
      Width           =   3075
   End
   Begin VB.ComboBox User 
      Height          =   315
      Left            =   1320
      OLEDropMode     =   1  'Manuell
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Tag             =   "3"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Gruppe 
      Height          =   285
      Left            =   2520
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   2
      Tag             =   "2"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Gebot 
      Height          =   285
      Left            =   1320
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   1
      Tag             =   "2"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   -360
   End
   Begin VB.TextBox Artikel 
      Height          =   285
      Left            =   1320
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   0
      Tag             =   "1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Fenster immer sichtbar"
      Height          =   435
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   6
      Tag             =   "6"
      ToolTipText     =   "Klick, um das Fenster oben links zu minimieren"
      Top             =   1785
      Width           =   1995
   End
   Begin VB.Label Label4 
      Caption         =   "Kommentar"
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   11
      Tag             =   "4"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Benutzer"
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   10
      Tag             =   "3"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Gebot / Gruppe"
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   9
      Tag             =   "2"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Artikel"
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   8
      Tag             =   "1"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmNeuerArtikel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'
' BOM default header
'
' this file is OpenSource.
' license model: GPL
' please respect the limitations
'
' main language: german
' compiled under VB6 SP5 german
'
' $author: hjs$
' $id: V 2.0.2 date 030303 hjs$
' $version: 2.0.2$
' $file: $
'
' last modified:
' &date: 030303$
'
' contact: visit http://de.groups.yahoo.com/group/BOMInfo
'
'*******************************************************
Option Explicit
Private mbCodeResize As Boolean
Private Sub SwapControls(ByVal iControl1 As Integer, ByVal iControl2 As Integer)
    
    Dim i As Integer
    Dim sTmp As String
    Dim iPosOld As Integer
    Dim iPosNew As Integer
    
    If iControl1 <> iControl2 Then
        i = 1
        sTmp = "," & gsNewItemWindowWidgetOrdner & ","
        
        iPosOld = InStr(1, sTmp, "," & CStr(iControl1) & ",")
        iPosNew = InStr(1, sTmp, "," & CStr(iControl2) & ",")
        
        sTmp = gsNewItemWindowWidgetOrdner
        gsNewItemWindowWidgetOrdner = ""
        
        Do While (i > 0)
            i = Val(GetBisZeichen(sTmp, ","))
            If i > 0 Then
                If i = iControl1 Then
                    'nichts
                ElseIf i = iControl2 Then
                    If iPosOld > iPosNew Then
                        gsNewItemWindowWidgetOrdner = gsNewItemWindowWidgetOrdner & "," & CStr(iControl1) & "," & CStr(i)
                    Else
                        gsNewItemWindowWidgetOrdner = gsNewItemWindowWidgetOrdner & "," & CStr(i) & "," & CStr(iControl1)
                    End If
                Else
                    gsNewItemWindowWidgetOrdner = gsNewItemWindowWidgetOrdner & "," & CStr(i)
                End If
            End If
        Loop
        gsNewItemWindowWidgetOrdner = Mid(gsNewItemWindowWidgetOrdner, 2)
        Call Form_Resize
    End If 'iControl1 <> iControl2
End Sub

Private Sub Artikel_DragDrop(Source As Control, X As Single, Y As Single)
    Call SwapControls(Source.Tag, Artikel.Tag)
End Sub

Private Sub Artikel_KeyDown(KeyCode As Integer, Shift As Integer)
' lg 14.05.03
  If Not (KeyCode = vbKeyMenu And Shift = vbAltMask) Then  '18 , 4
    If Timer.Enabled Then
      Timer.Enabled = False
      Call SetBackColor(&H8000000F)
      Me.Caption = gsarrLangTxt(223)
    End If
  End If
End Sub

Private Sub Artikel_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub btnOk_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Check1_Click()
  On Error Resume Next
  Call SetForeground(Check1.Value > 0, hWnd)
  Artikel.SetFocus
End Sub

Private Sub Check1_DragDrop(Source As Control, X As Single, Y As Single)
Call SwapControls(Source.Tag, Check1.Tag)
End Sub

Private Sub Check2_Click()
  On Error Resume Next
  Artikel.SetFocus
End Sub

Private Sub Check1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Check2_DragDrop(Source As Control, X As Single, Y As Single)
    Call SwapControls(Source.Tag, Check2.Tag)
End Sub

Private Sub Check2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Form_Activate()
  On Error Resume Next
  Artikel.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim v As Variant
    
    If KeyCode = vbKeyControl Then  '17
        Me.MousePointer = vbSizePointer
        For Each v In Me.Controls
            v.DragMode = 1
        Next
    End If
    
    If KeyCode = vbKeyEscape Then  '27
        Unload Me
    End If
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    Dim v As Variant
    
    If KeyCode = vbKeyControl Then  '17
        Me.MousePointer = 0
        For Each v In Me.Controls
            v.DragMode = 0
        Next
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Me.Icon = MyLoadResPicture(101, 16)
    Call SendMessage(Me.hWnd, WM_SETICON, 0, Me.Icon)
    Me.Caption = gsarrLangTxt(223)
    Label1.Caption = gsarrLangTxt(224)
    Label2.Caption = gsarrLangTxt(68) & "/" & gsarrLangTxt(69)
    Label3.Caption = gsarrLangTxt(360)
    Label4.Caption = gsarrLangTxt(220)
    Check1.Caption = gsarrLangTxt(225)
    Check2.Caption = gsarrLangTxt(373)
    btnOk.Caption = gsarrLangTxt(226)
    Check1.ToolTipText = gsarrLangTxt(318)
    Me.OLEDropMode = vbOLEDropManual
    Check1.Value = IIf(gbNewItemWindowAlwaysOnTop, vbChecked, vbUnchecked)
    Check2.Value = IIf(gbNewItemWindowKeepsValues, vbChecked, vbUnchecked)
    Call Check1_Click
    Me.Height = 2775
    Me.Left = glNeuerArtikelLeft
    Me.Top = glNeuerArtikelTop
    Me.Height = glNeuerArtikelHeight
    Me.Width = glNeuerArtikelWidth
    
    User.AddItem gsarrLangTxt(47)
    User.ListIndex = 0
  
    For i = 1 To UBound(gtarrUserArray)
        User.AddItem gtarrUserArray(i).UaUser
    Next i
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    Dim myNewHeight As Double
    Dim hRect As RECT
    Dim WindowSize As Integer
    Dim WidgetOrder(0 To 6) As Integer
    Dim WidgetHeight(0 To 6) As Integer
    Dim WidgetHeightSum(0 To 7) As Integer ' einer mehr
    Dim MaxWindowHeight As Double
    Dim i As Integer
    Dim sTmp As String
    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim sToDo As String
    Dim v As Variant
    Dim iLabelMaxWidth As Integer
    
    
    If Not mbCodeResize Then
        mbCodeResize = True
          
        Call GetClientRect(Me.hWnd, hRect)
        
        'die Höhe der Titelleiste berücksichtigen
        MaxWindowHeight = 2820 - 515 + Me.Height - (hRect.Bottom - hRect.Top) * Screen.TwipsPerPixelY
        MaxWindowHeight = MaxWindowHeight / Screen.TwipsPerPixelY
        MaxWindowHeight = Int(MaxWindowHeight)
        MaxWindowHeight = MaxWindowHeight * Screen.TwipsPerPixelY
        
        WidgetHeight(0) = 105 ' Rest
        WidgetHeight(1) = 345 ' Artikel
        WidgetHeight(2) = 345 ' Gebot / Gruppe
        WidgetHeight(3) = 390 ' User
        WidgetHeight(4) = 345 ' Kommentar
        WidgetHeight(5) = 330 ' KeepValues
        WidgetHeight(6) = 435 ' OK / Always on Top
        
        sTmp = gsNewItemWindowWidgetOrdner
        Y = WidgetHeight(0)
        Z = 0
        
        sToDo = "-1-2-3-4-5-6"
        For i = 1 To UBound(WidgetHeight)
            X = 9
            Do While (X > UBound(WidgetHeight))
                X = Val(GetBisZeichen(sTmp, ","))
            Loop
            If X = 0 Then
                X = Mid(sToDo, InStr(1, sToDo, "-") + 1, 1)
                gsNewItemWindowWidgetOrdner = gsNewItemWindowWidgetOrdner & "," & CStr(X)
            End If
            WidgetOrder(i) = X
            Select Case X
                Case 1
                    Artikel.Top = Y
                    Artikel.TabIndex = Z
                    Label1.Top = Y + 30
                    Z = Z + 1
                Case 2
                    Gebot.Top = Y
                    Gebot.TabIndex = Z
                    Gruppe.Top = Y
                    Gruppe.TabIndex = Z + 1
                    Label2.Top = Y + 30
                    Z = Z + 2
                Case 3
                    User.Top = Y
                    User.TabIndex = Z
                    Label3.Top = Y + 30
                    Z = Z + 1
                Case 4
                    Kommentar.Top = Y
                    Kommentar.TabIndex = Z
                    Label4.Top = Y + 30
                    Z = Z + 1
                Case 5
                    Check2.Top = Y
                    Check2.TabIndex = Z
                    Z = Z + 1
                Case 6
                    Check1.Top = Y
                    Check1.TabIndex = Z
                    btnOk.Top = Y
                    btnOk.TabStop = True
                    btnOk.TabIndex = Z + 1
                    Z = Z + 2
            End Select
            sToDo = Replace(sToDo, "-" & CStr(X), "+" & CStr(X))
            Y = Y + WidgetHeight(X)
        Next i
        
        If Left(gsNewItemWindowWidgetOrdner, 1) = "," Then gsNewItemWindowWidgetOrdner = Mid(gsNewItemWindowWidgetOrdner, 2)
        
        For i = UBound(WidgetHeight) To 0 Step -1
            WidgetHeightSum(i) = WidgetHeight(WidgetOrder(i)) + IIf(i < UBound(WidgetHeight), WidgetHeightSum(i + 1), 0)
        Next i
        
        If Me.Height >= MaxWindowHeight Then
            myNewHeight = MaxWindowHeight
            WindowSize = 6
        ElseIf Me.Height >= MaxWindowHeight - WidgetHeightSum(6) Then
            myNewHeight = MaxWindowHeight - WidgetHeightSum(6)
            WindowSize = 5
        ElseIf Me.Height >= MaxWindowHeight - WidgetHeightSum(5) Then
            myNewHeight = MaxWindowHeight - WidgetHeightSum(5)
            WindowSize = 4
        ElseIf Me.Height >= MaxWindowHeight - WidgetHeightSum(4) Then
            myNewHeight = MaxWindowHeight - WidgetHeightSum(4)
            WindowSize = 3
        ElseIf Me.Height >= MaxWindowHeight - WidgetHeightSum(3) Then
            myNewHeight = MaxWindowHeight - WidgetHeightSum(3)
            WindowSize = 2
        ElseIf Me.Height >= MaxWindowHeight - WidgetHeightSum(2) Then
            myNewHeight = MaxWindowHeight - WidgetHeightSum(2)
            WindowSize = 1
        Else
            myNewHeight = 0
            WindowSize = 0
        End If
          
        For Each v In Me.Controls
            v.Visible = True
        Next
        
        For i = WindowSize + 1 To UBound(WidgetHeight)
            Select Case WidgetOrder(i)
                Case 1
                    Artikel.Visible = False
                    Label1.Visible = False
                Case 2
                    Gebot.Visible = False
                    Gruppe.Visible = False
                    Label2.Visible = False
                Case 3
                    User.Visible = False
                    Label3.Visible = False
                Case 4
                    Kommentar.Visible = False
                    Label4.Visible = False
                Case 5
                    Check2.Visible = False
                Case 6
                    Check1.Visible = False
                    btnOk.Left = -10000
                    btnOk.TabStop = False
            End Select
        Next i
        
        
        For i = 1 To 9
            With Me.Controls("Label" & CStr(i))
                .AutoSize = True
                If .Visible And .Width > iLabelMaxWidth Then iLabelMaxWidth = .Width
            End With
        Next i
        
        Artikel.Left = Label1.Left + iLabelMaxWidth + 120
        Artikel.Width = Me.ScaleWidth - Artikel.Left - Label1.Left
        Gebot.Left = Artikel.Left
        Gebot.Width = (Artikel.Width - 120) / 5 * 3
        Gruppe.Left = Artikel.Left + Gebot.Width + 120
        Gruppe.Width = (Artikel.Width - 120) / 5 * 2
        User.Left = Artikel.Left
        User.Width = Artikel.Width
        Kommentar.Left = Artikel.Left
        Kommentar.Width = Artikel.Width
        If btnOk.TabStop Then btnOk.Left = Me.ScaleWidth - btnOk.Width - 120
        Check1.Width = Me.ScaleWidth * 3
        Check2.Width = Me.ScaleWidth * 3
        
        Me.Height = myNewHeight
        mbCodeResize = False
    End If 'Not CodeResize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    With Me
        glNeuerArtikelTop = .Top
        glNeuerArtikelLeft = .Left
        glNeuerArtikelHeight = .Height
        glNeuerArtikelWidth = .Width
    End With
    
End Sub

Private Sub Form_Deactivate()
    
    With Me
        glNeuerArtikelTop = .Top
        glNeuerArtikelLeft = .Left
        glNeuerArtikelHeight = .Height
        glNeuerArtikelWidth = .Width
    End With
    
End Sub

Private Sub Gebot_DragDrop(Source As Control, X As Single, Y As Single)
Call SwapControls(Source.Tag, Gebot.Tag)
End Sub

Private Sub Gebot_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Gruppe_DragDrop(Source As Control, X As Single, Y As Single)
    Call SwapControls(Source.Tag, Gruppe.Tag)
End Sub

Private Sub Gruppe_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Kommentar_DragDrop(Source As Control, X As Single, Y As Single)
    Call SwapControls(Source.Tag, Kommentar.Tag)
End Sub

Private Sub Kommentar_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
    Call SwapControls(Source.Tag, Label1.Tag)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
    Call SwapControls(Source.Tag, Label2.Tag)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
    Call SwapControls(Source.Tag, Label3.Tag)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
    Call SwapControls(Source.Tag, Label4.Tag)
End Sub

Private Sub Label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Label2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Label3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Label4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub btnOk_Click()
    Dim i As Integer
    Dim sTmp As String
    Dim strsave As String
    Dim arr As Variant
    
    strsave = Artikel.Text
    
    sTmp = Artikel.Text
    sTmp = Replace(sTmp, ":", ",")
    sTmp = Replace(sTmp, ";", ",")
    sTmp = Replace(sTmp, "/", ",")
    sTmp = Replace(sTmp, "-", ",")
    sTmp = Replace(sTmp, ".", ",")
    sTmp = Replace(sTmp, "+", ",")
    sTmp = Replace(sTmp, "|", ",")
    
    
    arr = Split(sTmp, ",")
    
    
    For i = LBound(arr) To UBound(arr)
        Artikel.Text = arr(i)
        Call HandleDragDropData(Artikel.Text)
    Next i
    Call HandleDragDropData(strsave)
    
    Timer.Enabled = False
    Timer.Enabled = True
    
    On Error Resume Next
    Artikel.SetFocus
    
End Sub

Private Sub btnOk_DragDrop(Source As Control, X As Single, Y As Single)
  SwapControls Source.Tag, btnOk.Tag
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Do_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  Dim str As String
  Dim i As Long
  
  For i = 1 To 255
    If Data.GetFormat(i) Then
      str = Data.GetData(i)
      If str > "" Then Exit For
    End If
  Next i
  
  Call HandleDragDropData(str)
  Timer.Enabled = False
  Timer.Enabled = True

End Sub

Private Sub HandleDragDropData(sData As String)
    On Error Resume Next
    Dim sBuffer As String
    Dim i As Integer
    Dim sArtikel As String
    Dim arr As Variant
    
    Call SetBackColor(vbRed)
    Me.Caption = gsarrLangTxt(227)
    sBuffer = sData
    
    sArtikel = GetItemFromUrl(sBuffer)
    If Len(sArtikel) > 0 Then ' schon was da
        arr = Array(sArtikel)
    Else
        arr = ResolveItemUrl(sBuffer)
    End If
    
    For i = LBound(arr) To UBound(arr)
        
        sArtikel = GetItemFromUrl(arr(i))
        
        If Len(sArtikel) > 0 And Not frmHaupt.CheckItemBlacklist(sArtikel) Then
            Call InsertArtikelBuff(MakeArtikelBuffText(sArtikel))
            Artikel.Text = sArtikel & gsarrLangTxt(228)
            Me.Caption = Artikel.Text
            Call SetBackColor(vbGreen)
        ElseIf LBound(arr) = UBound(arr) Then
            Me.Caption = "?"
        End If
        
    Next i
    
End Sub

Private Sub Timer_Timer()

  Timer.Enabled = False
  Call SetBackColor(&H8000000F)
  Me.Caption = gsarrLangTxt(223)
  Artikel.Text = ""
  If Check2.Value = vbUnchecked Then
    Gebot.Text = ""
    Gruppe.Text = ""
    User.ListIndex = 0
    Kommentar.Text = ""
  End If
  
End Sub

Private Sub GetTestArtikel()

    'MD-Marker , Sub wird nicht aufgerufen
    'L: Ist okay, wird nur zum Testen benutzt

'  Dim sBuffer As String
'  sBuffer = ShortPost("http://hub.ebay.de/buy")
'
'  Dim pos As Long
'  Dim pos2 As Long
'  Dim tmp As String
'  Dim arr As Variant
'  Dim sArtikel As String
'  Dim i As Integer
'
'  pos = InStr(pos + 1, sBuffer, "http://")
'  Do While (pos > 0)
'
'    pos2 = InStr(pos, sBuffer, """")
'    tmp = Mid(sBuffer, pos, pos2 - pos - 1)
'    If InStr(1, tmp, "_W0") > 0 Then
'
'      arr = ResolveItemUrl(tmp)
'
'      For i = LBound(arr) To UBound(arr)
'
'        sArtikel = GetItemFromUrl(arr(i))
'
'        If Len(sArtikel) > 0 Then
'          'Debug.Print sArtikel
'          InsertArtikelBuff MakeArtikelBuffText(sArtikel)
'          Artikel = sArtikel & gsarrLangTxt(228)
'        End If
'
'      Next i
'
'    End If
'
'    pos = InStr(pos + 1, sBuffer, "http://")
'  Loop

End Sub

Private Function GetGebot() As Double

  On Error Resume Next
  GetGebot = Gebot.Text

End Function

Private Function GetGruppe() As String

  On Error Resume Next
  GetGruppe = Trim(Gruppe.Text)

End Function

Private Function GetKommentar() As String

  On Error Resume Next
  Kommentar.Text = Replace(Kommentar.Text, vbTab, " ")
  GetKommentar = Trim(Kommentar.Text)

End Function

Private Function GetUser() As String

  On Error Resume Next
  If User.ListIndex > 0 Then
    GetUser = User.Text
  Else
    GetUser = ""
  End If

End Function

Private Function MakeArtikelBuffText(sArtikel As String) As String

  MakeArtikelBuffText = sArtikel & vbTab & GetGebot & vbTab & GetGruppe() & vbTab & GetUser() & vbTab & GetKommentar

End Function

Private Sub SetBackColor(lColor As Long)
    
    Dim v As Variant
    For Each v In Me.Controls
        If TypeName(v) = "Label" Or TypeName(v) = "CheckBox" Then
            v.BackColor = lColor
        End If
    Next
    Me.BackColor = lColor
    
End Sub

Private Sub User_DragDrop(Source As Control, X As Single, Y As Single)
    Call SwapControls(Source.Tag, User.Tag)
End Sub

Private Sub User_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Do_OLEDragDrop(0, Data, Effect, Button, Shift, X, Y)
End Sub
