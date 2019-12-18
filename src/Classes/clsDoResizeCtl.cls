VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoResizeCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
' $author: susi/hjs$
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

'thx to Susi
'*****************************************************************************
'*** Verbindet Control mit Resize-Konstanten; wird fuer Klasse DoResize
'*** benoetigt
'*****************************************************************************

Private moControl As Control
Private meResizeBehaviour As DoResizeEnum

Private mfInitLeft As Double '*** Abstaende zum FormRand bei Start
Private mfInitRight As Double
Private mfInitTop As Double
Private mfInitBottom As Double
Private mfInitWidth As Double
Private mfInitHeight As Double

Public Property Get Height() As Single
    Height = moControl.Height
End Property

Public Property Get Width() As Single
    Width = moControl.Width
End Property

Public Property Get Top() As Single
    Top = moControl.Top
End Property

Public Property Get Left() As Single
    Left = moControl.Left
End Property

Public Property Get ResizeBehaviour() As DoResizeEnum
    ResizeBehaviour = meResizeBehaviour
End Property

Public Property Let ResizeBehaviour(newVal As DoResizeEnum)
    meResizeBehaviour = newVal
End Property

Public Function SetvbControl(ctl As Control, vbForm As Object) As Boolean
    
    Set moControl = ctl
    On Error GoTo errhdl
    
    With moControl
       If TypeOf ctl Is Line Then
            mfInitLeft = .X1
            mfInitTop = .Y1
            mfInitRight = .X2
            mfInitBottom = .Y2
       Else
            mfInitLeft = .Left
            mfInitTop = .Top
            mfInitRight = vbForm.Width - .Width - .Left
            mfInitBottom = vbForm.Height - .Height - .Top
            mfInitWidth = .Width
            mfInitHeight = .Height
        End If
    End With
    SetvbControl = True
Exit Function
errhdl:
SetvbControl = False
End Function

Private Property Get ctlName() As String
    
On Error Resume Next
    
    If moControl Is Nothing Then
        ctlName = "Nothing"
    Else
        ctlName = moControl.Name
    End If
    
End Property

Public Function Resize(sfx As Double, sfy As Double) As Boolean

On Error GoTo ErrorMarke
 
If TypeOf moControl Is Line Then
    moControl.X1 = mfInitLeft * sfx
    moControl.Y1 = mfInitTop * sfy
    moControl.X2 = mfInitRight * sfx
    moControl.Y2 = mfInitBottom * sfy
       
Else
    moControl.Left = mfInitLeft * sfx
    moControl.Width = mfInitWidth * sfx
    moControl.Top = mfInitTop * sfy
    moControl.Height = mfInitHeight * sfy
End If
       
    Resize = False

    
Exit Function

ErrorMarke:
    'MsgBox Err.Description, vbExclamation, TypeName(Me) & ": Resize"
    Resize = True
End Function

Public Function ReStore() As Boolean

On Error GoTo ErrorMarke


If TypeOf moControl Is Line Then
    moControl.X1 = mfInitLeft
    moControl.Y1 = mfInitTop
    moControl.X2 = mfInitRight
    moControl.Y2 = mfInitBottom
       
Else
    moControl.Width = mfInitWidth
    moControl.Left = mfInitLeft
    moControl.Height = mfInitHeight
    moControl.Top = mfInitTop
End If
Exit Function

ErrorMarke:
    'MsgBox Err.Description, vbExclamation, TypeName(Me) & ": Resize"
    ReStore = True
End Function


