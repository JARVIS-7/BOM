VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoResize"
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
'*** Allgemeine Klasse zum Resizen von Formularen
'*****************************************************************************

Private moFrm As Object
Private moControls As Collection

Public Enum DoResizeEnum
    rszStickLeft = 1
    rszStickRight = 2
    rszStickTop = 4
    rszStickBottom = 8
End Enum

Public Property Get vbForm() As Object
    Set vbForm = moFrm
End Property

Public Property Let vbForm(ByVal newForm As Object)
    Set moFrm = newForm
End Property

Public Sub AddControl(oCtl As Control, eBehaviour As DoResizeEnum, oFrm As Object)
On Error GoTo ErrorMarke
Dim oCtlTmp As clsDoResizeCtl
Dim bOk As Boolean

Set oCtlTmp = New clsDoResizeCtl

With oCtlTmp
    bOk = .SetvbControl(oCtl, oFrm)
    .ResizeBehaviour = eBehaviour
End With
    
If bOk Then Call moControls.Add(oCtlTmp)

Done:
On Error GoTo 0
Set oCtlTmp = Nothing

Exit Sub

ErrorMarke:

Resume Done
End Sub

Public Function Resize(FormWidthAkt As Double, FormHeightAkt As Double, ByVal FormWidthOld As Double, ByVal FormHeightOld As Double) As Boolean

Dim ctl As clsDoResizeCtl
Dim fehler As Boolean
Dim sfx As Double
Dim sfy As Double

On Error GoTo ErrorMarke

sfx = FormWidthAkt / FormWidthOld
sfy = FormHeightAkt / FormHeightOld

    For Each ctl In moControls
        fehler = fehler Or ctl.Resize(sfx, sfy)
    Next ctl
    
    Resize = fehler

Exit Function

ErrorMarke:
    MsgBox Err.Description, vbExclamation, TypeName(Me) & ": Resize"
    Resize = True
End Function

'Originalstand wegretten

Public Function ReStore() As Boolean

Dim ctl As clsDoResizeCtl
Dim fehler As Boolean

On Error GoTo ErrorMarke

    For Each ctl In moControls
        fehler = fehler Or ctl.ReStore()
    Next ctl
    
    ReStore = fehler

Exit Function

ErrorMarke:
    MsgBox Err.Description, vbExclamation, TypeName(Me) & ": ReStore"
    ReStore = True
End Function

Private Sub Class_Initialize()
Set moControls = New Collection
End Sub


Private Sub Class_Terminate()
Set moControls = Nothing
End Sub


