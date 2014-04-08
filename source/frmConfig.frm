VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfig 
   Caption         =   "Select data source"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   -13200
   ClientWidth     =   7600
   OleObjectBlob   =   "frmConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------------------
' Description  : user interface for configuration
'------------------------------------------------------------------------

'Options
Option Explicit

Dim mblnCanceled As Boolean

Private Sub chkHasHeadline_Click()
    
    If Me.chkHasHeadline.Value = True Then
        p_ignoreHeadline
    Else
        p_includeHeadline
    End If
End Sub

'------------------------------------------------------------------------
' Description  : cancel dialog
'------------------------------------------------------------------------
Private Sub cmdCancel_Click()
    
    Me.canceled = True
    Me.Hide
End Sub
'------------------------------------------------------------------------
' Description  : start drawing
'------------------------------------------------------------------------
Private Sub cmdDrawChart_Click()

    Me.Hide
End Sub



'------------------------------------------------------------------------
' Description  : init form
'------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Me.canceled = False
End Sub
'------------------------------------------------------------------------
' Description  : say if dialog was canceled
'------------------------------------------------------------------------
Public Property Get canceled() As Variant
    
    canceled = mblnCanceled
End Property
'------------------------------------------------------------------------
' Description  : save if dialog was canceled
'------------------------------------------------------------------------
Public Property Let canceled(ByVal pblnCanceled As Variant)

    mblnCanceled = pblnCanceled
End Property
'------------------------------------------------------------------------
' Description  : reduce data range to ignore headline
'------------------------------------------------------------------------
Private Sub p_ignoreHeadline()

    Dim strStartAddress As String
    Dim lngStartRow As Long
    Dim strColAddress As String
    
    On Error GoTo error_handler
    Me.txtLabelRange.Text = p_reduceRangeByOneRow(Me.txtLabelRange.Text)
    Me.txtColorRange.Text = p_reduceRangeByOneRow(Me.txtColorRange.Text)
    Me.txtSizeRange.Text = p_reduceRangeByOneRow(Me.txtSizeRange.Text)
    Exit Sub
    
error_handler:
    basSystem.log_error "frmConfig.p_ignoreHeadline"
End Sub
'------------------------------------------------------------------------
' Description  : extend data range to include headline
'------------------------------------------------------------------------
Private Sub p_includeHeadline()

    On Error GoTo error_handler
    Me.txtLabelRange.Text = p_extendRangeByOneRow(Me.txtLabelRange.Text)
    Me.txtColorRange.Text = p_extendRangeByOneRow(Me.txtColorRange.Text)
    Me.txtSizeRange.Text = p_extendRangeByOneRow(Me.txtSizeRange.Text)
    Exit Sub
    
error_handler:
    basSystem.log_error "frmConfig.p_includeHeadline"
End Sub
'------------------------------------------------------------------------
' Description  : reduce size of a given range area by one row
' Parameters   : pstrRangeAddress   - a range area address (e.g. AB22:AB45)
' Returnvalue  : range area address as string (e.g. AB23:AB45)
'------------------------------------------------------------------------
Private Function p_reduceRangeByOneRow(pstrRangeAddress As String) As String
    
    Dim strStartAddress As String
    Dim lngStartRow As Long
    Dim strColAddress As String
    
    On Error GoTo error_handler
    If InStr(pstrRangeAddress, ":") > 0 Then
        strStartAddress = Left(pstrRangeAddress, InStr(pstrRangeAddress, ":") - 1)
        lngStartRow = basExcel.getRowFromAddress(strStartAddress) + 1
        strColAddress = basExcel.getColumnFromAddress(strStartAddress)
        p_reduceRangeByOneRow = strColAddress & lngStartRow & Right(pstrRangeAddress, InStr(pstrRangeAddress, ":") + 1)
    Else
        'return given address unchanged if it looks not liken an range address
        p_reduceRangeByOneRow = pstrRangeAddress
    End If
    Exit Function
    
error_handler:
    basSystem.log_error "frmConfig.p_reduceRangeByOneRow"
End Function

'------------------------------------------------------------------------
' Description  : extend size of a given range area by one row
' Parameters   : pstrRangeAddress   - a range area address (e.g. AB22:AB45)
' Returnvalue  : range area address as string (e.g. AB21:AB45)
'------------------------------------------------------------------------
Private Function p_extendRangeByOneRow(pstrRangeAddress As String) As String
    
    Dim strStartAddress As String
    Dim lngStartRow As Long
    Dim strColAddress As String
    
    On Error GoTo error_handler
    If InStr(pstrRangeAddress, ":") > 0 Then
        strStartAddress = Left(pstrRangeAddress, InStr(pstrRangeAddress, ":") - 1)
        lngStartRow = basExcel.getRowFromAddress(strStartAddress) - 1
        strColAddress = basExcel.getColumnFromAddress(strStartAddress)
        p_extendRangeByOneRow = strColAddress & lngStartRow & Right(pstrRangeAddress, InStr(pstrRangeAddress, ":") + 1)
    Else
        'return given address unchanged if it looks not liken an range address
        p_extendRangeByOneRow = pstrRangeAddress
    End If
    Exit Function
    
error_handler:
    basSystem.log_error "frmConfig.p_extendRangeByOneRow"
End Function

