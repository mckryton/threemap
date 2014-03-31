Attribute VB_Name = "basExcel"
'------------------------------------------------------------------------
' Description  : extension to excel functionality
'------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------------------
' Description  : get row part from a range address
'                   (e.g. get 34 from "B34"
' Parameters   : pstrAddress    - a single range address
' Returnvalue  : row number
'------------------------------------------------------------------------
Public Function getRowFromAddress(pstrAddress As String) As Long

    Dim lngColumn As Long
    Dim intCharCount As Integer
    
    On Error GoTo error_handler
    intCharCount = 1
    lngColumn = -1
    While Asc(Mid(pstrAddress, intCharCount, 1)) < 48 Or Asc(Mid(pstrAddress, intCharCount, 1)) > 57
        intCharCount = intCharCount + 1
    Wend
    lngColumn = CLng(Right(pstrAddress, Len(pstrAddress) - intCharCount + 1))
    getRowFromAddress = lngColumn
    Exit Function
    
error_handler:
    basSystem.log_error "basExcel.getRowFromAddress"
End Function

'------------------------------------------------------------------------
' Description  : get column part from a range address
'                   (e.g. get AB from "AB34"
' Parameters   : pstrAddress    - a single range address
' Returnvalue  : column char(s)
'------------------------------------------------------------------------
Public Function getColumnFromAddress(pstrAddress As String) As String

    Dim strRow As String
    Dim intCharCount As Integer
    
    On Error GoTo error_handler
    intCharCount = 1
    strRow = ""
    While Asc(Mid(pstrAddress, intCharCount, 1)) < 48 Or Asc(Mid(pstrAddress, intCharCount, 1)) > 57
        intCharCount = intCharCount + 1
    Wend
    strRow = Left(pstrAddress, intCharCount - 1)
    getColumnFromAddress = strRow
    Exit Function
    
error_handler:
    basSystem.log_error "basExcel.getRowFromAddress"
End Function
