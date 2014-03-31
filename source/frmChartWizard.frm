VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChartWizard 
   Caption         =   "Select data source"
   ClientHeight    =   4000
   ClientLeft      =   0
   ClientTop       =   -3080
   ClientWidth     =   7600
   OleObjectBlob   =   "frmChartWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChartWizard"
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

Private Sub TextBox3_Change()

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
