Attribute VB_Name = "basThreemap"
'------------------------------------------------------------------------
' Description  : This class contains all treemap specific procedures
'                   including all public start nodes
'------------------------------------------------------------------------

'Declarations
Dim mrgbLowerColor As Long        'shape color for color values < 0%
Dim mrgbUpperColor As Long        'shape color for color values > 0%
Dim mrgbDefaultColor As Long      'shape color if color value equals 0% or no color value defined
Dim mdblColorValueLimit As Double 'value where shape is set to full upper or lower color
Dim mfrmConfig As frmConfig       'reference to current config form

'Constants


'options
Option Explicit
'------------------------------------------------------------------------
' Description  : create a new sheet containing the treemap chart
' Parameters   :
' Returnvalue  : worksheet object interfacing the chart container
'------------------------------------------------------------------------
Private Function p_createChartSheet()

    Dim wshChartSheet As Worksheet          'new sheet
    Dim wshCurrentSheet As Worksheet        'an existing sheet in the active workbook
    Dim strSheetNumber As String            'number of the current treemap sheet as text
    Dim blnFoundNonNumber As Boolean        'flag is true when a treemap sheet name contains
                                            ' characters other then treemap and a number
    Dim intPosition As Integer              'counts characters
    Dim intNewNumber As Integer             'the number of the the new sheet
    Dim intDefaultLen As Integer            ' length of default name

    On Error GoTo error_handler
    intNewNumber = 1
    intDefaultLen = Len(cLangChartSheetName) + 1
    'create a new sheet in active workbook
    Set wshChartSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveSheet)
    basSystem.log "new sheet added"
    'set the name to treemap + number, where number is the highest not existing
    ' number for treemap sheets
    For Each wshCurrentSheet In ActiveWorkbook.Worksheets
        'looking for originally named treemap sheets
        If Left(wshCurrentSheet.Name, 7) = cLangChartSheetName Then
            'try to find the number from the rest of the name
            strSheetNumber = Mid(wshCurrentSheet.Name, intDefaultLen)
            'look for non number characters
            blnFoundNonNumber = False
            For intPosition = intDefaultLen To Len(wshCurrentSheet.Name)
                'wish I could use regex, instead have to check ascii codes to find
                ' non number characters
                If Asc(Mid(wshCurrentSheet.Name, intPosition, 1)) < 48 Or _
                        Asc(Mid(wshCurrentSheet.Name, intPosition, 1)) > 57 Then
                    blnFoundNonNumber = True
                End If
            Next
            'if name of the current sheet is a default name
            If Not blnFoundNonNumber And Len(strSheetNumber) > 0 Then
                'give the new sheet a higher number
                If CInt(strSheetNumber) >= intNewNumber Then
                    intNewNumber = CInt(strSheetNumber) + 1
                End If
            End If
        End If
    Next
    'set the default name for the new sheet
    wshChartSheet.Name = cLangChartSheetName & intNewNumber
    basSystem.log "new sheet named to " & cLangChartSheetName & intNewNumber
    'return the sheets object
    Set p_createChartSheet = wshChartSheet
    Exit Function
    
error_handler:
    basSystem.log_error "basThreemap.createChartSheet"
End Function
'------------------------------------------------------------------------
' Description  : START HERE - starting node for creating a new treemap chart
'------------------------------------------------------------------------
Public Sub createTreemap()
    
    Dim wshChartSheet As Worksheet             'reference to the chart sheet (this is a worksheet containing shapes)
    Dim wshTmpData As Worksheet                'temporary worksheet for chart data
    Dim wshDataSheet As Worksheet

    On Error GoTo error_handler
    'detect input data and fill config dialog with results
    Set wshDataSheet = ActiveSheet
    p_setupDataRangeInput
    'show dialog for chart configuration
    Config.Show
    If Not Config.canceled Then
        'put chart on a new sheet
        Set wshChartSheet = p_createChartSheet()
        'copy data to a temporary sheet to be able to manipulate data (e.g. sort)
        'Set wshTmpData = p_copyData(wshSampleData.Range("C4:C25"), wshSampleData.Range("B4:B25"), _
                            wshSampleData.Range("D4:D25"))
        Set wshTmpData = p_copyData(wshDataSheet.Range(Config.txtSizeRange.Text), _
                            wshDataSheet.Range(Config.txtLabelRange.Text), _
                            wshDataSheet.Range(Config.txtColorRange.Text))
        'create chart
        p_createChart wshChartSheet, wshTmpData
        'clean up - delete temporary data sheet
        p_cleanup wshTmpData
        'bring the result to front
        wshChartSheet.Activate
    End If
    Exit Sub
    
error_handler:
    basSystem.log_error "basThreemap.createTreemap"
End Sub
'------------------------------------------------------------------------
' Description  : creates a new chart at the treemap chart sheet based on rectangle shapes
' Parameters   : pwshChartSheet     - the sheet where to put the new chart on
'                pwshTmpData        - temporary worksheet for chart data
'                prngValue          - range containing the data for the size of the shapes
'------------------------------------------------------------------------
Private Sub p_createChart(pwshChartSheet As Worksheet, pwshTmpData As Worksheet)
    
    Dim shpBase As Shape                        'base shape for header and legend
    Dim lngX0 As Double, lngY0 As Double        'coordinates of top left corner for chart shape
    Dim lngWidth As Double, lngHeight As Double 'width and height for chart shape

    On Error GoTo error_handler
    'set dimensions for chart shape
    lngX0 = 5
    lngY0 = 5
    lngWidth = ActiveWindow.UsableWidth - 40
    lngHeight = ActiveWindow.UsableHeight - 25
    'add a base shape for heading and legend
'    Set shpBase = pwshChartSheet.Shapes.AddShape(msoShapeRectangle, lngX0, lngY0, lngWidth, lngHeight)
'    With shpBase
'        .Visible = msoTrue
'        .Fill.ForeColor.RGB = RGB(255, 255, 255)
'        .Fill.Transparency = 0
'        .Fill.Solid
'        .Line.Weight = 0.25
'        .Line.ForeColor.RGB = RGB(0, 0, 0)
'    End With
    'cluster data into three parts starting with the whole range of data
    p_clusterData pwshChartSheet, pwshTmpData, 1, pwshTmpData.Parent.Names(cRngIndex).RefersToRange.Rows.Count, _
                    lngX0, lngY0, lngWidth, lngHeight

    
    Exit Sub
    
error_handler:
    basSystem.log_error "basThreemap.p_createChart"
End Sub
'------------------------------------------------------------------------
' Description  : create a new sheet and copy chart data to this sheet
' Parameters   : prngData               - column containing chart data
'                prngDescriptionOrigin  - column containing description
'                prngColorValue         - column containing color data
' Returnvalue  : reference to the temporary data
'------------------------------------------------------------------------
Private Function p_copyData(prngData As Range, prngDescription As Range, Optional prngColorValue As Range)

    Dim wshTmpData As Worksheet             'temporary worksheet for chart data
    Dim wshCurrent As Worksheet             'any sheet in current workbook
    Dim dblDataRowCount As Double           'number of row in data column
    Dim blnValuesAreNegative As Boolean     'true if values are negative
    
    On Error GoTo error_handler
    'set defaults
    blnValuesAreNegative = False
    'look if name of the temporary sheet is already in use
    For Each wshCurrent In Worksheets
        If wshCurrent.Name = cTmpDataSheetName Then
            'if found backup existing sheet
            wshCurrent.Name = Format(Now(), "yyyymmddhhmmss") & "_" & cTmpDataSheetName
        End If
    Next
    'add new sheet
    Set wshTmpData = Worksheets.Add
    'name tmp data sheet
    wshTmpData.Name = cTmpDataSheetName
    'count data rows
    dblDataRowCount = prngData.Rows.Count
    'copy value data and set workbook name for later access
    prngData.Copy
    wshTmpData.Range("B2").PasteSpecial xlPasteValues
    wshTmpData.Parent.Names.Add cRngValues, "='" & cTmpDataSheetName & "'!" & _
                                wshTmpData.Range(Range("B2"), Range("B2").Offset(dblDataRowCount - 1)).Address
    'check if values are negative
    If Application.WorksheetFunction.Max(wshTmpData.Range(cRngValues)) < 0 Then
        blnValuesAreNegative = True
    End If
    'copy description data and set workbook name for later access
    prngDescription.Copy
    wshTmpData.Range("C2").PasteSpecial xlPasteValues
    wshTmpData.Parent.Names.Add cRngDescription, "='" & cTmpDataSheetName & "'!" & _
                                wshTmpData.Range(Range("C2"), Range("C2").Offset(dblDataRowCount - 1)).Address
    'if color data is given
    If Not IsMissing(prngColorValue) Then
        'copy color index data and set workbook name for later access
        prngColorValue.Copy
        wshTmpData.Range("D2").PasteSpecial xlPasteValues
        wshTmpData.Parent.Names.Add cRngColorData, "='" & cTmpDataSheetName & "'!" & _
                                    wshTmpData.Range(Range("D2"), Range("D2").Offset(dblDataRowCount - 1)).Address
        'find color value limit (max. value)
        basThreemap.p_ColorValueLimit = Application.WorksheetFunction.Max(Application.WorksheetFunction.Max(prngColorValue), _
                        Abs(Application.WorksheetFunction.Min(prngColorValue)))
    End If
    'sort data before creating index
    wshTmpData.Sort.SortFields.Clear
    'set head of data column as key sort field
    If blnValuesAreNegative Then
        'sort by absolout values
        wshTmpData.Sort.SortFields.Add Key:=Range("B2") _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortTextAsNumbers
    Else
        wshTmpData.Sort.SortFields.Add Key:=Range("B2") _
            , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
            xlSortTextAsNumbers
    End If
    With wshTmpData.Sort
        .SetRange Range(Range("B2"), Range("D2").Offset(dblDataRowCount - 1).Address)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'create index for to access data easier
    wshTmpData.Range("A2").Value = 1
    wshTmpData.Range("A3").Value = 2
    wshTmpData.Range("A2:A3").AutoFill wshTmpData.Range(Range("A2"), Range("A2").Offset(dblDataRowCount - 1)), _
                                xlFillDefault
    wshTmpData.Parent.Names.Add cRngIndex, "='" & cTmpDataSheetName & "'!" & _
                                wshTmpData.Range(Range("A2"), Range("A2").Offset(dblDataRowCount - 1)).Address
    'return the temporary data worksheet
    Set p_copyData = wshTmpData
    Exit Function
    
error_handler:
    basSystem.log_error "basThreemap.p_copyData"
End Function
'------------------------------------------------------------------------
' Description  : draw a single cluster shape in chart
' Parameters   : plngX0                 - X origin
'                plngY0                 - Y origin
'                plngWidth              - shape width
'                plngHeight             - shape height
'                pstrDescription        - record text
'                pdblColorValue         - color value
'------------------------------------------------------------------------
Private Sub p_drawClusterShape(pwshChartSheet As Worksheet, plngX0 As Double, plngY0 As Double, plngWidth As Double, _
                                plngHeight As Double, pstrDescription As String, Optional pdblColorValue As Double)
    
    Dim shpCluster As Shape                 'shape object for the single cluster shape
    Dim rgbShapeFillColor As Long           'shape fill color
    Dim rgbShapeFontColor As Long           'font color for shape
    Dim dblColorValue As Double             'normalized color value
    Dim intRedDiff As Long
    Dim intGreenDiff As Long
    Dim intBlueDiff As Long
    Dim intNewRed As Long
    Dim intNewGreen As Long
    Dim intNewBlue As Long

    On Error GoTo error_handler
    basSystem.logd "draw shape at x:" & plngX0 & " y:" & plngY0 & " w:" & plngWidth & " h:" & plngHeight & _
                    " desc: " & pstrDescription
    'set default colors white for background and black for font
    rgbShapeFillColor = basThreemap.p_DefaultColor
    rgbShapeFontColor = RGB(0, 0, 0)
    'calculate colors
    dblColorValue = pdblColorValue / basThreemap.p_ColorValueLimit
    rgbShapeFillColor = basThreemap.p_getBlockFillColor(dblColorValue)
    'create shape
    Set shpCluster = pwshChartSheet.Shapes.AddShape(msoShapeRectangle, plngX0, plngY0, plngWidth, plngHeight)
    'format shape
    With shpCluster
        .Visible = msoTrue
        .AlternativeText = pstrDescription
        .Fill.ForeColor.RGB = rgbShapeFillColor
        .Fill.Transparency = 0
        .Fill.Solid
        .Line.Weight = 0.25
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Text = pstrDescription
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    Exit Sub
    
error_handler:
    basSystem.log_error "basThreemap.p_drawClusterShape"
End Sub
'------------------------------------------------------------------------
' Description  : build group of three items each
' Parameters   : pwshChartSheet             - the sheet containing the treemap chart
'                pwshTmpData                - temporary worksheet for chart data
'                plngFromRecord             - id of the first record in a cluster range
'                plngToRecord               - id of the last record in a cluster range
'                plngClusterShapeX0         - X origin of the cluster range shape
'                plngClusterShapeY0         - Y origin of the cluster range shape
'                plngClusterRangeWidth      - width of the cluster range shape
'                plngClusterRangeHeight     - height of the cluster range shape
'------------------------------------------------------------------------
Private Sub p_clusterData(pwshChartSheet As Worksheet, pwshTmpData As Worksheet, plngFromRecord As Long, _
                            plngToRecord As Long, plngClusterShapeX0 As Double, plngClusterShapeY0 As Double, _
                            plngClusterRangeWidth As Double, plngClusterRangeHeight As Double)
    
    Dim dblClusterLimit As Double                   'max. size of a data cluster
    Dim rngCurrent As Range                         'a single cell to work with
    Dim rngClusterRange As Range                    'current data to be clustered
    Dim lngRecord As Long                           'record counter
    Dim lngClusterAreaSize As Long                  'number of records of a cluster
    Dim intCluster As Integer                       'cluster number 1 - 3
    Dim lngClusterStart As Long                     'first record in a cluster
    Dim lngClusterEnd As Long                       'last record in a cluster
    Dim lngClusterX0 As Double                      'X origin point for a cluster shape
    Dim lngClusterY0 As Double                      'Y origin point for a cluster shape
    Dim lngClusterWidth As Double                   'cluster shape width
    Dim lngClusterHeight As Double                  'cluster shape height
    Dim dblClusterSizePct As Double                 'size of a cluster in percent
    Dim dblClusterRangeValue As Double              'overall value of all three clusters
    Dim dblClusterValue As Double                   'value of one cluster
    Dim dblClusterRangeArea As Double               'shape size of all three clusters
    Dim dblClusterArea As Double                    'shape size of one cluster
    Dim dblColorValue As Double
    Dim strShapeLabel As String
    
    On Error GoTo error_handler
    basSystem.log "cluster data from record " & plngFromRecord & " to " & plngToRecord, cLogDebug
    intCluster = 1
    'calculate shape size of cluster range
    dblClusterRangeArea = plngClusterRangeWidth * plngClusterRangeHeight
    'recognize current data to be clustered by given index
    Set rngCurrent = pwshTmpData.Parent.Names(cRngIndex).RefersToRange.Find(plngFromRecord, LookAt:=xlWhole)
    Set rngClusterRange = Range(rngCurrent.Offset(, 1), rngCurrent.Offset(plngToRecord - plngFromRecord, 1))
    'get the overall value of this cluster range
    dblClusterRangeValue = Abs(Application.WorksheetFunction.Sum(rngClusterRange))
    'define a cluster limit
    dblClusterLimit = dblClusterRangeValue / 3
    basSystem.log "limit is " & dblClusterLimit, cLogDebug
    'init cluster value
    dblClusterValue = 0
    'save first record for the first cluster
    lngClusterStart = plngFromRecord
    'calculate cluster size (number of records in a cluster
    lngClusterAreaSize = plngToRecord - plngFromRecord + 1
    'get all records belonging to a cluster
    For lngRecord = 1 To lngClusterAreaSize
        'add the record value to the cluster value
         dblClusterValue = dblClusterValue + Abs(rngCurrent.Offset(, 1).Value)
        'if next record would succeed cluster limit or current record is last record then cluster is complete
         If ((intCluster < 3) And _
                        (Abs(Application.WorksheetFunction.Sum(Range(rngClusterRange.Cells(1, 1), _
                        rngClusterRange.Cells(lngRecord + 1, 1)))) _
                    > dblClusterLimit)) Or (lngRecord = lngClusterAreaSize) Then
            'save last record for this cluster
            lngClusterEnd = plngFromRecord + lngRecord - 1
            'calculate size of the cluster
            dblClusterSizePct = dblClusterValue / dblClusterRangeValue
            'determine shape size of this cluster
            dblClusterArea = dblClusterRangeArea * dblClusterSizePct
            'shape edge size depends on cluster id
            Select Case intCluster
                Case 1
                    'first cluster shape is on top and takes full width
                    'use given coordinates as origin for the first cluster
                    lngClusterX0 = plngClusterShapeX0
                    lngClusterY0 = plngClusterShapeY0
                    lngClusterWidth = plngClusterRangeWidth
                    'height depends on cluster area size
                    lngClusterHeight = dblClusterArea / plngClusterRangeWidth
                Case 2
                    'second cluster shape is left below first cluster shape
                    lngClusterY0 = lngClusterY0 + lngClusterHeight
                    'next line is just for a better understanding
                    lngClusterX0 = lngClusterX0
                    'width depends on cluster area size
                    lngClusterWidth = lngClusterWidth * _
                        (dblClusterArea / (dblClusterRangeArea - (lngClusterWidth * lngClusterHeight)))
                    lngClusterHeight = plngClusterRangeHeight - lngClusterHeight
                Case 3
                    'third cluster shape is right below first cluster shape
                    lngClusterX0 = lngClusterX0 + lngClusterWidth
                    'next line is just for a better understanding
                    lngClusterY0 = lngClusterY0
                    'height is the same as for cluster 2
                    lngClusterHeight = lngClusterHeight
                    'width of cluster three is the remaining width
                    lngClusterWidth = plngClusterRangeWidth - lngClusterWidth
            End Select
            'if cluster contains only one record, draw cluster
            If lngClusterEnd - lngClusterStart = 0 Then
                basSystem.log "draw cluster " & intCluster
                'draw cluster shape
                dblColorValue = Abs(rngCurrent.Offset(, 3).Value / basThreemap.p_ColorValueLimit)
                strShapeLabel = p_getShapeLabel(Config.txtDescriptionOutput.Text, rngCurrent.Offset(, 1).Value, _
                                    rngCurrent.Offset(, 2).Text, dblColorValue)
                p_drawClusterShape pwshChartSheet, lngClusterX0, lngClusterY0, lngClusterWidth, lngClusterHeight, _
                                strShapeLabel, rngCurrent.Offset(, 3).Value
            Else
                'cluster data by using this function recursive
                basSystem.log "recursive cluster " & intCluster & " from record " & lngClusterStart & " to " & lngClusterEnd & _
                                " for area x:" & lngClusterX0 & " y:" & lngClusterY0 & " w:" & lngClusterWidth & _
                                " h:" & lngClusterHeight
                p_clusterData pwshChartSheet, pwshTmpData, lngClusterStart, lngClusterEnd, _
                                lngClusterX0, lngClusterY0, lngClusterWidth, lngClusterHeight
                basSystem.log "return from cluster " & lngClusterStart & " to " & lngClusterEnd
            End If
            'start next cluster
            intCluster = intCluster + 1
            'reset cluster value
            dblClusterValue = 0
            'set new cluster start
            lngClusterStart = plngFromRecord + lngRecord
        End If
        'move to next record
        Set rngCurrent = rngCurrent.Offset(1)
    Next
    Exit Sub
    
error_handler:
    basSystem.log_error "basThreemap.p_clusterData"
End Sub
'------------------------------------------------------------------------
' Description  : cleanup temporary sheets and names
' Parameters   : pwshTmpData        - temporary worksheet for chart data
'------------------------------------------------------------------------
Private Sub p_cleanup(pwshTmpData As Worksheet)
    
    On Error GoTo error_handler
    basSystem.log ("cleanup names and temporary data sheet")
    'remove names from data sheet
    pwshTmpData.Parent.Names(cRngValues).Delete
    pwshTmpData.Parent.Names(cRngDescription).Delete
    pwshTmpData.Parent.Names(cRngColorData).Delete
    pwshTmpData.Parent.Names(cRngIndex).Delete
    'clean up - delete temporary data sheet
    Application.DisplayAlerts = False
    pwshTmpData.Delete
    Application.DisplayAlerts = True
    Exit Sub
    
error_handler:
    basSystem.log_error "basThreemap.p_cleanup"
End Sub
'------------------------------------------------------------------------
' Description  : read shape color for color values < 50%
' Parameters   :
' Returnvalue  : rgb color value
'------------------------------------------------------------------------
Private Property Get p_LowerColor() As Long
    
    On Error GoTo error_handler
    p_LowerColor = mrgbLowerColor
    Exit Property

error_handler:
    basSystem.log_error "basThreemap.Get p_LowerColor"
End Property
'------------------------------------------------------------------------
' Description  : set shape color for color values < 50%
' Parameters   :
'------------------------------------------------------------------------
Private Property Let p_LowerColor(ByVal prgbLowerColor As Long)
    
    On Error GoTo error_handler
    mrgbLowerColor = prgbLowerColor
    Exit Property

error_handler:
    basSystem.log_error "basThreemap.Let p_LowerColor"
End Property
'------------------------------------------------------------------------
' Description  : read shape color for color values > 50%
' Parameters   :
' Returnvalue  : rgb color value
'------------------------------------------------------------------------
Private Property Get p_UpperColor() As Long
    
    On Error GoTo error_handler
    p_UpperColor = mrgbUpperColor
    Exit Property

error_handler:
    basSystem.log_error "basThreemap.Get p_UpperColor"
End Property
'------------------------------------------------------------------------
' Description  : set shape color for color values > 50%
' Parameters   :
'------------------------------------------------------------------------
Private Property Let p_UpperColor(ByVal prgbUpperColor As Long)
    
    On Error GoTo error_handler
    mrgbUpperColor = prgbUpperColor
    Exit Property

error_handler:
    basSystem.log_error "basThreemap.Let p_UpperColor"
End Property
'------------------------------------------------------------------------
' Description  : read shape color for color values = 0%
' Parameters   :
' Returnvalue  : rgb color value
'------------------------------------------------------------------------
Private Property Get p_DefaultColor() As Long
    
    On Error GoTo error_handler
    p_DefaultColor = mrgbDefaultColor
    Exit Property

error_handler:
    basSystem.log_error "basThreemap.Get p_DefaultColor"
End Property
'------------------------------------------------------------------------
' Description  : set shape color for color values = 0%
' Parameters   :
'------------------------------------------------------------------------
Private Property Let p_DefaultColor(ByVal prgbDefaultColor As Long)
    
    On Error GoTo error_handler
    mrgbDefaultColor = prgbDefaultColor
    Exit Property

error_handler:
    basSystem.log_error "basThreemap.Let p_DefaultColor"
End Property
'------------------------------------------------------------------------
' Description  : read max. pct. value for color values
' Parameters   :
' Returnvalue  : pct. value
'------------------------------------------------------------------------
Private Property Get p_ColorValueLimit() As Double

    On Error GoTo error_handler
    p_ColorValueLimit = mdblColorValueLimit
    Exit Property

error_handler:
    basSystem.log_error "basThreemap.Get p_ColorValueLimit"
End Property
'------------------------------------------------------------------------
' Description  : set max. pct. value for color values
' Parameters   : pdblColorValueLimit  - new limit, where shape color is
'                                        set to Upper/Lower color
'------------------------------------------------------------------------
Private Property Let p_ColorValueLimit(ByVal pdblColorValueLimit As Double)

    On Error GoTo error_handler
    mdblColorValueLimit = pdblColorValueLimit
    Exit Property

error_handler:
    basSystem.log_error "basThreemap.Let p_ColorValueLimit"
End Property
'------------------------------------------------------------------------
' Description  : calculate color for a block by percentage of input value
' Parameters   :
' Returnvalue  : resulting rgb color
'------------------------------------------------------------------------
Private Function p_getBlockFillColor(pdblPctInputVal As Double) As Long

    Dim rgbBlockFillColor As Long

    On Error GoTo error_handler
    basSystem.logd "get color for value " & pdblPctInputVal
    If pdblPctInputVal > 0 Then
        rgbBlockFillColor = RGB(cBaseRed * (1 - pdblPctInputVal) + cPositiveRed * pdblPctInputVal, _
                                cBaseGreen * (1 - pdblPctInputVal) + cPositiveGreen * pdblPctInputVal, _
                                cBaseBlue * (1 - pdblPctInputVal) + cPositiveBlue * pdblPctInputVal)
        basSystem.logd "positiv color is r: " & cBaseRed * (1 - pdblPctInputVal) + cPositiveRed * pdblPctInputVal & _
                         " | g: " & cBaseGreen * (1 - pdblPctInputVal) + cPositiveGreen * pdblPctInputVal & _
                         " | b: " & cBaseBlue * (1 - pdblPctInputVal) + cPositiveBlue * pdblPctInputVal
    Else
        pdblPctInputVal = Abs(pdblPctInputVal)
        rgbBlockFillColor = RGB(cBaseRed * (1 - pdblPctInputVal) + cNegativeRed * pdblPctInputVal, _
                                cBaseGreen * (1 - pdblPctInputVal) + cNegativeGreen * pdblPctInputVal, _
                                cBaseBlue * (1 - pdblPctInputVal) + cNegativeBlue * pdblPctInputVal)
        basSystem.logd "negative color is r: " & cBaseRed * (1 - pdblPctInputVal) + cNegativeRed * pdblPctInputVal & _
                         " | g: " & cBaseGreen * (1 - pdblPctInputVal) + cNegativeGreen * pdblPctInputVal & _
                         " | b: " & cBaseBlue * (1 - pdblPctInputVal) + cNegativeBlue * pdblPctInputVal
    End If
    p_getBlockFillColor = rgbBlockFillColor
    Exit Function

error_handler:
    basSystem.log_error "basThreemap.Let p_getBlockFillColor"
End Function
'------------------------------------------------------------------------
' Description  : detect data table and fill config dialog with result
' Parameters   :
'------------------------------------------------------------------------
Private Sub p_setupDataRangeInput()

    Dim rngData As Range
    Dim strAddress As String

    On Error GoTo error_handler
    If TypeName(Selection) = "Range" Then
        Set rngData = Selection.CurrentRegion
        If rngData.Columns.Count = 1 And rngData.Rows.Count = 1 Then
            'didn't found data table
            Config.lblStatusbar.Caption = "Enter data range manually or select at least one cell from description column"
        ElseIf rngData.Columns.Count = 1 And rngData.Rows.Count > 1 Then
            'TODO:expect that only size is available, color and description are left
            
        ElseIf rngData.Columns.Count = 2 Then
            'TODO:expect that description column is missing
            
        Else
            'expect at least three columns: description, size and color values
            strAddress = rngData.Columns(1).Address
            strAddress = Replace(strAddress, "$", "")
            Config.txtLabelRange.Text = strAddress
            strAddress = rngData.Columns(2).Address
            strAddress = Replace(strAddress, "$", "")
            Config.txtSizeRange.Text = strAddress
            strAddress = rngData.Columns(3).Address
            strAddress = Replace(strAddress, "$", "")
            Config.txtColorRange.Text = strAddress
        End If
    Else
       Config.lblStatusbar.Caption = "Enter data range manually or select at least one cell from description column"
    End If
    Exit Sub

error_handler:
    basSystem.log_error "basThreemap.p_setupDataRangeInput"
End Sub
'------------------------------------------------------------------------
' Description  : build label for each shape
' Parameters   : pstrLabelPattern   - the input pattern from the config form
'                pdblSizeValue      - value from the size range
'                pstrLabel          - text from the label data range
'                pdblColorValue     - normalized color value in percent
' Returnvalue  : text to put on a single shape
'------------------------------------------------------------------------
Private Function p_getShapeLabel(pstrLabelPattern As String, pdblSizeValue As Double, Optional pstrLabel, Optional pdblColorValue)

    Dim strShapeLabel As String

    On Error GoTo error_handler
    'insert size
    strShapeLabel = Replace(pstrLabelPattern, "#size", pdblSizeValue)
    'insert label
    If Not IsMissing(pstrLabel) Then
        strShapeLabel = Replace(strShapeLabel, "#label", pstrLabel)
    End If
    'insert color percentage
    If Not IsMissing(pdblColorValue) Then
        strShapeLabel = Replace(strShapeLabel, "#colorpct", Format(pdblColorValue, "0.0%"))
    End If
    p_getShapeLabel = strShapeLabel
    Exit Function

error_handler:
    basSystem.log_error "basThreemap.p_getShapeLabel"
End Function
'------------------------------------------------------------------------
' Description  : access to current config form
' Parameters   :
'------------------------------------------------------------------------
Public Property Get Config() As frmConfig

    On Error GoTo error_handler
    If TypeName(mfrmConfig) = "Nothing" Then
        Set mfrmConfig = New frmConfig
    End If
    Set Config = mfrmConfig
    Exit Property

error_handler:
    basSystem.log_error "basThreemap.Config Get"
End Property

