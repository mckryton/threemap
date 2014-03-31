Attribute VB_Name = "basConstants"
'------------------------------------------------------------------------
' Description  : contains all global constants
'------------------------------------------------------------------------

'Options
Option Explicit

'log level (range is 1 to 100)
Global Const cLogDebug = 100
Global Const cLogInfo = 90
Global Const cLogWarning = 50
Global Const cLogError = 30
Global Const cLogCritical = 1

'current log level - decreasing log level means decreasing amount of messages
Global Const cCurrentLogLevel = 100

'name of the temporary data sheet
Global Const cTmpDataSheetName = "x@x-treemap-data"

'names used in workbook to define data ranges
Global Const cRngValues = "threemapValueData"
Global Const cRngDescription = "threemapDescriptionData"
Global Const cRngColorData = "threemapColorData"
Global Const cRngIndex = "threemapIndex"

'language specific constants
Global Const cLangChartSheetName = "treemap"

'chart colors
'color for positive values -> green
Global Const cPositiveRed = 0
Global Const cPositiveGreen = 255
Global Const cPositiveBlue = 0
'color for negative values -> red
Global Const cNegativeRed = 255
Global Const cNegativeGreen = 0
Global Const cNegativeBlue = 0
'color for zero values -> white
Global Const cBaseRed = 255
Global Const cBaseGreen = 255
Global Const cBaseBlue = 255
