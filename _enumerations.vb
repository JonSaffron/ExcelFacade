' ReSharper disable InconsistentNaming

Public Enum XlUnderlineStyle
    xlUnderlineStyleSingle = 2
    xlUnderlineStyleSingleAccounting = 4
    xlUnderlineStyleDoubleAccounting = 5
    xlUnderlineStyleDouble = -4119
    xlUnderlineStyleNone = -4142
End Enum

Public Enum XlInsertShiftDirection
    xlShiftDown = -4121
    xlShiftToRight = -4161
End Enum

Public Enum XlDeleteShiftDirection
    xlShiftToLeft = -4159
    xlShiftUp = -4162
End Enum

Public Enum XlPlacement
    xlMoveAndSize = 1
    xlMove = 2
    xlFreeFloating = 3
End Enum

Public Enum XlWindowState
    xlNormal = -4143
    xlMinimized = -4140
    xlMaximized = -4137
End Enum

Public Enum XlUpdateLinks
    xlUpdateLinksUserSetting = 1
    xlUpdateLinksNever = 2
    xlUpdateLinksAlways = 3
End Enum

Public Enum XlVAlign
    xlVAlignBottom = -4107
    xlVAlignCenter = -4108
    xlVAlignDistributed = -4117
    xlVAlignJustify = -4130
    xlVAlignTop = -4160
End Enum

Public Enum XlHAlign
    xlHAlignCenter = -4108
    xlHAlignCenterAcrossSelection = 7
    xlHAlignDistributed = -4117
    xlHAlignFill = 5
    xlHAlignGeneral = 1
    xlHAlignJustify = -4130
    xlHAlignLeft = -4131
    xlHAlignRight = -4152
End Enum

Public Enum XlColorIndex
    xlColorIndexAutomatic = -4105
    xlColorIndexNone = -4142
End Enum

Public Enum XlPattern
    xlPatternNone = -4142
    xlPatternAutomatic = -4105

    xlPatternSolid = 1
    xlPatternGray75 = -4126
    xlPatternGray50 = -4125 
    xlPatternGray25 = -4124
    xlPatternGray16 = 17
    xlPatternGray8 = 18

    xlPatternHorizontal = -4128
    xlPatternVertical = -4166
    xlPatternDown = -4121
    xlPatternUp = -4162
    xlPatternChecker = 9
    xlPatternSemiGray75 = 10

    xlPatternLightHorizontal = 11
    xlPatternLightVertical = 12
    xlPatternLightDown = 13
    xlPatternLightUp = 14
    xlPatternGrid = 15
    xlPatternCrissCross = 16

    xlPatternLinearGradient = 4000
    xlPatternRectangularGradient = 4001
End Enum

Public Enum XlWBATemplate
    xlWBATWorksheet = -4167
    xlWBATChart = -4109
End Enum

Public Enum XlBordersIndex
    xlDiagonalDown = 5
    xlDiagonalUp = 6
    xlEdgeBottom = 9
    xlEdgeLeft = 7
    xlEdgeRight = 10
    xlEdgeTop = 8
    xlInsideHorizontal = 12
    xlInsideVertical = 11
End Enum

Public Enum XlLineStyle
    xlContinuous = 1
    xlDash = -4115
    xlDashDot = 4
    xlDashDotDot = 5
    xlDot = -4118
    xlDouble = -4119
    xlLineStyleNone = -4142
    xlSlantDashDot = 13
End Enum

Public Enum XlBorderWeight
    xlHairline = 1
    xlMedium = -4138
    xlThick = 4
    xlThin = 2
End Enum

Public Enum XlPageOrientation
    xlLandscape = 2
    xlPortrait = 1
End Enum

Public Enum XlSheetType
    xlWorksheet = -4167
    xlChart = -4109
    xlDialogSheet = -4116
End Enum

Public Enum XlWindowView
    xlNormalView = 1
    xlPageBreakPreview = 2
End Enum

Public Enum XlPageBreakExtent
    xlPageBreakFull = 1
    xlPageBreakPartial = 2
End Enum

Public Enum XlPageBreak
    xlPageBreakAutomatic = -4105
    xlPageBreakManual = -4135
    xlPageBreakNone = -4142
End Enum

Public Enum XlSheetVisibility
    xlSheetVisible = -1
    xlSheetHidden = 0
    xlSheetVeryHidden = 2
End Enum

Public Enum XlPivotTableSourceType
    xlDatabase = 1
End Enum

Public Enum XlPivotFieldOrientation
    xlHidden = 0
    xlRowField = 1
    xlColumnField = 2
    xlPageField = 3
    xlDataField = 4
End Enum

Public Enum XlErrorChecks
    xlEvaluateToError = 1
    xlTextDate = 2
    xlNumberAsText = 3
    xlInconsistentFormula = 4
    xlOmittedCells = 5
    xlUnlockedFormulaCells = 6
    xlEmptyCellReferences = 7
    xlListDataValidation = 8
End Enum

Public Enum XlSubtotals
    xlAutomatic = 1
    xlSum = 2
    xlCount = 3
    xlAverage = 4
    xlMax = 5
    xlMin = 6
    xlProduct = 7
    xlCountNums = 8
    xlStDev = 9
    xlStDevP = 10
    xlVar = 11
    xlVarP = 12
End Enum

Public Enum MsoDocProperties
    msoPropertyTypeNumber = 1
    msoPropertyTypeBoolean = 2
    msoPropertyTypeDate = 3
    msoPropertyTypeString = 4
    msoPropertyTypeFloat = 5
End Enum

Public Enum XlFileFormat
    xlWorkbookNormal = -4143                ' Standard xls format - Applies to Office versions up to and including 2003

    xlExcel12 = 50                          ' Excel Binary Workbook in 2007 with or without macros: .xlsb
    xlOpenXMLWorkbook = 51                  ' Open xml workbook in 2007 without macros: .xlsx
    xlOpenXMLWorkbookMacroEnabled = 52      ' Open xml workbook in 2007 with or without macros: .xlsm
    xlExcel8 = 56                           ' Excel 97-2003 format in Excel 2007: .xls
End Enum

Public Enum XlPasteType
    xlPasteAll = -4104
    xlPasteAllExceptBorders = 7
    xlPasteColumnWidths = 8
    xlPasteComments = -4144
    xlPasteFormats = -4122
    xlPasteFormulas = -4123
    xlPasteFormulasAndNumberFormats = 11
    xlPasteValidation = 6
    xlPasteValues = -4163
    xlPasteValuesAndNumberFormats = 12
End Enum

Public Enum XlCutCopyMode
    xlNone = 0
    xlCopy = 1
    xlCut = 2
End Enum

Public Enum XlPaperSize
    xlPaperA3 = 8
    xlPaperA4 = 9
End Enum

Public Enum XlChartType
    xlXYScatterLines = 74
End Enum

Public Enum XlReferenceStyle
    xlA1 = 1
    xlR1C1 = -4150
End Enum

Public Enum XlChartLocation
    xlLocationAsNewSheet = 1
    xlLocationAsObject = 2
    xlLocationAutomatic = 3
End Enum

Public Enum XlTickMark
    xlTickMarkNone = -4142
    xlTickMarkInside = 2
    xlTickMarkOutside = 3
    xlTickMarkCross = 4
End Enum

Public Enum XlScaleType
    xlScaleLinear = -4132
    xlScaleLogarithmic = -4133
End Enum

Public Enum XlAxisType
    xlCategory = 1
    xlValue = 2
    xlSeriesAxis = 3
End Enum

Public Enum XlOrientation
    xlVertical = -4166
    xlUpward = -4171
    xlHorizontal = -4128
    xlDownward = -4170
End Enum
