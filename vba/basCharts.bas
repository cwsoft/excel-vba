Attribute VB_Name = "basCharts"
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Module: basCharts
' VBA routines to create and manipulate chart objects embedded on Excel worksheets.
' Main focus of this module is on the manipulation of Excel XY-Scatter chart objects.
'
' LICENSE: GNU General Public License 3.0
'
' @platform    Excel 2010 (Windows 7)
' @package     excel-vba (https://github.com/cwsoft/excel-vba)
' @author      cwsoft (http://cwsoft.de)
' @copyright   cwsoft
' @license     http://www.gnu.org/licenses/gpl-3.0.html
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' constants to enable debugging for this module
Private Const DEBUG_MODE = True
Private Const DISPLAY_ERRORS = True

' public types to store chart settings
Public Type CHART_AXIS_SCALES
   MinimumScaleIsAuto As Boolean
   MaximumScaleIsAuto As Boolean
   MinorUnitIsAuto As Boolean
   MajorUnitIsAuto As Boolean
   MinimumScale As Variant
   MaximumScale As Variant
   MinorUnit As Variant
   MajorUnit As Variant
End Type

Public Type CHART_SERIES_STYLES
   LC As Variant  ' line color (color index 1-60)
   LS As Variant  ' line style (allowed: members of xlLineStyle)
   LW As Variant  ' line weight (allowed: members of xlBorderWeight)
   MC As Variant  ' marker background/fill color (color index: 1-60)
   MFC As Variant ' marker foreground/border color (color index 1-60)
   MS As Variant  ' marker style (allowed: members of xlMarkerStyle)
   MW As Variant  ' marker size/weight (allowed: 2-72)
End Type
'''

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' VBA ROUTINES - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Sub errorHandler(err As ErrObject, Source As String)
   ' error handler for this module
   If err.Number = 0 Or Not DISPLAY_ERRORS Then Exit Sub
   
   MsgBox "An error occured in '" & Source & "'." & vbCr _
      & "Err " & CStr(err.Number) & ": " & err.Description & "." & vbCr & vbCr _
      & "Please contact the application author to fix the error." _
      , vbExclamation + vbOKOnly, "VBA Application Error"
End Sub

Sub addChart(wks As Worksheet _
   , Optional SrcRng As Range = Nothing _
   , Optional PlotBy = xlColumns _
   , Optional ChartName As String = "" _
   , Optional ChartType = xlXYScatterLines _
   , Optional TopLeftCell = "B2" _
   , Optional BottomRightCell = "J20" _
   )
   ' adds a new chart object to given worksheet
   Dim oChart As ChartObject
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   ' add new chart object
   Set oChart = wks.ChartObjects.Add( _
      Left:=wks.Range(TopLeftCell).Left, _
      Width:=(wks.Range(BottomRightCell).Offset(0, 1).Left - wks.Range(TopLeftCell).Left), _
      Top:=wks.Range(TopLeftCell).Top, _
      Height:=(wks.Range(BottomRightCell).Offset(1, 0).Top - wks.Range(TopLeftCell).Top))

   ' set chart source data, chart name and type
   If Not SrcRng Is Nothing Then oChart.Chart.SetSourceData Source:=SrcRng, PlotBy:=PlotBy
   If ChartName <> "" Then oChart.Name = ChartName
   oChart.Chart.ChartType = ChartType
   
errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.addChart")
End Sub

Sub addChartSeries(oChart As ChartObject, XRng As Range, YRng As Range, Label)
   ' adds a new data series to the given chart object
   Dim rng As Range
   Dim iArea As Integer
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   ' add new data series to given chart object
   iArea = 0
   For Each rng In XRng.Areas
      iArea = iArea + 1
      With oChart.Chart.SeriesCollection.NewSeries
         .XValues = XRng.Areas(iArea)
         .Values = YRng.Areas(iArea)
         
         ' set series label from string, range or array
         If TypeName(Label) = "String" Then
            .Name = "=""" & Label & """"
         ElseIf TypeName(Label) = "Range" Then
            ' build fully qualified label referece from label range (Excel 2010)
            .Name = "='" & Label.Areas(iArea).Parent.Name & "'!" & Label.Areas(iArea).Address
         ElseIf TypeName(Label) = "Variant()" Then
            ' read label from array
            .Name = "=""" & Label(iArea - 1) & """"
         End If
      End With
   Next

errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.addChartSeries")
End Sub

Sub addCustomChartType(oChart As ChartObject, CustomChartType)
   ' adds given chart as new custom chart type to Excel
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   ' add a new auto chart format
   With Application
      .DisplayAlerts = False
      .AddChartAutoFormat Chart:=oChart.Chart, Name:=CustomChartType
   End With

errHandler:
   Application.DisplayAlerts = True
   Call basCharts.errorHandler(err, Source:="basCharts.addCustomChartType")
End Sub

Sub deleteCharts(wks As Worksheet, Optional MatchName)
   ' deletes all chart objects on given worksheet (optional: only charts matching MatchName)
   Dim oChart As ChartObject
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   ' remove all charts embedded on given worksheet
   If wks.ChartObjects.Count = 0 Then Exit Sub
   If IsMissing(MatchName) Then
      wks.ChartObjects.Delete
   Else
      For Each oChart In wks.ChartObjects
         If oChart.Name Like MatchName Then oChart.Delete
      Next
   End If
  
errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.deleteCharts")
End Sub

Sub deleteChartSeries(oChart As ChartObject, Optional Series)
   ' deletes all chart series from given chart object (optional: only Series defined)
   Dim oSeries As Series
   Dim iNbrSeries As Integer
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   If oChart.Chart.SeriesCollection.Count = 0 Then Exit Sub
   ' delete all series of given chart if no specific series is defined
   If IsMissing(Series) Then
      For Each oSeries In oChart.Chart.SeriesCollection
         oSeries.Delete
      Next
      GoTo errHandler
   End If
   
   ' allow negative series numbers for reverse counting (-1:= last, -2: second last ...)
   If IsNumeric(Series) Then
      iNbrSeries = oChart.Chart.SeriesCollection.Count
      If Series < 0 Then Series = WorksheetFunction.Max(1, 1 + (iNbrSeries + Series))
      If Series > iNbrSeries Then Series = iNbrSeries
   End If
   
   ' delete defined Series of given chart object
   oChart.Chart.SeriesCollection(Series).Delete

errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.deleteChartSeries")
End Sub

Sub modifyChartSeries(oChart As ChartObject, Search, Replace, Optional Series)
   ' replaces search string in all chart series formulas of given chart (optional: only Series defined)
  Dim oSeries As Series
  Dim iNbrSeries As Integer
  Dim sSeriesFormula As String

   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   If oChart.Chart.SeriesCollection.Count = 0 Then Exit Sub
   Application.ScreenUpdating = False
   ' modify all chart series of given chart object
   If IsMissing(Series) Then
      For Each oSeries In oChart.Chart.SeriesCollection
         ' manipulate series formula of actual chart
         sSeriesFormula = WorksheetFunction.Substitute(oSeries.Formula, Search, Replace)
         oSeries.Formula = sSeriesFormula
      Next
      GoTo errHandler
   End If

   ' allow negative series numbers for reverse counting (-1:= last, -2: second last ...)
   If IsNumeric(Series) Then
      iNbrSeries = oChart.Chart.SeriesCollection.Count
      If Series < 0 Then Series = WorksheetFunction.Max(1, 1 + (iNbrSeries + Series))
      If Series > iNbrSeries Then Series = iNbrSeries
   End If
   
   ' modify defined Series of given chart object
   Set oSeries = oChart.Chart.SeriesCollection(Series)
   sSeriesFormula = WorksheetFunction.Substitute(oSeries.Formula, Search, Replace)
   oSeries.Formula = sSeriesFormula

errHandler:
   Application.ScreenUpdating = True
   Call basCharts.errorHandler(err, Source:="basCharts.modifyChartSeries")
End Sub

Sub setChartAxisScales(oChart As ChartObject, Axis, Scales As CHART_AXIS_SCALES)
   ' sets the scales of the requested axis type of the given chart
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   ' set chart axis scales
   With oChart.Chart.Axes(Axis)
      On Error Resume Next
      If Not IsEmpty(Scales.MinimumScale) Then .MinimumScale = Scales.MinimumScale
      If Not IsEmpty(Scales.MaximumScale) Then .MaximumScale = Scales.MaximumScale
      If Not IsEmpty(Scales.MinorUnit) Then .MinorUnit = Scales.MinorUnit
      If Not IsEmpty(Scales.MajorUnit) Then .MajorUnit = Scales.MajorUnit

      If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
      If Not IsEmpty(Scales.MinimumScaleIsAuto) Then .MinimumScaleIsAuto = Scales.MinimumScaleIsAuto
      If Not IsEmpty(Scales.MaximumScaleIsAuto) Then .MaximumScaleIsAuto = Scales.MaximumScaleIsAuto
      If Not IsEmpty(Scales.MinorUnitIsAuto) Then .MinorUnitIsAuto = Scales.MinorUnitIsAuto
      If Not IsEmpty(Scales.MajorUnitIsAuto) Then .MajorUnitIsAuto = Scales.MajorUnitIsAuto
   End With

errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.setChartAxisScales")
End Sub

Sub setChartCaptions(oChart As ChartObject, Optional Title, Optional XLabel, Optional YLabel)
   ' sets chart title and chart axis labels of given chart object
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   With oChart.Chart
      ' set chart title
      If Not IsMissing(Title) Then
         .HasTitle = Trim(Title) <> ""
         If .HasTitle Then .ChartTitle.Text = Title
      End If

      ' set chart x-axis label
      If Not IsMissing(XLabel) Then
         .Axes(xlCategory, xlPrimary).HasTitle = Trim(XLabel) <> ""
         If .Axes(xlCategory, xlPrimary).HasTitle Then .Axes(xlCategory, xlPrimary).AxisTitle.Text = XLabel
      End If
    
      ' set chart y-axis label
      If Not IsMissing(YLabel) Then
         .Axes(xlValue, xlPrimary).HasTitle = Trim(YLabel) <> ""
         If .Axes(xlValue, xlPrimary).HasTitle Then .Axes(xlValue, xlPrimary).AxisTitle.Text = YLabel
      End If
   End With

errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.setChartCaptions")
End Sub

Sub setChartLegendVisibility(oChart As ChartObject, Optional Visible)
   ' sets chart legend visibility of given chart object (optional: toggle status if Visible is not defined)
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   If IsMissing(Visible) Then
      ' toggle legend visibilty
      oChart.Chart.HasLegend = Not oChart.Chart.HasLegend
   Else
      ' set defined legend visibilty
      oChart.Chart.HasLegend = Visible
   End If

errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.setChartLegendVisibility")
End Sub

Sub setChartSeriesStyles(oChart As ChartObject, Series, Styles As CHART_SERIES_STYLES)
   ' sets series styles for a given chart object series
   Dim oSeries As Series
   Dim iNbrSeries As Integer
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   If oChart.Chart.SeriesCollection.Count = 0 Then Exit Sub
   ' allow negative series numbers for reverse counting (-1:= last, -2: second last ...)
   If IsNumeric(Series) Then
      iNbrSeries = oChart.Chart.SeriesCollection.Count
      If Series < 0 Then Series = WorksheetFunction.Max(1, 1 + (iNbrSeries + Series))
      If Series > iNbrSeries Then Series = iNbrSeries
   End If
      
   ' set chart object series styles
   Set oSeries = oChart.Chart.SeriesCollection(Series)
   With oSeries
      If Not IsEmpty(Styles.LC) Then .Border.ColorIndex = Styles.LC
      If Not IsEmpty(Styles.LS) Then .Border.lineStyle = Styles.LS
      If Not IsEmpty(Styles.LW) Then .Border.Weight = Styles.LW
      
      If Not IsEmpty(Styles.MC) Then .MarkerBackgroundColorIndex = Styles.MC
      If Not IsEmpty(Styles.MFC) Then .MarkerForegroundColorIndex = Styles.MFC
      If Not IsEmpty(Styles.MS) Then .markerStyle = Styles.MS
      If Not IsEmpty(Styles.MW) Then .markerSize = Styles.MW
   End With
   
errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.setChartSeriesStyles")
End Sub

Sub setCustomChartType(oChart As ChartObject, CustomChartType)
   ' sets given chart object to the defined custom chart type
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   oChart.Chart.ApplyCustomType ChartType:=xlUserDefined, TypeName:=CustomChartType

errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.setCustomChartType")
End Sub

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' VBA FUNCTIONS - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Function getChartSeriesStyles(oChart As ChartObject, Series) As CHART_SERIES_STYLES
   ' returns series styles for a given chart object series
   Dim oSeries As Series
   Dim iNbrSeries As Integer
   Dim Styles As CHART_SERIES_STYLES
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   If oChart.Chart.SeriesCollection.Count = 0 Then
      getChartSeriesStyles = Styles
      Exit Function
   End If
   ' allow negative series numbers for reverse counting (-1:= last, -2: second last ...)
   If IsNumeric(Series) Then
      iNbrSeries = oChart.Chart.SeriesCollection.Count
      If Series < 0 Then Series = WorksheetFunction.Max(1, 1 + (iNbrSeries + Series))
      If Series > iNbrSeries Then Series = iNbrSeries
   End If
      
   ' set chart object series styles
   Set oSeries = oChart.Chart.SeriesCollection(Series)
   With oSeries
      Styles.LC = .Border.ColorIndex
      Styles.LS = .Border.lineStyle
      Styles.LW = .Border.Weight
      
      Styles.MC = .MarkerBackgroundColorIndex
      Styles.MFC = .MarkerForegroundColorIndex
      Styles.MS = .markerStyle
      Styles.MW = .markerSize
   End With
   
errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.getChartSeriesStyles")
End Function

Function getChartAxisScales(oChart As ChartObject, Axis) As CHART_AXIS_SCALES
   ' returns the scales of the requested axis type of the given chart object
   Dim Scales As CHART_AXIS_SCALES
   
   If DEBUG_MODE Then On Error GoTo 0 Else On Error GoTo errHandler
   ' extract chart axis scales
   With oChart.Chart.Axes(Axis)
      Scales.MinimumScaleIsAuto = .MinimumScaleIsAuto
      Scales.MaximumScaleIsAuto = .MaximumScaleIsAuto
      Scales.MinorUnitIsAuto = .MinorUnitIsAuto
      Scales.MajorUnitIsAuto = .MajorUnitIsAuto
      
      Scales.MinimumScale = .MinimumScale
      Scales.MaximumScale = .MaximumScale
      Scales.MinorUnit = .MinorUnit
      Scales.MajorUnit = .MajorUnit
   End With

   getChartAxisScales = Scales

errHandler:
   Call basCharts.errorHandler(err, Source:="basCharts.getChartAxisScales")
End Function
