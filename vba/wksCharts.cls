VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wksCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Sheet: Chart-Demo
' Demo of the methods provided by the basCharts module.
'
' LICENSE: GNU General Public License 3.0
'
' @platform    Excel 2010 (Windows 7)
' @package     excel-vba (https://github.com/cwsoft/excel-vba)
' @author      cwsoft (http://cwsoft.de)
' @copyright   cwsoft
' @license     http://www.gnu.org/licenses/gpl-3.0.html
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' SHEET EVENT HANDLER - DON´T CHANGE ANYTHING BELOW THIS LINE UNLESS YOU KNOW WHAT YOU DO
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdAddChartObjects_Click()
   ' delete existing charts and create new ones
   Call basCharts.deleteCharts(wks:=Me, MatchName:="Chart [A-E]")
   
   Call basCharts.addChart(wks:=Me, SrcRng:=Me.Range("E3:I8"), PlotBy:=xlColumns, ChartName:="Chart A", TopLeftCell:="E10", BottomRightCell:="I22")
   Call basCharts.addChart(wks:=Me, SrcRng:=Me.Range("E3:E8,H3:I8"), PlotBy:=xlColumns, ChartName:="Chart B", TopLeftCell:="E24", BottomRightCell:="I36")
   
   Call basCharts.addChart(wks:=Me, SrcRng:=Me.Range("K3:M8"), PlotBy:=xlRows, ChartName:="Chart C", TopLeftCell:="K10", BottomRightCell:="O22")
   Call basCharts.addChart(wks:=Me, SrcRng:=Me.Range("K4:M4,K6:M6,K8:M8"), PlotBy:=xlRows, ChartName:="Chart D", TopLeftCell:="K24", BottomRightCell:="O36")
   
   Call basCharts.addChart(wks:=Me, ChartName:="Chart E", TopLeftCell:="Q10", BottomRightCell:="U22")
End Sub

Private Sub cmdDeleteAllChartObjects_Click()
   ' delete all charts objects on actual sheet
   Call basCharts.deleteCharts(wks:=Me)
End Sub

Private Sub cmdSetChartCaptions_Click()
   ' add chart captions
   Call basCharts.setChartCaptions(oChart:=Me.ChartObjects("Chart A"), Title:="Chart A", XLabel:="xA", YLabel:="yA")
   Call basCharts.setChartCaptions(oChart:=Me.ChartObjects("Chart B"), Title:="Chart B", XLabel:="xB", YLabel:="yB")
   Call basCharts.setChartCaptions(oChart:=Me.ChartObjects("Chart A"), Title:="Chart A+")
End Sub

Private Sub cmdCopyChartScales_Click()
   ' clone axis scales of both axis from Chart A to Chart B
   Dim Scales As CHART_AXIS_SCALES
   
   ' set x-axis scales
   Scales = basCharts.getChartAxisScales(oChart:=Me.ChartObjects("Chart A"), Axis:=xlCategory)
   Call basCharts.setChartAxisScales(oChart:=Me.ChartObjects("Chart B"), Axis:=xlCategory, Scales:=Scales)
   ' set y-axis scales
   Scales = basCharts.getChartAxisScales(oChart:=Me.ChartObjects("Chart A"), Axis:=xlValue)
   Call basCharts.setChartAxisScales(oChart:=Me.ChartObjects("Chart B"), Axis:=xlValue, Scales:=Scales)
End Sub

Private Sub cmdCopyChartStyles_Click()
   ' clone custom chart formats of Chart A to Chart B
   Call basCharts.addCustomChartType(oChart:=Me.ChartObjects("Chart A"), CustomChartType:="DummyFormat")
   Call basCharts.setCustomChartType(oChart:=Me.ChartObjects("Chart B"), CustomChartType:="DummyFormat")
End Sub

Private Sub cmdAddSeries_Click()
   ' add a chart series to a given chart
   Call basCharts.addChartSeries(oChart:=Me.ChartObjects("Chart E"), XRng:=Me.Range("E4:E8"), YRng:=Me.Range("I4:I8"), Label:=Me.Range("I3"))
   'Call basCharts.addChartSeries(oChart:=Me.ChartObjects("Chart E"), XRng:=Me.Range("E4:E8,E4:E8"), YRng:=Me.Range("F4:F8,G4:G8"), Label:=Me.Range("F3,G3"))
   Call basCharts.addChartSeries(oChart:=Me.ChartObjects("Chart E"), XRng:=Me.Range("E4:E8,E4:E8"), YRng:=Me.Range("F4:F8,G4:G8"), Label:=Array("test", "hallo"))
End Sub

Private Sub cmdsetSeriesStyles_Click()
   ' set chart series styles
   Dim SeriesStyles As CHART_SERIES_STYLES
   
   ' get styles of specified chart object series
   SeriesStyles = basCharts.getChartSeriesStyles(oChart:=Me.ChartObjects("Chart E"), Series:=-2)
   
   ' modify series styles and apply it
   SeriesStyles.LC = 3
   SeriesStyles.MC = 3
   SeriesStyles.MFC = 1
   SeriesStyles.MS = xlMarkerStyleCircle
   
   Call basCharts.setChartSeriesStyles(oChart:=Me.ChartObjects("Chart E"), Series:=-2, Styles:=SeriesStyles)
End Sub

Private Sub cmdDeleteSeries_Click()
   ' delete a chart series
   Call basCharts.deleteChartSeries(oChart:=Me.ChartObjects("Chart E"), Series:=-1)
End Sub

Private Sub cmdChangeSeriesSource_Click()
   ' modify the x-rng of all series in Chart E
   Call basCharts.modifyChartSeries(oChart:=Me.ChartObjects("Chart E"), Search:="$E$", Replace:="$G$")
   Call basCharts.modifyChartSeries(oChart:=Me.ChartObjects("Chart E"), Search:="$G$", Replace:="$H$", Series:=-1)
End Sub

Private Sub cmdToggleLegendVisibility_Click()
   ' hide legend of Chart D and toggle visibilty state of Chart E
   Call basCharts.setChartLegendVisibility(oChart:=Me.ChartObjects("Chart D"), Visible:=False)
   Call basCharts.setChartLegendVisibility(oChart:=Me.ChartObjects("Chart E"))
End Sub
