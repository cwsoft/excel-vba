# Excel-VBA (Excel 2010)

The `excel-vba` repository is a collection of handy Excel 2010 VBA modules for chart object manipulation, textfile handling (import, search, export) and more. The usage of the VBA routines and functions contained is shown in the demo sheets of the Excel 2010 `excel-vba.xlsm` file. There are now dependencies to other VBA projects, allowing you to use only the features you need.

----------------------------------------
## Package overview

- [basAPI](vba/basAPI.bas): Windows API routines to deal with user forms (e.g. borderless form)
- [basCharts](vba/basCharts.bas): Routines for the creation and manipulation of chart objects embedded on worksheets
- [basIO](vba/basIO.bas): Routines to import textfiles into sheets and to export a sheet as CSV file
- [basXSearch](vba/basXSearch.bas): Routines to search sheets for multiple search patterns
- [API](vba/wksAPI.cls), [frmAPI](vba/frmAPI.frm): Demo of the basAPI routines
- [Charts](vba/wksCharts.cls): Demo of the basChart routines
- [IO](vba/wksIO.cls): Demo of the basIO routines
- [xSearch](vba/wksXSearch.cls): Demo of the basXSearch routines

## Download
You can download the latest development branch of the `excel-vba` package using GitHubs [ZIP button](https://github.com/cwsoft/excel-vba/archive/master.zip).

## License
The `excel-vba` package is licensed under the [GNU General Public License 3.0](./LICENSE).

## Requirements
The `excel-vba` package was tested with Excel 2010 on a Windows 7 operating system. It may work for other platforms as well, but was not tested for.