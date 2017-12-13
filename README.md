# ExcelFacade
A quick and easy wrapper around Excel making life easier for .net developers.

I originally wrote this many years ago (before c# gained the `dynamic` keyword) for a project that 
produced reports in Excel that needed to be... "just so".

Over the years I've reused it several times on other projects, adding a little functionality here and there as required.

It doesn't cover the whole of the Excel COM interface, but certainly a whole lot of the commonly used stuff and lots of the less common too.

I've made it available here just in case it's of use to anyone else.

Example usage:

1. Starting Excel

```
using ExcelFacade;
using Application = ExcelFacade.Application;

public void CreateReport()
  {
  var excel = new Application();
  excel.Visible = true; // turn these off in release code to increase speed
  excel.Interactive = true;
  excel.ScreenUpdating = true;
  excel.DisplayAlerts = true;
    
  DoWork()
  }
```

2. Open an existing workbook

```
var workBook = excel.Workbooks.Add(filename);
var workSheet = workBook.Worksheets[1]; // all collections are 1 based, which is not the usual c# way
workSheet.Activate();

var r = workSheet.get_Range("K3");
r.Value = "*** DRAFT ***";
r.Font.Size = 20;
r.Font.Color = Color.Red;
r.HorizontalAlignment = XlHAlign.xlHAlignCenter;
```

3. Create a new workbook

```
var workBook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
var workSheet = workBook.Worksheets[1];

Range allCells = ws.Cells;
allCells.Font.Name = "Arial";
allCells.Font.Size = 10;
allCells.VerticalAlignment = XlVAlign.xlVAlignTop;
ws.DisplayPageBreaks = false; // helps speed things up
```

4. Page setup
```
public static void ExcelPageSetup(Worksheet ws)
  {
  // this routine will fail if no printer is available - check if excel.ActivePrinter.StartsWith("unknown") to avoid an exception
  Application excel = ws.Application;
  PageSetup ps = ws.PageSetup;

  const string rightHeader = "&\"Arial,Regular\"&14My Company Ltd";
  const string leftFooter = "&F";
  const string centerFooter = "&\"Arial,Regular\"&8Page &P of &N";
  const string rightFooter = "&\"Arial,Regular\"&8&D &T";

  decimal excelVersion = decimal.Parse(ws.Application.Version);
  try
    {
    if (excelVersion >= 14)
        ws.Application.PrintCommunication = false;

    ps.RightHeader = rightHeader;
    ps.LeftFooter = leftFooter;
    ps.CenterFooter = centerFooter;
    ps.RightFooter = rightFooter;

    double pts5mm = excel.CentimetersToPoints(0.5);
    double pts12mm = excel.CentimetersToPoints(1.2);
    ps.LeftMargin = pts5mm;
    ps.RightMargin = pts5mm;
    ps.TopMargin = pts12mm;
    ps.BottomMargin = pts12mm;
    ps.HeaderMargin = pts5mm;
    ps.FooterMargin = pts5mm;

    ps.Zoom = null;
    ps.FitToPagesWide = 1;
    ps.FitToPagesTall = null;
    ps.CenterHorizontally = true;

    ps.PrintTitleColumns = "";
    ps.PrintGridlines = false;
    ps.Orientation = XlPageOrientation.xlLandscape;
    ps.Draft = false;
    }
  finally
    {
    if (excelVersion >= 14)
        ws.Application.PrintCommunication = true;
    }
  }


