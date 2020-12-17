var reportPath = "~/Content/report.xlsx";

FileInfo newFile = new FileInfo(Server.MapPath(reportPath));

if (newFile.Exists)
{
    newFile.Delete();  // ensures we create a new workbook
    newFile = new FileInfo(Server.MapPath(reportPath));
}

using (ExcelPackage pck1 = new ExcelPackage(newFile))
{
    var wsData = pck1.Workbook.Worksheets.Add("Data");
        wsData = CreateWorkSheetData(lista, wsData);
        
        pck1.Save();
}

//after create worksheet 'Data' then we create worksheet with Pivot Table
Workbook workbook = new Workbook();
workbook.LoadFromFile(newFile.FullName);

Worksheet sheetData = workbook.Worksheets["Data"];
var lastRow = sheetData.Rows.Count();

CellRange dataRangeData = sheetData.Range["A1:M" + (lastRow - 1).ToString()];
PivotCache cacheData = workbook.PivotCaches.Add(dataRangeData);

workbook = CreatePivotTableWorkSheetConsolidatedMonth(workbook, cacheData);

workbook.SaveToFile(Server.MapPath(reportPath), ExcelVersion.Version2010);
