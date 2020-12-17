private ExcelWorksheet CreateWorkSheetData(IList<DailyReport> listDailyReport, ExcelWorksheet excelWorksheet)
{
    //header
    excelWorksheet.Cells["A1"].Value = "DAY";
    excelWorksheet.Cells["B1"].Value = "CLASSIFICATION";
    excelWorksheet.Cells["C1"].Value = "BANK";
    excelWorksheet.Cells["D1"].Value = "AMOUNT";
    excelWorksheet.Cells["E1"].Value = "TAX";
    excelWorksheet.Cells["F1"].Value = "NET VALUE";

    excelWorksheet.Cells["A1:F1"].Style.Font.Name = "Calibri";
    excelWorksheet.Cells["A1:F1"].Style.Font.Size = 11;
    excelWorksheet.Cells["A1:F1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
    excelWorksheet.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
    excelWorksheet.Cells["A1:F1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
    excelWorksheet.Cells["A1:F1"].Style.Font.Bold = true;

    int rowIndex = 2;
    
    foreach(DailyReport row in listDailyRport)
    {
        excelWorksheet.Cells["A" + rowIndex.ToString()].Value = row.Day;
        excelWorksheet.Cells["B" + rowIndex.ToString()].Value = row.Classification;
        excelWorksheet.Cells["C" + rowIndex.ToString()].Value = row.Bank;
        excelWorksheet.Cells["D" + rowIndex.ToString()].Value = row.Amount;
        excelWorksheet.Cells["E" + rowIndex.ToString()].Value = row.Tax;
        excelWorksheet.Cells["F" + rowIndex.ToString()].Value = row.NetValue;
        
        rowIndex++;
    }

    excelWorksheet.Cells["D2:F" + rowIndex.ToString()].Style.Numberformat.Format = "#,##0.00";
    excelWorksheet.Cells["A1:F" + rowIndex.ToString()].AutoFitColumns();   

    return excelWorksheet;
}

 private Workbook CreatePivotTableWorkSheetConsolidatedMonth(Workbook wbBaseData, PivotCache cacheData)
{
    Worksheet sheetConsolidatedMonth = wbBaseDados.CreateEmptySheet();
    sheetConsolidatedMonth.Name = "Consolidado_Mês";
    sheetConsolidatedMonth.MoveWorksheet(0);

    PivotTable pivotTable = sheetConsolidatedMonth.PivotTables.Add("Dinamic Table", wbBaseData.Worksheets["Data"].Range["A1"], cacheData);
    pivotTable.Cache.IsRefreshOnLoad = true;

    //row labels
    var r1 = pivotTable.PivotFields["CLASSIFICATION"];
    r1.Axis = AxisTypes.Row;

    var r2 = pivotTable.PivotFields["BANK"];
    r2.Axis = AxisTypes.Row;

    //data fields and set format
    pivotTable.DataFields.Add(pivotTable.PivotFields["AMOUNT"], "Sum of AMOUNT", SubtotalTypes.Sum);
    pivotTable.DataFields.Add(pivotTable.PivotFields["NET VALUE"], "Sum of NET VALUE", SubtotalTypes.Sum);
    pivotTable.DataFields.Add(pivotTable.PivotFields["TAX"], "Sum of TAX", SubtotalTypes.Sum);
    pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium9;

    //FILTER
    PivotReportFilter filter = new PivotReportFilter("DAY", true);
    pivotTable.ReportFilters.Add(filter);

    pivotTable.PivotFields["AMOUNT"].NumberFormat = "#,##0.00";
    pivotTable.PivotFields["NET VALUE"].NumberFormat = "#,##0.00";
    pivotTable.PivotFields["TAX"].NumberFormat = "#,##0.00";

    pivotTable.AllSubTotalTop = false;
    pivotTable.Options.RowHeaderCaption = "CLASSIFICATION";
    pivotTable.Options.ShowGridDropZone = false;
    pivotTable.Options.RowLayout = PivotTableLayoutType.Tabular;
    pivotTable.Options.ShowFieldList = false;
    pivotTable.CalculateData();
    
    sheetConsolidatedMonth.AutoFitColumn(1);
    sheetConsolidatedMonth.AutoFitColumn(2);
    sheetConsolidatedMonth.AutoFitColumn(3);
    sheetConsolidatedMonth.AutoFitColumn(4);
    sheetConsolidatedMonth.AutoFitColumn(5);

    return wbBaseDados;
}


