

class Excel {
    constructor() {
        this.oExcel = new ExcelApp();
        this.oExcel.Visible = true;

        if (!this.oExcel.ActiveWorkbook) this.oExcel.Workbooks.Add();
        //this.oExcel.ActiveWorkbook.Queries.FastCombine = true;
    }

    openWorkbook(filename) {
        this.oExcel.Workbooks.Open(filename);
	}

	addSheet() {
		this.oExcel.ActiveWorkbook.Sheets.Add();
	}

	nameSheet(name) {
		this.oExcel.ActiveWorkbook.ActiveSheet.Name = name;
	}

	saveAs(toFile, fileFormat) {
		this.oExcel.DisplayAlerts = false;

		fileFormat = fileFormat ? fileFormat : Excel.xlWorkbookNormal;
		this.oExcel.ActiveWorkbook.SaveAs(Filename = toFile, FileFormat = fileFormat);
	}

	save() {
		this.oExcel.ActiveWorkbook.Save();
	}

	close(saveChanges = false) {
		this.oExcel.ActiveWorkbook.Close(saveChanges);
		this.oExcel.Quit();
	}

    importCarat(filename) {
		this.oExcel.WorkBooks.OpenText(
			filename, // file
			1252, // origin
			1, // start row
			Excel.xlDelimited, // delimited
			Excel.xlTextQualifierNone, // text qualifier
			false, // consec delimiter
			false, // tab
			false, // semicolon
			false, // comma
			false, // space
			true, // other
			'^' // delim char
		);
	}

	importText(inFile, columnSpec) {
		// column spec is an object with
		// names : array of column names
		// types : array of column types
		// delimiter : delimiter char (default is ^)
		// rangeStart : where to insert data (default is A1)
		// tableName : what to name range (default is Report)

		var delim = '^';
		if (columnSpec.delimiter) delim = columnSpec.delimiter;

		var rangeStart = "A1";
		if (columnSpec.rangeStart) rangeStart = columnSpec.rangeStart;

		var tableName = "Report";
		if (columnSpec.tableName) tableName = columnSpec.tableName;

		if (columnSpec.columns) {
			columnSpec.types = new Array();
			columnSpec.names = new Array();
			for (var i = 0; i < columnSpec.columns.length; i++) {
				columnSpec.names[i] = columnSpec.columns[i].name;
				columnSpec.types[i] = columnSpec.columns[i].type;
			}
			columnSpec.names = columnSpec.names.join();
		}

		// create new workbook
		if (!this.oExcel.ActiveWorkbook) this.oExcel.Workbooks.Add();

		var query = this.oExcel.ActiveWorkbook.ActiveSheet.QueryTables.Add(Connection = "TEXT;" + inFile, Destination = this.oExcel.ActiveWorkbook.ActiveSheet.Range(rangeStart));
		query.Name = 'IMPORT';
		query.FieldNames = true;
		query.RowNumbers = false;
		query.FillAdjacentFormulas = false;
		query.PreserveFormatting = true;
		query.RefreshOnFileOpen = false;
		query.RefreshStyle = Excel.xlInsertDeleteCells;
		query.SavePassword = false;
		query.SaveData = true;
		query.AdjustColumnWidth = true;
		query.RefreshPeriod = 0;
		query.TextFilePromptOnRefresh = false;
		//query.TextFilePlatform = 437;
		//query.TextFilePlatform = 65000;
		query.TextFilePlatform = 2;
		query.TextFileStartRow = 1;
		query.TextFileParseType = Excel.xlDelimited;
		if (columnSpec.hasOwnProperty("quoted") && columnSpec.quoted) query.TextFileTextQualifier = Excel.xlTextQualifierNDoubleQuote;
		else query.TextFileTextQualifier = Excel.xlTextQualifierNone;
		query.TextFileConsecutiveDelimiter = false;
		query.TextFileTabDelimiter = false;
		query.TextFileSemicolonDelimiter = false;
		query.TextFileCommaDelimiter = false;
		query.TextFileSpaceDelimiter = false;
		query.TextFileOtherDelimiter = delim;
		//query.TextFileColumnDataTypes = JS2VBArray(columnSpec.types);
		query.TextFileColumnDataTypes = columnSpec.types.toVBArray();
		query.TextFileTrailingMinusNumbers = true;
		query.Refresh(BackgroundQuery = false);

		query.Delete();


		if (columnSpec.hasTable) {
			var tbl = this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects.Add(Excel.xlSrcRange, this.oExcel.ActiveWorkbook.ActiveSheet.Range(rangeStart).CurrentRegion, null, Excel.xlYes);
			tbl.Name = tableName;
		}

		this.nameColumns(columnSpec.names, rangeStart);
	}

	queryTable(qname, tname, unlink = false) {
		var odbcStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" + qname + ";Extended Properties=\"\"";
		var qlist = this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects.Add(3, odbcStr, null, Excel.xlYes, this.oExcel.ActiveWorkbook.ActiveSheet.Range("$A$1"));

		/*var qtable = this.oExcel.ActiveWorkbook.ActiveSheet.QueryTables.Add(
			Connection=odbcStr,
			Destination=this.oExcel.ActiveWorkbook.ActiveSheet.Range("$A$1")
		);*/

		var qtable = qlist.QueryTable;
		qtable.CommandType = 2; // xlCmdSql (XlCmdType)
		qtable.CommandText = "SELECT * FROM [" + qname + "]";
		qtable.RowNumbers = false;
		qtable.FillAdjacentFormulas = false;
		qtable.PreserveFormatting = true;
		qtable.RefreshOnFileOpen = false;
		qtable.BackgroundQuery = true;
		qtable.RefreshStyle = Excel.xlInsertDeleteCells;
		qtable.SavePassword = false;
		qtable.SaveData = true;
		qtable.AdjustColumnWidth = true;
		qtable.RefreshPeriod = 0;
		qtable.PreserveColumnInfo = true;
		qtable.ListObject.DisplayName = tname;
		qtable.Refresh(BackgroundQuery = false);

		if (unlink) qtable.ListObject.Unlink();

		return qtable;
	}

	formatify(toList = true, listName = "Report", size) {
		this.oExcel.ActiveWorkbook.ActiveSheet.Rows("1:1").Select();
		this.oExcel.Selection.Font.Bold = true;
		this.oExcel.ActiveWorkbook.ActiveSheet.Cells.Select();
		if (size) this.oExcel.Selection.Font.Size = size;
		this.oExcel.Selection.Font.Size = size;
		this.oExcel.ActiveWorkbook.ActiveSheet.Cells.EntireColumn.Autofit();
		this.oExcel.ActiveWorkbook.ActiveSheet.Range("A2").Select();
		this.oExcel.ActiveWindow.FreezePanes = true;

		if(toList) this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects.Add(Excel.xlSrcRange, this.oExcel.ActiveWorkbook.ActiveSheet.UsedRange, null, Excel.xlYes).Name = listName;
	}

	printify(area) {
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintTitleRows = "$1:$1";
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintGridlines = true;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.Orientation = Excel.xlLandscape;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.CenterHorizontally = true;

		if (area) this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintArea = area;
	}

	setColWidth(col, width) {
		this.oExcel.ActiveWorkbook.ActiveSheet.Columns(col + ":" + col).ColumnWidth = width;
	}

	setRowHeight(row, height) {
		this.oExcel.ActiveWorkbook.ActiveSheet.Rows(row + ":" + row).RowHeight = height;
	}

	addTopHeader() {
		this.oExcel.ActiveWorkbook.ActiveSheet.Rows("1:1").Select();
		this.oExcel.Selection.Insert; // Shift:=Excel.xlDown, CopyOrigin:=Excel.xlFormatFromLeftOrAbove; 
	}

	mergeCells(dat) {
		if (dat.range == null) return;
		if (dat.bkColor) this.oExcel.ActiveWorkbook.ActiveSheet.Range(dat.range).Font.ColorIndex = dat.valColor;
		if (dat.flgBold) this.oExcel.ActiveWorkbook.ActiveSheet.Range(dat.range).Font.Bold = dat.flgBold;
		if (dat.bkColor) this.oExcel.ActiveWorkbook.ActiveSheet.Range(dat.range).Interior.ColorIndex = dat.bkColor;
		var cells = dat.range.split(':');
		if (dat.title) this.oExcel.ActiveWorkbook.ActiveSheet.Range(cells[0]).value = dat.title; 
	}

	freezeCell(cellNum) {
		this.oExcel.ActiveWorkbook.ActiveSheet.Range(cellNum).Select();
		this.oExcel.ActiveWindow.FreezePanes = false;
		this.oExcel.ActiveWindow.FreezePanes = true;
	}

	hideColumns(cols) {
		if (cols == null) return;

		var clist = cols.split(',');
		for (var i = 0; i < clist.length; i++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.Columns(clist[i]).EntireColumn.Hidden = true;//.Select(); 
		}
	}

	hideTableColumns(colNames, rptnam = "Report") {
		if (colNames == null) return;

		var clist = colNames.split(',');
		for (var i = 0; i < clist.length; i++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(rptnam).ListColumns(clist[i]).Visible = true;
		}
	}

	getColumnLetter(n) {
		var colLetter = '';

		n++;
		while (n > 0) {
			n--;
			colLetter = String.fromCharCode(65 + (n % 26)) + colLetter;
			n = Math.floor(n / 26);
		}

		return colLetter;
	}

	nameColumns(names, rangeStart) {
		if (names == null) return;

		if (!rangeStart) rangeStart = "A1";

		var range = this.oExcel.ActiveWorkbook.ActiveSheet.Range(rangeStart);

		var clist = names.split(',');
		for (var i = 0; i < clist.length; i++) {
			range.Offset(0, i).FormulaR1C1 = clist[i].replace('|', '\n');
		}
	}

	showTotals(tblname = "Report") {
		this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(tblname).ShowTotals = true;
	}

	totalSum(colNames, rptnam = "Report", highlight) {
		if (colNames == null) return;
		this.showTotals(rptnam);

		var hlist = [];  // array of Headers to be highlighted in Totals row;  should be a subset of clist array
		if (highlight != null) hlist = highlight.split(',');

		var clist = colNames.split(',');
		for (var i = 0; i < clist.length; i++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(rptnam).ListColumns(clist[i]).TotalsCalculation = Excel.xlTotalsCalculationSum;
			// highlight row logic -- if a Column in the Totals Row is to be highlighted, the Highlight Name must match the Header value  (e.g. ‘LPI Accrual’ = ‘LPI Accrual’) 
			if (hlist.indexOf(clist[i]) > -1) {
				for (var w = 1; w <= this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(rptnam).ListColumns.count; w++) {
					if (this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(rptnam).ListColumns.Item(w) == clist[i]) {
						this.oExcel.ActiveWorkbook.ActiveSheet.Cells(this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(rptnam).ListRows.count + 2, w).Font.ColorIndex = 5;
					}
				}
			}
		}
	}

	totalAvg(colNames, prec, rptnam = "Report") {
		if (colNames == null) return;
		if (prec == null) prec = -1;
		this.showTotals(rptnam);

		var clist = colNames.split(',');
		for (var i = 0; i < clist.length; i++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(rptnam).ListColumns(clist[i]).TotalsCalculation = Excel.xlTotalsCalculationAverage;
		}
	}

	pageSetup(setup) {
		var sheet = this.oExcel.ActiveWorkbook.ActiveSheet;

		if (setup.orientation) sheet.PageSetup.Orientation = (setup.orientation == "portrait") ? Excel.xlPortrait : Excel.xlLandscape;
		if (setup.paperSize) sheet.PageSetup.PaperSize = (setup.paperSize == "ledger") ? Excel.xlLedger : setup.paperSize == "legal" ? Excel.xlLegal : Excel.xlLetter;
		if (setup.firstPage) sheet.PageSetup.FirstPageNumber = setup.firstPage;

		if (setup.centerHeader) sheet.PageSetup.CenterHeader = setup.centerHeader;
		if (setup.centerFooter) sheet.PageSetup.CenterFooter = setup.centerFooter;
		if (setup.leftHeader) sheet.PageSetup.LeftHeader = setup.leftHeader;
		if (setup.leftFooter) sheet.PageSetup.LeftFooter = setup.leftFooter;
		if (setup.rightHeader) sheet.PageSetup.RightHeader = setup.rightHeader;
		if (setup.rightFooter) sheet.PageSetup.RightFooter = setup.rightFooter;

		if (setup.leftMargin) sheet.PageSetup.LeftMargin = this.oExcel.InchesToPoints(setup.leftMargin);
		if (setup.rightMargin) sheet.PageSetup.RightMargin = this.oExcel.InchesToPoints(setup.rightMargin);
		if (setup.topMargin) sheet.PageSetup.TopMargin = this.oExcel.InchesToPoints(setup.topMargin);
		if (setup.bottomMargin) sheet.PageSetup.BottomMargin = this.oExcel.InchesToPoints(setup.bottomMargin);
		if (setup.headerMargin) sheet.PageSetup.HeaderMargin = this.oExcel.InchesToPoints(setup.headerMargin);
		if (setup.footerMargin) sheet.PageSetup.FooterMargin = this.oExcel.InchesToPoints(setup.footerMargin);


		if (setup.hasOwnProperty('fitToPagesWide')) {
			sheet.PageSetup.Zoom = false;
			sheet.PageSetup.FitToPagesWide = setup.fitToPagesWide;
		}
		if (setup.hasOwnProperty('fitToPagesTall')) {
			sheet.PageSetup.Zoom = false;
			sheet.PageSetup.FitToPagesTall = setup.fitToPagesTall;
		}
		if (setup.titleRows) sheet.PageSetup.PrintTitleRows = "$1:$" + setup.titleRows;
	}

	savePDF(filename) {
		this.oExcel.ActiveWorkbook.ActiveSheet.ExportAsFixedFormat(Type = this.oExcel.xlTypePDF, Filename = filename, Quality = 1);
	}

	toFormat(cols, format) {
		if (cols == null) return;

		var clist = cols.split('|');
		for (var i = 0; i < clist.length; i++) {
			if (clist[i].test(/^[0-9]+$/)) this.oExcel.ActiveWorkbook.ActiveSheet.Rows(clist[i] + ":" + clist[i]).NumberFormat = format; //.Select();
			else this.oExcel.ActiveWorkbook.ActiveSheet.Columns(clist[i] + ":" + clist[i]).NumberFormat = format; //.Select();
		}

		this.deselect();
	}

	toFormatCells(rows, cols, format) {
		if (cols == null || rows == null) return;
		var rlist = rows.split('|');
		var clist = cols.split('|');
		for (var i = 0; i < rlist.length; i++) {
			if (!rlist[i].test(/^[0-9]+$/)) continue;
			for (var k = 0; k < clist.length; k++) {
				this.oExcel.ActiveWorkbook.ActiveSheet.Cells(rlist[i], parseInt(clist[k])).NumberFormat = format;
			}
		}
		this.deselect();
	}

	printForm() {
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintTitleRows = "$1:$1";
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintGridlines = true;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.Orientation = Excel.xlPortrait;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.CenterHorizontally = true;

		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.LeftMargin = this.oExcel.Application.InchesToPoints(0.25);
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.RightMargin = this.oExcel.Application.InchesToPoints(0.25);
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.TopMargin = this.oExcel.Application.InchesToPoints(0.25);
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.BottomMargin = this.oExcel.Application.InchesToPoints(0.25);

		this.oExcel.ActiveWorkbook.ActiveSheet.PrintOut();
	}

	printit() {
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.PrintGridlines = true;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.Orientation = Excel.xlLandscape;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.LeftMargin = 0.25;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.RightMargin = 0.25;
		this.oExcel.ActiveWorkbook.ActiveSheet.PrintOut();
	}

	// pass plain text or RTF
	// &D = date, &P = page, &N = pages
	// &"Arial,Bold"some text = sets font to arial bold
	// &"-,Bold"some text = keep font, set to bold
	// &12 = set size to 12
	setHeader(left, center, right) {
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.LeftHeader = left;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.CenterHeader = center;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.RightHeader = right;
	}

	setFooter(left, center, right) {
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.LeftFooter = left;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.CenterFooter = center;
		this.oExcel.ActiveWorkbook.ActiveSheet.PageSetup.RightFooter = right;
	}

	setColumnWidth(col, width) {
		this.oExcel.ActiveWorkbook.ActiveSheet.Columns(col + ":" + col).ColumnWidth = width;
	}

	columnTotals(cols, tableName) {
		if (cols == null) return;

		//var rowCount = this.oExcel.ActiveWorkbook.ActiveSheet.usedRange.Rows.Count;

		this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(tableName ? tableName : "Report").ShowTotals = true;

		var clist = cols.split(',');
		for (var i = 0; i < clist.length; i++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(tableName ? tableName : "Report").ListColumns(clist[i]).TotalsCalculation = Excel.xlTotalsCalculationSum;
		}
	}

	deselect() {
		this.oExcel.ActiveWorkbook.ActiveSheet.Range("A2").Select();
	}

	setColumnStyle(cols, style) {
		if (cols == null) return;

		var clist = cols.split(',');
		for (var i = 0; i < clist.length; i++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.Columns(clist[i] + ":" + clist[i]).Style = style;//Select();
		}

		this.deselect();
	}

	setRowStyle(rows, style) {
		if (rows == null) return;

		var clist = rows.split(',');
		for (var i = 0; i < clist.length; i++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.Rows(clist[i] + ":" + clist[i]).Style = style;//Select();
		}

		this.deselect();
	}

	toStyle(cols, style) {
		if (cols == null) return;

		var clist = cols.split(',');
		for (var i = 0; i < clist.length; i++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.Columns(clist[i] + ":" + clist[i]).Style = style;//Select();
		}

		this.deselect();
	}

	toCurrency(cols) {
		this.toStyle(cols, "Currency");
	}

	toPercent(cols) {
		this.toStyle(cols, "Percent");
	}

	toDate(cols) {
		this.toStyle(cols, "Date");
	}

	toText(cols) {
		this.toStyle(cols, "Text");
	}

	colHighlight(cols) {
		if (cols == null) return;

		var clist = cols.split(',');
		for (var i = 0; i < clist.length; i++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.Columns(clist[i] + ":" + clist[i]).Interior.ColorIndex = 6;
			this.oExcel.ActiveWorkbook.ActiveSheet.Columns(clist[i] + ":" + clist[i]).Interior.Pattern = Excel.xlSolid;
		}
	}

	colBoldValues(cols) {
		if (cols == null) return;

		var clist = cols.split(',');
		for (var i = 0; i < clist.length; i++) {
			var fc = this.oExcel.ActiveWorkbook.ActiveSheet.Columns(clist[i] + ":" + clist[i]).FormatConditions;
			fc.Add(Excel.xlCellValue, Excel.xlGreater, "=1");
			fc(fc.Count).SetFirstPriority();
			fc(1).Font.Bold = true;
			fc(1).Font.Italic = false;
			fc(1).Font.TintAndShade = 0;
			fc(1).StopIfTrue = false;
		}
	}

	cellHighlight(rowNum, colNum, color, hlight, incr, cnt) {
		if (color == null || !Validator.isNumeric(color)) color = 6;
		if (colNum == null || rowNum == null) return;

		if (hlight == 'borderbold') {
			this.oExcel.ActiveWorkbook.ActiveSheet.Cells(rowNum, colNum).Font.ColorIndex = color;
			this.oExcel.ActiveWorkbook.ActiveSheet.Cells(rowNum, colNum).Borders.ColorIndex = color;
			this.oExcel.ActiveWorkbook.ActiveSheet.Cells(rowNum, colNum).Borders.Weight = 4;
		}
		else if (hlight == 'fontcolor') {
			this.oExcel.ActiveWorkbook.ActiveSheet.Cells(rowNum, colNum).Font.ColorIndex = color;
		}
		else {
			this.oExcel.ActiveWorkbook.ActiveSheet.Cells(rowNum, colNum).Interior.ColorIndex = color;
		}

		if (incr == null || cnt == null) return;
		for (var j = 1; j <= cnt; j++) {
			this.oExcel.ActiveWorkbook.ActiveSheet.Cells(rowNum + (j * incr), colNum).Interior.ColorIndex = color;
		}
	}

	getColNum(colName, tableName) {
		if (colName == null) return -1;
		var colNum = -1;
		WScript.StdOut.WriteLine('colName = ' + colName);
		WScript.StdOut.WriteLine('tot Cols = ' + this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(tableName ? tableName : "Report").ListColumns.count);
		for (var w = 1; w < this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(tableName ? tableName : "Report").ListColumns.count; w++) {
			if (this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(tableName ? tableName : "Report").ListColumns.Item(w) == colName) {
				colNum = w;
			}
		}

		return colNum;
	}

	numRows(tableName) {
		return this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(tableName ? tableName : "Report").ListRows.count; 
	}

	sumCol(colName, tableName) {
		var colNum = -1;

		var colNum = this.getColNum(colName, tableName ? tableName : "Report");
		if (colNum < 0) return 0;

		var tot = 0;
		for (var i = 2; i < this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(tableName ? tableName : "Report").ListRows.count + 2; i++) {

			if (Validator.isNumeric(this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).value)) {
				tot = tot + this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).value;
			}
		}
		
		return tot;
	}
	// for a specific column, do a comparison check on each value & if true, highlight the entire row (or the cell itself)
	//   note:  val can be an integer or a string;  type will be 'gte' [for >=], 'lte' [for <=], or 'equal'
	//   note2:  feat_type refers to what will be highlighted (possible values are 'row' and 'cell')
	//   note3:  valColor & bkColor both refer to Excel Color Index values;  default for bkColor is null (no background color)
	colCondFeatureData(dat) {

		if (dat.type == null) dat.type = 'equal';
		if (dat.type != 'equal' && dat.type != 'gte' && dat.type != 'lte' && dat.type != 'all') return;
		if (dat.feat_type == null) dat.feat_type = 'row';

		if (dat.colName == null) return;
		if (dat.val == null) dat.val = 0;
		if (dat.valColor == null) dat.valColor = 5;  	// default of Blue
		if (dat.flgBold == null) dat.flgBold = true;  	// default of Bold

		var colNum = -1;

		for (var w = 1; w < this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(dat.tableName ? dat.tableName : "Report").ListColumns.count; w++) {
			if (this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(dat.tableName ? dat.tableName : "Report").ListColumns.Item(w) == dat.colName) {
				colNum = w;
			}
		}
		if (colNum < 0) return;

		// begins in row 2 to exclude the Header record 
		for (var i = 2; i < this.oExcel.ActiveWorkbook.ActiveSheet.ListObjects(dat.tableName ? dat.tableName : "Report").ListRows.count + 2; i++) {
			if ((!Validator.isNumeric(dat.val) || (Validator.isNumeric(dat.val) && Validator.isNumeric(this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).value))) &&
				((dat.type == 'gte' && this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).value >= dat.val)   // highlights all values >= min value
					|| (dat.type == 'lte' && this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).value <= dat.val)   // highlights all values <= max value
					|| (dat.type == 'equal' && this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).value == dat.val)
					|| (dat.type == 'all'))) {
				if (dat.feat_type == 'row') {
					this.oExcel.ActiveWorkbook.ActiveSheet.Rows(i).Font.ColorIndex = dat.valColor;
					this.oExcel.ActiveWorkbook.ActiveSheet.Rows(i).Font.Bold = dat.flgBold;
					if (dat.bkColor) this.oExcel.ActiveWorkbook.ActiveSheet.Rows(i).Interior.ColorIndex = dat.bkColor;
				}
				else {
					this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).Font.ColorIndex = dat.valColor;
					this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).Font.Bold = dat.flgBold;
					if (dat.bkColor) this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).Interior.ColorIndex = dat.bkColor;
					if (dat.borderColor) {
						if (dat.borderLeft) {
							this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).Borders(Excel.xlEdgeLeft).ColorIndex = dat.borderColor;  // 7 corresponds to Left border;  10 for Rt;  8 for Top;  9 for Bottom
							this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).Borders(Excel.xlEdgeLeft).Weight = (dat.borderWeight) ? dat.borderWeight : 3;
						}
						else {
							this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).Borders.ColorIndex = dat.borderColor;
							this.oExcel.ActiveWorkbook.ActiveSheet.Cells(i, colNum).Borders.Weight = (dat.borderWeight) ? dat.borderWeight : 3;
						}
					}
				}
			}
		}
	}
}

Object.assign(Excel, {
	xlPortrait: 1,
	xlLandscape: 2,

	xlSolid: 1,
	xlThick: 4,
	xlDouble: -4119,
	xlNone: -4142,
	xlAutomatic: -4105,

	xlSrcExternal: 0,
	xlSrcRange: 1,
	xlSrcXml: 2,
	xlSrcQuery: 3,

	xlTop: -4160,
	xlCenter: -4108,
	xlBottom: -4107,
	xlLeft: -4131,
	xlRight: -4152,

	xlDown: -4121, // XlDirection constants 
	xlUp: -4162,
	xlToLeft: -4159,
	xlToRight: -4161,

	xlFormatFromLeftOrAbove: 0, 	// XlInsertFormatOrigin constants
	xlFormatFromRightOrBelow: 1,

	xlCellValue: 1,
	xlExpression: 2,

	xlEdgeLeft: 7,
	xlEdgeTop: 8,
	xlEdgeBottom: 9,
	xlEdgeRight: 10,

	xlBetween: 1,
	xlNotBetween: 2,
	xlEqual: 3,
	xlNotEqual: 4,
	xlGreater: 5,
	xlLess: 6,
	xlGreaterEqual: 7,
	xlLessEqual: 8,

	xlWhole: 1,
	xlPart: 2,

	xlLegal: 5,
	xlPaperLegal: 5,
	xlLetter: 1,
	xlPaperLetter: 1,
	xlPaperLetterSmall: 2,
	xlLedger: 4,
	xlPaperLedger: 4,
	xlPaper10x14: 16,
	xlPaper11x17: 17,
	xlPaperExecutive: 7,
	xlPaperFolio: 14,
	xlPaperNote: 18,
	xlPaperQuarto: 15,
	xlPaperStatement: 6,
	xlPaperTabloid: 3,
	xlPaperFanfoldUS: 39,
	xlPaperFanfoldStdGerman: 40,
	xlPaperFanfoldLegalGerman: 41,
	xlPaperA3: 8,
	xlPaperA4: 9,
	xlPaperA4Small: 10,
	xlPaperA5: 11,
	xlPaperB4: 12,
	xlPaperB5: 13,
	xlPaperCsheet: 24,
	xlPaperDsheet: 25,
	xlPaperEsheet: 26,
	xlPaperEnvelope10: 20,
	xlPaperEnvelope11: 21,
	xlPaperEnvelope12: 22,
	xlPaperEnvelope14: 23,
	xlPaperEnvelope9: 19,
	xlPaperEnvelopeB4: 33,
	xlPaperEnvelopeB5: 34,
	xlPaperEnvelopeB6: 35,
	xlPaperEnvelopeC3: 29,
	xlPaperEnvelopeC4: 30,
	xlPaperEnvelopeC5: 28,
	xlPaperEnvelopeC6: 31,
	xlPaperEnvelopeC65: 32,
	xlPaperEnvelopeDL: 27,
	xlPaperEnvelopeItaly: 36,
	xlPaperEnvelopeMonarch: 37,
	xlPaperEnvelopePersonal: 38,

	xlLandscape: 2,
	xlPortrait: 1,

	xlSourceWorkbook: 0,
	xlSourceSheet: 1,
	xlSourcePrintArea: 2,
	xlSourceAutoFilter: 3,
	xlSourceRange: 4,
	xlSourceChart: 5,
	xlSourcePivotTable: 6,
	xlSourceQuery: 7,

	xlDoNotSaveChanges: 2,
	xlSaveChanges: 1,

	xlNoChange: 1,
	xlShared: 2,
	xlNoChange: 3,

	xlSummaryAbove: 0,
	xlSummaryBelow: 1,

	xlDelimited: 1,
	xlFixedWidth: 2,

	xlInsertDeleteCells: 1,
	xlInsertEntireRows: 2,

	xlTextQualifierDoubleQuote: 1,
	xlTextQualifierSingleQuote: 2,
	xlTextQualifierNone: -4142,

	xlGeneralFormat: 1,
	xlTextFormat: 2,

	xlWorkbookNormal: -4143,
	xlWorkbookBinary: 50,
	xlExcel12: 50,
	xlWorkbookXLSX: 51,
	xlOpenXMLWorkbook: 51,
	xlWorkbookXLSM: 52,
	xlOpenXMLWorkbookMacroEnabled: 52,
	xlExcel8: 56,
	xlCSV: 6,
	xlTXT: -4158,
	xlPRN: 36,

	// for fixed files
	xlTypePDF: 0,
	xlTypeXPS: 1,

	xlGuess: 0,
	xlYes: 1,
	xlNo: 2,

	xlTotalsCalculationNone: 0,
	xlTotalsCalculationAverage: 2,
	xlTotalsCalculationCount: 3,
	xlTotalsCalculationCountNums: 4,
	xlTotalsCalculationMax: 6,
	xlTotalsCalculationMin: 5,
	xlTotalsCalculationStdDev: 7,
	xlTotalsCalculationSum: 1,
	xlTotalsCalculationVar: 8,

	xlAnd: 1,
	xlBottom10Items: 4,
	xlBottom10Percent: 6,
	xlFilterCellColor: 8,
	xlFilterDynamic: 11,
	xlFilterFontColor: 9,
	xlFilterValues: 7,
	xlOr: 2,
	xlTop10Items: 3,
	xlTop10Percent: 5
});
