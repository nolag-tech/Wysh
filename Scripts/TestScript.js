

Console.WriteLine("Got here.");

let fso = new FileSystemObject();
Console.WriteLine(fso);


!function(global) {
	global.doOutput = function() {
		Console.WriteLine("Have a nice day.");
	};
}(this);

doOutput();

oExcel = new ExcelApp();
oExcel.Visible = true;
oExcel.Workbooks.Add();
//oExcel.ActiveWorkbook.Queries.FastCombine = true;
oExcel.ActiveWorkbook.ActiveSheet.Name = "HelloSheet";
