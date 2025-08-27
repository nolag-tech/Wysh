using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;

namespace Wysh.Internal {
	public class ExcelProxy {
		private Application excel;

		public Application oExcel {
			get { return excel; }
		}

		public ExcelProxy() {
			excel = new Application();
			excel.Visible = true;

			if (excel.ActiveWorkbook == null) excel.Workbooks.Add();

			dynamic wb = excel.ActiveWorkbook;
			wb.Queries.FastCombine = true;
		}

		public void openWorkbook(string filename) {
			excel.Workbooks.Open(filename);
		}

		public void importCarat(string filename) {
			excel.Workbooks.OpenText(
				filename,
				1252,
				1,
				XlTextParsingType.xlDelimited,
				XlTextQualifier.xlTextQualifierNone,
				false,
				false,
				false,
				false,
				false,
				true,
				'^'
			);
		}

		public dynamic addQuery(string name, string query) {
			dynamic wb = excel.ActiveWorkbook;
			return wb.Queries.Add(name, query);
		}
		
	}
}
