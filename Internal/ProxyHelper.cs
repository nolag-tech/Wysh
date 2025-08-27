using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wysh.Internal {
	/*
	 * Some stuff doesn't work properly in the COM interfaces when pushed to ClearScript.
	 * This class contains calls that help to resolve some of the issues.
	 */
	public static class ProxyHelper {

		// Even though we are using the COM interface for Excel, ClearScript sees the underlying .NET interfaces.
		// The Queries collection of a Workbook is dynamic and therefore is invisible from within ClearScript.
		// This function solves the issue.
		public static dynamic ExcelAddQuery(Microsoft.Office.Interop.Excel.Application excel, string name, string query) {
			dynamic wb = excel.ActiveWorkbook;
			return wb.Queries.Add(name, query);
		}
	}
}
