using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ISpreadsheet.Implementation.EPPlus {
	/// <summary>
	/// Workbook implementation for Epplus (Excel 2077+, OpenDocument)
	/// </summary>
	public class EPPlusWorkbook : IWorkbook {

		public ExcelPackage Package { get; set; }
		public ExcelWorkbook Workbook { get; set; }
		private FileStream _stream;

		public static IWorkbook OpenFromFile(string filename) {
			var file = new FileStream(filename, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
			var package = new ExcelPackage(file);
			return new EPPlusWorkbook(package, file);
		}

		public static IWorkbook OpenFromStream(Stream stream) {
			var package = new ExcelPackage(stream);
			return new EPPlusWorkbook(package);
		}

		public void Dispose() {
			if (Package != null) {
				Package.Dispose();
				Package = null;

			}
			if (_stream != null) {
				_stream.Close();
				_stream.Dispose();
			}
		}

		public EPPlusWorkbook(ExcelPackage package, FileStream file = null) {
			_stream = file;
			Package = package;
			Workbook = Package.Workbook;
		}

		public IWorksheet GetSheet(string name) {
			var sheet = Workbook.Worksheets[name];
			return new EPPlusWorksheet(sheet);
		}

		public IWorksheet GetSheet(int num) {
			var sheet = Workbook.Worksheets[num];
			return new EPPlusWorksheet(sheet);
		}

		public IWorksheet[] Sheets {
			get {
				var list = new List<IWorksheet>();
				for (int i = 1; i <= Workbook.Worksheets.Count; i++) {
					list.Add(new EPPlusWorksheet(Workbook.Worksheets[i]));
				}
				return list.ToArray();
			}
		}
	}
}