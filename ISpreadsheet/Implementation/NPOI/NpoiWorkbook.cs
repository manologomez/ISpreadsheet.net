using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;

namespace ISpreadsheet.Implementation.NPOI {
	/// <summary>
	/// Workbook implementation for NPOI (Excel 2003 binary format)
	/// </summary>
	public class NpoiWorkbook : IWorkbook {

		public HSSFWorkbook Book { get; set; }
		private FormulaEvaluator _evaluator;

		public static IWorkbook OpenFromFile(string filename) {
			var file = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
			var fs = new POIFSFileSystem(file);
			var book = new HSSFWorkbook(fs);
			return new NpoiWorkbook(book);
		}

		public static IWorkbook OpenFromStream(Stream stream) {
			var book = new HSSFWorkbook(stream);
			return new NpoiWorkbook(book);
		}

		public void Dispose() {
			if (Book != null)
				Book.Dispose();
		}

		public NpoiWorkbook(HSSFWorkbook book) {
			_evaluator = book.GetCreationHelper().CreateFormulaEvaluator();
			Book = book;
		}

		public IWorksheet GetSheet(string name) {
			var sheet = Book.GetSheet(name);
			return new NpoiWorksheet(Book, sheet, _evaluator);
		}

		public IWorksheet GetSheet(int num) {
			var sheet = Book.GetSheetAt(num - 1);
			return new NpoiWorksheet(Book, sheet, _evaluator);
		}

		public IWorksheet[] Sheets {
			get {
				var lista = new List<IWorksheet>();
				for (int i = 0; i < Book.NumberOfSheets; i++) {
					lista.Add(new NpoiWorksheet(Book, Book.GetSheetAt(i), _evaluator));
				}
				return lista.ToArray();
			}
		}
	}
}