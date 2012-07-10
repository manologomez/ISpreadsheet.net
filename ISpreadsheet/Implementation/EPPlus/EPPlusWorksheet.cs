using System.Collections.Generic;
using ISpreadsheet.Utils;
using OfficeOpenXml;

namespace ISpreadsheet.Implementation.EPPlus {
	class EPPlusWorksheet : IWorksheet {
		public string Name { get; set; }
		public ExcelWorksheet Sheet { get; set; }

		public EPPlusWorksheet(ExcelWorksheet sheet) {
			Sheet = sheet;
			Name = sheet.Name;
		}

		public object GetCell(string address) {
			var range = Sheet.Cells[address];
			if (range.IsRichText)
				return range.RichText.Text;
			var value = range.Value;
			return value;
		}

		public object GetCell(int col, int row) {
			var range = Sheet.Cells[row, col];
			if (range.IsRichText)
				return range.RichText.Text;
			var value = range.Value;
			return value;
		}

		public object GetCell(string col, int row) {
			var add = string.Format("{0}{1}", col, row).ToUpper();
			return GetCell(add);
		}

		public object[] GetRow(int num) {
			var end = Sheet.Dimension.End.Column;
			var start = Sheet.Dimension.Start.Column;
			var list = new List<object>();
			for (var i = start; i <= end; i++) {
				list.Add(GetCell(i, num));
			}
			return list.ToArray();
		}

		public string GetString(string address) {
			var obj = GetCell(address);
			return obj == null ? null : obj.ToString();
		}

		public string GetString(dynamic col, int row) {
			object obj;
			if (col is int)
				obj = GetCell((int)col, row);
			else
				obj = GetCell(col.ToString(), row);
			return obj == null ? null : obj.ToString();
		}

		public IDictionary<string, object> GetRowMap(int num) {
			var end = Sheet.Dimension.End.Column;
			var start = Sheet.Dimension.Start.Column;
			var map = new Dictionary<string, object>();
			for (var i = start; i <= end; i++) {
				var letra = CellUtils.ColumnNumberToLetter(i);
				map[letra] = GetCell(i, num);
			}
			return map;
		}

		public int NumColumns {
			get {
				return Sheet.Dimension == null ? 0 : Sheet.Dimension.End.Column;
			}
		}

		public int NumRows {
			get {
				return Sheet.Dimension == null ? 0 : Sheet.Dimension.End.Row;
			}
		}

		public IWorksheet SetValue(string address, object value) {
			Sheet.SetValue(address, value);
			return this;
		}

		public IWorksheet SetValue(int col, int row, object value) {
			Sheet.SetValue(row, col, value);
			return this;
		}
	}
}