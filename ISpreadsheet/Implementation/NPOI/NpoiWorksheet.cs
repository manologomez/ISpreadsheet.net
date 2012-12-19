using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ISpreadsheet.Utils;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ISpreadsheet.Implementation.NPOI {
	public class NpoiWorksheet : IWorksheet {

		public HSSFWorkbook Book { get; set; }
		public ISheet Sheet { get; set; }

		public int Index { get; private set; }
		public string Name { get; set; }
		private int _maxcol = -1;
		private IFormulaEvaluator _evaluator;

		public NpoiWorksheet(HSSFWorkbook book, ISheet sheet) {
			sheet.ForceFormulaRecalculation = true;
			Book = book;
			Sheet = sheet;
			Index = book.GetSheetIndex(sheet);
			Name = book.GetSheetName(Index);
		}

		public object GetCell(string address) {
			//var dir = CeldasUtil.DireccionFromTexto(address);
			//return GetCell(dir.Column, dir.Row);
			var reference = new CellReference(address);
			return GetCell(reference.Col + 1, reference.Row + 1);
		}

		public object GetCell(string col, int row) {
			var add = string.Format("{0}{1}", col, row).ToUpper();
			return GetCell(add);
		}

		public object GetCell(int col, int row) {
			var rowobj = Sheet.GetRow(row - 1);
			if (rowobj == null)
				return null;
			var cell = rowobj.GetCell(col - 1);
			if (cell == null)
				return null;
			return GetValor(cell);
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

		public object[] GetRow(int num) {
			var list = new List<object>();
			var row = Sheet.GetRow(num - 1);
			if (row != null) {
				for (int i = 0; i <= row.LastCellNum; i++) {
					list.Add(GetValor(row.GetCell(i)));
				}
			}
			return list.ToArray();
		}

		public IDictionary<string, object> GetRowMap(int num) {
			var map = new Dictionary<string, object>();
			var row = Sheet.GetRow(num - 1);
			if (row != null) {
				for (var i = 0; i <= row.LastCellNum; i++) {
					var letra = CellUtils.ColumnNumberToLetter(i + 1);
					map[letra] = GetValor(row.GetCell(i));
				}
			}
			return map;
		}

		public int NumColumns {
			get { return CalculateMaxColumn(); }
		}

		public int NumRows {
			get { return Sheet.LastRowNum; }
		}

		/// <summary>
		/// returns the maximun number of cells in a row
		/// </summary>
		/// <returns></returns>
		public int CalculateMaxColumn() {
			if (_maxcol > -1)
				return _maxcol;
			_maxcol = 0;
			for (int i = Sheet.FirstRowNum; i <= Sheet.LastRowNum; i++) {
				var row = Sheet.GetRow(i);
				if (row == null)
					continue;
				if (row.LastCellNum > _maxcol)
					_maxcol = row.LastCellNum;
			}
			return _maxcol;
		}

		private readonly Regex _timeRegex = new Regex(@"h+:mm(:ss)?", RegexOptions.Compiled | RegexOptions.IgnoreCase);

		public object GetValor(ICell celda) {
			if (celda == null)
				return null;

			switch (celda.CellType) {
				case CellType.STRING:
					//return celda.RichStringCellValue;
					return celda.StringCellValue;
				case CellType.NUMERIC:
					if (DateUtil.IsCellDateFormatted(celda))
						return celda.DateCellValue;
					if (celda.CellStyle != null) {
						var format = celda.CellStyle.GetDataFormatString();
						//if (Regex.IsMatch(format, @"h+:mm(:ss)?"))
						if (_timeRegex.IsMatch(format))
							return celda.DateCellValue;
					}
					return celda.NumericCellValue;
				/*var iformat = celda.CellStyle.DataFormat;
				if(DateUtil.IsADateFormat(iformat, format)){
					Console.Write("fecha");
				}
				if (DateUtil.IsInternalDateFormat(iformat))
					return celda.DateCellValue;*/

				/*if (DateUtil.IsCellDateFormatted(celda) || DateUtil.IsCellInternalDateFormatted(celda)) {
					if (DateUtil.IsValidExcelDate(celda.NumericCellValue)) {
						var dt = celda.DateCellValue;
						return dt;
					}
				}
				return celda.NumericCellValue;*/
				case CellType.BOOLEAN:
					return celda.BooleanCellValue;
				case CellType.FORMULA:
					try {
						if (_evaluator == null)
							_evaluator = Book.GetCreationHelper().CreateFormulaEvaluator();
						var valorCelda = _evaluator.Evaluate(celda);
						return GetValorCelda(valorCelda, celda);
					} catch (Exception ex) {
						return "Error en formula " + celda.CellFormula;
					}
				default:

					break;
			}
			return null;
		}

		protected object GetValorCelda(CellValue valorCelda, ICell container) {
			if (valorCelda == null)
				return null;
			switch (valorCelda.CellType) {
				case CellType.STRING:
					//return celda.StringValorCelda;
					return container.StringCellValue;
				case CellType.NUMERIC:
					var numValue = container.NumericCellValue;
					if (DateUtil.IsCellDateFormatted(container)) {
						if (DateUtil.IsValidExcelDate(numValue)) {
							var dt = DateUtil.GetJavaDate(numValue);
							return dt;
						}
					}
					return numValue;
				case CellType.BOOLEAN:
					return container.BooleanCellValue;
				default:
					return null;
			}
		}

		public IWorksheet SetValue(int col, int row, object value) {
			if (value == null) return this;
			var irow = Sheet.GetRow(row - 1) ?? Sheet.CreateRow(row - 1);
			ICell celda = irow.GetCell(col - 1) ?? irow.CreateCell(col - 1);
			if (value is string)
				celda.SetCellValue(value.ToString());
			else if (value is int || value is decimal || value is float || value is double) {
				var dvalue = Convert.ToDouble(value);
				celda.SetCellValue(dvalue);
			} else if (value is DateTime)
				celda.SetCellValue((DateTime)value);
			else if (value is bool)
				celda.SetCellValue((bool)value);
			else
				celda.SetCellValue(value.ToString());
			return this;
		}

		public IWorksheet SetValue(string address, object value) {
			var reference = new CellReference(address);
			return SetValue(reference.Col + 1, reference.Row + 1, value);
		}
	}

}