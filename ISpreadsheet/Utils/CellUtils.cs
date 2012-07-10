using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ISpreadsheet.Utils {
	public class CellUtils {
		public static readonly Regex RegexAddress = new Regex(@"([a-zA-Z]+)(\d+)?", RegexOptions.Compiled);

		public static CellAddress AddressFromText(string texto) {
			if (string.IsNullOrEmpty(texto))
				return null;
			texto = texto.Trim();
			Match match = RegexAddress.Match(texto);
			if (!match.Success)
				return null;
			var dir = new CellAddress {
				Column = LetterToColumnNumber(match.Groups[1].Value),
				Row = Convert.ToInt32(match.Groups[2].Value),
				Text = texto.ToUpper()
			};
			return dir;
		}

		/// <summary>
		/// Dado un rango de lineas (inicio, fin) retorna una lista de cadenas de cada linea utilizando un separador
		/// y con una pista de parada si se encuentra la cadena
		/// </summary>
		/// <param name="sheet">Hoja de cálculo</param>
		/// <param name="start">Linea de inicio</param>
		/// <param name="end">Línea fin</param>
		/// <param name="separator">Cadena para separar entre celdas</param>
		/// <param name="stophint">Texto de parada si se encuentra en la primera celda</param>
		/// <returns></returns>
		public static List<string> TextLinesFromRange(IWorksheet sheet, int start, int end, string separator = "", string stophint = "") {
			end = Math.Min(sheet.NumRows, end);
			var list = new List<string>();
			for (var i = start; i <= end; i++) {
				var cells = sheet.GetRow(i).Where(x => x != null).Select(x => x.ToString()).ToArray();
				if (!string.IsNullOrEmpty(stophint)) {
					var first = cells.FirstOrDefault() ?? "";
					if (first.StartsWith(stophint))
						break;
				}
				var line = string.Join(separator, cells).Trim();
				list.Add(line);
			}
			return list;
		}

		public static string ColumnNumberToLetter(int value) {
			if (value < 0)
				return "";
			string alpha = "";
			while (value > 0) {
				int ascii = 64 + value - 26 * (value / 26);
				if (ascii == 64)
					ascii = 90;
				alpha = (char)ascii + alpha;
				value = value / 26;
				if (ascii == 90)
					value = value - 1;
			}
			return alpha;
		}

		public static int LetterToColumnNumber(string value) {
			if (string.IsNullOrEmpty(value))
				return -1;
			value = value.ToUpperInvariant();
			int num = 0;
			int tot = value.Length - 1;
			for (int i = 0; i <= tot; i++) {
				int ch = value[i] - 64;
				num += ch * Convert.ToInt32(Math.Pow(26, tot - i));
			}
			return num;
		}

		public static DateTime ExcelDate(int numero) {
			return ExcelDate(Convert.ToDouble(numero));
		}

		public static DateTime ExcelDate(float numero) {
			return ExcelDate(Convert.ToDouble(numero));
		}

		public static DateTime ExcelDate(decimal numero) {
			return ExcelDate(Convert.ToDouble(numero));
		}

		public static DateTime ExcelDate(double numero) {
			int datep = (int)numero;
			double timep = numero - datep;

			DateTime date = new DateTime(1900, 1, 1);
			datep = datep - 2;
			date = date.AddDays(datep);
			if (timep == 0) return date;
			timep = timep * 86400;
			date = date.AddSeconds(timep);
			return date;
		}

	}
}