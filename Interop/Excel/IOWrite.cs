using System;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace Excel
{
	public class IOWrite
	{
		private DataStruct _data;

		private InteropExcel.Application excel;

		public IOWrite (DataStruct data)
		{
			_data = data;
		}

		public bool exportTable ()
		{
			try
			{ 
				//Подготовка
				excel = new InteropExcel.ApplicationClass ();
				if (excel == null) return false;

				excel.Visible = false;

				InteropExcel.Workbook workbook = excel.Workbooks.Add ();
				if (workbook == null) return false;

				InteropExcel.Worksheet sheet = (InteropExcel.Worksheet) workbook.Worksheets [1];
				sheet.Name = "Таблица 1";  


				//Попълване на таблицата

				int i = 1;

				addRow ( new DataRow ("Първо име", "Фамилия", "Години"), i++, true, 40); i++;

				foreach (DataRow row in _data.table)
				{
					addRow ( row, i++, false, -1);
				}

				i++; addRow (new DataRow  ("Брой редове", "", _data.table.Count.ToString () ), i++, true, -1);

				//Запаметяване и затваряне
				workbook.SaveCopyAs ( getPath () );

				excel.DisplayAlerts = false; //Изключване всички съобщения на Excel

				workbook.Close ();
				excel.Quit ();

				return true;
			}catch{
			}
			return false;
		}
		public void addRow ( DataRow _dataRow, int _indexRow, bool isBold, int color )
		{
			try
			{
				InteropExcel.Range range;

				//Форматираме
				range = excel.Range ["A" + _indexRow.ToString(), "C" + _indexRow.ToString()];

				if ( color > 0 ) range.Interior.ColorIndex = color; //-1
				if ( isBold ) range.Font.Bold = isBold;

				//Въвеждаме данните клетка по клетка
				range = excel.Range ["A" + _indexRow.ToString(), "A" + _indexRow.ToString()];
				range.Value2 = _dataRow.firstName;

				range = excel.Range ["B" + _indexRow.ToString(), "B" + _indexRow.ToString()];
				range.Value2 = _dataRow.lastName; 

				range = excel.Range ["C" + _indexRow.ToString(), "C" + _indexRow.ToString()];
				range.Value2 = _dataRow.age; 

			}catch{
			}
		}
		public void runFile ()
		{
			try
			{
				System.Diagnostics.Process.Start ( getPath ());

			}catch{
			}
		}
		public string getPath ()
		{
			return System.IO.Path.Combine ( AppDomain.CurrentDomain.BaseDirectory, "Table1.xlsx");
		}

	}
}

