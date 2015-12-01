using System;

namespace Excel
{
	class MainClass
	{	

		public static void Main (string[] args)
		{
			DataStruct data = new DataStruct ();
			IOWrite write = new IOWrite (data);

			//Набиране на данни в основната таблица
			data.addRow ( "Георги", "Крумов", "71" );
			data.addRow ( "Невена", "Крумова", "68");

			//Проверка на таблицата
			data.printTable ();

			write.exportTable ();
			write.runFile ();



		}
	}
}
