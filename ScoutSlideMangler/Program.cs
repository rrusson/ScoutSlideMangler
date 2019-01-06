using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ScoutSlideMangler
{
	class Program
	{
		static void Main(string[] args)
		{
			ProcessMeritBadges();
		}


		static void ProcessMeritBadges()
		{
			DataTable data = ExcelReader.GetData(@"C:\Temp\COH.xlsx", "Merit Badge");
			var maker = new MeritBadgeSlideCreator();

			//int totalRows = data.Rows.Count;
			string nextName = GetName(data, 1);
			int i = 1;

			for (; nextName != null;)   // Exit when we're all out of rows
			{
				TroopPerson scout = GetScoutFromData(data, ref nextName, ref i);
				maker.SaveSlide(scout);
			}

			Console.WriteLine("Slides Generated");
			Console.ReadKey();
		}

		private static TroopPerson GetScoutFromData(DataTable data, ref string nextName, ref int i)
		{
			//EXPECTED FORMAT:
			//LastName		FirstName	Type	BadgeName
			var scout = new TroopPerson
			{
				FirstName = data.Rows[i][1].ToString().Trim(),
				LastName = data.Rows[i][0].ToString().Trim()
			};

			//First row is headers
			string thisScoutName = GetName(data, i);

			while (nextName == thisScoutName)
			{
				scout.MeritBadges.Add(data.Rows[i][3].ToString());
				i++;
				nextName = GetName(data, i);
			}

			Console.WriteLine(scout.LastName + ": " + scout.MeritBadges.Count);
			return scout;
		}

		static string GetName(DataTable table, int rowNumber)
		{
			if (rowNumber >= table.Rows.Count)
			{
				return null;
			}

			DataRow data = table.Rows[rowNumber];
			return data[1].ToString() + Environment.NewLine + data[0].ToString();
		}
	}
}
