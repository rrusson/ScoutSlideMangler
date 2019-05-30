using System;
using System.Data;

namespace ScoutSlideMangler
{
	public class MeritBadgeSlideMaker
	{
		private readonly DataTable _data = ExcelReader.GetData(@"C:\Temp\COH.xlsx", "Merit Badge");

		private readonly Zipper _zipper = new Zipper(@"C:\Temp\CourtOfHonor.pptx");


		public MeritBadgeSlideMaker()
		{
			//UNZIP .PPTX to receive new slides
			_zipper.UnzipFile();
		}

		public void ProcessMeritBadges()
		{
			var maker = new MeritBadgeSlideCreator(23);

			string nextName = GetName(1);
			int i = 1;

			for (; nextName != null;)   // Exit when we're all out of rows
			{
				var scout = new TroopPerson
				{
					FirstName = _data.Rows[i][1].ToString().Trim(),
					LastName = _data.Rows[i][0].ToString().Trim()
				};

				//First row is headers
				string thisScoutName = GetName(i);

				while (nextName == thisScoutName)
				{
					scout.MeritBadges.Add(_data.Rows[i][3].ToString());
					i++;
					nextName = GetName(i);
				}

				Console.WriteLine(scout.LastName + ": " + scout.MeritBadges.Count);
				maker.SaveSlide(scout);
			}

			//Re-Zip .PPTX with new slides
			_zipper.ReZipFile();
			Console.WriteLine("-------------ALL SLIDES PROCESSED--------------");
			Console.ReadKey();
		}

		private string GetName(int rowNumber)
		{
			if (rowNumber >= _data.Rows.Count)
			{
				return null;
			}

			DataRow data = _data.Rows[rowNumber];
			return data[1].ToString() + Environment.NewLine + data[0].ToString();
		}
	}
}

