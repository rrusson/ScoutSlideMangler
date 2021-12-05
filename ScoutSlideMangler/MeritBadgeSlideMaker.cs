using System;
using System.Data;
using System.Threading;

namespace ScoutSlideMangler
{
	public class MeritBadgeSlideMaker
	{
		private const int _firstSlideToInsert = 3;
		private readonly DataTable _data = ExcelReader.GetData(@"..\..\..\COH.xlsx", "Merit Badges");
		private readonly Zipper _zipper = new Zipper(@"..\..\..\CourtOfHonor.pptx");
		private readonly MeritBadgeSlideCreator _slideMaker = new MeritBadgeSlideCreator(_firstSlideToInsert);
		private int _counter = 1;           //Skip row 0, which contains headers

		public void ProcessMeritBadges()
		{
			bool hasData = true;

			//Unzip .PPTX to receive new slides
			try
			{
				_zipper.UnzipFile();
			}
			catch (System.IO.IOException)
			{
				Thread.Sleep(10);
				_zipper.UnzipFile();
			}

			while (hasData)   // Exit when we're all out of rows
			{
				hasData = CreateSlide();
			}

			//Re-Zip .PPTX with new slides
			_zipper.ReZipFile();

			Console.WriteLine("-------------ALL SLIDES PROCESSED--------------");
			Console.ReadKey();
		}

		private bool CreateSlide()
		{
			string thisScoutName = GetName(_counter);
			string nextName = thisScoutName;

			var scout = new TroopPerson
			{
				FirstName = _data.Rows[_counter][1].ToString().Trim(),
				LastName = _data.Rows[_counter][0].ToString().Trim()
			};

			while (nextName == thisScoutName)
			{
				scout.MeritBadges.Add(_data.Rows[_counter][3].ToString());
				_counter++;
				nextName = GetName(_counter);
			}

			Console.WriteLine(scout.LastName + ": " + scout.MeritBadges.Count);
			_slideMaker.SaveSlide(scout);

			return nextName != null;
		}

		private string GetName(int rowNumber)
		{
			if (_data?.Rows == null)
			{
				throw new FieldAccessException("Dude. There's no data. Check Excel tab name and format.");
			}

			if (rowNumber >= _data.Rows.Count)
			{
				return null;
			}

			DataRow data = _data.Rows[rowNumber];

			return $"{data[1]} {data[0]}";
		}
	}
}

