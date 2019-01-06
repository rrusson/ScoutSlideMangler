using System;

namespace ScoutSlideMangler
{
	class Program
	{
		static void Main(string[] args)
		{
			var worker = new SlideMangler();
			worker.CreateSlides(@"C:\Temp\COH.xlsx", "Merit Badge");

			Console.WriteLine(Environment.NewLine + "***   Slides Generated   ***");
			Console.ReadKey();
		}
	}
}
