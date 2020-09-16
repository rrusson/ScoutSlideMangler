using System.IO.Compression;
using System.Threading;

namespace ScoutSlideMangler
{
	public class Zipper
	{
		private const string outputDirectory = @"C:\Temp\CoH\";

		public string ZipFilePath { get; set; }


		public Zipper(string zipFilePath)
		{
			ZipFilePath = zipFilePath;
		}


		public void UnzipFile()
		{
			if (System.IO.Directory.Exists(outputDirectory))
			{
				try
				{
					//Fix for general flakiness of this method courtesy of:
					//http://zacharykniebel.com/blog/web-development/2013/june/21/solving-the-csharp-bug-when-recursively-deleting-directories
					System.IO.Directory.Delete(outputDirectory, true);
				}
				catch (System.IO.IOException)
				{
					Thread.Sleep(10);
					System.IO.Directory.Delete(outputDirectory, true);
				}
			}

			ZipFile.ExtractToDirectory(ZipFilePath, outputDirectory);
		}

		/// <summary>
		/// ZIPs working directory back into a final file (directory not deleted for diagnostic purposes)
		/// </summary>
		public void ReZipFile()
		{
			string pptxFile = @"C:\Temp\CourtOfHonor-DONE.pptx";

			if (System.IO.File.Exists(pptxFile))
			{
				try
				{
					System.IO.File.Delete(pptxFile);
				}
				catch (System.IO.IOException)
				{
					Thread.Sleep(0);
					System.IO.File.Delete(pptxFile);
				}
			}

			ZipFile.CreateFromDirectory(outputDirectory, pptxFile);
		}
	}
}
