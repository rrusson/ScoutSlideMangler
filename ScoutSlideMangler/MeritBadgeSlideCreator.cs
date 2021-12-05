using System.Collections.Generic;
using System.Text;

namespace ScoutSlideMangler
{
	//TODO: Create separate template.PPTX with formatted string placeholders for data (e.g. $"{BadgeName}") instead of hardcoded template here
	//TODO: Insert merit badge clip art based off badge name

	/// <summary>
	/// Inserts data from a TroopPerson object into an XML template and saves a data-filled XML document
	/// </summary>
	public class MeritBadgeSlideCreator
	{
		//NOTE: For some reason if the long "template" string is global, it DOES NOT WORK correctly when formatted
		private const string _postTemplate = "</p:txBody></p:sp></p:spTree><p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\"><p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"2365574124\"/></p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>";
		private int _slideNumber;

		/// <summary>
		/// CTOR: initializes class with the first slide number to save out (e.g. "slide42.xml")
		/// </summary>
		/// <param name="firstSlideNumber">The first slide number to insert/overwrite in the .PPTX</param>
		public MeritBadgeSlideCreator(int firstSlideNumber)
		{
			_slideNumber = firstSlideNumber;
		}

		/// <summary>
		/// Inserts data from <paramref name="scout"/> into a big, ugly XML template and saves it to a slide
		/// </summary>
		public void SaveSlide(TroopPerson scout)
		{
			string xmlDelcaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
			string preTemplate = $"<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"1111\" name=\"{scout.FirstName}{scout.LastName}\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"title\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"33866\" y=\"225954\"/><a:ext cx=\"4419601\" cy=\"1198469\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit fontScale=\"90000\"/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.FirstName}</a:t></a:r></a:p><a:p><a:pPr><a:defRPr sz=\"4000\" b=\"1\"><a:latin typeface=\"Copperplate Gothic Bold\"/><a:ea typeface=\"Copperplate Gothic Bold\"/><a:cs typeface=\"Copperplate Gothic Bold\"/><a:sym typeface=\"Copperplate Gothic Bold\"/></a:defRPr></a:pPr><a:r><a:t>{scout.LastName}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"1112\" name=\"Chess…\"/><p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"half\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"220133\" y=\"1828800\"/><a:ext cx=\"4204855\" cy=\"4724400\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>";

			string badgeData = GetBadgeMarkup(scout.MeritBadges);

			string slide = xmlDelcaration + string.Format(preTemplate, scout.FirstName, scout.LastName) + badgeData + _postTemplate;

			System.IO.File.WriteAllText($@"C:\Temp\CoH\ppt\slides\slide{_slideNumber}.xml", slide);
			_slideNumber++;
		}


		private string GetBadgeMarkup(List<string> badges)
		{
			var chunk = new StringBuilder();

			foreach (var badgeName in badges)
			{
				chunk.Append($"<a:p><a:pPr marL=\"342900\" indent=\"-342900\" algn=\"l\"><a:buFont typeface=\"Arial\" panose=\"020B0604020202020204\" pitchFamily=\"34\" charset=\"0\"/><a:buChar char=\"•\"/></a:pPr><a:r><a:rPr lang=\"en-US\" sz=\"2000\" dirty=\"0\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:effectLst/><a:latin typeface=\"Segoe UI Semibold\" panose=\"020B0702040204020203\" pitchFamily=\"34\" charset=\"0\"/><a:cs typeface=\"Segoe UI Semibold\" panose=\"020B0702040204020203\" pitchFamily=\"34\" charset=\"0\"/></a:rPr><a:t>{badgeName}</a:t></a:r></a:p>");
			}

			return chunk.ToString();
		}
	}
}
