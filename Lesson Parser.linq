<Query Kind="Program">
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <Namespace>DocumentFormat.OpenXml.Packaging</Namespace>
  <Namespace>DocumentFormat.OpenXml.Spreadsheet</Namespace>
</Query>

void Main()
{
	var lessonMappings = GetLessonMappings();
	var learningObjects = GetLearningObjectives();
	var lessonLinks = GetLessonLinks();

	var topicLOs = learningObjects.GroupBy(lo => lo.Topic).Select(g => new { Topic = g.Key, Ids = g.Select(x => x.Id)}).ToList();
	var tlMapping = new List<TopicLessonMapping>();
	foreach (var topicLO in topicLOs)
	{
		var map = new TopicLessonMapping() { Topic = topicLO.Topic, LearningObjectiveIds = new HashSet<string>(topicLO.Ids)};
		
		var lessonHash = new HashSet<Tuple<string,int>>();
		foreach (var id in topicLO.Ids)
		{
			var lessonsFound = lessonMappings.Where(lm => lm.LOIds.Contains(id)).Select(o => System.Tuple.Create(o.GradeNumber, o.LessonNumber));
			foreach (var item in lessonsFound)
			{
				lessonHash.Add(item);
			}
		}
		var lessons = lessonHash.Select(h => new Lesson { GradeNumber = h.Item1, LessonNumber = h.Item2}).OrderBy(x => x.GradeNumber).ThenBy(x => x.LessonNumber).ToList();
		foreach (var lesson in lessons)
		{
			lesson.Href = lessonLinks.SingleOrDefault(ll => ll.GradeNumber == lesson.GradeNumber && ll.LessonNumber == lesson.LessonNumber)?.Href;
		}
		map.LessonFiles = lessons;
		tlMapping.Add(map);
	}
	
	
	var scriptBuilder = new StringBuilder();
	foreach (var mapping in tlMapping)
	{
		scriptBuilder.AppendLine("|-");
		scriptBuilder.Append($"| '''[[{mapping.Topic}]]'''");
		var groupedFiles = mapping.LessonFiles
			.GroupBy(lf => {
				var grade = int.Parse(lf.GradeNumber);
				if (grade <= 4)
					return 0;
					else if (grade <= 8)
					return 1;
					else return 2;
			});
		scriptBuilder.Append(" || ");
		foreach (var group in groupedFiles)
		{
			foreach (var lesson in group)
			{
				scriptBuilder.Append($"{lesson.CreateLink()}");
				scriptBuilder.Append("<br />");
			}
			scriptBuilder.Append(" || ");
		}
		scriptBuilder.AppendLine();
	}
	scriptBuilder.ToString().Dump();
	
	
}



public List<LessonMapping> GetLessonMappings()
{
	var lessonMappings = new List<LessonMapping>();

	var rootFolder = @"C:\SS Curriculum\Phase 4 - Proofreading";
	var gradeFolders = new DirectoryInfo(rootFolder).EnumerateDirectories("Grade *").ToList();

	foreach (var gradeFolder in gradeFolders)
	{
		var gradeNumber = Regex.Match(gradeFolder.Name, @"Grade (?<grade>(\d){1,2})").Groups["grade"].Value;

		var lessonFiles = gradeFolder.EnumerateFiles("Lesson *.docx");
		foreach (var lessonFile in lessonFiles)
		{
			var lessonNumber = int.Parse(Regex.Match(lessonFile.Name, @"Lesson (?<lesson>(\d){1,2})").Groups["lesson"].Value);

			var filePath = lessonFile.FullName;
			string content = null;
			using (var doc = WordprocessingDocument.Open(filePath, false))
			{
				var mainPart = doc.MainDocumentPart;
				content = mainPart.Document.Body.InnerText;
			}

			var pattern = $@"G{gradeNumber}.LO(\d){{1,2}}";
			var loIdentifiers = Regex.Matches(content, pattern).Select(o => o.Value).Distinct();
			var mapping = new LessonMapping { GradeNumber = gradeNumber, LessonNumber = lessonNumber, LOIds = new HashSet<string>(loIdentifiers) };
			lessonMappings.Add(mapping);
		}
	}
	
	return lessonMappings.OrderBy(x => x.GradeNumber).ThenBy(x => x.LessonNumber).ToList();
}

public List<LessonLink> GetLessonLinks()
{
	var path = @"C:\SS Curriculum\Links to Lesson.xlsx";
	var links = new List<LessonLink>();
	using (var spreadSheet = SpreadsheetDocument.Open(path, false))
	{
		var wbPart = spreadSheet.WorkbookPart;
		var sheet = (Sheet)wbPart.Workbook.Sheets.First();

		var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
		var rows = wsPart.Worksheet.Descendants<Row>().ToList();
		var header = rows.First();
		var i = 0;
		foreach (var row in rows.Skip(1))
		{
			var cells = row.Elements<Cell>().ToList();
			if (cells.Count == 0)
			{
				break;
			}
			try
			{
				var link = new LessonLink { GradeNumber = GetString(wbPart, cells[0]), LessonNumber = GetValue(cells[1], i), Href = GetString(wbPart, cells[3]) };
				links.Add(link);

			}
			catch (Exception ex)
			{

			}
			i++;
		}
	}

	return links;
}

public List<LearningObjective> GetLearningObjectives()
{
	var path = @"C:\SS Curriculum\LO Mappings All.xlsx";
	var LOs = new List<LearningObjective>();

	using (var spreadSheet = SpreadsheetDocument.Open(path, false))
	{
		var wbPart = spreadSheet.WorkbookPart;
		var sheet = (Sheet)wbPart.Workbook.Sheets.First();

		var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
		var rows = wsPart.Worksheet.Descendants<Row>().ToList();
		var header = rows.First();

		foreach (var row in rows.Skip(1))
		{
			var cells = row.Elements<Cell>().ToList();

			var lo = new LearningObjective { Id = GetString(wbPart, cells[0]), Topic = GetString(wbPart, cells[1]), Description = GetString(wbPart, cells[2]) };
			LOs.Add(lo);
		}
	}

	return LOs;
}

public class Lesson
{
	public string GradeNumber { get; set; }
	public int LessonNumber { get; set; }
	public string Href { get; set; }
}

public class LessonMapping
{
	public string GradeNumber { get; set; }
	public int LessonNumber { get; set; }
	public HashSet<string> LOIds { get; set; }
}


private string GetString(WorkbookPart wbPart, Cell c)
{
	if (c.DataType != null && c.DataType == CellValues.SharedString)
	{
		var stringId = Convert.ToInt32(c.InnerText);
		return wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
	}
	else
	{
		return c.InnerText;
	}

}

public class LearningObjective
{
	public string Id { get; set; }
	public string Topic { get; set; }
	public string Description { get; set; }
}

public class TopicLessonMapping
{
	public string Topic { get; set; }
	public HashSet<string> LearningObjectiveIds { get; set; }
	public List<Lesson> LessonFiles { get; set; }
}

public class LessonLink
{
	public string Href { get; set; }
	public string GradeNumber { get; set; }
	public int LessonNumber { get; set; }
}

private int GetValue(Cell cell, int i)
{
	var rawValue = cell.InnerText;
	try
	{
		return (int)double.Parse(cell.InnerText);

	}
	catch
	{
		return 0;
	}
}

public static class LessonLinkExtensions
{
	public static string CreateLink(this Lesson lesson)
	{
		return $"[{lesson.Href} Grade {lesson.GradeNumber } Lesson {lesson.LessonNumber}]";
	}
}