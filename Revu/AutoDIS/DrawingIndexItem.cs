namespace MbrohDrawingIndexCreator;

public sealed class DrawingIndexItem
{
	public int SeqNo { get; set; }

	public string DrawingNumber { get; set; }

	public string Discipline { get; set; }

	public string Title { get; set; }

	public string SourcePdfPath { get; set; }

	public int SourcePageIndex { get; set; }

	public DrawingIndexItem()
	{
		DrawingNumber = string.Empty;
		Discipline = string.Empty;
		Title = string.Empty;
		SourcePdfPath = string.Empty;
	}
}
