namespace MbrohDrawingIndexCreator;

public sealed class DrawingIndexItem
{
	public int SeqNo { get; set; }

	public string DrawingNumber { get; set; }

	public string Discipline { get; set; }

	public string Title { get; set; }

	public string SourcePdfPath { get; set; }

	public int SourcePageIndex { get; set; }

	// Set when OCR was not confident about a field; the index highlights those cells so the user can check them.
	public bool DrawingNumberUncertain { get; set; }

	public bool DisciplineUncertain { get; set; }

	public bool TitleUncertain { get; set; }

	// Why the entry was flagged, shown in the review dialog.
	public string Note { get; set; } = string.Empty;

	public double PageWidth { get; set; }

	public double PageHeight { get; set; }

	public DrawingIndexItem()
	{
		DrawingNumber = string.Empty;
		Discipline = string.Empty;
		Title = string.Empty;
		SourcePdfPath = string.Empty;
	}
}
