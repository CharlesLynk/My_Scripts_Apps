using System.Text;
using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

public class RedactionTests
{
	private const string TwoRectangles = "1 0 0 rg\n10 10 10 10 re f\n50 50 10 10 re f\n";

	[Fact]
	public void RemovesPathFullyInsideRedactionRectangle_AndKeepsPathOutsideIt()
	{
		byte[] input = Encoding.ASCII.GetBytes(TwoRectangles);

		byte[]? output = PdfContentRedactor.Process(input, 0, 0, 30, 30, out int removed);

		Assert.Equal(1, removed);
		Assert.NotNull(output);
		string text = Encoding.ASCII.GetString(output!);
		Assert.DoesNotContain("10 10 10 10 re", text);
		Assert.Contains("50 50 10 10 re", text);
	}

	[Fact]
	public void CornerOrderOfRedactionRectangleDoesNotMatter()
	{
		byte[] input = Encoding.ASCII.GetBytes(TwoRectangles);

		PdfContentRedactor.Process(input, 30, 30, 0, 0, out int removed);

		Assert.Equal(1, removed);
	}

	[Fact]
	public void ReportsNothingRemovedWhenNoPathIsInside()
	{
		byte[] input = Encoding.ASCII.GetBytes(TwoRectangles);

		PdfContentRedactor.Process(input, 200, 200, 300, 300, out int removed);

		Assert.Equal(0, removed);
	}
}
