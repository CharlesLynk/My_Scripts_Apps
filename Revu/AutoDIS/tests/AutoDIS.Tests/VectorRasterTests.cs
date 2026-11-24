using System;
using System.Collections.Generic;
using System.Linq;
using Tesseract;
using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

public sealed class VectorRasterTests
{
	private static VectorShape Line(double x1, double y1, double x2, double y2, byte gray = 0)
	{
		VectorShape shape = new VectorShape { Stroked = true, StrokeGray = gray };
		shape.Add(new[] { (x1, y1), (x2, y2) }, false);
		return shape;
	}

	private static VectorShape Rectangle(double x1, double y1, double x2, double y2, byte fillGray)
	{
		VectorShape shape = new VectorShape { Filled = true, FillGray = fillGray };
		shape.Add(new[] { (x1, y1), (x2, y1), (x2, y2), (x1, y2) }, true);
		return shape;
	}

	private static int Dark(GrayImage image, int x, int y) => image.Pixels[y * image.Width + x] < 128 ? 1 : 0;

	[Fact]
	public void AStrokeIsDrawnWithTheRequestedPenWidth()
	{
		// 72 dpi: one point is one pixel. A 4 pt pen across the middle of a 100 x 40 pt region.
		GrayImage image = VectorRaster.Render(new[] { Line(0, 20, 100, 20) }, 0, 0, 100, 40, 72.0, 4.0);

		Assert.Equal(100, image.Width);
		Assert.Equal(40, image.Height);
		int thickness = Enumerable.Range(0, 40).Sum(y => Dark(image, 50, y));
		Assert.InRange(thickness, 3, 5);
		Assert.Equal(0, Dark(image, 50, 5));
		Assert.Equal(0, Dark(image, 50, 35));
	}

	[Fact]
	public void PageCoordinatesRunUpwardButImagesRunDownward()
	{
		// A stroke near the top of the region (high y in the page) lands in the top rows of the image.
		GrayImage image = VectorRaster.Render(new[] { Line(0, 36, 100, 36) }, 0, 0, 100, 40, 72.0, 2.0);

		Assert.Equal(1, Dark(image, 50, 3));
		Assert.Equal(0, Dark(image, 50, 36));
	}

	[Fact]
	public void ShapesOutsideTheRegionAreIgnored()
	{
		GrayImage image = VectorRaster.Render(new[] { Line(0, 500, 100, 500), Line(900, 0, 900, 40) }, 0, 0, 100, 40, 72.0, 2.0);

		Assert.Equal(0.0, VectorRaster.InkFraction(image));
	}

	[Fact]
	public void AWhiteFillDrawnLaterHidesWhatIsUnderIt()
	{
		GrayImage image = VectorRaster.Render(new[] { Line(0, 20, 100, 20), Rectangle(40, 10, 60, 30, 255) }, 0, 0, 100, 40, 72.0, 4.0);

		Assert.Equal(1, Dark(image, 20, 20));
		Assert.Equal(0, Dark(image, 50, 20));
		Assert.Equal(1, Dark(image, 80, 20));
	}

	[Fact]
	public void ABlackFillIsInkAndHolesStayOpen()
	{
		VectorShape frame = new VectorShape { Filled = true, FillGray = 0 };
		frame.Add(new[] { (10.0, 10.0), (90.0, 10.0), (90.0, 30.0), (10.0, 30.0) }, true);
		frame.Add(new[] { (30.0, 15.0), (70.0, 15.0), (70.0, 25.0), (30.0, 25.0) }, true);

		GrayImage image = VectorRaster.Render(new[] { frame }, 0, 0, 100, 40, 72.0, 1.0);

		Assert.Equal(1, Dark(image, 15, 20));
		Assert.Equal(0, Dark(image, 50, 20));
	}

	[Fact]
	public void ALongDiagonalAcrossALargeRegionIsFast()
	{
		System.Diagnostics.Stopwatch timer = System.Diagnostics.Stopwatch.StartNew();
		GrayImage image = VectorRaster.Render(new[] { Line(0, 0, 400, 60) }, 0, 0, 400, 60, 300.0, 0.7);
		timer.Stop();

		Assert.True(timer.ElapsedMilliseconds < 1000, $"took {timer.ElapsedMilliseconds} ms");
		Assert.InRange(VectorRaster.InkFraction(image), 0.001, 0.05);
	}

	[Fact]
	public void PenWidthFollowsTheLetterHeight()
	{
		// Eight-point letters get a pen of about 0.72 pt; tiny or huge shapes are clamped.
		List<VectorShape> letters = Enumerable.Range(0, 9).Select(i => Line(i * 10, 0, i * 10 + 4, 8)).ToList();

		Assert.InRange(VectorRaster.ChoosePenPoints(letters, 0, 0, 100, 20), 0.6, 0.8);
		Assert.Equal(VectorRaster.DefaultPenPoints, VectorRaster.ChoosePenPoints(new List<VectorShape>(), 0, 0, 100, 20));
		List<VectorShape> tiny = Enumerable.Range(0, 9).Select(i => Line(i * 10, 0, i * 10 + 1, 1.2)).ToList();
		Assert.Equal(VectorRaster.MinimumPenPoints, VectorRaster.ChoosePenPoints(tiny, 0, 0, 100, 20));
	}

	[Fact]
	public void BlockLettersDrawnFromStrokesAreReadByTheOcrEngine()
	{
		// "HELLO 2" from straight strokes, drawn at the size of title-block lettering (about 8 pt tall).
		List<VectorShape> shapes = new List<VectorShape>();
		double x = 4;
		void Stroke(params (double X, double Y)[] points)
		{
			VectorShape shape = new VectorShape { Stroked = true, StrokeGray = 0 };
			shape.Add(points, false);
			shapes.Add(shape);
		}
		void H() { Stroke((x, 2), (x, 10)); Stroke((x + 5, 2), (x + 5, 10)); Stroke((x, 6), (x + 5, 6)); x += 8; }
		void E() { Stroke((x + 5, 10), (x, 10), (x, 2), (x + 5, 2)); Stroke((x, 6), (x + 4, 6)); x += 8; }
		void L() { Stroke((x, 10), (x, 2), (x + 5, 2)); x += 8; }
		void O() { Stroke((x, 2), (x + 5, 2), (x + 5, 10), (x, 10), (x, 2)); x += 8; }
		H(); E(); L(); L(); O();

		GrayImage image = VectorRaster.Render(shapes, 0, 0, 52, 12, 300.0, 0.8);
		OcrEngineResult result = OcrSupport.TryCreateEngine(OcrSupport.DefaultTessdataDirectory, OcrSupport.FastModel);
		Assert.NotNull(result.Engine);
		using TesseractEngine engine = result.Engine!;
		OcrFieldResult read = OcrFieldRecognizer.Read(engine, image, OcrFieldKind.Title, new OcrReadOptions(new[] { OcrVariant.Plain, OcrVariant.Thick }, 1));

		Assert.Equal("HELLO", read.Text.Trim().ToUpperInvariant().Replace("\n", " "));
	}
}
