using System;
using System.Drawing;
using System.IO;
using Tesseract;
using Xunit;

#pragma warning disable CA1416 // Windows-only System.Drawing is intentional; the app and tests run on Windows.

namespace MbrohDrawingIndexCreator.Tests;

public sealed class OcrImageTests
{
	private static GrayImage Letters(int letterHeight)
	{
		GrayImage image = new GrayImage(letterHeight * 12, letterHeight * 3);
		for (int letter = 0; letter < 8; letter++)
		{
			for (int y = letterHeight; y < letterHeight * 2; y++)
			{
				for (int x = 0; x < letterHeight / 5; x++)
				{
					image.Pixels[y * image.Width + letterHeight + letter * (letterHeight / 2 + letterHeight / 5) + x] = 0;
				}
			}
		}
		return image;
	}

	[Fact]
	public void TallLettersAreShrunkToTheWorkingHeightAndSmallOnesAreLeftAlone()
	{
		GrayImage tall = Letters(100);
		Assert.InRange(tall.TypicalLetterHeight(), 99, 101);
		GrayImage shrunk = tall.NormalizeLineHeight(34, 1.3, out int before, out bool didShrink);
		Assert.True(didShrink);
		Assert.InRange(before, 99, 101);
		Assert.InRange(shrunk.TypicalLetterHeight(), 30, 38);

		GrayImage small = Letters(30);
		Assert.Same(small, small.NormalizeLineHeight(34, 1.3, out _, out bool smallShrunk));
		Assert.False(smallShrunk);
	}

	[Fact]
	public void RemoveLongLines_ErasesBordersButKeepsLetters()
	{
		GrayImage image = new GrayImage(400, 100);
		for (int x = 0; x < 400; x++)
		{
			image.Pixels[2 * 400 + x] = 0;
			image.Pixels[3 * 400 + x] = 0;
		}
		for (int y = 30; y < 70; y++)
		{
			for (int x = 100; x < 106; x++)
			{
				image.Pixels[y * 400 + x] = 0;
			}
		}

		GrayImage cleaned = image.RemoveLongLines();

		Assert.Equal(255, cleaned.Pixels[2 * 400 + 200]);
		Assert.Equal(255, cleaned.Pixels[3 * 400 + 50]);
		Assert.Equal(0, cleaned.Pixels[50 * 400 + 103]);
	}

	[Fact]
	public void RemoveLongLines_StrictEdgesDropsArcFragmentsTouchingTheTopAndBottom()
	{
		GrayImage image = new GrayImage(300, 100);
		image.Pixels[0 * 300 + 20] = 0;
		image.Pixels[1 * 300 + 21] = 0;
		for (int y = 40; y < 60; y++)
		{
			image.Pixels[y * 300 + 150] = 0;
		}

		GrayImage cleaned = image.RemoveLongLines(0.25, 0.85, strictEdges: true);

		Assert.Equal(255, cleaned.Pixels[0 * 300 + 20]);
		Assert.Equal(0, cleaned.Pixels[50 * 300 + 150]);
	}

	[Fact]
	public void Thicken_WidensDarkStrokes()
	{
		GrayImage image = new GrayImage(20, 20);
		for (int y = 5; y < 15; y++)
		{
			image.Pixels[y * 20 + 10] = 0;
		}

		GrayImage thick = image.Thicken();

		Assert.Equal(0, thick.Pixels[10 * 20 + 9]);
		Assert.Equal(0, thick.Pixels[10 * 20 + 11]);
	}

	[Fact]
	public void Deslant_StraightensALeaningStroke()
	{
		GrayImage image = new GrayImage(60, 40);
		for (int y = 0; y < 40; y++)
		{
			int x = 20 + (int)Math.Round(0.2 * (39 - y));
			image.Pixels[y * 60 + x] = 0;
		}

		GrayImage straight = image.Deslant(0.2);

		int firstColumn = -1;
		for (int y = 0; y < 40; y++)
		{
			for (int x = 0; x < straight.Width; x++)
			{
				if (straight.Pixels[y * straight.Width + x] == 0)
				{
					Assert.True(firstColumn < 0 || firstColumn == x, "the stroke should be vertical after de-slanting");
					firstColumn = x;
				}
			}
		}
		Assert.True(firstColumn >= 0);
	}

	[Fact]
	public void ToBmp_ProducesALoadableImageOfTheSameSize()
	{
		GrayImage image = new GrayImage(33, 7);

		using MemoryStream stream = new MemoryStream(image.ToBmp());
		using Bitmap bitmap = new Bitmap(stream);

		Assert.Equal(33, bitmap.Width);
		Assert.Equal(7, bitmap.Height);
	}

	[Fact]
	public void Recognizer_ReadsATitleThroughABorderLine()
	{
		OcrEngineResult fast = OcrSupport.TryCreateEngine(OcrSupport.DefaultTessdataDirectory, OcrSupport.FastModel);
		using TesseractEngine? engine = fast.Engine;
		Assert.NotNull(engine);
		GrayImage field = RenderText("DESIGN CRITERIA", 900, 160, withBorder: true);

		OcrFieldResult result = OcrFieldRecognizer.Read(engine!, field, OcrFieldKind.Title, OcrPresets.Get(OcrPreset.Balanced).OptionsFor(OcrFieldKind.Title));

		Assert.Contains("CRITERIA", result.Text.ToUpperInvariant());
	}

	[Fact]
	public void Recognizer_ReadsASheetIdAfterCleanup()
	{
		OcrEngineResult fast = OcrSupport.TryCreateEngine(OcrSupport.DefaultTessdataDirectory, OcrSupport.FastModel);
		using TesseractEngine? engine = fast.Engine;
		Assert.NotNull(engine);
		GrayImage field = RenderText("G-16", 400, 140, withBorder: false);

		OcrFieldResult result = OcrFieldRecognizer.Read(engine!, field, OcrFieldKind.SheetId, OcrPresets.Get(OcrPreset.Balanced).OptionsFor(OcrFieldKind.SheetId));

		Assert.Equal("G-16", OcrTextCleanup.CleanSheetId(result.Text, expectHyphen: true));
	}

	[Theory]
	[InlineData(0)]
	[InlineData(1)]
	[InlineData(2)]
	public void EveryFastModelPreset_ReadsATitle(int presetIndex)
	{
		OcrEngineResult fast = OcrSupport.TryCreateEngine(OcrSupport.DefaultTessdataDirectory, OcrSupport.FastModel);
		using TesseractEngine? engine = fast.Engine;
		Assert.NotNull(engine);
		GrayImage field = RenderText("DESIGN CRITERIA II", 1000, 160, withBorder: true);

		OcrFieldResult result = OcrFieldRecognizer.Read(engine!, field, OcrFieldKind.Title, OcrPresets.Get((OcrPreset)presetIndex).OptionsFor(OcrFieldKind.Title));

		Assert.Contains("CRITERIA", result.Text.ToUpperInvariant());
	}

	[Fact]
	public void Darken_MakesFaintPixelsDarkerAndLeavesBlackAndWhiteAlone()
	{
		GrayImage image = new GrayImage(3, 1, new byte[] { 0, 200, 255 });

		GrayImage darker = image.Darken(3);

		Assert.Equal(0, darker.Pixels[0]);
		Assert.True(darker.Pixels[1] < 200);
		Assert.Equal(255, darker.Pixels[2]);
	}
	private static GrayImage RenderText(string text, int width, int height, bool withBorder)
	{
		using Bitmap bitmap = new Bitmap(width, height);
		using (Graphics graphics = Graphics.FromImage(bitmap))
		using (Font font = new Font("Arial", 56f, FontStyle.Italic, GraphicsUnit.Pixel))
		{
			graphics.Clear(Color.White);
			graphics.DrawString(text, font, Brushes.Black, 30f, 30f);
			if (withBorder)
			{
				graphics.FillRectangle(Brushes.Black, 0, 0, width, 3);
				graphics.FillRectangle(Brushes.Black, 0, height - 3, width, 3);
			}
		}
		return GrayImage.FromBitmap(bitmap);
	}
}
