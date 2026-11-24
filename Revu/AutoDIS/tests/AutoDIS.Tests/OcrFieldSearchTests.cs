using System;
using Xunit;

namespace MbrohDrawingIndexCreator.Tests;

public sealed class OcrFieldSearchTests
{
	// A white image with filled dark rectangles (x, y, width, height).
	private static GrayImage Page(int width, int height, params (int x, int y, int w, int h)[] marks)
	{
		GrayImage image = new GrayImage(width, height);
		Array.Fill(image.Pixels, (byte)255);
		foreach ((int x, int y, int w, int h) in marks)
		{
			for (int row = y; row < y + h; row++)
			{
				for (int column = x; column < x + w; column++)
				{
					image.Pixels[row * width + column] = 0;
				}
			}
		}
		return image;
	}

	private static bool Dark(GrayImage image, int x, int y) => image.Pixels[y * image.Width + x] < 128;

	[Fact]
	public void TheLineNearestTheExpectedPlaceIsKept()
	{
		// The expected line is in the middle (rows 40-60); a taller line of a neighbour sits lower (rows 80-110).
		GrayImage image = Page(200, 120, (20, 40, 100, 20), (10, 80, 180, 30));

		GrayImage kept = image.SelectNearestLine(0.5);

		Assert.True(Dark(kept, 50, 50));
		Assert.False(Dark(kept, 50, 95));
	}

	[Fact]
	public void LinesCutByTheEdgeArePassedOverWhileAWholeLineExists()
	{
		// A sliver of a neighbour touches the bottom edge and happens to be the nearest to the anchor; the whole line wins.
		GrayImage image = Page(200, 100, (20, 30, 100, 20), (10, 95, 180, 5));

		GrayImage kept = image.SelectNearestLine(0.97);

		Assert.True(Dark(kept, 50, 40));
		Assert.False(Dark(kept, 50, 97));
	}

	[Fact]
	public void AllLinesCutByAnEdgeStillGiveAnAnswer()
	{
		GrayImage image = Page(200, 100, (20, 0, 100, 12), (10, 90, 180, 10));

		GrayImage kept = image.SelectNearestLine(0.0);

		Assert.True(Dark(kept, 50, 5));
	}

	[Fact]
	public void CroppingToInkCentresShortLetteringWithAMargin()
	{
		GrayImage image = Page(400, 60, (150, 20, 40, 20));

		GrayImage cropped = image.CropToInk();

		Assert.InRange(cropped.Width, 40, 120);
		Assert.InRange(cropped.Height, 20, 60);
		Assert.True(Dark(cropped, cropped.Width / 2, cropped.Height / 2));
	}

	[Fact]
	public void AMultiLineFieldDropsOnlyShortSlivers()
	{
		// Two full lines and a sliver: the lines stay, the sliver goes.
		GrayImage image = Page(200, 120, (10, 5, 4, 3), (20, 20, 150, 20), (20, 60, 150, 20));

		GrayImage kept = image.KeepTextLines(false);

		Assert.True(Dark(kept, 50, 30));
		Assert.True(Dark(kept, 50, 70));
		Assert.False(Dark(kept, 11, 6));
	}

	[Fact]
	public void ABlankImageIsLeftAlone()
	{
		GrayImage image = Page(100, 40);

		Assert.Same(image, image.CropToInk());
		Assert.Same(image, image.SelectNearestLine(0.5));
	}
}
