using System;
using System.Collections.Generic;
using System.Linq;

namespace MbrohDrawingIndexCreator
{
	// One drawn shape (a stroked and/or filled path) already flattened to straight segments, in page units (points, origin bottom-left).
	internal sealed class VectorShape
	{
		internal readonly List<(double X, double Y)[]> Polylines = new List<(double, double)[]>();
		internal readonly List<bool> Closed = new List<bool>();
		internal bool Stroked;
		internal bool Filled;
		// 0 = black, 255 = white. A white fill hides what was drawn under it (CAD "wipeouts"), exactly as on paper.
		internal byte StrokeGray;
		internal byte FillGray;
		internal double MinX = double.MaxValue;
		internal double MinY = double.MaxValue;
		internal double MaxX = double.MinValue;
		internal double MaxY = double.MinValue;

		internal void Add((double X, double Y)[] points, bool closed)
		{
			if (points.Length == 0)
			{
				return;
			}
			Polylines.Add(points);
			Closed.Add(closed);
			foreach ((double x, double y) in points)
			{
				MinX = Math.Min(MinX, x);
				MinY = Math.Min(MinY, y);
				MaxX = Math.Max(MaxX, x);
				MaxY = Math.Max(MaxY, y);
			}
		}

		internal bool Intersects(double x1, double y1, double x2, double y2, double margin) => MaxX >= x1 - margin && MinX <= x2 + margin && MaxY >= y1 - margin && MinY <= y2 + margin;
	}

	// Draws a small region of a page from its vector shapes straight into a gray image. CAD plots draw their title-block lettering as
	// strokes, so no PDF renderer is needed to see it, and this is many times quicker than drawing the whole page.
	// Stroke width is chosen here, not taken from the file: PDF line widths are scaled by the page's own transform, which the shape data
	// does not carry, and OCR wants a consistent, fairly thin pen anyway.
	internal static class VectorRaster
	{
		internal const double MinimumPenPoints = 0.35;
		internal const double MaximumPenPoints = 1.0;
		internal const double DefaultPenPoints = 0.7;

		// A pen about a tenth of the letter height reads well: the typical glyph is the median height of the shapes in the region.
		// Frame lines and boxes (anything spanning most of the region) are not lettering and are ignored: counted, they made the pen as thick
		// as the box and turned small lettering into blobs.
		internal static double ChoosePenPoints(IEnumerable<VectorShape> shapes, double x1, double y1, double x2, double y2)
		{
			double regionWidth = x2 - x1;
			double regionHeight = y2 - y1;
			List<double> heights = shapes.Where(s => s.Stroked && s.Intersects(x1, y1, x2, y2, 0.0) && (s.MaxX - s.MinX) < 0.6 * regionWidth && (s.MaxY - s.MinY) < 0.6 * regionHeight)
				.Select(s => s.MaxY - s.MinY).Where(h => h >= 1.0 && h <= 40.0).OrderBy(h => h).ToList();
			if (heights.Count < 3)
			{
				return DefaultPenPoints;
			}
			return Math.Clamp(heights[heights.Count / 2] * 0.09, MinimumPenPoints, MaximumPenPoints);
		}

		internal static GrayImage Render(IReadOnlyList<VectorShape> shapes, double x1, double y1, double x2, double y2, double dpi, double penPoints)
		{
			double scale = dpi / 72.0;
			int width = Math.Max(1, (int)Math.Ceiling((x2 - x1) * scale));
			int height = Math.Max(1, (int)Math.Ceiling((y2 - y1) * scale));
			GrayImage image = new GrayImage(width, height);
			Array.Fill(image.Pixels, (byte)255);
			double half = Math.Max(0.5, penPoints * scale / 2.0);
			foreach (VectorShape shape in shapes)
			{
				if (!shape.Intersects(x1, y1, x2, y2, penPoints))
				{
					continue;
				}
				List<(double X, double Y)[]> device = shape.Polylines.Select(line => line.Select(p => ((p.X - x1) * scale, (y2 - p.Y) * scale)).ToArray()).ToList();
				if (shape.Filled)
				{
					FillPolygons(image, device.Where(d => d.Length >= 3).ToList(), shape.FillGray);
				}
				if (!shape.Stroked)
				{
					continue;
				}
				for (int k = 0; k < device.Count; k++)
				{
					(double X, double Y)[] line = device[k];
					for (int i = 0; i + 1 < line.Length; i++)
					{
						StrokeSegment(image, line[i], line[i + 1], half, shape.StrokeGray);
					}
					if (shape.Closed[k] && line.Length > 2)
					{
						StrokeSegment(image, line[^1], line[0], half, shape.StrokeGray);
					}
				}
			}			return image;
		}

		// Share of the image that is not white; zero ink means nothing was drawn there (or it was all hidden).
		internal static double InkFraction(GrayImage image)
		{
			if (image.Pixels.Length == 0)
			{
				return 0.0;
			}
			int dark = 0;
			foreach (byte value in image.Pixels)
			{
				if (value < 128)
				{
					dark++;
				}
			}
			return (double)dark / image.Pixels.Length;
		}

		private static void Blend(GrayImage image, int x, int y, double coverage, byte gray)
		{
			if (coverage <= 0.0)
			{
				return;
			}
			int index = y * image.Width + x;
			double under = image.Pixels[index];
			image.Pixels[index] = (byte)Math.Round(under + (gray - under) * Math.Min(1.0, coverage));
		}

		// A segment with a round pen: coverage falls off over one pixel around the pen's edge. Only pixels near the segment are visited,
		// so a long line across a large region costs its length, not its bounding box.
		private static void StrokeSegment(GrayImage image, (double X, double Y) a, (double X, double Y) b, double half, byte gray)
		{
			double dx = b.X - a.X;
			double dy = b.Y - a.Y;
			double lengthSquared = dx * dx + dy * dy;
			double reach = half + 1.0;
			bool horizontalMajor = Math.Abs(dx) >= Math.Abs(dy);
			if (horizontalMajor)
			{
				double slope = Math.Abs(dx) < 1e-9 ? 0.0 : Math.Abs(dy / dx);
				int xStart = Math.Max(0, (int)Math.Floor(Math.Min(a.X, b.X) - reach));
				int xEnd = Math.Min(image.Width - 1, (int)Math.Ceiling(Math.Max(a.X, b.X) + reach));
				double spread = reach * Math.Sqrt(1.0 + slope * slope) + 1.0;
				for (int x = xStart; x <= xEnd; x++)
				{
					double centerY = Math.Abs(dx) < 1e-9 ? a.Y : a.Y + (Math.Clamp(x + 0.5, Math.Min(a.X, b.X), Math.Max(a.X, b.X)) - a.X) / dx * dy;
					int yStart = Math.Max(0, (int)Math.Floor(centerY - spread));
					int yEnd = Math.Min(image.Height - 1, (int)Math.Ceiling(centerY + spread));
					for (int y = yStart; y <= yEnd; y++)
					{
						Blend(image, x, y, CoverageAt(x + 0.5, y + 0.5, a, dx, dy, lengthSquared, half), gray);
					}
				}
			}
			else
			{
				double slope = Math.Abs(dy) < 1e-9 ? 0.0 : Math.Abs(dx / dy);
				int yStart = Math.Max(0, (int)Math.Floor(Math.Min(a.Y, b.Y) - reach));
				int yEnd = Math.Min(image.Height - 1, (int)Math.Ceiling(Math.Max(a.Y, b.Y) + reach));
				double spread = reach * Math.Sqrt(1.0 + slope * slope) + 1.0;
				for (int y = yStart; y <= yEnd; y++)
				{
					double centerX = a.X + (Math.Clamp(y + 0.5, Math.Min(a.Y, b.Y), Math.Max(a.Y, b.Y)) - a.Y) / dy * dx;
					int xStart = Math.Max(0, (int)Math.Floor(centerX - spread));
					int xEnd = Math.Min(image.Width - 1, (int)Math.Ceiling(centerX + spread));
					for (int x = xStart; x <= xEnd; x++)
					{
						Blend(image, x, y, CoverageAt(x + 0.5, y + 0.5, a, dx, dy, lengthSquared, half), gray);
					}
				}
			}
		}

		private static double CoverageAt(double px, double py, (double X, double Y) a, double dx, double dy, double lengthSquared, double half)
		{
			double ox = px - a.X;
			double oy = py - a.Y;
			double t = lengthSquared < 1e-9 ? 0.0 : Math.Clamp((ox * dx + oy * dy) / lengthSquared, 0.0, 1.0);
			double ex = ox - t * dx;
			double ey = oy - t * dy;
			double distance = Math.Sqrt(ex * ex + ey * ey);
			return Math.Clamp(half + 0.5 - distance, 0.0, 1.0);
		}

		// Even-odd scanline fill over all of a shape's outlines together (holes in glyph outlines stay open).
		private static void FillPolygons(GrayImage image, List<(double X, double Y)[]> polygons, byte gray)
		{
			if (polygons.Count == 0)
			{
				return;
			}
			double minY = polygons.Min(poly => poly.Min(p => p.Y));
			double maxY = polygons.Max(poly => poly.Max(p => p.Y));
			int yStart = Math.Max(0, (int)Math.Floor(minY));
			int yEnd = Math.Min(image.Height - 1, (int)Math.Ceiling(maxY));
			List<double> crossings = new List<double>();
			for (int y = yStart; y <= yEnd; y++)
			{
				double sampleY = y + 0.5;
				crossings.Clear();
				foreach ((double X, double Y)[] polygon in polygons)
				{
					for (int i = 0; i < polygon.Length; i++)
					{
						(double X, double Y) a = polygon[i];
						(double X, double Y) b = polygon[(i + 1) % polygon.Length];
						if ((a.Y <= sampleY && b.Y > sampleY) || (b.Y <= sampleY && a.Y > sampleY))
						{
							crossings.Add(a.X + (sampleY - a.Y) / (b.Y - a.Y) * (b.X - a.X));
						}
					}
				}
				crossings.Sort();
				for (int k = 0; k + 1 < crossings.Count; k += 2)
				{
					int xStart = Math.Max(0, (int)Math.Round(crossings[k]));
					int xEnd = Math.Min(image.Width - 1, (int)Math.Round(crossings[k + 1]) - 1);
					for (int x = xStart; x <= xEnd; x++)
					{
						image.Pixels[y * image.Width + x] = gray;
					}
				}
			}
		}
	}
}