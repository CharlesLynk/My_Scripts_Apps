using System;
using System.Collections.Generic;
using System.Linq;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Graphics;
using UglyToad.PdfPig.Graphics.Colors;

namespace MbrohDrawingIndexCreator
{
	// Collects the vector shapes of a page that lie near a region, ready for VectorRaster. Declines (with a reason) when the page is not
	// really a vector drawing there, so the caller falls back to drawing the page as a picture.
	internal static class PdfPigVectorSource
	{
		private const double RegionMargin = 2.0;

		internal static bool TryCollect(Page page, double x1, double y1, double x2, double y2, out List<VectorShape> shapes, out string? reason)
			=> TryCollect(page, new[] { (x1, y1, x2, y2) }, out shapes, out reason);

		// Only the shapes that touch one of the given areas are collected (the field boxes), not the drawing around them.
		internal static bool TryCollect(Page page, IReadOnlyList<(double X1, double Y1, double X2, double Y2)> areas, out List<VectorShape> shapes, out string? reason)
		{
			shapes = new List<VectorShape>();
			reason = null;
			foreach (IPdfImage image in page.GetImages())
			{
				foreach ((double x1, double y1, double x2, double y2) in areas)
				{
					double areaSize = Math.Max(1.0, (x2 - x1) * (y2 - y1));
					double overlapWidth = Math.Min(x2, image.Bounds.Right) - Math.Max(x1, image.Bounds.Left);
					double overlapHeight = Math.Min(y2, image.Bounds.Top) - Math.Max(y1, image.Bounds.Bottom);
					if (overlapWidth > 0 && overlapHeight > 0 && overlapWidth * overlapHeight >= 0.1 * areaSize)
					{
						reason = "a picture covers the title block";
						return false;
					}
				}
			}
			IReadOnlyList<PdfPath> paths = page.Paths;
			if (paths.Count == 0)
			{
				reason = "the page has no vector drawing";
				return false;
			}
			foreach (PdfPath path in paths)
			{
				if (!path.IsStroked && !path.IsFilled)
				{
					continue;
				}
				VectorShape? shape = null;
				foreach (PdfSubpath subpath in path)
				{
					(double X, double Y)[]? points = Flatten(subpath, out bool closed);
					if (points == null || !areas.Any(a => Touches(points, a.X1, a.Y1, a.X2, a.Y2)))
					{
						continue;
					}
					shape ??= NewShape(path);
					shape.Add(points, closed);
				}
				if (shape != null)
				{
					shapes.Add(shape);
				}
			}
			if (shapes.Count == 0)
			{
				reason = "nothing is drawn in the title block";
				return false;
			}
			return true;
		}

		private static VectorShape NewShape(PdfPath path)
		{
			return new VectorShape
			{
				Stroked = path.IsStroked,
				Filled = path.IsFilled,
				StrokeGray = ToGray(path.StrokeColor),
				FillGray = ToGray(path.FillColor)
			};
		}

		private static byte ToGray(IColor? color)
		{
			if (color == null)
			{
				return 0;
			}
			try
			{
				(double r, double g, double b) = color.ToRGBValues();
				return (byte)Math.Clamp(Math.Round((0.299 * r + 0.587 * g + 0.114 * b) * 255.0), 0, 255);
			}
			catch (Exception ex) when (ex is InvalidOperationException || ex is NotSupportedException || ex is ArgumentException)
			{
				return 0;
			}
		}

		private static bool Touches((double X, double Y)[] points, double x1, double y1, double x2, double y2)
		{
			double minX = double.MaxValue;
			double minY = double.MaxValue;
			double maxX = double.MinValue;
			double maxY = double.MinValue;
			foreach ((double x, double y) in points)
			{
				minX = Math.Min(minX, x);
				minY = Math.Min(minY, y);
				maxX = Math.Max(maxX, x);
				maxY = Math.Max(maxY, y);
			}
			return maxX >= x1 - RegionMargin && minX <= x2 + RegionMargin && maxY >= y1 - RegionMargin && minY <= y2 + RegionMargin;
		}

		private static (double X, double Y)[]? Flatten(PdfSubpath subpath, out bool closed)
		{
			closed = false;
			List<(double X, double Y)> points = new List<(double, double)>();
			foreach (var command in subpath.Commands)
			{
				switch (command)
				{
					case PdfSubpath.Move move:
						points.Add((move.Location.X, move.Location.Y));
						break;
					case PdfSubpath.Line line:
						if (points.Count == 0)
						{
							points.Add((line.From.X, line.From.Y));
						}
						points.Add((line.To.X, line.To.Y));
						break;
					case PdfSubpath.BezierCurve curve:
						if (points.Count == 0)
						{
							points.Add((curve.StartPoint.X, curve.StartPoint.Y));
						}
						foreach (PdfSubpath.Line segment in curve.ToLines(8))
						{
							points.Add((segment.To.X, segment.To.Y));
						}
						break;
					case PdfSubpath.Close:
						closed = true;
						break;
				}
			}
			return points.Count == 0 ? null : points.ToArray();
		}
	}
}
