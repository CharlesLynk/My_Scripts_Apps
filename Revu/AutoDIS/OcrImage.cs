using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace MbrohDrawingIndexCreator
{
	// 8-bit grayscale image (0 = black, 255 = white) with the clean-up steps that help Tesseract read
	// CAD title blocks: border lines removed, thin strokes thickened, italics straightened.
	internal sealed class GrayImage
	{
		internal int Width { get; }

		internal int Height { get; }

		internal byte[] Pixels { get; }

		internal GrayImage(int width, int height, byte[]? pixels = null)
		{
			Width = Math.Max(1, width);
			Height = Math.Max(1, height);
			Pixels = pixels ?? CreateFilled(Width * Height, 255);
		}

		private static byte[] CreateFilled(int length, byte value)
		{
			byte[] data = new byte[length];
			Array.Fill(data, value);
			return data;
		}

#pragma warning disable CA1416 // Windows-only System.Drawing is intentional; the app and tests run on Windows.
		internal static GrayImage FromBitmap(Bitmap bitmap)
		{
			int width = bitmap.Width;
			int height = bitmap.Height;
			BitmapData data = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
			try
			{
				byte[] raw = new byte[Math.Abs(data.Stride) * height];
				Marshal.Copy(data.Scan0, raw, 0, raw.Length);
				byte[] gray = new byte[width * height];
				for (int y = 0; y < height; y++)
				{
					int row = y * data.Stride;
					for (int x = 0; x < width; x++)
					{
						int i = row + x * 4;
						gray[y * width + x] = (byte)((raw[i] * 29 + raw[i + 1] * 150 + raw[i + 2] * 77) >> 8);
					}
				}
				return new GrayImage(width, height, gray);
			}
			finally
			{
				bitmap.UnlockBits(data);
			}
		}
#pragma warning restore CA1416

		internal GrayImage Crop(int x, int y, int width, int height)
		{
			int x0 = Math.Clamp(x, 0, Width - 1);
			int y0 = Math.Clamp(y, 0, Height - 1);
			int w = Math.Clamp(width, 1, Width - x0);
			int h = Math.Clamp(height, 1, Height - y0);
			byte[] data = new byte[w * h];
			for (int row = 0; row < h; row++)
			{
				Buffer.BlockCopy(Pixels, (y0 + row) * Width + x0, data, row * w, w);
			}
			return new GrayImage(w, h, data);
		}

		// Keeps the real lines of lettering and erases slivers of neighbouring text that the field box catches (the tops of the title's
		// capitals inside a discipline box, the bottoms of a line above). Rows of ink are grouped into lines; a single-line field keeps only
		// its tallest line, a multi-line field drops lines much shorter than its tallest.
		internal GrayImage KeepTextLines(bool singleLine)
		{
			List<(int top, int bottom)> bands = FindLines(OtsuThreshold());
			if (bands.Count < 2)
			{
				return this;
			}
			int tallest = bands.Max(b => b.bottom - b.top + 1);
			byte[] result = (byte[])Pixels.Clone();
			bool keptOne = false;
			foreach ((int top, int bottom) in bands)
			{
				int height = bottom - top + 1;
				bool keep = singleLine ? (height == tallest && !keptOne) : height >= 0.35 * tallest;
				if (keep)
				{
					keptOne = true;
					continue;
				}
				Array.Fill(result, (byte)255, top * Width, (bottom - top + 1) * Width);
			}
			return new GrayImage(Width, Height, result);
		}

		// For a single-line field searched with a generous margin (so lettering drawn a little high or low is still inside): of the lines of ink,
		// keeps the one whose centre is nearest where the field is expected (anchorFraction: 0 = top edge, 1 = bottom edge). Lines cut by the
		// top or bottom edge are neighbours' lettering and are passed over while a whole line is available.
		internal GrayImage SelectNearestLine(double anchorFraction) => SelectLine(anchorFraction, 0) ?? this;

		// The rank-th nearest line (0 = nearest); null when there are not that many lines. Later ranks are for the case where the nearest
		// line did not read as the wanted field: the field may be one of its neighbours (a discipline written above its box).
		internal GrayImage? SelectLine(double anchorFraction, int rank)
		{
			int threshold = OtsuThreshold();
			List<(int top, int bottom)> bands = FindLines(threshold);
			if (bands.Count < 2)
			{
				return rank == 0 ? this : null;
			}
			double anchor = anchorFraction * Height;
			List<(int top, int bottom)> whole = bands.Where(b => b.top > 0 && b.bottom < Height - 1 && b.bottom - b.top + 1 >= Math.Max(3, Height * 0.06)).ToList();
			List<(int top, int bottom)> candidates = whole.Count > 0 ? whole : bands;
			List<(int top, int bottom)> ordered = candidates.OrderBy(b => Math.Abs((b.top + b.bottom) / 2.0 - anchor)).ToList();
			if (rank >= ordered.Count)
			{
				return null;
			}
			(int top, int bottom) chosen = ordered[rank];
			byte[] result = (byte[])Pixels.Clone();
			foreach ((int top, int bottom) in bands)
			{
				if (top == chosen.top)
				{
					continue;
				}
				Array.Fill(result, (byte)255, top * Width, (bottom - top + 1) * Width);
			}
			return new GrayImage(Width, Height, result);
		}
		private List<(int top, int bottom)> FindLines(int threshold)
		{
			int[] ink = new int[Height];
			for (int y = 0; y < Height; y++)
			{
				for (int x = 0; x < Width; x++)
				{
					if (Pixels[y * Width + x] <= threshold)
					{
						ink[y]++;
					}
				}
			}
			int gap = Math.Max(2, Height / 40);
			List<(int top, int bottom)> bands = new List<(int, int)>();
			int bandTop = -1;
			int lastInk = -1;
			for (int y = 0; y < Height; y++)
			{
				if (ink[y] == 0)
				{
					continue;
				}
				if (bandTop < 0)
				{
					bandTop = y;
				}
				else if (y - lastInk > gap)
				{
					bands.Add((bandTop, lastInk));
					bandTop = y;
				}
				lastInk = y;
			}
			if (bandTop >= 0)
			{
				bands.Add((bandTop, lastInk));
			}
			return bands;
		}

		// Trims the empty space around the lettering (keeping a margin of about half a letter height). A short word in a wide field -
		// a four-letter discipline in a box built for twenty - is otherwise lost in the white, and lettering drawn a little off-centre
		// ends up centred, so a field box that does not quite match the sheet still reads.
		internal GrayImage CropToInk(int minimumMargin = 6, double marginFraction = 0.5)
		{
			int threshold = OtsuThreshold();
			int minX = Width;
			int maxX = -1;
			int minY = Height;
			int maxY = -1;
			for (int y = 0; y < Height; y++)
			{
				for (int x = 0; x < Width; x++)
				{
					if (Pixels[y * Width + x] <= threshold)
					{
						minX = Math.Min(minX, x);
						maxX = Math.Max(maxX, x);
						minY = Math.Min(minY, y);
						maxY = Math.Max(maxY, y);
					}
				}
			}
			if (maxX < 0)
			{
				return this;
			}
			int margin = Math.Max(minimumMargin, (int)((maxY - minY + 1) * marginFraction));
			int left = Math.Max(0, minX - margin);
			int top = Math.Max(0, minY - margin);
			int right = Math.Min(Width - 1, maxX + margin);
			int bottom = Math.Min(Height - 1, maxY + margin);
			return left == 0 && top == 0 && right == Width - 1 && bottom == Height - 1 ? this : Crop(left, top, right - left + 1, bottom - top + 1);
		}

		// Labels the blobs of ink (8-connected). Used to tell letters from the debris of scanned frames.
		private List<(int x0, int y0, int x1, int y1, int area)> FindBlobs(int threshold)
		{
			bool[] seen = new bool[Pixels.Length];
			List<(int, int, int, int, int)> blobs = new List<(int, int, int, int, int)>();
			Stack<int> stack = new Stack<int>();
			for (int start = 0; start < Pixels.Length; start++)
			{
				if (seen[start] || Pixels[start] > threshold)
				{
					continue;
				}
				int x0 = Width;
				int y0 = Height;
				int x1 = 0;
				int y1 = 0;
				int area = 0;
				seen[start] = true;
				stack.Push(start);
				while (stack.Count > 0)
				{
					int p = stack.Pop();
					int y = p / Width;
					int x = p - y * Width;
					area++;
					x0 = Math.Min(x0, x);
					x1 = Math.Max(x1, x);
					y0 = Math.Min(y0, y);
					y1 = Math.Max(y1, y);
					for (int dy = -1; dy <= 1; dy++)
					{
						int ny = y + dy;
						if (ny < 0 || ny >= Height)
						{
							continue;
						}
						for (int dx = -1; dx <= 1; dx++)
						{
							int nx = x + dx;
							if (nx < 0 || nx >= Width)
							{
								continue;
							}
							int q = ny * Width + nx;
							if (!seen[q] && Pixels[q] <= threshold)
							{
								seen[q] = true;
								stack.Push(q);
							}
						}
					}
				}
				blobs.Add((x0, y0, x1, y1, area));
			}
			return blobs;
		}

		private void Erase(byte[] pixels, int x0, int y0, int x1, int y1)
		{
			for (int y = y0; y <= y1; y++)
			{
				Array.Fill(pixels, (byte)255, y * Width + x0, x1 - x0 + 1);
			}
		}

		// For a single line of lettering cut from a scan: keeps the run of letters and clears what lies beyond it (frame lines that are not
		// quite straight, dots, a border bar at either end) and any speck or line inside it. A run is letters of similar height, separated by
		// no more than a few letter heights; the run with the most letters wins.
		internal GrayImage KeepMainCluster()
		{
			int threshold = OtsuThreshold();
			List<(int x0, int y0, int x1, int y1, int area)> blobs = FindBlobs(threshold);
			List<int> heights = blobs.Select(b => b.y1 - b.y0 + 1).Where(h => h >= 6).OrderBy(h => h).ToList();
			if (heights.Count < 3)
			{
				return this;
			}
			double letterHeight = heights[heights.Count * 3 / 4];
			List<(int x0, int y0, int x1, int y1, int area)> letters = blobs.Where(b => b.y1 - b.y0 + 1 >= 0.45 * letterHeight && b.y1 - b.y0 + 1 <= 1.5 * letterHeight && b.x1 - b.x0 + 1 <= 4.0 * letterHeight)
				.OrderBy(b => b.x0).ToList();
			if (letters.Count < 2)
			{
				return this;
			}
			List<List<(int x0, int y0, int x1, int y1, int area)>> clusters = new List<List<(int, int, int, int, int)>>();
			int reach = int.MinValue;
			foreach ((int x0, int y0, int x1, int y1, int area) letter in letters)
			{
				if (clusters.Count == 0 || letter.x0 - reach > 2.5 * letterHeight)
				{
					clusters.Add(new List<(int, int, int, int, int)>());
				}
				clusters[^1].Add(letter);
				reach = Math.Max(reach, letter.x1);
			}
			List<(int x0, int y0, int x1, int y1, int area)> main = clusters.OrderByDescending(c => c.Count).First();
			int left = Math.Max(0, (int)(main.Min(b => b.x0) - 0.6 * letterHeight));
			int right = Math.Min(Width - 1, (int)(main.Max(b => b.x1) + 0.6 * letterHeight));
			int top = main.Min(b => b.y0);
			int bottom = main.Max(b => b.y1);
			byte[] result = (byte[])Pixels.Clone();
			foreach ((int x0, int y0, int x1, int y1, int area) blob in blobs)
			{
				int bw = blob.x1 - blob.x0 + 1;
				int bh = blob.y1 - blob.y0 + 1;
				bool outside = blob.x1 < left || blob.x0 > right;
				// A speck level with the letters may be a full stop; one above or below them is debris.
				bool speck = blob.area < 0.04 * letterHeight * letterHeight && bh < 0.35 * letterHeight && bw < 0.35 * letterHeight && (blob.y1 < top || blob.y0 > bottom);
				bool line = bw > 6.0 * letterHeight || bh > 1.6 * letterHeight;
				if (outside || speck || line)
				{
					Erase(result, blob.x0, blob.y0, blob.x1, blob.y1);
				}
			}
			return new GrayImage(Width, Height, result);
		}

		// For a title cut from a scan: clears the sheet's frame that the cut catches along its edges (double lines, slightly crooked and
		// dotted, that a straight-line test misses). A band of columns near the left or right edge (or rows near the top or bottom) that has
		// ink on most of its length is the frame; lettering never fills the whole height of a title field.
		internal GrayImage RemoveFrameBands()
		{
			int threshold = OtsuThreshold();
			byte[] result = (byte[])Pixels.Clone();
			bool changed = false;
			int window = Math.Max(3, Math.Min(Width, Height) / 60);
			int zoneX = Math.Max(window + 2, (int)(Width * 0.05));
			int zoneY = Math.Max(window + 2, (int)(Height * 0.12));
			for (int zone = 0; zone < 2; zone++)
			{
				int from = zone == 0 ? 0 : Width - zoneX;
				for (int x = from; x < from + zoneX - window; x++)
				{
					int covered = 0;
					for (int y = 0; y < Height; y++)
					{
						for (int k = 0; k < window; k++)
						{
							if (Pixels[y * Width + x + k] <= threshold)
							{
								covered++;
								break;
							}
						}
					}
					if (covered >= 0.8 * Height)
					{
						for (int y = 0; y < Height; y++)
						{
							Array.Fill(result, (byte)255, y * Width + x, window);
						}
						changed = true;
					}
				}
			}
			for (int zone = 0; zone < 2; zone++)
			{
				int from = zone == 0 ? 0 : Height - zoneY;
				for (int y = from; y < from + zoneY - window; y++)
				{
					int covered = 0;
					for (int x = 0; x < Width; x++)
					{
						for (int k = 0; k < window; k++)
						{
							if (Pixels[(y + k) * Width + x] <= threshold)
							{
								covered++;
								break;
							}
						}
					}
					if (covered >= 0.8 * Width)
					{
						for (int k = 0; k < window; k++)
						{
							Array.Fill(result, (byte)255, (y + k) * Width, Width);
						}
						changed = true;
					}
				}
			}
			return changed ? new GrayImage(Width, Height, result) : this;
		}

		// For a title cut from a scan: clears thin tall bars hugging the left or right edge (the sheet's frame, not lettering).
		internal GrayImage RemoveEdgeBars()
		{
			int threshold = OtsuThreshold();
			List<(int x0, int y0, int x1, int y1, int area)> blobs = FindBlobs(threshold);
			int letterHeight = TypicalLetterHeight();
			if (letterHeight <= 0)
			{
				return this;
			}
			byte[] result = (byte[])Pixels.Clone();
			bool changed = false;
			foreach ((int x0, int y0, int x1, int y1, int area) blob in blobs)
			{
				int bw = blob.x1 - blob.x0 + 1;
				int bh = blob.y1 - blob.y0 + 1;
				bool atEdge = blob.x1 < Width * 0.03 || blob.x0 > Width * 0.97;
				if (atEdge && bh >= 0.6 * letterHeight && bw <= 0.3 * letterHeight)
				{
					Erase(result, blob.x0, blob.y0, blob.x1, blob.y1);
					changed = true;
				}
			}
			return changed ? new GrayImage(Width, Height, result) : this;
		}

		// Height in pixels of the capital letters: the upper quartile of the heights of the ink blobs (letters, not dots or hyphens), which holds
		// even when two lines of a title touch. 0 when there is no ink.
		internal int TypicalLetterHeight()
		{
			int threshold = OtsuThreshold();
			bool[] seen = new bool[Pixels.Length];
			List<int> heights = new List<int>();
			Stack<int> stack = new Stack<int>();
			for (int start = 0; start < Pixels.Length; start++)
			{
				if (seen[start] || Pixels[start] > threshold)
				{
					continue;
				}
				int minY = Height;
				int maxY = 0;
				seen[start] = true;
				stack.Push(start);
				while (stack.Count > 0)
				{
					int p = stack.Pop();
					int y = p / Width;
					int x = p - y * Width;
					minY = Math.Min(minY, y);
					maxY = Math.Max(maxY, y);
					for (int dy = -1; dy <= 1; dy++)
					{
						int ny = y + dy;
						if (ny < 0 || ny >= Height)
						{
							continue;
						}
						for (int dx = -1; dx <= 1; dx++)
						{
							int nx = x + dx;
							if (nx < 0 || nx >= Width)
							{
								continue;
							}
							int q = ny * Width + nx;
							if (!seen[q] && Pixels[q] <= threshold)
							{
								seen[q] = true;
								stack.Push(q);
							}
						}
					}
				}
				int height = maxY - minY + 1;
				if (height >= 4 && height < Height * 0.9)
				{
					heights.Add(height);
				}
			}
			if (heights.Count == 0)
			{
				return 0;
			}
			heights.Sort();
			return heights[Math.Min(heights.Count - 1, heights.Count * 3 / 4)];
		}
		// Smaller copy, each new pixel the average of the pixels it covers (so thin strokes stay as lighter gray rather than vanishing).
		internal GrayImage Shrink(double scale)
		{
			if (scale >= 1.0)
			{
				return this;
			}
			int w = Math.Max(1, (int)Math.Round(Width * scale));
			int h = Math.Max(1, (int)Math.Round(Height * scale));
			byte[] data = new byte[w * h];
			for (int y = 0; y < h; y++)
			{
				int y0 = (int)((long)y * Height / h);
				int y1 = Math.Max(y0 + 1, (int)((long)(y + 1) * Height / h));
				for (int x = 0; x < w; x++)
				{
					int x0 = (int)((long)x * Width / w);
					int x1 = Math.Max(x0 + 1, (int)((long)(x + 1) * Width / w));
					int sum = 0;
					for (int yy = y0; yy < y1; yy++)
					{
						int row = yy * Width;
						for (int xx = x0; xx < x1; xx++)
						{
							sum += Pixels[row + xx];
						}
					}
					data[y * w + x] = (byte)(sum / ((y1 - y0) * (x1 - x0)));
				}
			}
			return new GrayImage(w, h, data);
		}

		// The reader rescales every line of text to a fixed working height of its own (about 36 px), so pixels beyond that add no detail, only noise and
		// time (and the layout analysis before it copes worse with huge letters). Letters taller than the useful height are shrunk to it;
		// smaller lettering is left alone. before is the measured letter height.
		internal GrayImage NormalizeLineHeight(int target, double tolerance, out int before, out bool shrunk)
		{
			before = TypicalLetterHeight();
			shrunk = false;
			if (before <= target * tolerance)
			{
				return this;
			}
			shrunk = true;
			return Shrink((double)target / before);
		}

		internal GrayImage Pad(int margin)
		{
			int w = Width + margin * 2;
			int h = Height + margin * 2;
			byte[] data = CreateFilled(w * h, 255);
			for (int row = 0; row < Height; row++)
			{
				Buffer.BlockCopy(Pixels, row * Width, data, (row + margin) * w + margin, Width);
			}
			return new GrayImage(w, h, data);
		}

		// Otsu's method: the gray level that best separates ink from paper.
		internal int OtsuThreshold()
		{
			int[] histogram = new int[256];
			foreach (byte p in Pixels)
			{
				histogram[p]++;
			}
			long total = Pixels.Length;
			double sumAll = 0;
			for (int i = 0; i < 256; i++)
			{
				sumAll += i * (double)histogram[i];
			}
			double sumBackground = 0;
			long weightBackground = 0;
			double bestVariance = -1;
			int best = 128;
			for (int t = 0; t < 256; t++)
			{
				weightBackground += histogram[t];
				if (weightBackground == 0)
				{
					continue;
				}
				long weightForeground = total - weightBackground;
				if (weightForeground == 0)
				{
					break;
				}
				sumBackground += t * (double)histogram[t];
				double meanBackground = sumBackground / weightBackground;
				double meanForeground = (sumAll - sumBackground) / weightForeground;
				double variance = (double)weightBackground * weightForeground * (meanBackground - meanForeground) * (meanBackground - meanForeground);
				if (variance > bestVariance)
				{
					bestVariance = variance;
					best = t;
				}
			}
			return best;
		}

		// Erases long thin dark runs (table borders and rules that cross a field) so they are not read as "|", "I" or "_".
		// Letter strokes are far shorter than the field, so they survive.
		internal GrayImage RemoveLongLines(double minHorizontalFraction = 0.25, double minVerticalFraction = 0.85, bool strictEdges = false)
		{
			int threshold = OtsuThreshold();
			// Looser than Otsu so the anti-aliased fringe of a line is erased with it.
			int inkLimit = threshold + (255 - threshold) / 2;
			byte[] result = (byte[])Pixels.Clone();
			bool[] erase = new bool[Pixels.Length];
			int minRunX = Math.Max(12, (int)(Width * minHorizontalFraction));
			int minRunY = Math.Max(12, (int)(Height * minVerticalFraction));
			for (int y = 0; y < Height; y++)
			{
				int x = 0;
				while (x < Width)
				{
					if (Pixels[y * Width + x] > inkLimit)
					{
						x++;
						continue;
					}
					int start = x;
					while (x < Width && Pixels[y * Width + x] <= inkLimit)
					{
						x++;
					}
					if (x - start >= minRunX)
					{
						for (int i = start; i < x; i++)
						{
							erase[y * Width + i] = true;
						}
					}
				}
			}
			for (int x = 0; x < Width; x++)
			{
				int y = 0;
				while (y < Height)
				{
					if (Pixels[y * Width + x] > inkLimit)
					{
						y++;
						continue;
					}
					int start = y;
					while (y < Height && Pixels[y * Width + x] <= inkLimit)
					{
						y++;
					}
					if (y - start >= minRunY)
					{
						for (int i = start; i < y; i++)
						{
							erase[i * Width + x] = true;
						}
					}
				}
			}
			for (int i = 0; i < erase.Length; i++)
			{
				if (erase[i])
				{
					result[i] = 255;
				}
			}
			return new GrayImage(Width, Height, result).RemoveEdgeComponents(inkLimit, 0.35, 0.8, strictEdges);
		}

		// Erases connected ink that touches the crop edge and spans much of it: the arcs of the DWG NO. oval, box corners and
		// neighbouring rules. Letters sit inside the crop and are left alone.
		// With strictEdges (single-line fields) anything touching the top or bottom edge goes, however small.
		internal GrayImage RemoveEdgeComponents(int inkLimit, double minWidthFraction = 0.35, double minHeightFraction = 0.8, bool strictEdges = false)
		{
			byte[] result = (byte[])Pixels.Clone();
			bool[] seen = new bool[Pixels.Length];
			int[] stack = new int[Pixels.Length];
			List<int> members = new List<int>();
			for (int start = 0; start < Pixels.Length; start++)
			{
				if (seen[start] || Pixels[start] > inkLimit)
				{
					continue;
				}
				members.Clear();
				int top = 0;
				stack[top++] = start;
				seen[start] = true;
				int minX = Width, maxX = 0, minY = Height, maxY = 0;
				while (top > 0)
				{
					int p = stack[--top];
					members.Add(p);
					int px = p % Width;
					int py = p / Width;
					minX = Math.Min(minX, px);
					maxX = Math.Max(maxX, px);
					minY = Math.Min(minY, py);
					maxY = Math.Max(maxY, py);
					for (int dy = -1; dy <= 1; dy++)
					{
						for (int dx = -1; dx <= 1; dx++)
						{
							int nx = px + dx;
							int ny = py + dy;
							if (nx < 0 || ny < 0 || nx >= Width || ny >= Height)
							{
								continue;
							}
							int n = ny * Width + nx;
							if (!seen[n] && Pixels[n] <= inkLimit)
							{
								seen[n] = true;
								stack[top++] = n;
							}
						}
					}
				}
				bool touchesEdge = minX == 0 || minY == 0 || maxX == Width - 1 || maxY == Height - 1;
				bool large = (maxX - minX + 1) >= Width * minWidthFraction || (maxY - minY + 1) >= Height * minHeightFraction;
				if ((touchesEdge && large) || (strictEdges && (minY == 0 || maxY == Height - 1)))
				{
					foreach (int member in members)
					{
						result[member] = 255;
					}
				}
			}
			return new GrayImage(Width, Height, result);
		}

		// Grayscale erosion (spreads dark pixels), which thickens hairline CAD strokes.
		internal GrayImage Thicken()
		{
			byte[] horizontal = new byte[Pixels.Length];
			for (int y = 0; y < Height; y++)
			{
				for (int x = 0; x < Width; x++)
				{
					byte v = Pixels[y * Width + x];
					if (x > 0 && Pixels[y * Width + x - 1] < v)
					{
						v = Pixels[y * Width + x - 1];
					}
					if (x < Width - 1 && Pixels[y * Width + x + 1] < v)
					{
						v = Pixels[y * Width + x + 1];
					}
					horizontal[y * Width + x] = v;
				}
			}
			byte[] result = new byte[Pixels.Length];
			for (int y = 0; y < Height; y++)
			{
				for (int x = 0; x < Width; x++)
				{
					byte v = horizontal[y * Width + x];
					if (y > 0 && horizontal[(y - 1) * Width + x] < v)
					{
						v = horizontal[(y - 1) * Width + x];
					}
					if (y < Height - 1 && horizontal[(y + 1) * Width + x] < v)
					{
						v = horizontal[(y + 1) * Width + x];
					}
					result[y * Width + x] = v;
				}
			}
			return new GrayImage(Width, Height, result);
		}

		// Darkens faint anti-aliased pixels (gamma > 1) without widening strokes, so hairline strokes survive binarisation.
		internal GrayImage Darken(double gamma)
		{
			byte[] lookup = new byte[256];
			for (int i = 0; i < 256; i++)
			{
				lookup[i] = (byte)Math.Clamp(Math.Round(255.0 * Math.Pow(i / 255.0, gamma)), 0, 255);
			}
			byte[] result = new byte[Pixels.Length];
			for (int i = 0; i < result.Length; i++)
			{
				result[i] = lookup[Pixels[i]];
			}
			return new GrayImage(Width, Height, result);
		}

		// Undoes a rightward lean: rows nearer the top move left. slope 0.2 is about 11 degrees.
		internal GrayImage Deslant(double slope)
		{
			int extra = (int)Math.Ceiling(slope * Height);
			int newWidth = Width + extra;
			byte[] result = CreateFilled(newWidth * Height, 255);
			for (int y = 0; y < Height; y++)
			{
				int shift = (int)Math.Round(slope * (Height - 1 - y));
				for (int x = 0; x < Width; x++)
				{
					int dest = x + extra - shift;
					if (dest >= 0 && dest < newWidth)
					{
						result[y * newWidth + dest] = Pixels[y * Width + x];
					}
				}
			}
			return new GrayImage(newWidth, Height, result);
		}

		// Uncompressed 8-bit BMP, which Tesseract's loader reads without any encode/decode cost.
		internal byte[] ToBmp()
		{
			int rowSize = (Width + 3) & ~3;
			int pixelBytes = rowSize * Height;
			int headerBytes = 14 + 40 + 1024;
			using MemoryStream stream = new MemoryStream(headerBytes + pixelBytes);
			using BinaryWriter writer = new BinaryWriter(stream);
			writer.Write((byte)'B');
			writer.Write((byte)'M');
			writer.Write(headerBytes + pixelBytes);
			writer.Write(0);
			writer.Write(headerBytes);
			writer.Write(40);
			writer.Write(Width);
			writer.Write(Height);
			writer.Write((short)1);
			writer.Write((short)8);
			writer.Write(0);
			writer.Write(pixelBytes);
			writer.Write(11811);
			writer.Write(11811);
			writer.Write(256);
			writer.Write(0);
			for (int i = 0; i < 256; i++)
			{
				writer.Write((byte)i);
				writer.Write((byte)i);
				writer.Write((byte)i);
				writer.Write((byte)0);
			}
			byte[] row = new byte[rowSize];
			for (int y = Height - 1; y >= 0; y--)
			{
				Buffer.BlockCopy(Pixels, y * Width, row, 0, Width);
				writer.Write(row);
			}
			writer.Flush();
			return stream.ToArray();
		}
	}
}
