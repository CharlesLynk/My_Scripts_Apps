using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;

namespace MbrohDrawingIndexCreator;

internal static class PdfContentRedactor
{
	private readonly struct Matrix(double P_0, double P_1, double P_2, double P_3, double P_4, double P_5)
	{
		public readonly double A = P_0;

		public readonly double B = P_1;

		public readonly double C = P_2;

		public readonly double D = P_3;

		public readonly double E = P_4;

		public readonly double F = P_5;

		public static readonly Matrix Identity = new Matrix(1.0, 0.0, 0.0, 1.0, 0.0, 0.0);

		public Matrix Times(Matrix P_0)
		{
			return new Matrix(A * P_0.A + B * P_0.C, A * P_0.B + B * P_0.D, C * P_0.A + D * P_0.C, C * P_0.B + D * P_0.D, E * P_0.A + F * P_0.C + P_0.E, E * P_0.B + F * P_0.D + P_0.F);
		}

		public (double x, double y) Apply(double P_0, double P_1)
		{
			return (x: P_0 * A + P_1 * C + E, y: P_0 * B + P_1 * D + F);
		}
	}

	private sealed class Op
	{
		public string Name = "";

		public List<double> Nums = new List<double>();

		public string? Operand;

		public int Start;

		public int End;
	}

	private sealed class Subpath
	{
		public int Start;

		public int End;

		public bool AllInside = true;

		public bool Closed;
	}

	public static int RemoveContentInRect(PdfDocument P_0, PdfPage P_1, double P_2, double P_3, double P_4, double P_5)
	{
		List<PdfDictionary> list = new List<PdfDictionary>();
		PdfItem? pdfItem = P_1.Elements["/Contents"];
		if (pdfItem is PdfReference pdfReference)
		{
			pdfItem = pdfReference.Value;
		}
		if (pdfItem is PdfArray pdfArray)
		{
			foreach (PdfItem element in pdfArray.Elements)
			{
				if (((element is PdfReference pdfReference2) ? pdfReference2.Value : element) is PdfDictionary { Stream: not null } pdfDictionary)
				{
					list.Add(pdfDictionary);
				}
			}
		}
		else if (pdfItem is PdfDictionary { Stream: not null } pdfDictionary2)
		{
			list.Add(pdfDictionary2);
		}
		if (list.Count == 0)
		{
			return 0;
		}
		PdfDictionary? xobjects;
		using (MemoryStream memoryStream = new MemoryStream())
		{
			foreach (PdfDictionary item in list)
			{
				byte[] unfilteredValue = item.Stream.UnfilteredValue;
				memoryStream.Write(unfilteredValue, 0, unfilteredValue.Length);
				memoryStream.WriteByte(10);
			}
			xobjects = P_1.Elements.GetDictionary("/Resources")?.Elements.GetDictionary("/XObject");
			byte[]? array = Process(memoryStream.ToArray(), P_2, P_3, P_4, P_5, out var num, XObjectBounds);
			if (num == 0 || array == null)
			{
				return 0;
			}
			PdfDictionary pdfDictionary3 = new PdfDictionary(P_0);
			pdfDictionary3.CreateStream(Deflate(array));
			pdfDictionary3.Elements.SetName("/Filter", "/FlateDecode");
			P_0.Internals.AddObject(pdfDictionary3);
			P_1.Elements["/Contents"] = pdfDictionary3.Reference;
			return num;
		}
		(double[] BBox, double[] Matrix)? XObjectBounds(string name)
		{
			PdfItem? pdfItem2 = xobjects?.Elements[name];
			if (pdfItem2 is PdfReference pdfReference3)
			{
				pdfItem2 = pdfReference3.Value;
			}
			if (!(pdfItem2 is PdfDictionary pdfDictionary4))
			{
				return null;
			}
			string name2 = pdfDictionary4.Elements.GetName("/Subtype");
			if (name2 == "/Image")
			{
				return (new double[4] { 0.0, 0.0, 1.0, 1.0 }, new double[6] { 1.0, 0.0, 0.0, 1.0, 0.0, 0.0 });
			}
			if (name2 != "/Form")
			{
				return null;
			}
			double[]? array2 = ReadNumbers(pdfDictionary4.Elements.GetArray("/BBox"), 4);
			if (array2 == null)
			{
				return null;
			}
			double[] array3 = ReadNumbers(pdfDictionary4.Elements.GetArray("/Matrix"), 6) ?? new double[6] { 1.0, 0.0, 0.0, 1.0, 0.0, 0.0 };
			return (array2, array3);
		}
	}

	private static double[]? ReadNumbers(PdfArray? P_0, int P_1)
	{
		if (P_0 == null || P_0.Elements.Count < P_1)
		{
			return null;
		}
		double[] array = new double[P_1];
		for (int i = 0; i < P_1; i++)
		{
			PdfItem pdfItem = P_0.Elements[i];
			if (pdfItem is PdfReference pdfReference)
			{
				pdfItem = pdfReference.Value;
			}
			if (pdfItem is PdfInteger pdfInteger)
			{
				array[i] = pdfInteger.Value;
				continue;
			}
			if (pdfItem is PdfReal pdfReal)
			{
				array[i] = pdfReal.Value;
				continue;
			}
			return null;
		}
		return array;
	}

	private static byte[] Deflate(byte[] P_0)
	{
		using MemoryStream memoryStream = new MemoryStream();
		using (ZLibStream zLibStream = new ZLibStream(memoryStream, CompressionLevel.Optimal, true))
		{
			zLibStream.Write(P_0, 0, P_0.Length);
		}
		return memoryStream.ToArray();
	}

	internal static byte[]? Process(byte[] P_0, double P_1, double P_2, double P_3, double P_4, out int P_5, Func<string, (double[] BBox, double[] Matrix)?>? P_6 = null)
	{
		P_5 = 0;
		double minX = Math.Min(P_1, P_3) - 0.5;
		double maxX = Math.Max(P_1, P_3) + 0.5;
		double minY = Math.Min(P_2, P_4) - 0.5;
		double maxY = Math.Max(P_2, P_4) + 0.5;
		List<(int Start, int End, byte[]? Replacement)> list = new List<(int Start, int End, byte[]? Replacement)>();
		Matrix ctm = Matrix.Identity;
		int num = 0;
		Stack<(Matrix, int)> stack = new Stack<(Matrix, int)>();
		Matrix matrix = Matrix.Identity;
		Matrix matrix2 = Matrix.Identity;
		double num2 = 0.0;
		int num3 = -1;
		int num4 = 0;
		List<Subpath> subpaths = new List<Subpath>();
		Subpath? current = null;
		int pathStart = -1;
		bool hasClip = false;
		bool hasDependentSubpath = false;
		bool Inside((double x, double y) p)
		{
			return p.x >= minX && p.x <= maxX && p.y >= minY && p.y <= maxY;
		}
		void AddPoint(double x, double y)
		{
			if (current == null)
			{
				return;
			}
			if (!Inside(ctm.Apply(x, y)))
			{
				current.AllInside = false;
			}
		}
		void BeginSubpath(Op op, bool dependent)
		{
			if (pathStart < 0)
			{
				pathStart = op.Start;
			}
			current = new Subpath
			{
				Start = op.Start,
				End = op.End
			};
			subpaths.Add(current);
			if (!dependent)
			{
				hasDependentSubpath = true;
			}
		}
		void ContinueSubpath(Op op)
		{
			if (current == null || current.Closed)
			{
				BeginSubpath(op, dependent: false);
			}
			current!.End = op.End;
		}
		void ResetPath()
		{
			subpaths.Clear();
			current = null;
			pathStart = -1;
			hasClip = false;
			hasDependentSubpath = false;
		}
		foreach (Op item5 in Tokenize(P_0))
		{
			List<double> nums = item5.Nums;
			string name = item5.Name;
			int length;
			if (name != null)
			{
				length = name.Length;
				if (length != 1)
				{
					if (length == 2)
					{
						char c = name[1];
						if ((uint)c <= 68u)
						{
							if (c != '*')
							{
								if (c == 'D' && name == "TD")
								{
									goto IL_01c1;
								}
							}
							else if (name == "T*")
							{
								goto IL_01c1;
							}
						}
						else if (c != 'T')
						{
							if (c != 'd')
							{
								if (c == 'm' && name == "Tm")
								{
									goto IL_01c1;
								}
							}
							else if (name == "Td")
							{
								goto IL_01c1;
							}
						}
						else if (name == "BT" || name == "ET")
						{
							goto IL_01c1;
						}
					}
				}
				else
				{
					char c = name[0];
					if (c == '"' || c == '\'')
					{
						goto IL_01c1;
					}
				}
			}
			bool flag = false;
			goto IL_01c9;
			IL_01c1:
			flag = true;
			goto IL_01c9;
			IL_08c2:
			if (!hasClip && item5.Name != "n" && subpaths.Count > 0)
			{
				int num5 = 0;
				foreach (Subpath item6 in subpaths)
				{
					if (item6.AllInside)
					{
						num5++;
					}
				}
				if (num5 == subpaths.Count)
				{
					list.Add((pathStart, item5.End, null));
					P_5++;
				}
				else if (num5 > 0 && !hasDependentSubpath)
				{
					foreach (Subpath item7 in subpaths)
					{
						if (item7.AllInside)
						{
							list.Add((item7.Start, item7.End, null));
							P_5++;
						}
					}
				}
			}
			ResetPath();
			continue;
			IL_01c9:
			if (flag)
			{
				num3 = -1;
			}
			name = item5.Name;
			if (name == null)
			{
				continue;
			}
			length = name.Length;
			if (length != 1)
			{
				if (length != 2)
				{
					continue;
				}
				char c = name[1];
				if ((uint)c <= 84u)
				{
					if ((uint)c <= 68u)
					{
						if (c != '*')
						{
							if (c == 'D' && name == "TD" && nums.Count >= 2)
							{
								num2 = 0.0 - nums[nums.Count - 1];
								matrix2 = new Matrix(1.0, 0.0, 0.0, 1.0, nums[nums.Count - 2], nums[nums.Count - 1]).Times(matrix2);
								matrix = matrix2;
							}
							continue;
						}
						switch (name)
						{
						case "W*":
							break;
						case "f*":
						case "B*":
						case "b*":
							goto IL_08c2;
						case "T*":
							matrix2 = new Matrix(1.0, 0.0, 0.0, 1.0, 0.0, 0.0 - num2).Times(matrix2);
							matrix = matrix2;
							continue;
						default:
							continue;
						}
						goto IL_08b5;
					}
					if (c != 'J')
					{
						switch (c)
						{
						case 'T':
							if (name == "BT")
							{
								matrix = (matrix2 = Matrix.Identity);
							}
							break;
						case 'L':
							if (name == "TL" && nums.Count >= 1)
							{
								num2 = nums[nums.Count - 1];
							}
							break;
						}
						continue;
					}
					if (!(name == "TJ"))
					{
						continue;
					}
				}
				else
				{
					if ((uint)c > 106u)
					{
						switch (c)
						{
						case 'o':
						{
							if (!(name == "Do") || P_6 == null || item5.Operand == null)
							{
								break;
							}
							(double[], double[])? tuple = P_6(item5.Operand);
							if (tuple.HasValue)
							{
								double[] item = tuple.Value.Item1;
								double[] item2 = tuple.Value.Item2;
								Matrix matrix3 = new Matrix(item2[0], item2[1], item2[2], item2[3], item2[4], item2[5]).Times(ctm);
								if (Inside(matrix3.Apply(item[0], item[1])) && Inside(matrix3.Apply(item[2], item[1])) && Inside(matrix3.Apply(item[0], item[3])) && Inside(matrix3.Apply(item[2], item[3])))
								{
									list.Add((item5.Start, item5.End, null));
									P_5++;
								}
							}
							break;
						}
						case 'm':
							if (!(name == "cm"))
							{
								if (name == "Tm" && nums.Count >= 6)
								{
									matrix2 = new Matrix(nums[nums.Count - 6], nums[nums.Count - 5], nums[nums.Count - 4], nums[nums.Count - 3], nums[nums.Count - 2], nums[nums.Count - 1]);
									matrix = matrix2;
								}
							}
							else if (nums.Count >= 6)
							{
								ctm = new Matrix(nums[nums.Count - 6], nums[nums.Count - 5], nums[nums.Count - 4], nums[nums.Count - 3], nums[nums.Count - 2], nums[nums.Count - 1]).Times(ctm);
							}
							break;
						case 'r':
							if (name == "Tr" && nums.Count >= 1)
							{
								num = (int)nums[nums.Count - 1];
							}
							break;
						}
						continue;
					}
					if (c == 'd')
					{
						if (name == "Td" && nums.Count >= 2)
						{
							matrix2 = new Matrix(1.0, 0.0, 0.0, 1.0, nums[nums.Count - 2], nums[nums.Count - 1]).Times(matrix2);
							matrix = matrix2;
						}
						continue;
					}
					if (c == 'e')
					{
						if (name == "re")
						{
							BeginSubpath(item5, true);
							if (nums.Count >= 4)
							{
								double num6 = nums[nums.Count - 4];
								double num7 = nums[nums.Count - 3];
								double num8 = nums[nums.Count - 2];
								double num9 = nums[nums.Count - 1];
								AddPoint(num6, num7);
								AddPoint(num6 + num8, num7);
								AddPoint(num6, num7 + num9);
								AddPoint(num6 + num8, num7 + num9);
							}
							current!.Closed = true;
						}
						continue;
					}
					if (c != 'j' || !(name == "Tj"))
					{
						continue;
					}
				}
				goto IL_0bd7;
			}
			switch (name[0])
			{
			case 'q':
				stack.Push((ctm, num));
				continue;
			case 'Q':
				if (stack.Count > 0)
				{
					(ctm, num) = stack.Pop();
				}
				continue;
			case 'm':
				BeginSubpath(item5, true);
				if (nums.Count >= 2)
				{
					AddPoint(nums[nums.Count - 2], nums[nums.Count - 1]);
				}
				continue;
			case 'l':
				ContinueSubpath(item5);
				if (nums.Count >= 2)
				{
					AddPoint(nums[nums.Count - 2], nums[nums.Count - 1]);
				}
				continue;
			case 'c':
				ContinueSubpath(item5);
				if (nums.Count >= 6)
				{
					AddPoint(nums[nums.Count - 6], nums[nums.Count - 5]);
					AddPoint(nums[nums.Count - 4], nums[nums.Count - 3]);
					AddPoint(nums[nums.Count - 2], nums[nums.Count - 1]);
				}
				continue;
			case 'v':
			case 'y':
				ContinueSubpath(item5);
				if (nums.Count >= 4)
				{
					AddPoint(nums[nums.Count - 4], nums[nums.Count - 3]);
					AddPoint(nums[nums.Count - 2], nums[nums.Count - 1]);
				}
				continue;
			case 'h':
				if (current != null)
				{
					current.End = item5.End;
					current.Closed = true;
				}
				continue;
			case 'W':
				break;
			case 'B':
			case 'F':
			case 'S':
			case 'b':
			case 'f':
			case 'n':
			case 's':
				goto IL_08c2;
			case '"':
			case '\'':
				goto IL_0bd7;
			default:
				continue;
			}
			goto IL_08b5;
			IL_08b5:
			hasClip = true;
			continue;
			IL_0bd7:
			if (item5.Name != "Tj" && item5.Name != "TJ")
			{
				matrix2 = new Matrix(1.0, 0.0, 0.0, 1.0, 0.0, 0.0 - num2).Times(matrix2);
				matrix = matrix2;
			}
			if (num3 >= 0)
			{
				(int, int, byte[]?) tuple3 = list[num3];
				int item3 = tuple3.Item1;
				int item4 = tuple3.Item2;
				list[num3] = (item3, item4, HiddenTextReplacement(P_0, item3, item4, num4));
				num3 = -1;
			}
			if (num != 3 && Inside(matrix.Times(ctm).Apply(0.0, 0.0)))
			{
				list.Add((item5.Start, item5.End, null));
				P_5++;
				num3 = list.Count - 1;
				num4 = num;
			}
		}
		if (P_5 == 0)
		{
			return null;
		}
		list.Sort(((int Start, int End, byte[]? Replacement) a, (int Start, int End, byte[]? Replacement) b) => a.Start.CompareTo(b.Start));
		using MemoryStream memoryStream = new MemoryStream(P_0.Length);
		int num10 = 0;
		foreach (var (num11, num12, array) in list)
		{
			if (num11 >= num10)
			{
				memoryStream.Write(P_0, num10, num11 - num10);
				if (array != null)
				{
					memoryStream.Write(array, 0, array.Length);
				}
				else
				{
					memoryStream.WriteByte(32);
				}
				num10 = num12;
			}
		}
		memoryStream.Write(P_0, num10, P_0.Length - num10);
		return memoryStream.ToArray();
	}

	private static byte[] HiddenTextReplacement(byte[] P_0, int P_1, int P_2, int P_3)
	{
		List<byte> list = new List<byte>();
		list.AddRange(Encoding.ASCII.GetBytes("3 Tr\n"));
		for (int i = P_1; i < P_2; i++)
		{
			list.Add(P_0[i]);
		}
		list.AddRange(Encoding.ASCII.GetBytes("\n" + P_3.ToString(CultureInfo.InvariantCulture) + " Tr\n"));
		return list.ToArray();
	}

	private static bool IsWhite(byte P_0)
	{
		if (P_0 != 0 && P_0 != 9 && P_0 != 10 && P_0 != 12 && P_0 != 13)
		{
			return P_0 == 32;
		}
		return true;
	}

	private static bool IsDelim(byte P_0)
	{
		if (P_0 != 40 && P_0 != 41 && P_0 != 60 && P_0 != 62 && P_0 != 91 && P_0 != 93 && P_0 != 123 && P_0 != 125 && P_0 != 47)
		{
			return P_0 == 37;
		}
		return true;
	}

	private static IEnumerable<Op> Tokenize(byte[] P_0)
	{
		int i = 0;
		int len = P_0.Length;
		Op op = new Op
		{
			Start = -1
		};
		int num = 0;
		while (i < len)
		{
			byte b = P_0[i];
			if (IsWhite(b))
			{
				i++;
				continue;
			}
			switch (b)
			{
			case 37:
				for (; i < len && P_0[i] != 10 && P_0[i] != 13; i++)
				{
				}
				continue;
			case 40:
			{
				MarkOperand(i);
				int num2 = 0;
				for (; i < len; i++)
				{
					if (P_0[i] == 92)
					{
						i++;
					}
					else if (P_0[i] == 40)
					{
						num2++;
					}
					else if (P_0[i] == 41)
					{
						num2--;
						if (num2 == 0)
						{
							i++;
							break;
						}
					}
				}
				continue;
			}
			case 60:
				MarkOperand(i);
				if (i + 1 < len && P_0[i + 1] == 60)
				{
					num++;
					i += 2;
				}
				else
				{
					for (; i < len && P_0[i] != 62; i++)
					{
					}
					i++;
				}
				continue;
			case 62:
				if (i + 1 < len && P_0[i + 1] == 62)
				{
					num--;
					i += 2;
				}
				else
				{
					i++;
				}
				continue;
			case 91:
				MarkOperand(i);
				num++;
				i++;
				continue;
			case 93:
				num--;
				i++;
				continue;
			case 41:
			case 123:
			case 125:
				i++;
				continue;
			case 47:
			{
				MarkOperand(i);
				int num3 = i;
				for (i++; i < len && !IsWhite(P_0[i]) && !IsDelim(P_0[i]); i++)
				{
				}
				if (num == 0)
				{
					op.Operand = Encoding.ASCII.GetString(P_0, num3, i - num3);
				}
				continue;
			}
			}
			int num4 = i;
			for (; i < len && !IsWhite(P_0[i]) && !IsDelim(P_0[i]); i++)
			{
			}
			string text = Encoding.ASCII.GetString(P_0, num4, i - num4);
			if (text.Length > 0 && (char.IsDigit(text[0]) || text[0] == '-' || text[0] == '+' || text[0] == '.'))
			{
				MarkOperand(num4);
				if (num == 0 && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var item))
				{
					op.Nums.Add(item);
				}
				continue;
			}
			if (num <= 0)
			{
				switch (text)
				{
				case "true":
				case "false":
				case "null":
					break;
				default:
					if (op.Start < 0)
					{
						op.Start = num4;
					}
					op.Name = text;
					op.End = i;
					if (text == "BI")
					{
						int num5 = IndexOfKeyword(P_0, i, "ID");
						if (num5 < 0)
						{
							op.End = len;
							i = len;
						}
						else
						{
							int j = num5 + 3;
							int num6 = -1;
							for (; j + 1 < len; j++)
							{
								if (P_0[j] == 69 && P_0[j + 1] == 73 && IsWhite(P_0[j - 1]) && (j + 2 >= len || IsWhite(P_0[j + 2]) || IsDelim(P_0[j + 2])))
								{
									num6 = j;
									break;
								}
							}
							i = ((num6 < 0) ? len : (num6 + 2));
							op.End = i;
						}
					}
					yield return op;
					op = new Op
					{
						Start = -1
					};
					num = 0;
					continue;
				}
			}
			MarkOperand(num4);
		}
		void MarkOperand(int start)
		{
			if (op.Start < 0)
			{
				op.Start = start;
			}
		}
	}

	private static int IndexOfKeyword(byte[] P_0, int P_1, string P_2)
	{
		for (int i = P_1; i + P_2.Length <= P_0.Length; i++)
		{
			bool flag = true;
			for (int j = 0; j < P_2.Length; j++)
			{
				if (P_0[i + j] != P_2[j])
				{
					flag = false;
					break;
				}
			}
			if (flag)
			{
				bool num = i == 0 || IsWhite(P_0[i - 1]);
				bool flag2 = i + P_2.Length >= P_0.Length || IsWhite(P_0[i + P_2.Length]);
				if (num & flag2)
				{
					return i;
				}
			}
		}
		return -1;
	}
}
