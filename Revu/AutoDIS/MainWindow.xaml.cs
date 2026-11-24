using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Markup;
using Microsoft.UI.Xaml.Media;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using Tesseract;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Graphics;
using WinRT;
using WinRT.Interop;
using Windows.ApplicationModel;
using Windows.Foundation;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.System;
using Windows.UI;

namespace MbrohDrawingIndexCreator;

public sealed partial class MainWindow : Window
{
	private sealed class ProgressTracker
	{
		private readonly IProgress<(int current, int total, string status)>? _progress;

		public int Total { get; set; }

		public int Current { get; private set; }

		public ProgressTracker(IProgress<(int current, int total, string status)>? P_0, int P_1)
		{
			_progress = P_0;
			Total = P_1;
		}

		private readonly object _gate = new object();

		public void Report(string P_0)
		{
			lock (_gate)
			{
				Current++;
				_progress?.Report((Current, Total, P_0));
			}
		}

		// A new status line without moving the bar (stages that have no countable steps).
		public void Status(string P_0)
		{
			lock (_gate)
			{
				_progress?.Report((Current, Total, P_0));
			}
		}

		// For callers on several threads: the message is built from the count this call claims.
		public void Report(Func<int, string> P_0)
		{
			lock (_gate)
			{
				Current++;
				_progress?.Report((Current, Total, P_0(Current)));
			}
		}
	}

	private sealed class NaturalFileNameComparer : IComparer<string>
	{
		public static readonly NaturalFileNameComparer Instance = new NaturalFileNameComparer();

		public int Compare(string? P_0, string? P_1)
		{
			if (P_0 == null || P_1 == null)
			{
				return string.Compare(P_0, P_1, StringComparison.OrdinalIgnoreCase);
			}
			int i = 0;
			int j = 0;
			while (i < P_0.Length && j < P_1.Length)
			{
				char c = P_0[i];
				char c2 = P_1[j];
				if (char.IsDigit(c) && char.IsDigit(c2))
				{
					int num = i;
					int num2 = j;
					for (; i < P_0.Length && char.IsDigit(P_0[i]); i++)
					{
					}
					for (; j < P_1.Length && char.IsDigit(P_1[j]); j++)
					{
					}
					string text = P_0.Substring(num, i - num).TrimStart('0');
					string text2 = P_1.Substring(num2, j - num2).TrimStart('0');
					if (text.Length != text2.Length)
					{
						return text.Length.CompareTo(text2.Length);
					}
					int num3 = string.CompareOrdinal(text, text2);
					if (num3 != 0)
					{
						return num3;
					}
				}
				else
				{
					int num4 = char.ToUpperInvariant(c).CompareTo(char.ToUpperInvariant(c2));
					if (num4 != 0)
					{
						return num4;
					}
					i++;
					j++;
				}
			}
			return (P_0.Length - i).CompareTo(P_1.Length - j);
		}
	}

	private sealed class TablePagination
	{
		public List<List<List<(bool isHeader, string headerText, DrawingIndexItem? row)>>> Pages = new List<List<List<(bool, string, DrawingIndexItem)>>>();

		public double HeaderHeight;

		public double RowHeight;

		public double HeaderFontSize;

		public double CellFontSize;
	}

	private sealed class SystemFontResolver : IFontResolver, IFontResolverMarker
	{
		public FontResolverInfo ResolveTypeface(string P_0, bool P_1, bool P_2)
		{
			string text = "arial";
			if (!string.IsNullOrWhiteSpace(P_0))
			{
				string text2 = P_0.ToLowerInvariant();
				if (text2.Contains("helvetica") || text2.Contains("arial") || text2.Contains("sans"))
				{
					text = "arial";
				}
				else if (text2.Contains("times") || text2.Contains("serif"))
				{
					text = "times";
				}
				else if (text2.Contains("courier") || text2.Contains("monospace"))
				{
					text = "courier";
				}
			}
			text += (P_1 ? "#b" : "#r");
			return new FontResolverInfo(text);
		}

		public byte[] GetFont(string P_0)
		{
			try
			{
				string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
				string text;
				if (P_0.ToLowerInvariant().StartsWith("times"))
				{
					text = Path.Combine(folderPath, P_0.ToLowerInvariant().Contains("#b") ? "timesbd.ttf" : "times.ttf");
				}
				else
				{
					text = (P_0.ToLowerInvariant().StartsWith("courier") ? Path.Combine(folderPath, P_0.ToLowerInvariant().Contains("#b") ? "courbd.ttf" : "cour.ttf") : Path.Combine(folderPath, P_0.ToLowerInvariant().Contains("#b") ? "arialbd.ttf" : "arial.ttf"));
				}
				if (File.Exists(text))
				{
					return File.ReadAllBytes(text);
				}
			}
			catch
			{
			}
			return Array.Empty<byte>();
		}
	}

	private readonly record struct FieldRect
	{
		public double X1 { get; }

		public double Y1 { get; }

		public double X2 { get; }

		public double Y2 { get; }

		public FieldRect(double P_0, double P_1, double P_2, double P_3)
		{
			X1 = P_0;
			Y1 = P_1;
			X2 = P_2;
			Y2 = P_3;
		}
	}

	private readonly record struct TextHit
	{
		public string Text { get; }

		public double Left { get; }

		public double Bottom { get; }

		public double Right { get; }

		public double Top { get; }

		public TextHit(string P_0, double P_1, double P_2, double P_3, double P_4)
		{
			Text = P_0;
			Left = P_1;
			Bottom = P_2;
			Right = P_3;
			Top = P_4;
		}
	}

	private sealed class BookmarkNode
	{
		public string Title = string.Empty;

		public int? SourcePageIndex;

		public List<BookmarkNode> Children = new List<BookmarkNode>();

		public int? ForcedOutputPageIndex;
	}

	private CancellationTokenSource? _generationCts;
	private bool _closeAfterRun;

	private static readonly ConditionalWeakTable<PdfSharp.Pdf.PdfDocument, Dictionary<string, PdfDictionary>> _fontResourceCache = new ConditionalWeakTable<PdfSharp.Pdf.PdfDocument, Dictionary<string, PdfDictionary>>();

	private static readonly (double r, double g, double b) NotFoundHighlightColor = (r: 1.0, g: 1.0, b: 0.0);

	// OCR was not confident about this cell: orange, so it stands apart from the yellow "not found" cells.
	private static readonly (double r, double g, double b) UncertainHighlightColor = (r: 1.0, g: 0.72, b: 0.3);

	private static OcrPreset OcrPresetSetting = LoadOcrPreset();

	private static ResourceLevel ResourceLevelSetting = LoadResourceLevel();

	// Fixed image resolution for reading title blocks (0 = automatic, the preset decides).
	// An example sheet number typed by the user (optional). It is not saved: the format belongs to one drawing set, not to the program.
	private static string IdExampleSetting = string.Empty;

	private static OcrPresetSettings CurrentOcrPreset => OcrPresets.Get(OcrPresetSetting);

	public MainWindow()
	{
		InitializeComponent();
		Title = "Mbroh Drawing Index Creator";
		Task.Run(() => FileHygiene.SweepStale(Path.GetTempPath(), "AutoDIS_*.pdf", TimeSpan.FromDays(1)));
		try
		{
			ExtendsContentIntoTitleBar = true;
			SetTitleBar(AppTitleBar);
			AppWindowTitleBar titleBar = AppWindow.TitleBar;
			Windows.UI.Color value = Windows.UI.Color.FromArgb(255, 255, 255, 255);
			Windows.UI.Color value2 = Windows.UI.Color.FromArgb(255, 229, 229, 229);
			Windows.UI.Color value3 = Windows.UI.Color.FromArgb(255, 204, 204, 204);
			Windows.UI.Color value4 = Windows.UI.Color.FromArgb(255, 11, 31, 65);
			titleBar.BackgroundColor = value;
			titleBar.ForegroundColor = value4;
			titleBar.ButtonBackgroundColor = value;
			titleBar.ButtonForegroundColor = value4;
			titleBar.ButtonHoverBackgroundColor = value2;
			titleBar.ButtonHoverForegroundColor = value4;
			titleBar.ButtonPressedBackgroundColor = value3;
			titleBar.ButtonPressedForegroundColor = value4;
			titleBar.ButtonInactiveBackgroundColor = value;
			titleBar.ButtonInactiveForegroundColor = value4;
		}
		catch
		{
		}
		try
		{
			string text = Path.Combine(GetPackageContentDirectory(), "Assets", "AppIcon.ico");
			if (File.Exists(text))
			{
				AppWindow.SetIcon(text);
			}
		}
		catch
		{
		}
		AppDomain.CurrentDomain.UnhandledException += (object s, System.UnhandledExceptionEventArgs ev) =>
		{
			try
			{
				Exception ex = ev.ExceptionObject as Exception;
				File.WriteAllText(GetLogFilePath("unhandled_exception.txt"), ex?.ToString() ?? ev.ToString());
			}
			catch
			{
			}
		};
		TaskScheduler.UnobservedTaskException += (object? s, UnobservedTaskExceptionEventArgs ev) =>
		{
			try
			{
				File.WriteAllText(GetLogFilePath("unobserved_task_exception.txt"), ev.Exception?.ToString() ?? ev.ToString());
			}
			catch
			{
			}
			ev.SetObserved();
		};
		AppWindow.Closing += (AppWindow s, AppWindowClosingEventArgs ev) =>
		{
			if (_generationCts != null) { ev.Cancel = true; _closeAfterRun = true; GenerationStatusText.Text = "Stopping..."; _generationCts.Cancel(); }
		};
	}

	private void RootGrid_Loaded(object sender, RoutedEventArgs e)
	{
		FocusSink.Focus(FocusState.Programmatic);
	}

	private static string GetLogDirectory()
	{
		try
		{
			return ApplicationData.Current.LocalFolder.Path;
		}
		catch
		{
			return AppContext.BaseDirectory;
		}
	}

	private static string GetLogFilePath(string P_0)
	{
		return Path.Combine(GetLogDirectory(), P_0);
	}

	private static string GetOrCreateOutputFolder(string P_0)
	{
		string text = Path.Combine(Path.GetDirectoryName(P_0) ?? Environment.CurrentDirectory, "Output");
		Directory.CreateDirectory(text);
		return text;
	}

	private static string GetPackageContentDirectory()
	{
		try
		{
			return Package.Current.InstalledLocation.Path;
		}
		catch
		{
			return AppContext.BaseDirectory;
		}
	}

	private static string GetAppVersion()
	{
		try
		{
			PackageVersion version = Package.Current.Id.Version;
			return $"{version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
		}
		catch
		{
			Version version2 = Assembly.GetExecutingAssembly().GetName().Version;
			return (version2 != null) ? version2.ToString() : "unknown (unpackaged build)";
		}
	}

	private async void SettingsMenuItem_Click(object? sender, RoutedEventArgs e)
	{
		ComboBox presetBox = new ComboBox
		{
			Header = "OCR speed / detail",
			HorizontalAlignment = HorizontalAlignment.Stretch
		};
		foreach (OcrPresetSettings preset in OcrPresets.Presets)
		{
			presetBox.Items.Add(preset.DisplayName);
		}
		presetBox.SelectedIndex = (int)OcrPresetSetting;
		TextBlock presetDescription = new TextBlock
		{
			TextWrapping = TextWrapping.Wrap,
			Text = CurrentOcrPreset.Description
		};
		presetBox.SelectionChanged += (object s, SelectionChangedEventArgs args) =>
		{
			if (presetBox.SelectedIndex >= 0)
			{
				presetDescription.Text = OcrPresets.Get((OcrPreset)presetBox.SelectedIndex).Description;
			}
		};
		ComboBox resourceBox = new ComboBox
		{
			Header = "Computer resources",
			HorizontalAlignment = HorizontalAlignment.Stretch
		};
		foreach (ResourceLevel level in Enum.GetValues<ResourceLevel>())
		{
			resourceBox.Items.Add(ResourcePlanner.DisplayName(level));
		}
		resourceBox.SelectedIndex = (int)ResourceLevelSetting;
		HardwareSnapshot hardwareNow = HardwareProbe.Probe();
		TextBlock resourceDescription = new TextBlock
		{
			TextWrapping = TextWrapping.Wrap
		};
		void ShowResourceDescription()
		{
			ResourceLevel selected = (ResourceLevel)Math.Max(0, resourceBox.SelectedIndex);
			ResourcePlan plan = ResourcePlanner.Plan(selected, hardwareNow);
			resourceDescription.Text = ResourcePlanner.Description(selected) + "\n" + ResourcePlanner.Summary(plan, hardwareNow);
		}
		ShowResourceDescription();
		resourceBox.SelectionChanged += (object s, SelectionChangedEventArgs args) => ShowResourceDescription();
		CheckBox reviewBox = new CheckBox
		{
			Content = "Review the discipline headings (with a picture of each) before the index is saved",
			IsChecked = ReviewBeforeSaveSetting
		};
		TextBox idExampleBox = new TextBox
		{
			Header = "Sheet number example (optional)",
			PlaceholderText = "for example HH-A-01",
			Text = IdExampleSetting,
			HorizontalAlignment = HorizontalAlignment.Stretch
		};
		StackPanel stackPanel = new StackPanel
		{
			Spacing = 8.0
		};
		stackPanel.Children.Add(new TextBlock
		{
			Text = "Some sheets have title blocks that cannot be read as text (scanned pages, or CAD lettering drawn as lines). For those, the app reads the title block as an image (OCR). This setting trades speed for detail. Balanced suits most drawing sets; choose Quality or Ultra for scanned drawings or sets with thin or italic lettering throughout, and Fast or Fastest for clean sets when you need the quickest result.",
			TextWrapping = TextWrapping.Wrap
		});
		stackPanel.Children.Add(presetBox);
		stackPanel.Children.Add(presetDescription);
		stackPanel.Children.Add(new TextBlock
		{
			Text = "Entries that OCR is not sure about are highlighted in orange in the index so you can check them.",
			TextWrapping = TextWrapping.Wrap
		});
		stackPanel.Children.Add(new TextBlock
		{
			Text = "How much of this computer a run may use. The program checks what the computer has and takes more or less of it; it does not change the result, only how fast it arrives and how much else you can do meanwhile.",
			TextWrapping = TextWrapping.Wrap
		});
		stackPanel.Children.Add(resourceBox);
		stackPanel.Children.Add(resourceDescription);
		stackPanel.Children.Add(reviewBox);
		stackPanel.Children.Add(new TextBlock
		{
			Text = "If the sheet numbers in this set are hard to read, type one of them here as an example. Only its pattern is used (letters, digits and dashes, not their count), so HH-A-01 also covers WT-A-104. Numbers read in a different format are corrected toward it where possible and highlighted. The example is used for this session only; clear it for the next drawing set.",
			TextWrapping = TextWrapping.Wrap
		});
		stackPanel.Children.Add(idExampleBox);
		await ShowPopupWindowAsync("Settings", null, stackPanel, "Save", "Cancel", 720.0, 760.0, onPrimary: () =>
		{
			if (presetBox.SelectedIndex >= 0)
			{
				OcrPresetSetting = (OcrPreset)presetBox.SelectedIndex;
				SaveOcrPreset(OcrPresetSetting);
			}
			if (resourceBox.SelectedIndex >= 0)
			{
				ResourceLevelSetting = (ResourceLevel)resourceBox.SelectedIndex;
				SaveResourceLevel(ResourceLevelSetting);
			}
			IdExampleSetting = (idExampleBox.Text ?? string.Empty).Trim();
			ReviewBeforeSaveSetting = reviewBox.IsChecked == true;
			SaveReviewBeforeSave(ReviewBeforeSaveSetting);
		});
	}
	private async void AboutMenuItem_Click(object? sender, RoutedEventArgs e)
	{
		await new ContentDialog
		{
			Title = "About MbrohDrawingIndexCreator",
			Content = "Version " + GetAppVersion(),
			CloseButtonText = "OK",
			XamlRoot = Content.XamlRoot
		}.ShowAsync();
	}

	private async void TemplateHelpButton_Click(object? sender, RoutedEventArgs e)
	{
		string instructionsPath = Path.Combine(GetPackageContentDirectory(), "Assets", "AutoDISTemplateInstructions.docx");
		if (!File.Exists(instructionsPath))
		{
			await new ContentDialog
			{
				Title = "Instructions not found",
				Content = "Could not find the template instructions file at:\n" + instructionsPath,
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
			return;
		}
		try
		{
			await Launcher.LaunchFileAsync(await StorageFile.GetFileFromPathAsync(instructionsPath));
		}
		catch (Exception ex)
		{
			await new ContentDialog
			{
				Title = "Could not open instructions",
				Content = ex.Message + "\n\nFile location:\n" + instructionsPath,
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
		}
	}

	private async void OpenLogFileMenuItem_Click(object? sender, RoutedEventArgs e)
	{
		string logPath = GetLogFilePath("export_log.txt");
		if (!File.Exists(logPath))
		{
			await new ContentDialog
			{
				Title = "Log file not found",
				Content = "No log file exists yet at:\n" + logPath + "\n\nIt's created the first time you generate a drawing index.",
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
			return;
		}
		try
		{
			await Launcher.LaunchFileAsync(await StorageFile.GetFileFromPathAsync(logPath));
		}
		catch (Exception ex)
		{
			await new ContentDialog
			{
				Title = "Could not open log file",
				Content = ex.Message + "\n\nLog file location:\n" + logPath,
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
		}
	}

	private static string GetPdfBaseFontName(string P_0, bool P_1)
	{
		string text = (P_0 ?? string.Empty).ToLowerInvariant();
		if (text.Contains("helvetica"))
		{
			if (!P_1)
			{
				return "Helvetica";
			}
			return "Helvetica-Bold";
		}
		if (text.Contains("times"))
		{
			if (!P_1)
			{
				return "Times-Roman";
			}
			return "Times-Bold";
		}
		if (text.Contains("courier"))
		{
			if (!P_1)
			{
				return "Courier";
			}
			return "Courier-Bold";
		}
		if (!P_1)
		{
			return "Arial";
		}
		return "Arial,Bold";
	}

	private static bool IsStandardPdf14Font(string P_0)
	{
		string text = (P_0 ?? string.Empty).ToLowerInvariant();
		if (!text.Contains("helvetica") && !text.Contains("times"))
		{
			return text.Contains("courier");
		}
		return true;
	}

	private static PdfDictionary GetOrCreateFontResource(PdfSharp.Pdf.PdfDocument P_0, string P_1, bool P_2 = false)
	{
		Dictionary<string, PdfDictionary> value = _fontResourceCache.GetValue(P_0, (PdfSharp.Pdf.PdfDocument _) => new Dictionary<string, PdfDictionary>());
		string text = $"{P_1}|{P_2}";
		if (value.TryGetValue(text, out var value2))
		{
			return value2;
		}
		PdfDictionary pdfDictionary;
		if (IsStandardPdf14Font(P_1))
		{
			pdfDictionary = new PdfDictionary(P_0);
			pdfDictionary.Elements.SetName("/Type", "/Font");
			pdfDictionary.Elements.SetName("/Subtype", "/Type1");
			pdfDictionary.Elements.SetName("/BaseFont", "/" + GetPdfBaseFontName(P_1, P_2));
			pdfDictionary.Elements.SetName("/Encoding", "/WinAnsiEncoding");
			P_0.Internals.AddObject(pdfDictionary);
		}
		else
		{
			pdfDictionary = EmbedNonStandardFont(P_0, P_1, P_2);
		}
		value[text] = pdfDictionary;
		return pdfDictionary;
	}

	private static void EnsureAcroFormDefaultResources(PdfSharp.Pdf.PdfDocument P_0, string P_1)
	{
		PdfDictionary orCreateFontResource = GetOrCreateFontResource(P_0, P_1);
		PdfDictionary orCreateFontResource2 = GetOrCreateFontResource(P_0, P_1, true);
		PdfDictionary pdfDictionary = new PdfDictionary(P_0);
		pdfDictionary.Elements["/" + GetPdfBaseFontName(P_1, false)] = orCreateFontResource.Reference;
		pdfDictionary.Elements["/" + GetPdfBaseFontName(P_1, true)] = orCreateFontResource2.Reference;
		PdfDictionary pdfDictionary2 = new PdfDictionary(P_0);
		pdfDictionary2.Elements["/Font"] = pdfDictionary;
		PdfDictionary pdfDictionary3 = new PdfDictionary(P_0);
		pdfDictionary3.Elements["/DR"] = pdfDictionary2;
		P_0.Internals.Catalog.Elements["/AcroForm"] = pdfDictionary3;
	}

	private static PdfDictionary EmbedNonStandardFont(PdfSharp.Pdf.PdfDocument P_0, string P_1, bool P_2)
	{
		PdfDictionary pdfDictionary = null;
		PdfSharp.Pdf.PdfPage pdfPage = P_0.AddPage();
		try
		{
			using (XGraphics xGraphics = XGraphics.FromPdfPage(pdfPage))
			{
				XPdfFontOptions pdfOptions = new XPdfFontOptions(PdfFontEncoding.WinAnsi);
				XFont font = new XFont(P_1, 10.0, P_2 ? XFontStyleEx.Bold : XFontStyleEx.Regular, pdfOptions);
				StringBuilder stringBuilder = new StringBuilder();
				for (int i = 32; i <= 255; i++)
				{
					if (i < 128 || i > 159)
					{
						stringBuilder.Append((char)i);
					}
				}
				char[] array = new char[8] { '…', '‘', '’', '“', '”', '•', '–', '—' };
				foreach (char c in array)
				{
					stringBuilder.Append(c);
				}
				xGraphics.DrawString(stringBuilder.ToString(), font, XBrushes.Black, 0.0, 10.0);
			}
			PdfDictionary pdfDictionary2 = (pdfPage.Elements["/Resources"] as PdfDictionary) ?? ((pdfPage.Elements["/Resources"] as PdfReference)?.Value as PdfDictionary);
			PdfDictionary pdfDictionary3 = (pdfDictionary2?.Elements["/Font"] as PdfDictionary) ?? ((pdfDictionary2?.Elements["/Font"] as PdfReference)?.Value as PdfDictionary);
			if (pdfDictionary3 != null)
			{
				foreach (string key in pdfDictionary3.Elements.Keys)
				{
					PdfItem pdfItem = pdfDictionary3.Elements[key];
					pdfDictionary = ((pdfItem as PdfReference)?.Value as PdfDictionary) ?? (pdfItem as PdfDictionary);
					if (pdfDictionary != null)
					{
						break;
					}
				}
			}
		}
		finally
		{
			P_0.Pages.Remove(pdfPage);
		}
		if (pdfDictionary == null)
		{
			pdfDictionary = new PdfDictionary(P_0);
			pdfDictionary.Elements.SetName("/Type", "/Font");
			pdfDictionary.Elements.SetName("/Subtype", "/Type1");
			pdfDictionary.Elements.SetName("/BaseFont", "/" + GetPdfBaseFontName(P_1, P_2));
			pdfDictionary.Elements.SetName("/Encoding", "/WinAnsiEncoding");
			P_0.Internals.AddObject(pdfDictionary);
		}
		return pdfDictionary;
	}

	private static string EscapePdfLiteralString(string P_0)
	{
		return P_0.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
	}

	private static byte[] EncodeWinAnsi(string P_0)
	{
		byte[] array = new byte[P_0.Length];
		for (int i = 0; i < P_0.Length; i++)
		{
			char c = P_0[i];
			byte[] array2 = array;
			int num = i;
			array2[num] = c switch
			{
				'–' => (byte)150, 
				'—' => (byte)151, 
				'‘' => (byte)145, 
				'’' => (byte)146, 
				'“' => (byte)147, 
				'”' => (byte)148, 
				'•' => (byte)149, 
				'…' => (byte)133, 
				_ => (byte)((c > 'ÿ') ? 63 : ((byte)c)), 
			};
		}
		return array;
	}

	private static PdfDictionary BuildFreeTextAppearance(PdfSharp.Pdf.PdfDocument P_0, string P_1, double P_2, double P_3, double P_4, int P_5, string P_6, bool P_7 = false, (double r, double g, double b)? P_8 = null, (double a, double b, double c, double d, double e, double f)? P_9 = null, double? P_10 = null, double? P_11 = null, bool P_12 = false, bool P_13 = false)
	{
		PdfDictionary orCreateFontResource = GetOrCreateFontResource(P_0, P_6, P_7);
		double num = P_10 ?? P_2;
		double num2 = P_11 ?? P_3;
		XFont xFont = new XFont(P_6, P_4, P_7 ? XFontStyleEx.Bold : XFontStyleEx.Regular);
		using PdfSharp.Pdf.PdfDocument pdfDocument = new PdfSharp.Pdf.PdfDocument();
		using XGraphics xGraphics = XGraphics.FromPdfPage(pdfDocument.AddPage());
		List<string> list = ((!P_12) ? new List<string> { P_1 } : WrapTextToWidth(P_1, num - 4.0, xGraphics, xFont));
		double num3 = P_4 * 1.15;
		StringBuilder stringBuilder = new StringBuilder();
		StringBuilder stringBuilder2;
		StringBuilder.AppendInterpolatedStringHandler appendInterpolatedStringHandler;
		if (P_8.HasValue)
		{
			stringBuilder2 = stringBuilder;
			StringBuilder stringBuilder3 = stringBuilder2;
			appendInterpolatedStringHandler = new StringBuilder.AppendInterpolatedStringHandler(6, 3, stringBuilder2);
			appendInterpolatedStringHandler.AppendFormatted(P_8.Value.r, "0.###");
			appendInterpolatedStringHandler.AppendLiteral(" ");
			appendInterpolatedStringHandler.AppendFormatted(P_8.Value.g, "0.###");
			appendInterpolatedStringHandler.AppendLiteral(" ");
			appendInterpolatedStringHandler.AppendFormatted(P_8.Value.b, "0.###");
			appendInterpolatedStringHandler.AppendLiteral(" rg\n");
			stringBuilder3.Append(ref appendInterpolatedStringHandler);
			stringBuilder2 = stringBuilder;
			StringBuilder stringBuilder4 = stringBuilder2;
			appendInterpolatedStringHandler = new StringBuilder.AppendInterpolatedStringHandler(9, 2, stringBuilder2);
			appendInterpolatedStringHandler.AppendLiteral("0 0 ");
			appendInterpolatedStringHandler.AppendFormatted(P_2, "0.##");
			appendInterpolatedStringHandler.AppendLiteral(" ");
			appendInterpolatedStringHandler.AppendFormatted(P_3, "0.##");
			appendInterpolatedStringHandler.AppendLiteral(" re\n");
			stringBuilder4.Append(ref appendInterpolatedStringHandler);
			stringBuilder.Append("f\n");
		}
		if (P_9.HasValue)
		{
			(double, double, double, double, double, double) value = P_9.Value;
			stringBuilder.Append(FormattableString.Invariant($"q\n{value.Item1:0.######} {value.Item2:0.######} {value.Item3:0.######} {value.Item4:0.######} {value.Item5:0.##} {value.Item6:0.##} cm\n"));
		}
		stringBuilder.Append("BT\n");
		stringBuilder2 = stringBuilder;
		StringBuilder stringBuilder5 = stringBuilder2;
		appendInterpolatedStringHandler = new StringBuilder.AppendInterpolatedStringHandler(8, 1, stringBuilder2);
		appendInterpolatedStringHandler.AppendLiteral("/F1 ");
		appendInterpolatedStringHandler.AppendFormatted(P_4, "0.##");
		appendInterpolatedStringHandler.AppendLiteral(" Tf\n");
		stringBuilder5.Append(ref appendInterpolatedStringHandler);
		stringBuilder.Append("0 0 0 rg\n");
		double num4 = (P_13 ? ((num2 + (double)(list.Count - 1) * num3 - GetCapHeightRatio(xFont) * P_4) / 2.0) : Math.Max(0.0, num2 - P_4 - 2.0));
		double num5 = 0.0;
		double num6 = 0.0;
		for (int i = 0; i < list.Count; i++)
		{
			string text = list[i];
			double num7 = 2.0;
			if (P_5 != 0)
			{
				double width = xGraphics.MeasureString(text, xFont).Width;
				switch (P_5)
				{
				case 1:
					num7 = Math.Max(2.0, (num - width) / 2.0);
					break;
				case 2:
					num7 = Math.Max(2.0, num - width - 2.0);
					break;
				}
			}
			double num8 = ((i == 0) ? num4 : (num4 - (double)i * num3));
			stringBuilder2 = stringBuilder;
			StringBuilder stringBuilder6 = stringBuilder2;
			appendInterpolatedStringHandler = new StringBuilder.AppendInterpolatedStringHandler(5, 2, stringBuilder2);
			appendInterpolatedStringHandler.AppendFormatted(num7 - num5, "0.##");
			appendInterpolatedStringHandler.AppendLiteral(" ");
			appendInterpolatedStringHandler.AppendFormatted(num8 - num6, "0.##");
			appendInterpolatedStringHandler.AppendLiteral(" Td\n");
			stringBuilder6.Append(ref appendInterpolatedStringHandler);
			num5 = num7;
			num6 = num8;
			stringBuilder.Append('(').Append(EscapePdfLiteralString(text)).Append(") Tj\n");
		}
		stringBuilder.Append("ET\n");
		if (P_9.HasValue)
		{
			stringBuilder.Append("Q\n");
		}
		PdfDictionary pdfDictionary = new PdfDictionary(P_0);
		pdfDictionary.Elements.SetName("/Type", "/XObject");
		pdfDictionary.Elements.SetName("/Subtype", "/Form");
		pdfDictionary.Elements.SetInteger("/FormType", 1);
		PdfArray pdfArray = new PdfArray(P_0);
		pdfArray.Elements.Add(new PdfReal(0.0));
		pdfArray.Elements.Add(new PdfReal(0.0));
		pdfArray.Elements.Add(new PdfReal(P_2));
		pdfArray.Elements.Add(new PdfReal(P_3));
		pdfDictionary.Elements["/BBox"] = pdfArray;
		PdfDictionary pdfDictionary2 = new PdfDictionary(P_0);
		PdfDictionary pdfDictionary3 = new PdfDictionary(P_0);
		pdfDictionary3.Elements["/F1"] = orCreateFontResource.Reference;
		pdfDictionary2.Elements["/Font"] = pdfDictionary3;
		pdfDictionary.Elements["/Resources"] = pdfDictionary2;
		pdfDictionary.CreateStream(EncodeWinAnsi(stringBuilder.ToString()));
		P_0.Internals.AddObject(pdfDictionary);
		return pdfDictionary;
	}

	private static List<string> WrapTextToWidth(string P_0, double P_1, XGraphics P_2, XFont P_3)
	{
		string[] array = P_0.Split(' ');
		List<string> list = new List<string>();
		StringBuilder stringBuilder = new StringBuilder();
		string[] array2 = array;
		foreach (string text in array2)
		{
			string text2 = ((stringBuilder.Length == 0) ? text : (stringBuilder?.ToString() + " " + text));
			if (stringBuilder.Length > 0 && P_2.MeasureString(text2, P_3).Width > P_1)
			{
				list.Add(stringBuilder.ToString());
				stringBuilder.Clear();
				stringBuilder.Append(text);
			}
			else
			{
				stringBuilder.Clear();
				stringBuilder.Append(text2);
			}
		}
		if (stringBuilder.Length > 0)
		{
			list.Add(stringBuilder.ToString());
		}
		if (list.Count <= 0)
		{
			return new List<string> { P_0 };
		}
		return list;
	}

	private static double GetCapHeightRatio(XFont P_0)
	{
		XFontMetrics metrics = P_0.Metrics;
		if (metrics.CapHeight <= 0 || metrics.UnitsPerEm <= 0)
		{
			return 0.72;
		}
		return (double)metrics.CapHeight / (double)metrics.UnitsPerEm;
	}

	private static double FitFontSizeToBox(string P_0, string P_1, bool P_2, double P_3, double P_4, double P_5, double P_6 = 6.0)
	{
		double num = P_3 - 4.0;
		XFontStyleEx style = (P_2 ? XFontStyleEx.Bold : XFontStyleEx.Regular);
		using PdfSharp.Pdf.PdfDocument pdfDocument = new PdfSharp.Pdf.PdfDocument();
		using XGraphics xGraphics = XGraphics.FromPdfPage(pdfDocument.AddPage());
		double capHeightRatio = GetCapHeightRatio(new XFont(P_1, 10.0, style));
		for (double num2 = Math.Floor(P_4 * P_5 / capHeightRatio * 4.0) / 4.0; num2 >= P_6; num2 -= 0.25)
		{
			XFont xFont = new XFont(P_1, num2, style);
			List<string> list = WrapTextToWidth(P_0, num, xGraphics, xFont);
			bool flag = true;
			foreach (string item in list)
			{
				if (xGraphics.MeasureString(item, xFont).Width > num)
				{
					flag = false;
					break;
				}
			}
			if (flag && (double)(list.Count - 1) * num2 * 1.15 + capHeightRatio * num2 <= P_4 * P_5)
			{
				return num2;
			}
		}
		return P_6;
	}

	private static string GetCssTextAlign(int P_0)
	{
		return P_0 switch
		{
			1 => "center", 
			2 => "right", 
			_ => "left", 
		};
	}

	private static string BuildDefaultStyleCss(string P_0, double P_1, int P_2, bool P_3)
	{
		string text = (P_3 ? " font-weight:bold;" : "");
		return $"font: '{P_0}' {P_1:0.##}pt;{text} text-align:{GetCssTextAlign(P_2)}; line-height:{P_1 * 1.15:0.##}pt; color:#000000";
	}

	private static string EscapeXmlText(string P_0)
	{
		return P_0.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
	}

	private static PdfDictionary? AddFreeTextAnnotation(PdfSharp.Pdf.PdfDocument P_0, PdfSharp.Pdf.PdfPage P_1, string? P_2, double P_3, double P_4, double P_5, double P_6, double P_7, string P_8, int P_9 = 0, bool P_10 = false, (double r, double g, double b)? P_11 = null, double? P_12 = null, (double a, double b, double c, double d, double e, double f)? P_13 = null, double? P_14 = null, double? P_15 = null, bool P_16 = false, bool P_17 = false)
	{
		if (string.IsNullOrEmpty(P_2))
		{
			return null;
		}
		PdfDictionary pdfDictionary = new PdfDictionary(P_0);
		pdfDictionary.Elements.SetName("/Type", "/Annot");
		pdfDictionary.Elements.SetName("/Subtype", "/FreeText");
		pdfDictionary.Elements.SetString("/Contents", P_2);
		pdfDictionary.Elements.SetString("/DA", $"/{GetPdfBaseFontName(P_8, P_10)} {P_7:0.##} Tf 0 0 0 rg");
		pdfDictionary.Elements.SetInteger("/F", 132);
		pdfDictionary.Elements.SetInteger("/Q", P_9);
		string text = BuildDefaultStyleCss(P_8, P_7, P_9, P_10);
		pdfDictionary.Elements.SetString("/DS", text);
		pdfDictionary.Elements.SetString("/RC", "<?xml version=\"1.0\"?><body xmlns:xfa=\"http://www.xfa.org/schema/xfa-data/1.0/\" xfa:contentType=\"text/html\" xfa:APIVersion=\"BluebeamPDFRevu:2018\" xfa:spec=\"2.2.0\" style=\"" + text + "\" xmlns=\"http://www.w3.org/1999/xhtml\"><p>" + EscapeXmlText(P_2) + "</p></body>");
		PdfArray pdfArray = new PdfArray(P_0);
		pdfArray.Elements.Add(new PdfInteger(0));
		pdfArray.Elements.Add(new PdfInteger(0));
		pdfArray.Elements.Add(new PdfInteger(0));
		pdfDictionary.Elements["/Border"] = pdfArray;
		if (P_11.HasValue)
		{
			PdfArray pdfArray2 = new PdfArray(P_0);
			pdfArray2.Elements.Add(new PdfReal(P_11.Value.r));
			pdfArray2.Elements.Add(new PdfReal(P_11.Value.g));
			pdfArray2.Elements.Add(new PdfReal(P_11.Value.b));
			pdfDictionary.Elements["/IC"] = pdfArray2;
		}
		double num = P_12 ?? P_1.Height.Point;
		double x = P_1.MediaBox.X1;
		double y = P_1.MediaBox.Y1;
		PdfArray pdfArray3 = new PdfArray(P_0);
		pdfArray3.Elements.Add(new PdfReal(P_3 + x));
		pdfArray3.Elements.Add(new PdfReal(num - P_4 - P_6 + y));
		pdfArray3.Elements.Add(new PdfReal(P_3 + P_5 + x));
		pdfArray3.Elements.Add(new PdfReal(num - P_4 + y));
		pdfDictionary.Elements["/Rect"] = pdfArray3;
		try
		{
			PdfDictionary pdfDictionary2 = BuildFreeTextAppearance(P_0, P_2, P_5, P_6, P_7, P_9, P_8, P_10, P_11, P_13, P_14, P_15, P_16, P_17);
			PdfDictionary pdfDictionary3 = new PdfDictionary(P_0);
			pdfDictionary3.Elements["/N"] = pdfDictionary2.Reference;
			pdfDictionary.Elements["/AP"] = pdfDictionary3;
		}
		catch
		{
		}
		P_0.Internals.AddObject(pdfDictionary);
		PdfArray pdfArray4 = P_1.Elements["/Annots"] as PdfArray;
		if (pdfArray4 == null)
		{
			pdfArray4 = new PdfArray(P_0);
			P_1.Elements["/Annots"] = pdfArray4;
		}
		pdfArray4.Elements.Add((PdfItem)(((object)pdfDictionary.Reference) ?? ((object)pdfDictionary)));
		return pdfDictionary;
	}

	private static PdfDictionary AddLinkAnnotation(PdfSharp.Pdf.PdfDocument P_0, PdfSharp.Pdf.PdfPage P_1, double P_2, double P_3, double P_4, double P_5)
	{
		PdfDictionary pdfDictionary = new PdfDictionary(P_0);
		pdfDictionary.Elements.SetName("/Type", "/Annot");
		pdfDictionary.Elements.SetName("/Subtype", "/Link");
		pdfDictionary.Elements.SetInteger("/F", 4);
		PdfArray pdfArray = new PdfArray(P_0);
		pdfArray.Elements.Add(new PdfReal(1.0));
		pdfArray.Elements.Add(new PdfReal(0.647));
		pdfArray.Elements.Add(new PdfReal(0.0));
		pdfDictionary.Elements["/C"] = pdfArray;
		PdfDictionary pdfDictionary2 = new PdfDictionary(P_0);
		pdfDictionary2.Elements.SetName("/Type", "/Border");
		pdfDictionary2.Elements.SetReal("/W", 1.0);
		pdfDictionary2.Elements.SetName("/S", "/U");
		pdfDictionary.Elements["/BS"] = pdfDictionary2;
		PdfArray pdfArray2 = new PdfArray(P_0);
		pdfArray2.Elements.Add(new PdfInteger(0));
		pdfArray2.Elements.Add(new PdfInteger(0));
		pdfArray2.Elements.Add(new PdfInteger(0));
		pdfDictionary.Elements["/Border"] = pdfArray2;
		double point = P_1.Height.Point;
		double x = P_1.MediaBox.X1;
		double y = P_1.MediaBox.Y1;
		PdfArray pdfArray3 = new PdfArray(P_0);
		pdfArray3.Elements.Add(new PdfReal(P_2 + x));
		pdfArray3.Elements.Add(new PdfReal(point - P_3 - P_5 + y));
		pdfArray3.Elements.Add(new PdfReal(P_2 + P_4 + x));
		pdfArray3.Elements.Add(new PdfReal(point - P_3 + y));
		pdfDictionary.Elements["/Rect"] = pdfArray3;
		P_0.Internals.AddObject(pdfDictionary);
		PdfArray pdfArray4 = P_1.Elements["/Annots"] as PdfArray;
		if (pdfArray4 == null)
		{
			pdfArray4 = new PdfArray(P_0);
			P_1.Elements["/Annots"] = pdfArray4;
		}
		pdfArray4.Elements.Add((PdfItem)(((object)pdfDictionary.Reference) ?? ((object)pdfDictionary)));
		return pdfDictionary;
	}

	private static List<(string Key, List<DrawingIndexItem> Items)> ChunkConsecutiveByDiscipline(List<DrawingIndexItem> P_0)
	{
		List<(string, List<DrawingIndexItem>)> list = new List<(string, List<DrawingIndexItem>)>();
		string text = null;
		List<DrawingIndexItem> list2 = null;
		foreach (DrawingIndexItem item in P_0)
		{
			string text2 = (string.IsNullOrWhiteSpace(item.Discipline) ? "[DISCIPLINE NOT FOUND]" : item.Discipline);
			if (list2 == null || !string.Equals(text2, text, StringComparison.OrdinalIgnoreCase))
			{
				text = text2;
				list2 = new List<DrawingIndexItem>();
				list.Add((text, list2));
			}
			list2.Add(item);
		}
		return list;
	}

	private static TablePagination BuildTablePagination(IEnumerable<DrawingIndexItem> P_0, double P_1, double P_2, double P_3, double P_4, double[] P_5, int P_6, int P_7, string P_8, double? P_9 = null)
	{
		List<DrawingIndexItem> list = P_0.ToList();
		double num = P_2 - P_3 - P_4;
		int num2 = 20;
		string logFilePath = GetLogFilePath("export_log.txt");
		bool flag = false;
		int num3 = 0;
		try
		{
			using PdfSharp.Pdf.PdfDocument pdfDocument = new PdfSharp.Pdf.PdfDocument();
			PdfSharp.Pdf.PdfPage pdfPage = pdfDocument.AddPage();
			pdfPage.Width = XUnit.FromPoint(P_1);
			pdfPage.Height = XUnit.FromPoint(P_2);
			using XGraphics xGraphics = XGraphics.FromPdfPage(pdfPage);
			XFont xFont = new XFont(P_8, P_6);
			XFont xFont2 = new XFont(P_8, P_7);
			int num4 = P_5.Length;
			double num5 = MeasureWrappedHeight(xGraphics, "SEQ. NO.", xFont, P_5[0] - 8.0);
			double num6 = ((num4 > 1) ? MeasureWrappedHeight(xGraphics, "DWG. NO.", xFont, P_5[1] - 8.0) : 0.0);
			double num7 = ((num4 > 2) ? MeasureWrappedHeight(xGraphics, "DESCRIPTION", xFont, P_5[2] - 8.0) : 0.0);
			double num8 = Math.Max(Math.Max(num5, num6), num7);
			List<double> list2 = new List<double>();
			foreach (DrawingIndexItem item in list)
			{
				double num9 = MeasureWrappedHeight(xGraphics, item.SeqNo.ToString(), xFont2, P_5[0] - 8.0);
				double num10 = MeasureWrappedHeight(xGraphics, item.DrawingNumber ?? string.Empty, xFont2, P_5[1] - 8.0);
				double num11 = MeasureWrappedHeight(xGraphics, item.Title ?? string.Empty, xFont2, P_5[2] - 8.0);
				double num12 = Math.Max(Math.Max(num9, num10), num11);
				num12 += 6.0;
				list2.Add(num12);
			}
			double num13 = num8 + 6.0;
			int num14 = 0;
			for (int i = 0; i < list2.Count && !(num13 + list2[i] > num); i++)
			{
				num13 += list2[i];
				num14++;
			}
			num2 = ((num14 < 5) ? 20 : num14);
			num3 = num14;
			flag = true;
		}
		catch (Exception ex)
		{
			num2 = 20;
			try
			{
				File.AppendAllText(logFilePath, $"{DateTime.Now}: BuildTablePagination measurement failed: {ex}\n");
			}
			catch
			{
			}
		}
		if (!flag && list.Count > num2)
		{
			try
			{
				File.AppendAllText(logFilePath, $"{DateTime.Now}: BuildTablePagination used fallback rowsPerColumn={num2} for totalRows={list.Count}\n");
			}
			catch
			{
			}
		}
		else if (flag && num3 != num2 && list.Count > num2)
		{
			try
			{
				File.AppendAllText(logFilePath, $"{DateTime.Now}: BuildTablePagination measured rowsPerColumn={num3} totalRows={list.Count}\n");
			}
			catch
			{
			}
		}
		List<(string, List<DrawingIndexItem>)> list3 = ChunkConsecutiveByDiscipline(list);
		int num15 = 1;
		foreach (var item2 in list3)
		{
			foreach (DrawingIndexItem item3 in item2.Item2)
			{
				item3.SeqNo = num15++;
			}
		}
		double num16 = P_7;
		double num17 = P_6;
		List<List<List<(bool, string, DrawingIndexItem)>>> list4 = new List<List<List<(bool, string, DrawingIndexItem)>>>();
		list4.Add(new List<List<(bool, string, DrawingIndexItem)>>
		{
			new List<(bool, string, DrawingIndexItem)>(),
			new List<(bool, string, DrawingIndexItem)>()
		});
		int num18 = 0;
		int num19 = 0;
		XFont xFont3 = new XFont(P_8, (int)num17);
		XFont xFont4 = new XFont(P_8, (int)num16);
		double height = xFont3.GetHeight();
		double num20 = xFont4.GetHeight() + 4.0;
		double num21 = P_2 - P_3 - (P_9 ?? P_4);
		double[] array = new double[2]
		{
			num - (height + 6.0),
			num21 - (height + 6.0)
		};
		double num22 = array[num19];
		double num23 = height + 12.0;
		HashSet<string> hashSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		foreach (var item4 in list3)
		{
			string text;
			if (hashSet.Add(item4.Item1))
			{
				(text, _) = item4;
			}
			else
			{
				text = item4.Item1 + " (CONTINUED)";
			}
			string text2 = text;
			string text3 = item4.Item1 + " (CONTINUED)";
			bool flag2 = item4.Item2.Count > 0 && num23 + num20 > num22;
			if ((num23 > num22) | flag2)
			{
				if (num19 == 0)
				{
					num19 = 1;
				}
				else
				{
					list4.Add(new List<List<(bool, string, DrawingIndexItem)>>
					{
						new List<(bool, string, DrawingIndexItem)>(),
						new List<(bool, string, DrawingIndexItem)>()
					});
					num18++;
					num19 = 0;
				}
				num22 = array[num19];
			}
			list4[num18][num19].Add((true, text2, item4.Item2.FirstOrDefault(r => r.DisciplineUncertain)));
			num22 -= num23;
			foreach (DrawingIndexItem item5 in item4.Item2)
			{
				if (num20 > num22)
				{
					if (num19 == 0)
					{
						num19 = 1;
					}
					else
					{
						list4.Add(new List<List<(bool, string, DrawingIndexItem)>>
						{
							new List<(bool, string, DrawingIndexItem)>(),
							new List<(bool, string, DrawingIndexItem)>()
						});
						num18++;
						num19 = 0;
					}
					num22 = array[num19];
					list4[num18][num19].Add((true, text3, item4.Item2.FirstOrDefault(r => r.DisciplineUncertain)));
					num22 -= num23;
				}
				list4[num18][num19].Add((false, null, item5));
				num22 -= num20;
			}
		}
		if (list4.Count > 1)
		{
			try
			{
				File.AppendAllText(logFilePath, $"{DateTime.Now}: BuildTablePagination content spans {list4.Count} pages; caller will add extra output pages.\n");
			}
			catch
			{
			}
		}
		return new TablePagination
		{
			Pages = list4,
			HeaderHeight = height,
			RowHeight = num20,
			HeaderFontSize = num17,
			CellFontSize = num16
		};
	}

	private static double GetDrawingIndexTitleReservedHeight()
	{
		return 36.0;
	}

	private static void PlaceTableColumnAnnotations(PdfSharp.Pdf.PdfDocument P_0, PdfSharp.Pdf.PdfPage P_1, List<List<(bool isHeader, string headerText, DrawingIndexItem? row)>> P_2, double P_3, double P_4, double[] P_5, TablePagination P_6, string P_7, Dictionary<DrawingIndexItem, (PdfSharp.Pdf.PdfPage page, PdfDictionary annot)>? P_8 = null, bool P_9 = false, double P_10 = 12.0)
	{
		double[] array = new double[2]
		{
			P_3,
			P_3 + P_5.Sum() + P_10
		};
		double headerHeight = P_6.HeaderHeight;
		double rowHeight = P_6.RowHeight;
		double headerFontSize = P_6.HeaderFontSize;
		double cellFontSize = P_6.CellFontSize;
		using PdfSharp.Pdf.PdfDocument pdfDocument = new PdfSharp.Pdf.PdfDocument();
		PdfSharp.Pdf.PdfPage page = pdfDocument.AddPage();
		XGraphics measGfx = XGraphics.FromPdfPage(page);
		try
		{
			XFont xFont = new XFont(P_7, (int)cellFontSize);
			XFont xFont2 = new XFont(P_7, (int)cellFontSize, XFontStyleEx.Bold);
			XFont xFont3 = new XFont(P_7, (int)headerFontSize, XFontStyleEx.Bold);
			double num = 20.0;
			double num2 = num + 4.0;
			double num3 = P_5.Sum() * 2.0 + P_10;
			AddFreeTextAnnotation(P_0, P_1, "DRAWING INDEX", P_3, P_4 - GetDrawingIndexTitleReservedHeight(), num3, num2, num, P_7, 1, true);
			for (int i = 0; i < 2; i++)
			{
				if (P_2[i].Count == 0)
				{
					continue;
				}
				double num4 = array[i];
				double num5 = P_4;
				AddFreeTextAnnotation(P_0, P_1, "SEQ. NO.", num4, num5, P_5[0], headerHeight, headerFontSize, P_7, 0, true);
				AddFreeTextAnnotation(P_0, P_1, "DWG. NO.", num4 + P_5[0], num5, P_5[1], headerHeight, headerFontSize, P_7, 0, true);
				AddFreeTextAnnotation(P_0, P_1, "DESCRIPTION", num4 + P_5[0] + P_5[1], num5, P_5[2], headerHeight, headerFontSize, P_7, 0, true);
				num5 += headerHeight + 6.0;
				foreach (var item in P_2[i])
				{
					if (item.isHeader)
					{
						num5 += 6.0;
						bool flag = item.headerText.StartsWith("[DISCIPLINE NOT FOUND]", StringComparison.Ordinal);
						bool headerUncertain = !flag && item.row != null && item.row.DisciplineUncertain;
						double num6 = MeasureBoxWidth(item.headerText, xFont3, P_5.Sum());
						AddFreeTextAnnotation(P_0, P_1, item.headerText, num4, num5, num6, headerHeight + 4.0, headerFontSize, P_7, 0, true, flag ? new (double, double, double)?(NotFoundHighlightColor) : (headerUncertain ? new (double, double, double)?(UncertainHighlightColor) : (((double, double, double)?)null)));
						num5 += headerHeight + 6.0;
						continue;
					}
					double num7 = MeasureBoxWidth(item.row.SeqNo.ToString(), xFont, P_5[0]);
					AddFreeTextAnnotation(P_0, P_1, item.row.SeqNo.ToString(), num4, num5, num7, rowHeight, cellFontSize, P_7);
					bool idNotFound = item.row.DrawingNumber == "[NOT FOUND]";
					bool idUncertain = !idNotFound && item.row.DrawingNumberUncertain;
					bool flag2 = idNotFound || idUncertain;
					double num8 = MeasureBoxWidth(item.row.DrawingNumber, flag2 ? xFont2 : xFont, P_5[1]);
					PdfDictionary pdfDictionary = AddFreeTextAnnotation(P_0, P_1, item.row.DrawingNumber, num4 + P_5[0], num5, num8, rowHeight, cellFontSize, P_7, 0, flag2, idNotFound ? new (double, double, double)?(NotFoundHighlightColor) : (idUncertain ? new (double, double, double)?(UncertainHighlightColor) : (((double, double, double)?)null)));
					if (P_9 && pdfDictionary != null)
					{
						double num9 = MeasureLinkWidth(item.row.DrawingNumber, flag2 ? xFont2 : xFont, num8);
						PdfDictionary pdfDictionary2 = AddLinkAnnotation(P_0, P_1, num4 + P_5[0], num5, num9, rowHeight);
						if (P_8 != null)
						{
							P_8[item.row] = (P_1, pdfDictionary2);
						}
					}
					bool titleNotFound = item.row.Title == "[DESCRIPTION NOT FOUND]";
					bool titleUncertain = !titleNotFound && item.row.TitleUncertain;
					bool flag3 = titleNotFound || titleUncertain;
					double num10 = MeasureBoxWidth(item.row.Title, flag3 ? xFont2 : xFont, P_5[2]);
					AddFreeTextAnnotation(P_0, P_1, item.row.Title, num4 + P_5[0] + P_5[1], num5, num10, rowHeight, cellFontSize, P_7, 0, flag3, titleNotFound ? new (double, double, double)?(NotFoundHighlightColor) : (titleUncertain ? new (double, double, double)?(UncertainHighlightColor) : (((double, double, double)?)null)));
					num5 += rowHeight;
				}
			}
		}
		finally
		{
			if (measGfx != null)
			{
				((IDisposable)measGfx).Dispose();
			}
		}
		double MeasureBoxWidth(string? text, XFont font, double val)
		{
			if (string.IsNullOrEmpty(text))
			{
				return Math.Min(val, 10.0);
			}
			double num11 = measGfx.MeasureString(text, font).Width + 24.0;
			return Math.Min(val, Math.Max(num11, 10.0));
		}
		double MeasureLinkWidth(string? text, XFont font, double val)
		{
			if (string.IsNullOrEmpty(text))
			{
				return Math.Min(val, 10.0);
			}
			double num11 = measGfx.MeasureString(text, font).Width + 4.0;
			return Math.Min(val, Math.Max(num11, 10.0));
		}
	}

	private static double GetTableAreaWidth(double P_0, double P_1, bool P_2, double? P_3 = null)
	{
		if (P_2)
		{
			return P_0;
		}
		if (P_3.HasValue)
		{
			return Math.Max(P_0 * 0.5, P_3.Value - 15.0);
		}
		return P_0 * 0.8;
	}

	private static readonly Dictionary<(string Text, string Font, double Size, XFontStyleEx Style, double Width), double> WrappedHeightCache = new Dictionary<(string Text, string Font, double Size, XFontStyleEx Style, double Width), double>();

	private static double MeasureWrappedHeight(XGraphics P_0, string P_1, XFont P_2, double P_3)
	{
		if (string.IsNullOrEmpty(P_1))
		{
			return 0.0;
		}
		var key = (P_1, P_2.Name, P_2.Size, P_2.Style, P_3);
		if (WrappedHeightCache.TryGetValue(key, out double cached))
		{
			return cached;
		}
		if (WrappedHeightCache.Count > 50000)
		{
			WrappedHeightCache.Clear();
		}
		double result = MeasureWrappedHeightUncached(P_0, P_1, P_2, P_3);
		WrappedHeightCache[key] = result;
		return result;
	}

	private static double MeasureWrappedHeightUncached(XGraphics P_0, string P_1, XFont P_2, double P_3)
	{
		if (string.IsNullOrEmpty(P_1))
		{
			return 0.0;
		}
		string[] array = P_1.Split(' ');
		StringBuilder stringBuilder = new StringBuilder();
		double height = P_2.GetHeight();
		double num = 0.0;
		string[] array2 = array;
		foreach (string text in array2)
		{
			string text2 = ((stringBuilder.Length == 0) ? text : (stringBuilder?.ToString() + " " + text));
			if (P_0.MeasureString(text2, P_2).Width > P_3 && stringBuilder.Length > 0)
			{
				num += height;
				stringBuilder.Clear();
				stringBuilder.Append(text);
				continue;
			}
			if (stringBuilder.Length > 0)
			{
				stringBuilder.Append(' ');
			}
			stringBuilder.Append(text);
		}
		if (stringBuilder.Length > 0)
		{
			num += height;
		}
		return num;
	}

	private static void RunWithFileLockRetry(Action P_0, string P_1, string P_2, int P_3 = 5, int P_4 = 1000)
	{
		for (int i = 1; i <= P_3; i++)
		{
			try
			{
				P_0();
				break;
			}
			catch (IOException ex) when (i < P_3)
			{
				try
				{
					File.AppendAllText(P_1, $"{P_2} failed (attempt {i}/{P_3}), likely a transient lock (e.g. OneDrive sync) -- retrying in {P_4}ms: {ex.Message}\n");
				}
				catch
				{
				}
				Thread.Sleep(P_4);
			}
		}
	}

	private static bool TryRepairMissingStreamLengths(string P_0, string P_1)
	{
		try
		{
			byte[] array = File.ReadAllBytes(P_0);
			Encoding latin = Encoding.Latin1;
			string text = latin.GetString(array);
			if (Regex.IsMatch(text, "/Type\\s*/ObjStm") || Regex.IsMatch(text, "/Type\\s*/XRef"))
			{
				return false;
			}
			if (Regex.Matches(text, "\\btrailer\\b").Count != 1)
			{
				return false;
			}
			Match match = Regex.Match(text, "trailer\\s*<<(?<body>.*?)>>", RegexOptions.Singleline);
			if (!match.Success)
			{
				return false;
			}
			string value = match.Groups["body"].Value;
			Match match2 = Regex.Match(value, "/Root\\s+(\\d+)\\s+(\\d+)\\s+R");
			if (!match2.Success)
			{
				return false;
			}
			Match match3 = Regex.Match(value, "/Info\\s+(\\d+)\\s+(\\d+)\\s+R");
			bool repairedAny = false;
			Regex regex = new Regex("(?<num>\\d+)[ \\t]+(?<gen>\\d+)[ \\t]+obj(?<body>.*?)endobj", RegexOptions.Singleline);
			string text2 = regex.Replace(text, (Match m) =>
			{
				string value3 = m.Groups["body"].Value;
				int num6 = value3.IndexOf("stream", StringComparison.Ordinal);
				if (num6 < 0)
				{
					return m.Value;
				}
				string text3 = value3.Substring(0, num6);
				string text4 = value3.Substring(num6);
				int num7 = "stream".Length;
				if (num7 < text4.Length && text4[num7] == '\r')
				{
					num7++;
				}
				if (num7 < text4.Length && text4[num7] == '\n')
				{
					num7++;
				}
				int num8 = text4.IndexOf("endstream", num7, StringComparison.Ordinal);
				if (num8 < 0)
				{
					return m.Value;
				}
				int num9 = num8;
				if (num9 > num7 && text4[num9 - 1] == '\n')
				{
					num9--;
				}
				if (num9 > num7 && text4[num9 - 1] == '\r')
				{
					num9--;
				}
				int num10 = num9 - num7;
				if (Regex.IsMatch(text3, "/Length\\s+\\d+\\s+\\d+\\s+R"))
				{
					return m.Value;
				}
				Match match5 = Regex.Match(text3, "/Length\\s+(\\d+)\\b");
				if (match5.Success && int.Parse(match5.Groups[1].Value) == num10)
				{
					return m.Value;
				}
				string text5;
				if (match5.Success)
				{
					text5 = text3.Substring(0, match5.Index) + "/Length " + num10 + text3.Substring(match5.Index + match5.Length);
				}
				else
				{
					int num11 = text3.LastIndexOf(">>", StringComparison.Ordinal);
					if (num11 < 0)
					{
						return m.Value;
					}
					text5 = text3.Substring(0, num11) + "/Length " + num10 + " " + text3.Substring(num11);
				}
				repairedAny = true;
				return m.Groups["num"].Value + " " + m.Groups["gen"].Value + " obj" + text5 + text4 + "endobj";
			});
			if (!repairedAny)
			{
				return false;
			}
			Dictionary<int, (long, int)> dictionary = new Dictionary<int, (long, int)>();
			int num = 0;
			foreach (Match item in regex.Matches(text2))
			{
				int num2 = int.Parse(item.Groups["num"].Value);
				int num3 = int.Parse(item.Groups["gen"].Value);
				dictionary[num2] = (item.Index, num3);
				if (num2 > num)
				{
					num = num2;
				}
			}
			StringBuilder stringBuilder = new StringBuilder(text2);
			long num4 = latin.GetByteCount(text2);
			stringBuilder.Append("xref\r\n");
			stringBuilder.Append("0 ").Append(num + 1).Append("\r\n");
			stringBuilder.Append("0000000000 65535 f\r\n");
			for (int num5 = 1; num5 <= num; num5++)
			{
				if (dictionary.TryGetValue(num5, out var value2))
				{
					stringBuilder.Append(value2.Item1.ToString("D10")).Append(' ').Append(value2.Item2.ToString("D5"))
						.Append(" n\r\n");
				}
				else
				{
					stringBuilder.Append("0000000000 65535 f\r\n");
				}
			}
			stringBuilder.Append("trailer\r\n<< /Size ").Append(num + 1).Append(" /Root ")
				.Append(match2.Groups[1].Value)
				.Append(' ')
				.Append(match2.Groups[2].Value)
				.Append(" R");
			if (match3.Success)
			{
				stringBuilder.Append(" /Info ").Append(match3.Groups[1].Value).Append(' ')
					.Append(match3.Groups[2].Value)
					.Append(" R");
			}
			stringBuilder.Append(" >>\r\nstartxref\r\n").Append(num4).Append("\r\n%%EOF");
			File.WriteAllBytes(P_0, latin.GetBytes(stringBuilder.ToString()));
			try
			{
				File.AppendAllText(P_1, $"{DateTime.Now}: Repaired stream(s) with missing/incorrect /Length and rebuilt xref table for {P_0}\n");
			}
			catch
			{
			}
			return true;
		}
		catch (Exception ex)
		{
			try
			{
				File.AppendAllText(P_1, $"{DateTime.Now}: TryRepairMissingStreamLengths failed: {ex}\n");
			}
			catch
			{
			}
			return false;
		}
	}

	private static void CreateOverlayOnlyPdf(string P_0, string P_1, IEnumerable<DrawingIndexItem> P_2, string P_3, string P_4 = "Arial")
	{
		try
		{
			using UglyToad.PdfPig.PdfDocument pdfDocument = UglyToad.PdfPig.PdfDocument.Open(P_0);
			UglyToad.PdfPig.Content.Page? page = pdfDocument.GetPages().FirstOrDefault() ?? throw new InvalidOperationException("Template has no pages");
			double width = page.Width;
			double height = page.Height;
			PdfSharp.Pdf.PdfDocument outDoc = new PdfSharp.Pdf.PdfDocument();
			PdfSharp.Pdf.PdfPage pdfPage = outDoc.AddPage();
			pdfPage.Width = XUnit.FromPoint(width);
			pdfPage.Height = XUnit.FromPoint(height);
			using XGraphics xGraphics = XGraphics.FromPdfPage(pdfPage);
			double num = 40.0;
			double num2 = 20.0;
			double num3 = 54.0;
			double[] array = new double[4] { 72.0, 180.0, 129.6, 216.0 };
			XFont font = new XFont(P_4, 30.0);
			XFont font2 = new XFont(P_4, 12.0);
			XPen pen = new XPen(XColors.Black, 0.5);
			double num4 = num;
			double num5 = num2;
			xGraphics.DrawRectangle(XBrushes.LightGray, num4, num5, array.Sum(), num3);
			xGraphics.DrawRectangle(pen, num4, num5, array.Sum(), num3);
			xGraphics.DrawString("SEQ. NO.", font, XBrushes.Black, new XRect(num4 + 4.0, num5 + 2.0, array[0], num3), XStringFormats.TopLeft);
			num4 += array[0];
			xGraphics.DrawString("DWG. NO.", font, XBrushes.Black, new XRect(num4 + 4.0, num5 + 2.0, array[1], num3), XStringFormats.TopLeft);
			num4 += array[1];
			xGraphics.DrawString("DISCIPLINE", font, XBrushes.Black, new XRect(num4 + 4.0, num5 + 2.0, array[2], num3), XStringFormats.TopLeft);
			num4 += array[2];
			xGraphics.DrawString("DRAWING TITLE", font, XBrushes.Black, new XRect(num4 + 4.0, num5 + 2.0, array[3], num3), XStringFormats.TopLeft);
			num5 += num3;
			foreach (DrawingIndexItem item in P_2)
			{
				num4 = num;
				xGraphics.DrawRectangle(pen, num4, num5, array.Sum(), num3);
				xGraphics.DrawString(item.SeqNo.ToString(), font2, XBrushes.Black, new XRect(num4 + 4.0, num5 + 2.0, array[0], num3), XStringFormats.TopLeft);
				num4 += array[0];
				xGraphics.DrawString(item.DrawingNumber, font2, XBrushes.Black, new XRect(num4 + 4.0, num5 + 2.0, array[1], num3), XStringFormats.TopLeft);
				num4 += array[1];
				xGraphics.DrawString(item.Discipline, font2, XBrushes.Black, new XRect(num4 + 4.0, num5 + 2.0, array[2], num3), XStringFormats.TopLeft);
				num4 += array[2];
				xGraphics.DrawString(item.Title, font2, XBrushes.Black, new XRect(num4 + 4.0, num5 + 2.0, array[3], num3), XStringFormats.TopLeft);
				num5 += num3;
				if (num5 + num3 > height - 60.0)
				{
					break;
				}
			}
			RunWithFileLockRetry(() =>
			{
				outDoc.Save(P_1);
			}, P_3, "Save " + P_1 + " (fallback overlay)");
			File.AppendAllText(P_3, "Fallback overlay PDF saved\n");
		}
		catch (Exception ex)
		{
			File.AppendAllText(P_3, "CreateOverlayOnlyPdf error: " + ex?.ToString() + "\n");
			throw;
		}
	}

	private static string? GenerateDrawingIndexCSV(string P_0, FieldRect P_1, FieldRect P_2, FieldRect P_3, string P_4, IProgress<(int current, int total, string status)> P_5, CancellationToken P_6, string P_7 = "DrawingIndex_", bool P_8 = true, string P_9 = "Arial", string? P_10 = null, string? P_11 = null, bool P_12 = false, bool P_13 = false, OcrRunStats? P_14 = null)
	{
		P_5.Report((0, 0, "Scanning folder for sheet PDFs..."));
		if (!Directory.Exists(Path.GetDirectoryName(P_0) ?? string.Empty))
		{
			return null;
		}
		string text = $"__AutoDIS_CSVTemp_{Guid.NewGuid():N}_";
		List<DrawingIndexItem> item;
		try
		{
			bool flag = P_12;
			bool flag2 = P_13;
			item = GenerateDrawingIndexCore(P_0, P_1, P_2, P_3, false, false, P_4, P_5, P_6, text, false, P_9, P_7, P_10, null, false, true, flag, null, flag2, P_14).finalRows;
		}
		finally
		{
			try
			{
				foreach (string item2 in Directory.EnumerateFiles(GetOrCreateOutputFolder(P_0), text + "*.pdf"))
				{
					try
					{
						File.Delete(item2);
					}
					catch
					{
					}
				}
			}
			catch
			{
			}
		}
		if (item == null)
		{
			return null;
		}
		P_5.Report((item.Count, item.Count, "Writing CSV file..."));
		string text2 = P_11 ?? Path.Combine(GetOrCreateOutputFolder(P_0), P_8 ? $"{P_7}{DateTime.Now:yyyyMMdd_HHmmss}.csv" : (P_7 + ".csv"));
		try
		{
			StringBuilder stringBuilder = new StringBuilder();
			bool anyUncertain = item.Any(r => r.DrawingNumberUncertain || r.DisciplineUncertain || r.TitleUncertain);
			stringBuilder.AppendLine(anyUncertain ? "SEQ. NO.,DISCIPLINE,DWG. NO.,DESCRIPTION,CHECK" : "SEQ. NO.,DISCIPLINE,DWG. NO.,DESCRIPTION");
			foreach (DrawingIndexItem item3 in item)
			{
				string text3 = EscapeCsvField(item3.SeqNo.ToString());
				string text4 = EscapeCsvField(string.IsNullOrWhiteSpace(item3.Discipline) ? "[DISCIPLINE NOT FOUND]" : item3.Discipline);
				string text5 = EscapeCsvField(item3.DrawingNumber ?? string.Empty);
				string text6 = EscapeCsvField(item3.Title ?? string.Empty);
				StringBuilder stringBuilder2 = stringBuilder;
				StringBuilder.AppendInterpolatedStringHandler appendInterpolatedStringHandler = new StringBuilder.AppendInterpolatedStringHandler(3, 4, stringBuilder2);
				appendInterpolatedStringHandler.AppendFormatted(text3);
				appendInterpolatedStringHandler.AppendLiteral(",");
				appendInterpolatedStringHandler.AppendFormatted(text4);
				appendInterpolatedStringHandler.AppendLiteral(",");
				appendInterpolatedStringHandler.AppendFormatted(text5);
				appendInterpolatedStringHandler.AppendLiteral(",");
				appendInterpolatedStringHandler.AppendFormatted(text6);
				if (anyUncertain)
				{
					List<string> toCheck = new List<string>();
					if (item3.DisciplineUncertain)
					{
						toCheck.Add("DISCIPLINE");
					}
					if (item3.DrawingNumberUncertain)
					{
						toCheck.Add("DWG. NO.");
					}
					if (item3.TitleUncertain)
					{
						toCheck.Add("DESCRIPTION");
					}
					appendInterpolatedStringHandler.AppendLiteral(",");
					appendInterpolatedStringHandler.AppendFormatted(EscapeCsvField(string.Join("; ", toCheck)));
				}
				stringBuilder2.AppendLine(ref appendInterpolatedStringHandler);
			}
			File.WriteAllText(text2, stringBuilder.ToString(), Encoding.UTF8);
			File.AppendAllText(P_4, $"{DateTime.Now}: CSV written to {text2}\n");
			return text2;
		}
		catch (Exception ex)
		{
			File.AppendAllText(P_4, $"{DateTime.Now}: Failed to write CSV: {ex}\n");
			throw;
		}
	}

	private static string EscapeCsvField(string P_0)
	{
		if (string.IsNullOrEmpty(P_0))
		{
			return "\"\"";
		}
		if (P_0.Contains(',') || P_0.Contains('"') || P_0.Contains('\n'))
		{
			return "\"" + P_0.Replace("\"", "\"\"") + "\"";
		}
		return P_0;
	}

	private async void PickPdfButton_Click(object? sender, RoutedEventArgs e)
	{
		FileOpenPicker fileOpenPicker = new FileOpenPicker
		{
			FileTypeFilter = { ".pdf" }
		};
		nint windowHandle = WindowNative.GetWindowHandle(this);
		InitializeWithWindow.Initialize(fileOpenPicker, windowHandle);
		StorageFile storageFile = await fileOpenPicker.PickSingleFileAsync();
		if ((object)storageFile != null)
		{
			if (HasNonTemplateContent(storageFile.Path))
			{
				PdfPathTextBox.Text = string.Empty;
				await new ContentDialog
				{
					Title = "Non-template content detected",
					Content = "The selected PDF has non-template content. Select a document that is primarily blank other than the title block etc.",
					CloseButtonText = "OK",
					XamlRoot = Content.XamlRoot
				}.ShowAsync();
			}
			else
			{
				PdfPathTextBox.Text = storageFile.Path;
			}
		}
	}

	private async void PickCombinedSetButton_Click(object? sender, RoutedEventArgs e)
	{
		FileOpenPicker fileOpenPicker = new FileOpenPicker
		{
			FileTypeFilter = { ".pdf" }
		};
		nint windowHandle = WindowNative.GetWindowHandle(this);
		InitializeWithWindow.Initialize(fileOpenPicker, windowHandle);
		StorageFile storageFile = await fileOpenPicker.PickSingleFileAsync();
		if ((object)storageFile != null)
		{
			CombinedSetPathTextBox.Text = storageFile.Path;
		}
	}

	private void ClearCombinedSetButton_Click(object? sender, RoutedEventArgs e)
	{
		CombinedSetPathTextBox.Text = string.Empty;
	}

	private void IncludeBookmarksCheckBox_CheckedChanged(object? sender, RoutedEventArgs e)
	{
		if (!(BookmarkSourceComboBox == null))
		{
			BookmarkSourceComboBox.IsEnabled = IncludeBookmarksCheckBox.IsChecked == true;
		}
	}

	private static bool HasNonTemplateContent(string P_0)
	{
		try
		{
			using UglyToad.PdfPig.PdfDocument pdfDocument = UglyToad.PdfPig.PdfDocument.Open(P_0);
			UglyToad.PdfPig.Content.Page page = pdfDocument.GetPages().FirstOrDefault();
			if (page == null)
			{
				return false;
			}
			double width = page.Width;
			double height = page.Height;
			double num = width * 0.2;
			double num2 = width * 0.8;
			double num3 = height * 0.2;
			double num4 = height * 0.8;
			foreach (Word word in page.GetWords())
			{
				IEnumerable<Letter> letters = word.Letters;
				List<Letter> list = (letters ?? Enumerable.Empty<Letter>()).ToList();
				if (list.Any())
				{
					double num5 = list.Average((Letter l) => l.GlyphRectangle.Left + l.GlyphRectangle.Width / 2.0);
					double num6 = list.Average((Letter l) => l.GlyphRectangle.Bottom + l.GlyphRectangle.Height / 2.0);
					if (num5 >= num && num5 <= num2 && num6 >= num3 && num6 <= num4)
					{
						return true;
					}
				}
			}
		}
		catch
		{
		}
		return false;
	}

	private static bool TryGetTemplateFieldRects(string P_0, out FieldRect P_1, out FieldRect P_2, out FieldRect P_3, out FieldRect? P_4, out string P_5)
	{
		P_1 = default;
		P_2 = default;
		P_3 = default;
		P_4 = null;
		P_5 = string.Empty;
		try
		{
			using PdfSharp.Pdf.PdfDocument pdfDocument = PdfReader.Open(P_0, PdfDocumentOpenMode.Import);
			if (pdfDocument.PageCount == 0)
			{
				P_5 = "Template has no pages.";
				return false;
			}
			PdfSharp.Pdf.PdfPage pdfPage = pdfDocument.Pages[0];
			double x = pdfPage.MediaBox.X1;
			double y = pdfPage.MediaBox.Y1;
			PdfItem pdfItem = pdfPage.Elements["/Annots"];
			PdfArray pdfArray = (pdfItem as PdfArray) ?? ((pdfItem as PdfReference)?.Value as PdfArray);
			if (pdfArray == null)
			{
				P_5 = "No form fields found in the template.";
				return false;
			}
			bool flag = false;
			bool flag2 = false;
			bool flag3 = false;
			foreach (PdfItem element in pdfArray.Elements)
			{
				PdfDictionary pdfDictionary = ((element as PdfReference)?.Value as PdfDictionary) ?? (element as PdfDictionary);
				if (pdfDictionary == null || !string.Equals(pdfDictionary.Elements.GetName("/Subtype"), "/Widget", StringComparison.OrdinalIgnoreCase))
				{
					continue;
				}
				string text = pdfDictionary.Elements.GetString("/T");
				if (string.IsNullOrWhiteSpace(text))
				{
					continue;
				}
				PdfItem pdfItem2 = pdfDictionary.Elements["/Rect"];
				PdfArray pdfArray2 = (pdfItem2 as PdfArray) ?? ((pdfItem2 as PdfReference)?.Value as PdfArray);
				if (pdfArray2 != null && pdfArray2.Elements.Count >= 4)
				{
					double num = pdfArray2.Elements.GetReal(0) - x;
					double num2 = pdfArray2.Elements.GetReal(1) - y;
					double num3 = pdfArray2.Elements.GetReal(2) - x;
					double num4 = pdfArray2.Elements.GetReal(3) - y;
					FieldRect fieldRect = new FieldRect(Math.Min(num, num3), Math.Min(num2, num4), Math.Max(num, num3), Math.Max(num2, num4));
					if (string.Equals(text, "SheetID", StringComparison.OrdinalIgnoreCase))
					{
						P_1 = fieldRect;
						flag = true;
					}
					else if (string.Equals(text, "SheetDisc", StringComparison.OrdinalIgnoreCase))
					{
						P_2 = fieldRect;
						flag2 = true;
					}
					else if (string.Equals(text, "SheetTitle", StringComparison.OrdinalIgnoreCase))
					{
						P_3 = fieldRect;
						flag3 = true;
					}
					else if (string.Equals(text, "SheetNumber", StringComparison.OrdinalIgnoreCase))
					{
						P_4 = fieldRect;
					}
				}
			}
			List<string> list = new List<string>();
			if (!flag)
			{
				list.Add("SheetID");
			}
			if (!flag2)
			{
				list.Add("SheetDisc");
			}
			if (!flag3)
			{
				list.Add("SheetTitle");
			}
			if (list.Count > 0)
			{
				P_5 = "The following required form field(s) were not found in the template: " + string.Join(", ", list) + ".";
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			P_5 = "Could not read template form fields: " + ex.Message;
			return false;
		}
	}

	private static void RemoveTemplateFormFieldWidgets(PdfSharp.Pdf.PdfDocument P_0, PdfSharp.Pdf.PdfPage P_1)
	{
		PdfItem pdfItem = P_1.Elements["/Annots"];
		PdfArray pdfArray = (pdfItem as PdfArray) ?? ((pdfItem as PdfReference)?.Value as PdfArray);
		if (pdfArray == null)
		{
			return;
		}
		HashSet<string> hashSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "SheetID", "SheetDisc", "SheetTitle", "SheetNumber" };
		PdfArray pdfArray2 = new PdfArray(P_0);
		foreach (PdfItem element in pdfArray.Elements)
		{
			PdfDictionary pdfDictionary = ((element as PdfReference)?.Value as PdfDictionary) ?? (element as PdfDictionary);
			bool flag = false;
			if (pdfDictionary != null)
			{
				string name = pdfDictionary.Elements.GetName("/Subtype");
				string text = pdfDictionary.Elements.GetString("/T");
				if (string.Equals(name, "/Widget", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(text) && hashSet.Contains(text))
				{
					flag = true;
				}
			}
			if (!flag)
			{
				pdfArray2.Elements.Add(element);
			}
		}
		P_1.Elements["/Annots"] = pdfArray2;
	}

	private static void SetIndexPageFormFieldValues(PdfSharp.Pdf.PdfDocument P_0, PdfSharp.Pdf.PdfPage P_1, string P_2, string P_3, string P_4, double P_5 = 1.0)
	{
		PdfItem pdfItem = P_1.Elements["/Annots"];
		PdfArray pdfArray = (pdfItem as PdfArray) ?? ((pdfItem as PdfReference)?.Value as PdfArray);
		if (pdfArray == null)
		{
			return;
		}
		(double, double, double, double)? tuple = null;
		(double, double, double, double)? tuple2 = null;
		HashSet<string> hashSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "SheetID", "SheetDisc", "SheetTitle", "SheetNumber" };
		PdfArray pdfArray2 = new PdfArray(P_0);
		foreach (PdfItem element in pdfArray.Elements)
		{
			PdfDictionary pdfDictionary = ((element as PdfReference)?.Value as PdfDictionary) ?? (element as PdfDictionary);
			bool flag = false;
			if (pdfDictionary != null)
			{
				string name = pdfDictionary.Elements.GetName("/Subtype");
				string text = pdfDictionary.Elements.GetString("/T");
				if (string.Equals(name, "/Widget", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(text) && hashSet.Contains(text))
				{
					flag = true;
					PdfItem pdfItem2 = pdfDictionary.Elements["/Rect"];
					PdfArray pdfArray3 = (pdfItem2 as PdfArray) ?? ((pdfItem2 as PdfReference)?.Value as PdfArray);
					if (pdfArray3 != null && pdfArray3.Elements.Count >= 4)
					{
						(double, double, double, double) value = (pdfArray3.Elements.GetReal(0), pdfArray3.Elements.GetReal(1), pdfArray3.Elements.GetReal(2), pdfArray3.Elements.GetReal(3));
						if (string.Equals(text, "SheetID", StringComparison.OrdinalIgnoreCase))
						{
							tuple = value;
						}
						else if (string.Equals(text, "SheetTitle", StringComparison.OrdinalIgnoreCase))
						{
							tuple2 = value;
						}
					}
				}
			}
			if (!flag)
			{
				pdfArray2.Elements.Add(element);
			}
		}
		P_1.Elements["/Annots"] = pdfArray2;
		double num = Math.Max(6.0, 20.0 * P_5);
		double point = P_1.Height.Point;
		double x = P_1.MediaBox.X1;
		double y = P_1.MediaBox.Y1;
		if (tuple.HasValue)
		{
			(double, double, double, double) value2 = tuple.Value;
			double num2 = value2.Item1 - x;
			double num3 = point - (value2.Item4 - y);
			AddFreeTextAnnotation(P_0, P_1, P_2, num2, num3, value2.Item3 - value2.Item1, value2.Item4 - value2.Item2, num, P_4, 1);
		}
		if (tuple2.HasValue)
		{
			(double, double, double, double) value3 = tuple2.Value;
			double num4 = value3.Item1 - x;
			double num5 = point - (value3.Item4 - y);
			AddFreeTextAnnotation(P_0, P_1, P_3, num4, num5, value3.Item3 - value3.Item1, value3.Item4 - value3.Item2, num, P_4, 1);
		}
	}

	private static void SetSheetNumberFormFieldValue(PdfSharp.Pdf.PdfDocument P_0, PdfSharp.Pdf.PdfPage P_1, FieldRect P_2, int P_3, int P_4, string P_5, int P_6, double P_7, double P_8)
	{
		double num = P_8;
		double num2 = P_2.X2 - P_2.X1;
		double num3 = P_2.Y2 - P_2.Y1;
		FieldRect fieldRect = P_2;
		int normalizedRotate = (P_6 % 360 + 360) % 360;
		(double, double, double, double, double, double)? tuple = normalizedRotate switch
		{
			90 => ((double, double, double, double, double, double)?)(0.0, 1.0, -1.0, 0.0, num3, 0.0), 
			180 => (-1.0, 0.0, 0.0, -1.0, num2, num3), 
			270 => (0.0, -1.0, 1.0, 0.0, 0.0, num2), 
			_ => null, 
		};
		if (normalizedRotate != 0)
		{
			(double x, double y) tuple2 = InverseTransformPoint(P_2.X1, P_2.Y1);
			double item = tuple2.x;
			double item2 = tuple2.y;
			(double x, double y) tuple3 = InverseTransformPoint(P_2.X2, P_2.Y2);
			double item3 = tuple3.x;
			double item4 = tuple3.y;
			fieldRect = new FieldRect(Math.Min(item, item3), Math.Min(item2, item4), Math.Max(item, item3), Math.Max(item2, item4));
		}
		int num4 = 0;
		try
		{
			double x = P_1.MediaBox.X1;
			double y = P_1.MediaBox.Y1;
			num4 = PdfContentRedactor.RemoveContentInRect(P_0, P_1, fieldRect.X1 + x, fieldRect.Y1 + y, fieldRect.X2 + x, fieldRect.Y2 + y);
		}
		catch
		{
			num4 = 0;
		}
		(double, double, double)? tuple4 = ((num4 > 0) ? (((double, double, double)?)null) : new (double, double, double)?((1.0, 1.0, 1.0)));
		string text = $" {P_3} OF {P_4}";
		double x2 = fieldRect.X1;
		double num5 = fieldRect.X2 - fieldRect.X1;
		double num6 = fieldRect.Y2 - fieldRect.Y1;
		double num7 = num - fieldRect.Y2;
		double num8 = Math.Max(6.0, FitFontSizeToBox(text, P_5, false, num2, num3, 0.6) * 0.7);
		AddFreeTextAnnotation(P_0, P_1, text, x2, num7, num5, num6, num8, P_5, 1, false, tuple4, num, tuple, num2, num3, true, true);
		(double x, double y) InverseTransformPoint(double num10, double num9)
		{
			return normalizedRotate switch
			{
				90 => (x: P_7 - num9, y: num10), 
				180 => (x: P_7 - num10, y: P_8 - num9), 
				270 => (x: num9, y: P_8 - num10), 
				_ => (x: num10, y: num9), 
			};
		}
	}

	private static string DeduplicateStruckGlyphs(IReadOnlyList<Letter> P_0)
	{
		List<Letter> list = P_0.OrderBy((Letter l) => l.GlyphRectangle.Left).ToList();
		StringBuilder stringBuilder = new StringBuilder();
		Letter letter = null;
		foreach (Letter item in list)
		{
			if (letter != null && item.Value == letter.Value)
			{
				double num = Math.Abs(item.GlyphRectangle.Left - letter.GlyphRectangle.Left);
				double num2 = Math.Abs(item.GlyphRectangle.Bottom - letter.GlyphRectangle.Bottom);
				double num3 = Math.Max(0.75, item.GlyphRectangle.Width * 0.3);
				if (num < num3 && num2 < num3)
				{
					continue;
				}
			}
			stringBuilder.Append(item.Value);
			letter = item;
		}
		return stringBuilder.ToString();
	}

	private static string GetSettingsFilePath()
	{
		return GetLogFilePath("settings.txt");
	}

	private static Dictionary<string, string> LoadSettingsFile()
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		try
		{
			string settingsFilePath = GetSettingsFilePath();
			if (File.Exists(settingsFilePath))
			{
				string[] array = File.ReadAllLines(settingsFilePath);
				for (int i = 0; i < array.Length; i++)
				{
					string[] array2 = array[i].Split('=', 2);
					if (array2.Length == 2)
					{
						dictionary[array2[0].Trim()] = array2[1].Trim();
					}
				}
			}
		}
		catch
		{
		}
		return dictionary;
	}

	private static void SaveSettingsFile(Dictionary<string, string> P_0)
	{
		try
		{
			StringBuilder stringBuilder = new StringBuilder();
			foreach (KeyValuePair<string, string> item in P_0)
			{
				stringBuilder.Append(item.Key).Append('=').Append(item.Value)
					.Append('\n');
			}
			File.WriteAllText(GetSettingsFilePath(), stringBuilder.ToString());
		}
		catch
		{
		}
	}

	private static OcrPreset LoadOcrPreset()
	{
		return LoadSettingsFile().TryGetValue("OcrPreset", out string value) ? OcrPresets.ParseOrDefault(value) : OcrPresets.Default;
	}

	private static ResourceLevel LoadResourceLevel()
	{
		return LoadSettingsFile().TryGetValue("ResourceLevel", out string value) ? ResourcePlanner.ParseOrDefault(value) : ResourcePlanner.Default;
	}

	private static void SaveResourceLevel(ResourceLevel level)
	{
		Dictionary<string, string> dictionary = LoadSettingsFile();
		dictionary["ResourceLevel"] = level.ToString();
		SaveSettingsFile(dictionary);
	}

	private static void SaveOcrPreset(OcrPreset preset)
	{
		Dictionary<string, string> dictionary = LoadSettingsFile();
		dictionary["OcrPreset"] = preset.ToString();
		SaveSettingsFile(dictionary);
	}
	// Checked before the user picks an output file, so a broken OCR install is reported up front instead of producing a silently incomplete index.
	private async Task<bool> ConfirmOcrAvailableAsync(string logPath)
	{
		OcrEngineSet? engines = await Task.Run(() => OcrSupport.TryCreateEngineSet(OcrSupport.DefaultTessdataDirectory, false, out _));
		if (engines != null)
		{
			engines.Dispose();
			return true;
		}
		string reason = await Task.Run(() =>
		{
			OcrSupport.TryCreateEngineSet(OcrSupport.DefaultTessdataDirectory, false, out string? why);
			return why ?? "unknown error";
		});
		AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR preflight failed: {reason}");
		ContentDialog dialog = new ContentDialog
		{
			Title = "OCR is not available",
			Content = "Sheets whose title block has no readable text (scanned or CAD-vector sheets) cannot be read without OCR, and their index entries will be marked [NOT FOUND].\n\nReason: " + reason + "\n\nReinstalling the app normally fixes this.",
			PrimaryButtonText = "Continue without OCR",
			CloseButtonText = "Cancel",
			DefaultButton = ContentDialogButton.Close,
			XamlRoot = Content.XamlRoot
		};
		return await dialog.ShowAsync() == ContentDialogResult.Primary;
	}

	private async Task ReportOcrOutcomeAsync(OcrRunStats stats, string logPath)
	{
		AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR summary: {stats.Describe()}");
		AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR image scale: {OcrFieldRecognizer.DescribeScaleStats()}");
		if (!stats.HasProblem)
		{
			return;
		}
		List<string> lines = new List<string>();
		if (!stats.EngineAvailable)
		{
			lines.Add($"OCR was needed on {stats.SheetsNeedingOcr} sheet(s) but was not available, so those entries are marked [NOT FOUND] / [DESCRIPTION NOT FOUND].");
		}
		else
		{
			lines.Add($"{stats.SheetsNeedingOcr} sheet(s) had to be read by OCR.");
			if (stats.SheetsStillMissing > 0)
			{
				lines.Add($"{stats.SheetsStillMissing} still have a missing field; those are marked in yellow ([NOT FOUND]).");
			}
			if (stats.LowConfidenceCount > 0)
			{
				lines.Add($"{stats.LowConfidenceCount} were read with low confidence; those cells are highlighted in orange in the index. Please check them against the sheets.");
			}
			if (stats.RenderFailures > 0)
			{
				lines.Add($"{stats.RenderFailures} sheet image(s) could not be rendered for OCR.");
			}
		}
		await new ContentDialog
		{
			Title = "Please review some entries",
			Content = string.Join("\n\n", lines) + "\n\nDetails were written to:\n" + logPath,
			CloseButtonText = "OK",
			XamlRoot = Content.XamlRoot
		}.ShowAsync();
	}

	private static string FindTextInRect(UglyToad.PdfPig.Content.Page P_0, FieldRect P_1, out List<TextHit> P_2, double P_3 = 10.0)
	{
		string text = FindTextInRectCore(P_0, P_1, P_3, 4.0, 4.0, out P_2);
		if (!string.IsNullOrWhiteSpace(text))
		{
			return text;
		}
		return FindTextInRectCore(P_0, P_1, P_3, 4.0, 30.0, out P_2);
	}

	private static string StripLeadingDuplicateFieldByPosition(string P_0, List<TextHit> P_1, string P_2, List<TextHit> P_3)
	{
		if (string.IsNullOrWhiteSpace(P_0) || string.IsNullOrWhiteSpace(P_2))
		{
			return P_0;
		}
		if (P_1.Count == 0 || P_3.Count == 0)
		{
			return P_0;
		}
		string text = P_2.Trim();
		if (P_0.Length <= text.Length)
		{
			return P_0;
		}
		if (!P_0.StartsWith(text, StringComparison.OrdinalIgnoreCase))
		{
			return P_0;
		}
		string text2 = P_0.Substring(text.Length);
		if (text2.Length == 0 || !char.IsWhiteSpace(text2[0]))
		{
			return P_0;
		}
		int num = Math.Min(P_3.Count, P_1.Count);
		List<TextHit> list = P_1.Take(num).ToList();
		double num2 = list.Min((TextHit h) => h.Left);
		double num3 = list.Max((TextHit h) => h.Right);
		double num4 = list.Min((TextHit h) => h.Bottom);
		double num5 = list.Max((TextHit h) => h.Top);
		double num6 = P_3.Min((TextHit h) => h.Left);
		double num7 = P_3.Max((TextHit h) => h.Right);
		double num8 = P_3.Min((TextHit h) => h.Bottom);
		double num9 = P_3.Max((TextHit h) => h.Top);
		if (!(num2 < num7) || !(num3 > num6) || !(num4 < num9) || !(num5 > num8))
		{
			return P_0;
		}
		return text2.TrimStart();
	}

	private static bool HasOverlappingHits(List<TextHit> P_0)
	{
		for (int i = 0; i < P_0.Count; i++)
		{
			for (int j = i + 1; j < P_0.Count; j++)
			{
				TextHit textHit = P_0[i];
				TextHit textHit2 = P_0[j];
				if (textHit.Left < textHit2.Right && textHit.Right > textHit2.Left && textHit.Bottom < textHit2.Top && textHit.Top > textHit2.Bottom)
				{
					return true;
				}
			}
		}
		return false;
	}

	private static TextHit ClosestHitToRect(List<TextHit> P_0, FieldRect P_1)
	{
		return P_0.OrderBy(DistanceToRect).First();
		double DistanceToRect(TextHit h)
		{
			double num = (h.Left + h.Right) / 2.0;
			double num2 = (h.Bottom + h.Top) / 2.0;
			double num3 = Math.Max(Math.Max(P_1.X1 - num, num - P_1.X2), 0.0);
			double num4 = Math.Max(Math.Max(P_1.Y1 - num2, num2 - P_1.Y2), 0.0);
			return Math.Sqrt(num3 * num3 + num4 * num4);
		}
	}

	private static string FlattenEmbeddedNewlines(string P_0)
	{
		if (string.IsNullOrEmpty(P_0))
		{
			return P_0;
		}
		return Regex.Replace(Regex.Replace(P_0, "\\s*[\\r\\n]+\\s*", " "), " {2,}", " ").Trim();
	}

	private static string NormalizeDashesToEnDash(string P_0)
	{
		return P_0.Replace('-', '–').Replace('—', '–');
	}

	private static string CollapseConsecutiveDashes(string P_0)
	{
		if (string.IsNullOrEmpty(P_0))
		{
			return P_0;
		}
		return Regex.Replace(P_0, "–{2,}", "–");
	}

	private static readonly ConditionalWeakTable<UglyToad.PdfPig.Content.Page, List<Word>> PageWordCache = new ConditionalWeakTable<UglyToad.PdfPig.Content.Page, List<Word>>();

	private static List<Word> GetPageWords(UglyToad.PdfPig.Content.Page P_0)
	{
		return PageWordCache.GetValue(P_0, (UglyToad.PdfPig.Content.Page p) => p.GetWords().ToList());
	}

	private static readonly string[] ExcludedHeaderTexts = new string[6] { "DRAWING NO.", "SHEET NO.", "SHEET ID", "DWG NO.", "SHEETID", "SHEETNO." };

	private static readonly Regex DashOnlyRegex = new Regex("^[-–—_]{2,}$");

	private readonly record struct WordGeometry(double Cy, double Cx, string Text, double Height, double Left, double Right, double Top, double Bottom);

	private static readonly ConditionalWeakTable<UglyToad.PdfPig.Content.Page, List<WordGeometry>> PageWordGeometryCache = new ConditionalWeakTable<UglyToad.PdfPig.Content.Page, List<WordGeometry>>();

	private static List<WordGeometry> GetPageWordGeometry(UglyToad.PdfPig.Content.Page P_0)
	{
		return PageWordGeometryCache.GetValue(P_0, BuildPageWordGeometry);
	}

	private static List<WordGeometry> BuildPageWordGeometry(UglyToad.PdfPig.Content.Page P_0)
	{
		List<WordGeometry> result = new List<WordGeometry>();
		foreach (Word word in GetPageWords(P_0))
		{
			IEnumerable<Letter> letters = word.Letters;
			List<Letter> list2 = (letters ?? Enumerable.Empty<Letter>()).ToList();
			if (!list2.Any())
			{
				continue;
			}
			double top = list2.Max((Letter l) => l.GlyphRectangle.Bottom + l.GlyphRectangle.Height);
			double bottom = list2.Min((Letter l) => l.GlyphRectangle.Bottom);
			double left = list2.Min((Letter l) => l.GlyphRectangle.Left);
			double right = list2.Max((Letter l) => l.GlyphRectangle.Left + l.GlyphRectangle.Width);
			double cx = list2.Average((Letter l) => l.GlyphRectangle.Left + l.GlyphRectangle.Width / 2.0);
			double cy = (top + bottom) / 2.0;
			string text = DeduplicateStruckGlyphs(list2);
			if (ExcludedHeaderTexts.Any((string p) => string.Equals(text, p, StringComparison.OrdinalIgnoreCase)) || DashOnlyRegex.IsMatch(text))
			{
				continue;
			}
			result.Add(new WordGeometry(cy, cx, text, top - bottom, left, right, top, bottom));
		}
		return result;
	}

	private static string FindTextInRectCore(UglyToad.PdfPig.Content.Page P_0, FieldRect P_1, double P_2, double P_3, double P_4, out List<TextHit> P_5)
	{
		FieldRect fieldRect = new FieldRect(Math.Max(0.0, P_1.X1 - P_2), Math.Max(0.0, P_1.Y1 - P_4), P_1.X2 + P_2, P_1.Y2 + P_3);
		List<(double, double, string, double, double, double, double, double)> list = new List<(double, double, string, double, double, double, double, double)>();
		foreach (WordGeometry entry in GetPageWordGeometry(P_0))
		{
			if (entry.Cx >= fieldRect.X1 && entry.Cx <= fieldRect.X2 && entry.Cy >= fieldRect.Y1 && entry.Cy <= fieldRect.Y2)
			{
				list.Add((entry.Cy, entry.Cx, entry.Text, entry.Height, entry.Left, entry.Right, entry.Top, entry.Bottom));
			}
		}
		if (list.Count == 0)
		{
			P_5 = new List<TextHit>();
			return string.Empty;
		}
		List<(double, double, string, double, double, double, double, double)> list3 = list.OrderByDescending(((double cy, double cx, string text, double height, double left, double right, double top, double bottom) h) => h.cy).ToList();
		double num7 = Math.Max(2.0, list3.Average(((double cy, double cx, string text, double height, double left, double right, double top, double bottom) h) => h.height) * 0.6);
		List<List<(double, double, string, double, double, double, double, double)>> list4 = new List<List<(double, double, string, double, double, double, double, double)>>();
		foreach (var item in list3)
		{
			object obj;
			if (list4.Count <= 0)
			{
				obj = null;
			}
			else
			{
				obj = list4[list4.Count - 1];
			}
			List<(double, double, string, double, double, double, double, double)> list5 = (List<(double, double, string, double, double, double, double, double)>)obj;
			if (list5 != null && Math.Abs(list5[0].Item1 - item.Item1) <= num7)
			{
				list5.Add(item);
				continue;
			}
			list4.Add(new List<(double, double, string, double, double, double, double, double)> { item });
		}
		Regex seqOfPattern = new Regex("^(SEQ\\.?\\s*)?\\d+\\s*OF\\s*\\d+\\.?$", RegexOptions.IgnoreCase);
		List<(double, double, string, double, double, double, double, double)> list6 = list4.Where((List<(double cy, double cx, string text, double height, double left, double right, double top, double bottom)> line) =>
		{
			string text2 = string.Join(" ", from h in line
				orderby h.cx
				select h.text).Trim();
			return !seqOfPattern.IsMatch(text2);
		}).ToList().SelectMany((List<(double cy, double cx, string text, double height, double left, double right, double top, double bottom)> line) => line.OrderBy(((double cy, double cx, string text, double height, double left, double right, double top, double bottom) h) => h.cx))
			.ToList();
		P_5 = list6.Select(((double cy, double cx, string text, double height, double left, double right, double top, double bottom) h) => new TextHit(h.text, h.left, h.bottom, h.right, h.top)).ToList();
		return string.Join(" ", list6.Select(((double cy, double cx, string text, double height, double left, double right, double top, double bottom) h) => h.text)).Trim();
	}

	private static double EstimateTitleBlockTopY(string P_0, double P_1)
	{
		try
		{
			using UglyToad.PdfPig.PdfDocument pdfDocument = UglyToad.PdfPig.PdfDocument.Open(P_0);
			UglyToad.PdfPig.Content.Page page = pdfDocument.GetPages().FirstOrDefault();
			if (page == null)
			{
				return P_1;
			}
			double num = page.Height * 0.22;
			double num2 = P_1;
			bool flag = false;
			foreach (Word word in page.GetWords())
			{
				IEnumerable<Letter> letters = word.Letters;
				List<Letter> list = (letters ?? Enumerable.Empty<Letter>()).ToList();
				if (list.Any())
				{
					flag = true;
					double num3 = list.Max((Letter l) => l.GlyphRectangle.Bottom + l.GlyphRectangle.Height);
					if (num3 <= num && num3 > num2)
					{
						num2 = num3;
					}
				}
			}
			if (!flag)
			{
				return Math.Max(P_1, page.Height * 0.1);
			}
			return num2;
		}
		catch
		{
			return P_1;
		}
	}

	private static double EstimateLeftMarginX(string P_0, double P_1)
	{
		try
		{
			using UglyToad.PdfPig.PdfDocument pdfDocument = UglyToad.PdfPig.PdfDocument.Open(P_0);
			UglyToad.PdfPig.Content.Page page = pdfDocument.GetPages().FirstOrDefault();
			if (page == null)
			{
				return 90.0;
			}
			double num = page.Width * 0.08;
			double num2 = page.Height * 0.97;
			double num3 = 0.0;
			foreach (Word word in page.GetWords())
			{
				IEnumerable<Letter> letters = word.Letters;
				List<Letter> list = (letters ?? Enumerable.Empty<Letter>()).ToList();
				if (list.Any() && !(list.Min((Letter l) => l.GlyphRectangle.Left) > num) && !(list.Min((Letter l) => l.GlyphRectangle.Bottom) < P_1) && !(list.Max((Letter l) => l.GlyphRectangle.Bottom + l.GlyphRectangle.Height) > num2))
				{
					double num4 = list.Max((Letter l) => l.GlyphRectangle.Left + l.GlyphRectangle.Width);
					if (num4 > num3)
					{
						num3 = num4;
					}
				}
			}
			return (num3 > 0.0) ? (num3 + 15.0) : 90.0;
		}
		catch
		{
			return 90.0;
		}
	}

	private static double EstimateTopMarginY(string P_0, double? P_1 = null)
	{
		try
		{
			using UglyToad.PdfPig.PdfDocument pdfDocument = UglyToad.PdfPig.PdfDocument.Open(P_0);
			UglyToad.PdfPig.Content.Page page = pdfDocument.GetPages().FirstOrDefault();
			if (page == null)
			{
				return 60.0;
			}
			double num = (P_1.HasValue ? (P_1.Value - 40.0) : (page.Height * 0.9));
			double num2 = 1.7976931348623157E+308;
			foreach (Word word in page.GetWords())
			{
				IEnumerable<Letter> letters = word.Letters;
				List<Letter> list = (letters ?? Enumerable.Empty<Letter>()).ToList();
				if (list.Any() && !list.Any((Letter l) => l.TextOrientation != TextOrientation.Horizontal))
				{
					double num3 = list.Min((Letter l) => l.GlyphRectangle.Bottom);
					if (!(num3 < num) && num3 < num2)
					{
						num2 = num3;
					}
				}
			}
			if (num2 == 1.7976931348623157E+308)
			{
				return 60.0;
			}
			double num4 = page.Height - num2 + 10.0;
			return Math.Max(60.0, num4);
		}
		catch
		{
			return 60.0;
		}
	}

	private static (double Left, double Bottom, double Right, double Top, double PageWidth, double PageHeight)? EstimateBorderRect(string P_0)
	{
		try
		{
			using UglyToad.PdfPig.PdfDocument pdfDocument = UglyToad.PdfPig.PdfDocument.Open(P_0);
			UglyToad.PdfPig.Content.Page page = pdfDocument.GetPages().FirstOrDefault();
			if (page == null)
			{
				return null;
			}
			double num = page.Width * page.Height;
			double num2 = 0.0;
			(double, double, double, double)? tuple = null;
			foreach (PdfPath path in page.Paths)
			{
				if (!path.IsStroked)
				{
					continue;
				}
				UglyToad.PdfPig.Core.PdfRectangle? boundingRectangle = path.GetBoundingRectangle();
				if (boundingRectangle.HasValue)
				{
					UglyToad.PdfPig.Core.PdfRectangle value = boundingRectangle.Value;
					double num3 = value.Width * value.Height;
					if (!(num3 < num * 0.5) && !(num3 > num * 0.985) && num3 > num2)
					{
						num2 = num3;
						tuple = (value.Left, value.Bottom, value.Right, value.Top);
					}
				}
			}
			if (tuple.HasValue)
			{
				(double, double, double, double) value2 = tuple.Value;
				double? num4 = FindNestedLine(true, value2.Item1, 1);
				double num5;
				if (!num4.HasValue)
				{
					(num5, _, _, _) = value2;
				}
				else
				{
					num5 = num4.GetValueOrDefault();
				}
				double num6 = FindNestedLine(true, value2.Item3, -1) ?? value2.Item3;
				double num7 = FindNestedLine(false, value2.Item4, -1) ?? value2.Item4;
				double num8 = FindNestedLine(false, value2.Item2, 1) ?? value2.Item2;
				return (num5, num8, num6, num7, page.Width, page.Height);
			}
			double num9 = 1.7976931348623157E+308;
			double num10 = 1.7976931348623157E+308;
			double num11 = -1.7976931348623157E+308;
			double num12 = -1.7976931348623157E+308;
			bool flag = false;
			foreach (PdfPath path2 in page.Paths)
			{
				if (!path2.IsStroked)
				{
					continue;
				}
				UglyToad.PdfPig.Core.PdfRectangle? boundingRectangle2 = path2.GetBoundingRectangle();
				if (boundingRectangle2.HasValue)
				{
					UglyToad.PdfPig.Core.PdfRectangle value3 = boundingRectangle2.Value;
					if (!(value3.Width <= page.Width * 0.4) || !(value3.Height <= page.Height * 0.4))
					{
						flag = true;
						num9 = Math.Min(num9, value3.Left);
						num10 = Math.Min(num10, value3.Bottom);
						num11 = Math.Max(num11, value3.Right);
						num12 = Math.Max(num12, value3.Top);
					}
				}
			}
			if (!flag)
			{
				return null;
			}
			double num13 = (num11 - num9) * (num12 - num10);
			if (num13 < num * 0.5 || num13 > num * 0.985)
			{
				return null;
			}
			return (num9, num10, num11, num12, page.Width, page.Height);
			double? FindNestedLine(bool flag2, double num17, int num16)
			{
				double? result = null;
				foreach (PdfPath path3 in page.Paths)
				{
					if (path3.IsStroked)
					{
						UglyToad.PdfPig.Core.PdfRectangle? boundingRectangle3 = path3.GetBoundingRectangle();
						if (boundingRectangle3.HasValue)
						{
							UglyToad.PdfPig.Core.PdfRectangle value4 = boundingRectangle3.Value;
							double num14;
							if (flag2)
							{
								if (value4.Height < page.Height * 0.4)
								{
									continue;
								}
								num14 = (value4.Left + value4.Right) / 2.0;
							}
							else
							{
								if (value4.Width < page.Width * 0.4)
								{
									continue;
								}
								num14 = (value4.Bottom + value4.Top) / 2.0;
							}
							double num15 = ((num16 > 0) ? (num14 - num17) : (num17 - num14));
							if (!(num15 < 1.0) && !(num15 > 60.0) && (!result.HasValue || ((num16 > 0) ? (num14 > result.Value) : (num14 < result.Value))))
							{
								result = num14;
							}
						}
					}
				}
				return result;
			}
		}
		catch
		{
			return null;
		}
	}

	private static (double Left, double Bottom, double Right, double Top, double PageWidth, double PageHeight)? EstimateBorderRectFromBitmap(string P_0, string? P_1 = null)
	{
		try
		{
			int w;
			int hgt;
			int stride;
			byte[] bytes;
			int minGapPx;
			int maxBandPx;
			using (UglyToad.PdfPig.PdfDocument pdfDocument = UglyToad.PdfPig.PdfDocument.Open(P_0))
			{
				UglyToad.PdfPig.Content.Page page = pdfDocument.GetPages().FirstOrDefault();
				if (page == null)
				{
					return null;
				}
				double width = page.Width;
				double height = page.Height;
				if (width <= 0.0 || height <= 0.0)
				{
					return null;
				}
				FieldRect fieldRect = new FieldRect(0.0, 0.0, width, height);
				GrayImage? picture = PdfiumRender.TryRegion(P_0, 0, fieldRect.X1, fieldRect.Y1, fieldRect.X2, fieldRect.Y2, 100.0, out string? drawProblem);
				if (picture == null)
				{
					Log("full-page render failed: " + drawProblem);
					return null;
				}
				w = picture.Width;
				hgt = picture.Height;
				if (w < 10 || hgt < 10)
				{
					return null;
				}
				{
					stride = w;
					bytes = picture.Pixels;
					int num = Math.Max(1, (int)((double)w * 0.15));
					int num2 = Math.Max(1, (int)((double)hgt * 0.15));
					double num3 = 1.3888888888888888;
					minGapPx = Math.Max(1, (int)Math.Round(3.0 * num3));
					maxBandPx = Math.Max(minGapPx, (int)Math.Round(60.0 * num3));
					var (num4, flag) = FindBorderLine(ColDarkFraction, 0, num, 1, w);
					var (num5, flag2) = FindBorderLine(ColDarkFraction, w - 1, num, -1, w);
					var (num6, flag3) = FindBorderLine(RowDarkFraction, 0, num2, 1, hgt);
					var (num7, flag4) = FindBorderLine(RowDarkFraction, hgt - 1, num2, -1, hgt);
					if (flag | flag2 | flag3 | flag4)
					{
						Log($"double-line margin band detected on: {(flag ? "left " : "")}{(flag2 ? "right " : "")}{(flag3 ? "top " : "")}{(flag4 ? "bottom " : "")}" + "-- using the inner line on those edges");
					}
					if (!num4.HasValue || !num5.HasValue || !num6.HasValue || !num7.HasValue)
					{
						Log($"no line found within the scan band on: {((!num4.HasValue) ? "left " : "")}{((!num5.HasValue) ? "right " : "")}{((!num6.HasValue) ? "top " : "")}{((!num7.HasValue) ? "bottom " : "")}(scanned {num}px in from left/right, {num2}px in from top/bottom, at {100.0} DPI)");
						return null;
					}
					if (num5.Value <= num4.Value || num7.Value <= num6.Value)
					{
						Log($"detected lines don't form a valid rectangle: leftPx={num4}, rightPx={num5}, topPx={num6}, bottomPx={num7}");
						return null;
					}
					double num8 = 1.3888888888888888;
					double num9 = (double)num4.Value / num8;
					double num10 = (double)num5.Value / num8;
					double num11 = height - (double)num6.Value / num8;
					double num12 = height - (double)num7.Value / num8;
					double num13 = (num10 - num9) * (num11 - num12);
					double num14 = width * height;
					if (num13 < num14 * 0.5 || num13 > num14 * 0.985)
					{
						Log($"detected rectangle failed the plausibility check: Left={num9:0.##}, Bottom={num12:0.##}, Right={num10:0.##}, Top={num11:0.##} (area {num13:0.##} vs page area {num14:0.##}, need 50%-98.5%)");
						return null;
					}
					Log($"found border: Left={num9:0.##}, Bottom={num12:0.##}, Right={num10:0.##}, Top={num11:0.##}");
					return (num9, num12, num10, num11, width, height);
				}
			}
			double ColDarkFraction(int x)
			{
				int num15 = Math.Max(1, hgt / 500);
				int num16 = 0;
				int num17 = 0;
				for (int i = 0; i < hgt; i += num15)
				{
					if (IsDark(x, i))
					{
						num16++;
					}
					num17++;
				}
				if (num17 <= 0)
				{
					return 0.0;
				}
				return (double)num16 / (double)num17;
			}
			(int? Pos, bool UsedSecondLine) FindBorderLine(Func<int, double> func, int num18, int num16, int num19, int num20)
			{
				int? num15 = null;
				for (int i = 0; i < num16; i++)
				{
					int num17 = num18 + i * num19;
					if (num17 < 0 || num17 >= num20)
					{
						break;
					}
					if (func(num17) >= 0.6)
					{
						num15 = num17;
						break;
					}
				}
				if (!num15.HasValue)
				{
					return (Pos: null, UsedSecondLine: false);
				}
				for (int j = minGapPx; j <= maxBandPx; j++)
				{
					int num21 = num15.Value + j * num19;
					if (num21 < 0 || num21 >= num20)
					{
						break;
					}
					if (func(num21) >= 0.6)
					{
						return (Pos: num21, UsedSecondLine: true);
					}
				}
				return (Pos: num15, UsedSecondLine: false);
			}
			bool IsDark(int num17, int num16)
			{
				return bytes[num16 * stride + num17] < 150;
			}
			double RowDarkFraction(int y)
			{
				int num15 = Math.Max(1, w / 500);
				int num16 = 0;
				int num17 = 0;
				for (int i = 0; i < w; i += num15)
				{
					if (IsDark(i, y))
					{
						num16++;
					}
					num17++;
				}
				if (num17 <= 0)
				{
					return 0.0;
				}
				return (double)num16 / (double)num17;
			}
		}
		catch
		{
			return null;
		}
		void Log(string text)
		{
			if (P_1 != null)
			{
				try
				{
					File.AppendAllText(P_1, $"{DateTime.Now}: EstimateBorderRectFromBitmap: {text}\n");
				}
				catch
				{
				}
			}
		}
	}

	private string GetSelectedFontFamily()
	{
		_ = FontComboBox.SelectedItem;
		return (((ContentControl)null)?.Content as string) ?? "Arial";
	}

	private async void GenerateDrawingIndexButton_Click(object? sender, RoutedEventArgs e)
	{
		Diagnostics.BeginRun("drawing-index"); // DIAG
		var totalTimer = Diagnostics.Start(); // DIAG
		var templateTimer = Diagnostics.Start(); // DIAG
		string templatePath = PdfPathTextBox.Text;
		if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
		{
			await new ContentDialog
			{
				Title = "No template selected",
				Content = (string.IsNullOrWhiteSpace(templatePath) ? "Please select a template PDF file before generating." : "The selected template PDF file could not be found. Please browse for it again."),
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
			return;
		}
		string combinedSetPath = (string.IsNullOrWhiteSpace(CombinedSetPathTextBox.Text) ? null : CombinedSetPathTextBox.Text);
		if (combinedSetPath != null && !File.Exists(combinedSetPath))
		{
			await new ContentDialog
			{
				Title = "Combined sheet set not found",
				Content = "The selected combined sheet set PDF could not be found. Please browse for it again, or clear it to scan the template's folder for individual sheet PDFs instead.",
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
			return;
		}
		string logPath = GetLogFilePath("export_log.txt");
		string normalizedTemplateTempFile = NormalizeTemplateRotationIfNeeded(templatePath, logPath);
		if (normalizedTemplateTempFile != null)
		{
			templatePath = normalizedTemplateTempFile;
		}
		if (!TryGetTemplateFieldRects(templatePath, out FieldRect drawingRect, out FieldRect disciplineRect, out FieldRect descriptionRect, out FieldRect? sheetNumberRect, out string content))
		{
			await new ContentDialog
			{
				Title = "Template form fields not found",
				Content = content,
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
			if (normalizedTemplateTempFile == null)
			{
				return;
			}
			try
			{
				File.Delete(normalizedTemplateTempFile);
				return;
			}
			catch
			{
				return;
			}
		}
		try
		{
			File.AppendAllText(logPath, $"{DateTime.Now}: Field rects: SheetID=[{drawingRect.X1:0.##}, {drawingRect.Y1:0.##}, {drawingRect.X2:0.##}, {drawingRect.Y2:0.##}], SheetDisc=[{disciplineRect.X1:0.##}, {disciplineRect.Y1:0.##}, {disciplineRect.X2:0.##}, {disciplineRect.Y2:0.##}], SheetTitle=[{descriptionRect.X1:0.##}, {descriptionRect.Y1:0.##}, {descriptionRect.X2:0.##}, {descriptionRect.Y2:0.##}]\n");
		}
		catch
		{
		}
		Diagnostics.Split("template prep (rotation + fields)", templateTimer); // DIAG
		if (!await ConfirmOcrAvailableAsync(logPath))
		{
			if (normalizedTemplateTempFile != null)
			{
				FileHygiene.DeleteQuietly(normalizedTemplateTempFile);
			}
			return;
		}
		OcrRunStats ocrStats = new OcrRunStats();
		bool includeSheetsInOutput = IncludeSheetsCheckBox.IsChecked == true;
		bool includeHyperlinks = IncludeHyperlinksCheckBox.IsChecked == true;
		bool includeBookmarks = IncludeBookmarksCheckBox.IsChecked == true;
		bool bookmarksFromGeneratedIndex = BookmarkSourceComboBox.SelectedIndex == 1;
		bool shortGeneralNumbers = ShortGeneralNumbersCheckBox.IsChecked == true;
		string outputPrefix = "Drawing Index";
		bool appendDT = true;
		string selectedFontFamily = GetSelectedFontFamily();
		string suggestedFileName = (appendDT ? (outputPrefix ?? "") : outputPrefix);
		FileSavePicker fileSavePicker = new FileSavePicker
		{
			SuggestedFileName = suggestedFileName,
			FileTypeChoices = { 
			{
				"PDF Document",
				(IList<string>)new List<string> { ".pdf" }
			} }
		};
		nint windowHandle = WindowNative.GetWindowHandle(this);
		InitializeWithWindow.Initialize(fileSavePicker, windowHandle);
		StorageFile storageFile = await fileSavePicker.PickSaveFileAsync();
		if (storageFile == null)
		{
			if (normalizedTemplateTempFile == null)
			{
				return;
			}
			try
			{
				File.Delete(normalizedTemplateTempFile);
				return;
			}
			catch
			{
				return;
			}
		}
		string chosenOutputPath = storageFile.Path;
		string partialOutputPath = FileHygiene.PartialPathFor(chosenOutputPath);
		// The index is built in a local file and copied to its destination at the end (see GenerateDrawingIndexCore).
		string workOutputPath = Path.Combine(Path.GetTempPath(), $"AutoDIS_work_{Guid.NewGuid():N}_{Path.GetFileName(chosenOutputPath)}");
		_generationCts = new CancellationTokenSource();
		_reviewHost = this;
		CancellationToken token = _generationCts.Token;
		Progress<(int current, int total, string status)> progress = new Progress<(int, int, string)>(((int current, int total, string status) p) =>
		{
			GenerationStatusText.Text = p.status;
			if (p.total > 0)
			{
				GenerationProgressBar.IsIndeterminate = false;
				GenerationProgressBar.Maximum = p.total;
				GenerationProgressBar.Value = p.current;
			}
			else
			{
				GenerationProgressBar.IsIndeterminate = true;
			}
		});
		totalTimer?.Restart(); // DIAG
		GenerateButton.IsEnabled = false;
		GenerateCSVButton.IsEnabled = false;
		CancelGenerationButton.IsEnabled = true;
		PdfPathTextBox.IsEnabled = false;
		BrowsePdfButton.IsEnabled = false;
		CombinedSetPathTextBox.IsEnabled = false;
		BrowseCombinedSetButton.IsEnabled = false;
		ClearCombinedSetButton.IsEnabled = false;
		IncludeSheetsCheckBox.IsEnabled = false;
		IncludeHyperlinksCheckBox.IsEnabled = false;
		ShortGeneralNumbersCheckBox.IsEnabled = false;
		IncludeBookmarksCheckBox.IsEnabled = false;
		BookmarkSourceComboBox.IsEnabled = false;
		FontComboBox.IsEnabled = false;
		SettingsMenuBarItem.IsEnabled = false;
		GenerationProgressPanel.Visibility = Visibility.Visible;
		GenerationProgressBar.IsIndeterminate = true;
		GenerationStatusText.Text = "Starting...";
		string outPath = null;
		bool cancelled = false;
		Exception failure = null;
		try
		{
			outPath = (await Task.Run(() => GenerateDrawingIndexCore(templatePath, drawingRect, disciplineRect, descriptionRect, includeSheetsInOutput, includeHyperlinks, logPath, progress, token, outputPrefix, appendDT, selectedFontFamily, null, combinedSetPath, workOutputPath, includeBookmarks, bookmarksFromGeneratedIndex, normalizedTemplateTempFile != null, sheetNumberRect, shortGeneralNumbers, ocrStats), token)).Item1;
			if (outPath != null)
			{
				var copyBackTimer = Diagnostics.Start(); // DIAG
				string copyOut = await Task.Run(() => CopyFileWithProgress(workOutputPath, partialOutputPath, token, progress, "Copying the finished file to its destination"), token);
				AppLog.TryAppend(logPath, $"{DateTime.Now}: Copied the finished file to its destination: {copyOut}");
				Diagnostics.Split("copy result to destination", copyBackTimer, copyOut); // DIAG
				outPath = partialOutputPath;
			}
		}
		catch (OperationCanceledException)
		{
			cancelled = true;
			try
			{
				File.AppendAllText(logPath, $"{DateTime.Now}: Generation cancelled; no output file was written\n");
			}
			catch
			{
			}
		}
		catch (Exception ex2)
		{
			failure = ex2;
			try
			{
				File.AppendAllText(logPath, $"{DateTime.Now}: GenerateDrawingIndexButton_Click failed: {ex2}\n");
			}
			catch
			{
			}
		}
		finally
		{
			GenerateButton.IsEnabled = true;
			GenerateCSVButton.IsEnabled = true;
			CancelGenerationButton.IsEnabled = false;
			PdfPathTextBox.IsEnabled = true;
			BrowsePdfButton.IsEnabled = true;
			CombinedSetPathTextBox.IsEnabled = true;
			BrowseCombinedSetButton.IsEnabled = true;
			ClearCombinedSetButton.IsEnabled = true;
			IncludeSheetsCheckBox.IsEnabled = true;
			IncludeHyperlinksCheckBox.IsEnabled = true;
			ShortGeneralNumbersCheckBox.IsEnabled = true;
			IncludeBookmarksCheckBox.IsEnabled = true;
			BookmarkSourceComboBox.IsEnabled = IncludeBookmarksCheckBox.IsChecked == true;
			FontComboBox.IsEnabled = true;
			SettingsMenuBarItem.IsEnabled = true;
			GenerationProgressPanel.Visibility = Visibility.Collapsed;
			_generationCts?.Dispose();
			_generationCts = null;
			_reviewHost = null;
			if (_closeAfterRun) { Close(); }
			Diagnostics.Split("total", totalTimer); // DIAG
			Diagnostics.EndRun(); // DIAG
			FileHygiene.DeleteQuietly(workOutputPath);
			if (normalizedTemplateTempFile != null)
			{
				try
				{
					File.Delete(normalizedTemplateTempFile);
				}
				catch
				{
				}
			}
		}
		if (cancelled || failure != null || outPath == null)
		{
			FileHygiene.DeleteQuietly(partialOutputPath);
		}
		else
		{
			outPath = await CommitOutputAsync(partialOutputPath, chosenOutputPath, logPath);
			if (outPath == null)
			{
				return;
			}
		}
		if (cancelled)
		{
			return;
		}
		if (failure == null)
		{
			await ReportOcrOutcomeAsync(ocrStats, logPath);
			try
			{
				if (outPath != null && File.Exists(outPath))
				{
					await Launcher.LaunchFileAsync(await StorageFile.GetFileFromPathAsync(outPath));
				}
				return;
			}
			catch
			{
				return;
			}
		}
		await new ContentDialog
		{
			Title = "Failed to generate drawing index",
			Content = failure.Message + "\n\nDetails were written to:\n" + logPath,
			CloseButtonText = "OK",
			XamlRoot = Content.XamlRoot
		}.ShowAsync();
	}

	private async void GenerateCSVButton_Click(object? sender, RoutedEventArgs e)
	{
		string templatePath = PdfPathTextBox.Text;
		if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
		{
			await new ContentDialog
			{
				Title = "No template selected",
				Content = (string.IsNullOrWhiteSpace(templatePath) ? "Please select a template PDF file before generating." : "The selected template PDF file could not be found. Please browse for it again."),
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
			return;
		}
		string combinedSetPath = (string.IsNullOrWhiteSpace(CombinedSetPathTextBox.Text) ? null : CombinedSetPathTextBox.Text);
		if (combinedSetPath != null && !File.Exists(combinedSetPath))
		{
			await new ContentDialog
			{
				Title = "Combined sheet set not found",
				Content = "The selected combined sheet set PDF could not be found. Please browse for it again, or clear it to scan the template's folder for individual sheet PDFs instead.",
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
			return;
		}
		string logPath = GetLogFilePath("export_log.txt");
		string normalizedTemplateTempFile = NormalizeTemplateRotationIfNeeded(templatePath, logPath);
		if (normalizedTemplateTempFile != null)
		{
			templatePath = normalizedTemplateTempFile;
		}
		if (!TryGetTemplateFieldRects(templatePath, out FieldRect drawingRect, out FieldRect disciplineRect, out FieldRect descriptionRect, out FieldRect? _, out string content))
		{
			await new ContentDialog
			{
				Title = "Template form fields not found",
				Content = content,
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
			if (normalizedTemplateTempFile == null)
			{
				return;
			}
			try
			{
				File.Delete(normalizedTemplateTempFile);
				return;
			}
			catch
			{
				return;
			}
		}
		try
		{
			File.AppendAllText(logPath, $"{DateTime.Now}: CSV generation started\n");
		}
		catch
		{
		}
		if (!await ConfirmOcrAvailableAsync(logPath))
		{
			if (normalizedTemplateTempFile != null)
			{
				FileHygiene.DeleteQuietly(normalizedTemplateTempFile);
			}
			return;
		}
		OcrRunStats ocrStats = new OcrRunStats();
		string outputPrefix = "Drawing Index";
		bool appendDT = true;
		string selectedFontFamily = GetSelectedFontFamily();
		bool shortGeneralNumbers = ShortGeneralNumbersCheckBox.IsChecked == true;
		string suggestedFileName = (appendDT ? $"{outputPrefix}{DateTime.Now:yyyyMMdd}" : outputPrefix);
		FileSavePicker fileSavePicker = new FileSavePicker
		{
			SuggestedFileName = suggestedFileName,
			FileTypeChoices = { 
			{
				"CSV File",
				(IList<string>)new List<string> { ".csv" }
			} }
		};
		nint windowHandle = WindowNative.GetWindowHandle(this);
		InitializeWithWindow.Initialize(fileSavePicker, windowHandle);
		StorageFile storageFile = await fileSavePicker.PickSaveFileAsync();
		if (storageFile == null)
		{
			if (normalizedTemplateTempFile == null)
			{
				return;
			}
			try
			{
				File.Delete(normalizedTemplateTempFile);
				return;
			}
			catch
			{
				return;
			}
		}
		string chosenOutputPath = storageFile.Path;
		string partialOutputPath = FileHygiene.PartialPathFor(chosenOutputPath);
		_generationCts = new CancellationTokenSource();
		_reviewHost = this;
		CancellationToken token = _generationCts.Token;
		Progress<(int current, int total, string status)> progress = new Progress<(int, int, string)>(((int current, int total, string status) p) =>
		{
			GenerationStatusText.Text = p.status;
			if (p.total > 0)
			{
				GenerationProgressBar.IsIndeterminate = false;
				GenerationProgressBar.Maximum = p.total;
				GenerationProgressBar.Value = p.current;
			}
			else
			{
				GenerationProgressBar.IsIndeterminate = true;
			}
		});
		GenerateButton.IsEnabled = false;
		GenerateCSVButton.IsEnabled = false;
		CancelGenerationButton.IsEnabled = true;
		PdfPathTextBox.IsEnabled = false;
		BrowsePdfButton.IsEnabled = false;
		CombinedSetPathTextBox.IsEnabled = false;
		BrowseCombinedSetButton.IsEnabled = false;
		ClearCombinedSetButton.IsEnabled = false;
		IncludeSheetsCheckBox.IsEnabled = false;
		IncludeHyperlinksCheckBox.IsEnabled = false;
		ShortGeneralNumbersCheckBox.IsEnabled = false;
		IncludeBookmarksCheckBox.IsEnabled = false;
		BookmarkSourceComboBox.IsEnabled = false;
		FontComboBox.IsEnabled = false;
		SettingsMenuBarItem.IsEnabled = false;
		GenerationProgressPanel.Visibility = Visibility.Visible;
		GenerationProgressBar.IsIndeterminate = true;
		GenerationStatusText.Text = "Starting...";
		FileHygiene.SweepStale(GetOrCreateOutputFolder(templatePath), "__AutoDIS_CSVTemp_*.pdf", TimeSpan.FromHours(1));
		Diagnostics.BeginRun("csv-index"); // DIAG
		var totalTimer = Diagnostics.Start(); // DIAG
		string outPath = null;
		bool cancelled = false;
		Exception failure = null;
		try
		{
			outPath = await Task.Run(() => GenerateDrawingIndexCSV(templatePath, drawingRect, disciplineRect, descriptionRect, logPath, progress, token, outputPrefix, appendDT, selectedFontFamily, combinedSetPath, partialOutputPath, normalizedTemplateTempFile != null, shortGeneralNumbers, ocrStats), token);
		}
		catch (OperationCanceledException)
		{
			cancelled = true;
			try
			{
				File.AppendAllText(logPath, $"{DateTime.Now}: Generation cancelled; no output file was written\n");
			}
			catch
			{
			}
		}
		catch (Exception ex2)
		{
			failure = ex2;
			try
			{
				File.AppendAllText(logPath, $"{DateTime.Now}: GenerateCSVButton_Click failed: {ex2}\n");
			}
			catch
			{
			}
		}
		finally
		{
			GenerateButton.IsEnabled = true;
			GenerateCSVButton.IsEnabled = true;
			CancelGenerationButton.IsEnabled = false;
			PdfPathTextBox.IsEnabled = true;
			BrowsePdfButton.IsEnabled = true;
			CombinedSetPathTextBox.IsEnabled = true;
			BrowseCombinedSetButton.IsEnabled = true;
			ClearCombinedSetButton.IsEnabled = true;
			IncludeSheetsCheckBox.IsEnabled = true;
			IncludeHyperlinksCheckBox.IsEnabled = true;
			ShortGeneralNumbersCheckBox.IsEnabled = true;
			IncludeBookmarksCheckBox.IsEnabled = true;
			BookmarkSourceComboBox.IsEnabled = IncludeBookmarksCheckBox.IsChecked == true;
			FontComboBox.IsEnabled = true;
			SettingsMenuBarItem.IsEnabled = true;
			GenerationProgressPanel.Visibility = Visibility.Collapsed;
			_generationCts?.Dispose();
			_generationCts = null;
			_reviewHost = null;
			if (_closeAfterRun) { Close(); }
			Diagnostics.Split("total", totalTimer); // DIAG
			Diagnostics.EndRun(); // DIAG
			if (normalizedTemplateTempFile != null)
			{
				try
				{
					File.Delete(normalizedTemplateTempFile);
				}
				catch
				{
				}
			}
		}
		if (cancelled || failure != null || outPath == null)
		{
			FileHygiene.DeleteQuietly(partialOutputPath);
		}
		else
		{
			outPath = await CommitOutputAsync(partialOutputPath, chosenOutputPath, logPath);
			if (outPath == null)
			{
				return;
			}
		}
		if (cancelled)
		{
			return;
		}
		if (failure == null)
		{
			await ReportOcrOutcomeAsync(ocrStats, logPath);
			try
			{
				if (outPath != null && File.Exists(outPath))
				{
					await Launcher.LaunchFileAsync(await StorageFile.GetFileFromPathAsync(outPath));
				}
				return;
			}
			catch
			{
				return;
			}
		}
		await new ContentDialog
		{
			Title = "Failed to generate CSV index",
			Content = failure.Message + "\n\nDetails were written to:\n" + logPath,
			CloseButtonText = "OK",
			XamlRoot = Content.XamlRoot
		}.ShowAsync();
	}

	private async Task<string?> CommitOutputAsync(string partialPath, string finalPath, string logPath)
	{
		try
		{
			FileHygiene.Commit(partialPath, finalPath);
			AppLog.TryAppend(logPath, $"{DateTime.Now}: Output written to {finalPath}");
			return finalPath;
		}
		catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
		{
			AppLog.TryAppend(logPath, $"{DateTime.Now}: Could not replace {finalPath}; finished output kept at {partialPath}: {ex.Message}");
			await new ContentDialog
			{
				Title = "Output file is locked",
				Content = $"{Path.GetFileName(finalPath)} could not be replaced, most likely because it is open in another program.\n\nThe finished index was kept as:\n{partialPath}\n\nClose the other program, then rename the kept file to the name you want (remove the .partial ending).",
				CloseButtonText = "OK",
				XamlRoot = Content.XamlRoot
			}.ShowAsync();
			return null;
		}
	}

	private void CancelGenerationButton_Click(object? sender, RoutedEventArgs e)
	{
		_generationCts?.Cancel();
		GenerationStatusText.Text = "Cancelling...";
	}

	// Large sheet sets live on a network share, where the many small reads and writes of building the index were the slowest part
	// of a run. The set is copied once, sequentially, to a local working folder and everything else works from that copy.
	private static (string? outPath, List<DrawingIndexItem>? finalRows) GenerateDrawingIndexCore(string P_0, FieldRect P_1, FieldRect P_2, FieldRect P_3, bool P_4, bool P_5, string P_6, IProgress<(int current, int total, string status)> P_7, CancellationToken P_8, string P_9 = "DrawingIndex_Output_", bool P_10 = true, string P_11 = "Arial", string? P_12 = null, string? P_13 = null, string? P_14 = null, bool P_15 = false, bool P_16 = true, bool P_17 = false, FieldRect? P_18 = null, bool P_19 = false, OcrRunStats? P_20 = null)
	{
		string? workFolder = null;
		ProcessPriorityClass? previousPriority = null;
		TimeSpan gcPause = TimeSpan.Zero;
		int gcGen2 = 0;
		long gcAllocated = 0;
		try
		{
			long setBytes = !string.IsNullOrWhiteSpace(P_13) && File.Exists(P_13) ? new FileInfo(P_13).Length : 0L;
			HardwareSnapshot hardware = HardwareProbe.Probe();
			ResourcePlan plan = ResourcePlanner.Plan(ResourceLevelSetting, hardware, setBytes);
			gcPause = GC.GetTotalPauseDuration();
			gcGen2 = GC.CollectionCount(2);
			gcAllocated = GC.GetTotalAllocatedBytes();
			_resourcePlan = plan;
			RenderGate = new SemaphoreSlim(plan.RenderConcurrency);
			_renderPool = new RenderPool(plan.RenderProcesses, () => new ProcessRenderConnection(m => AppLog.TryAppend(P_6, $"{DateTime.Now}: {m}")), m => AppLog.TryAppend(P_6, $"{DateTime.Now}: {m}"));
			AppLog.TryAppend(P_6, $"{DateTime.Now}: Resources - {ResourcePlanner.Summary(plan, hardware)}");
			try
			{
				using Process self = Process.GetCurrentProcess();
				previousPriority = self.PriorityClass;
				self.PriorityClass = plan.Priority;
			}
			catch (Exception ex) when (ex is InvalidOperationException || ex is System.ComponentModel.Win32Exception)
			{
				previousPriority = null;
			}
			if (!string.IsNullOrWhiteSpace(P_13) && File.Exists(P_13))
			{
				workFolder = Path.Combine(Path.GetTempPath(), "AutoDIS_work_" + Guid.NewGuid().ToString("N"));
				Directory.CreateDirectory(workFolder);
				string localCopy = Path.Combine(workFolder, Path.GetFileName(P_13));
				var copyTimer = Diagnostics.Start(); // DIAG
				string copyIn = CopyFileWithProgress(P_13, localCopy, P_8, P_7, "Copying the sheet set to a local working file");
				AppLog.TryAppend(P_6, $"{DateTime.Now}: Copied the sheet set to a local working file: {copyIn}");
				Diagnostics.Split("copy sheet set locally", copyTimer, Path.GetFileName(P_13) + " " + copyIn); // DIAG
				P_13 = localCopy;
			}
			return GenerateDrawingIndexCoreInner(P_0, P_1, P_2, P_3, P_4, P_5, P_6, P_7, P_8, P_9, P_10, P_11, P_12, P_13, P_14, P_15, P_16, P_17, P_18, P_19, P_20);
		}
		finally
		{
			RenderPool? pool = _renderPool;
			_renderPool = null;
			if (pool != null)
			{
				pool.Dispose();
				AppLog.TryAppend(P_6, $"{DateTime.Now}: Page-drawing helpers: {pool.ProcessesStarted} process(es) started, {pool.Rendered} page(s) drawn by them, {pool.Declined} drawn in the main program instead");
			}
			AppLog.TryAppend(P_6, $"{DateTime.Now}: Memory management during this run: {(GC.GetTotalPauseDuration() - gcPause).TotalSeconds:0.0} s paused for garbage collection, {GC.CollectionCount(2) - gcGen2} full collections, {(GC.GetTotalAllocatedBytes() - gcAllocated) / (1024 * 1024)} MB allocated");
			if (previousPriority.HasValue)
			{
				try
				{
					using Process self = Process.GetCurrentProcess();
					self.PriorityClass = previousPriority.Value;
				}
				catch (Exception ex) when (ex is InvalidOperationException || ex is System.ComponentModel.Win32Exception)
				{
				}
			}
			if (workFolder != null)
			{
				try
				{
					Directory.Delete(workFolder, recursive: true);
				}
				catch (IOException)
				{
				}
				catch (UnauthorizedAccessException)
				{
				}
			}
		}
	}
	private static (string? outPath, List<DrawingIndexItem>? finalRows) GenerateDrawingIndexCoreInner(string P_0, FieldRect P_1, FieldRect P_2, FieldRect P_3, bool P_4, bool P_5, string P_6, IProgress<(int current, int total, string status)> P_7, CancellationToken P_8, string P_9 = "DrawingIndex_Output_", bool P_10 = true, string P_11 = "Arial", string? P_12 = null, string? P_13 = null, string? P_14 = null, bool P_15 = false, bool P_16 = true, bool P_17 = false, FieldRect? P_18 = null, bool P_19 = false, OcrRunStats? P_20 = null)
	{
			OcrRunStats ocrStats = P_20 ?? new OcrRunStats();
		bool flag = !string.IsNullOrWhiteSpace(P_13) && File.Exists(P_13);
		var setupTimer = Diagnostics.Start(); // DIAG
		P_7.Report((0, 0, flag ? "Opening combined sheet set..." : "Scanning folder for sheet PDFs..."));
		string text = Path.GetDirectoryName(P_0) ?? string.Empty;
		if (!Directory.Exists(text))
		{
			return (outPath: null, finalRows: null);
		}
		byte[]? combinedBytes = flag ? File.ReadAllBytes(P_13) : null;
		using UglyToad.PdfPig.PdfDocument pdfDocument = (flag ? UglyToad.PdfPig.PdfDocument.Open(combinedBytes) : null);
		int num = (flag ? pdfDocument.NumberOfPages : 0);
		List<(string, int)> list;
		List<(string, int)> skippedFiles = new List<(string, int)>();
		if (flag)
		{
			list = (from i in Enumerable.Range(0, pdfDocument.NumberOfPages)
				select (P_13, i: i)).ToList();
			File.AppendAllText(P_6, $"{DateTime.Now}: Reading combined sheet set '{Path.GetFileName(P_13)}' ({pdfDocument.NumberOfPages} pages) directly, in page order\n");
		}
		else
		{
			string effectiveExcludePrefix = P_12 ?? P_9;
			List<string> list2 = (from p in Directory.EnumerateFiles(text, "*.pdf", SearchOption.TopDirectoryOnly)
				where !Path.GetFullPath(p).Equals(Path.GetFullPath(P_0), StringComparison.OrdinalIgnoreCase)
				where !Path.GetFileName(p).StartsWith(effectiveExcludePrefix, StringComparison.OrdinalIgnoreCase)
				select p).OrderBy((string p) => p, NaturalFileNameComparer.Instance).ToList();
			list = list2.Select((string p) => (p: p, 0)).ToList();
		}
		int progressFactor = (!P_4) ? 1 : 2;
		Diagnostics.Split("scan sources", setupTimer, $"{list.Count} candidate(s)"); // DIAG
		ProgressTracker progressTracker = new ProgressTracker(P_7, list.Count * progressFactor);
		OcrEngineSet? probeEngines = TryCreateOcrEngines(P_6, ocrStats, out OcrEngineSet? largeEngines);
		Diagnostics.Split("ocr engine init", setupTimer); // DIAG
		_templatePageSize = ReadPageSize(P_0);
		SheetRead[] sheetReads = ReadSheets(list, flag, combinedBytes, P_1, P_2, P_3, probeEngines, ocrStats, P_6, progressTracker, P_8);
		combinedBytes = null;
		pdfDocument?.Dispose();
		progressTracker.Status("Checking the index for patterns and likely mistakes...");
		bool willSecondLook = largeEngines != null && CurrentOcrPreset.SecondLook;
		var assembled = AssembleSheets(sheetReads, list, P_19, ocrStats, P_6, !willSecondLook);
		Diagnostics.Split("index analysis", setupTimer); // DIAG
		if (willSecondLook)
		{
			if (assembled.flagged.Count > 0)
			{
				SecondLook(assembled.flagged, list, P_1, P_2, P_3, largeEngines, ocrStats, P_6, progressTracker, P_8);
				Diagnostics.Split("second look (large model)", setupTimer, $"{assembled.flagged.Count} sheet(s)"); // DIAG
			}
			else
			{
				largeEngines!.Dispose();
			}
			assembled = AssembleSheets(sheetReads, list, P_19, ocrStats, P_6, true);
		}
		else
		{
			largeEngines?.Dispose();
		}
		List<DrawingIndexItem> list5 = assembled.rows;
		ReviewEntries(list5, P_2, P_3, P_6, P_7, P_8);
		List<(string, int)> sheets = assembled.sheets;
		skippedFiles.AddRange(assembled.skipped);		list = sheets;
		Diagnostics.Split("sheet loop (all sheets)", setupTimer, $"{sheets.Count} kept, {skippedFiles.Count} skipped"); // DIAG
		progressTracker.Total = list.Count * progressFactor;
		if (skippedFiles.Count > 0)
		{
			try
			{
				string text2 = string.Join(", ", skippedFiles.Select(((string name, int pages) s) => s.name + " (" + ((s.pages < 0) ? "unreadable" : (s.pages + " pages")) + ")"));
				File.AppendAllText(P_6, $"{DateTime.Now}: Skipped {skippedFiles.Count} non-single-page PDF(s) not treated as drawing sheets: {text2}\n");
			}
			catch
			{
			}
		}
		bool flag3 = false;
		try
		{
			using UglyToad.PdfPig.PdfDocument pdfDocument4 = UglyToad.PdfPig.PdfDocument.Open(P_0);
			UglyToad.PdfPig.Content.Page page2 = pdfDocument4.GetPages().FirstOrDefault();
			if (page2 != null)
			{
				double num5 = Math.Min(P_1.X1, Math.Min(P_2.X1, P_3.X1));
				double num6 = Math.Max(P_1.Y2, Math.Max(P_2.Y2, P_3.Y2));
				double num7 = Math.Min(P_1.Y2, Math.Min(P_2.Y2, P_3.Y2));
				double num8 = page2.Width * page2.Height;
				double num9 = 0.0;
				double num10 = 0.0;
				foreach (PdfPath path in page2.Paths)
				{
					if (!path.IsStroked)
					{
						continue;
					}
					UglyToad.PdfPig.Core.PdfRectangle? boundingRectangle = path.GetBoundingRectangle();
					if (!boundingRectangle.HasValue)
					{
						continue;
					}
					UglyToad.PdfPig.Core.PdfRectangle value = boundingRectangle.Value;
					if (!(value.Width * value.Height >= num8 * 0.5))
					{
						if (value.Width <= 20.0 && value.Height > num9 && Math.Abs(value.Left - num5) <= 60.0)
						{
							num9 = value.Height;
						}
						if (value.Height <= 20.0 && value.Width > num10 && value.Bottom >= num6 - 20.0 && value.Bottom <= num6 + 80.0)
						{
							num10 = value.Width;
						}
					}
				}
				double num11 = num9 / page2.Height;
				double num12 = num10 / page2.Width;
				flag3 = (!(num11 > 0.4) || !(num11 > num12)) && ((num12 > 0.4 && num12 >= num11) || num6 - num7 < page2.Height * 0.1);
			}
		}
		catch (Exception ex) when (ex is not OperationCanceledException)
		{
			AppLog.TryAppend(P_6, $"{DateTime.Now}: Title block flag detection failed: {ex.Message}");
		}
		double num13 = Math.Max(P_1.Y2, Math.Max(P_2.Y2, P_3.Y2));
		double num14 = (flag3 ? EstimateTitleBlockTopY(P_0, num13) : 0.0);
		Diagnostics.Split("title block scan", setupTimer, flag3 ? "title block detected" : "no title block"); // DIAG
		(double, double, double, double, double, double)? tuple2 = EstimateBorderRect(P_0);
		(double, double, double, double, double, double)? tuple3 = tuple2 ?? EstimateBorderRectFromBitmap(P_0, P_6);
		Diagnostics.Split("border detection", setupTimer, tuple2.HasValue ? "vector" : (tuple3.HasValue ? "bitmap fallback" : "not found")); // DIAG
		double num15 = EstimateLeftMarginX(P_0, num14);
		double num16 = EstimateTopMarginY(P_0, tuple3?.Item4);
		Diagnostics.Split("margin estimates", setupTimer); // DIAG
		double num17 = num15;
		if (tuple3.HasValue)
		{
			(double, double, double, double, double, double) value2 = tuple3.Value;
			num15 = Math.Max(num15, value2.Item1 + 15.0);
			num16 = Math.Max(num16, value2.Item6 - value2.Item4 + 15.0);
			num17 = Math.Max(num17, value2.Item5 - value2.Item3 + 15.0);
		}
		try
		{
			string text7;
			if (tuple2.HasValue)
			{
				text7 = "vector-path detection";
			}
			else
			{
				text7 = (tuple3.HasValue ? "bitmap dark-line scan fallback" : "none (vector single-rectangle, vector line-envelope, and bitmap dark-line scan all failed)");
			}
			File.AppendAllText(P_6, tuple3.HasValue ? $"{DateTime.Now}: Border detected via {text7}: Left={tuple3.Value.Item1:0.##}, Bottom={tuple3.Value.Item2:0.##}, Right={tuple3.Value.Item3:0.##}, Top={tuple3.Value.Item4:0.##}\n" : $"{DateTime.Now}: No border detected -- {text7}\n");
		}
		catch
		{
		}
		double num19;
		double num20;
		if (flag3)
		{
			double num18 = num14 + 30.0;
			if (tuple3.HasValue)
			{
				num18 = Math.Max(num18, tuple3.Value.Item2 + 30.0);
			}
			num19 = num18;
			num20 = num18;
		}
		else
		{
			double num21 = (tuple3.HasValue ? (tuple3.Value.Item2 + 30.0) : 30.0);
			double num22 = num13 + 20.0;
			num19 = num21;
			num20 = Math.Max(num21, num22);
		}
		double? num23 = (flag3 ? ((double?)null) : new double?(Math.Min(P_1.X1, Math.Min(P_2.X1, P_3.X1))));
		double num24 = num16;
		double num25 = num19;
		double num26 = num20;
		double num27 = num15;
		double num28 = num17;
		try
		{
			File.AppendAllText(P_6, $"{DateTime.Now}: Estimated title block top Y={num14:0.##} (fields max Y2={num13:0.##}); effective margins top={num24:0.##}, bottom(left)={num25:0.##}, bottom(right)={num26:0.##}, left={num27:0.##}, right={num28:0.##}\n");
		}
		catch
		{
		}
		P_8.ThrowIfCancellationRequested();
		string text8 = P_14 ?? Path.Combine(GetOrCreateOutputFolder(P_0), P_10 ? (P_9 ?? "") : (P_9 + ".pdf"));
		bool flag4 = list.Count > 0;
		var exportTimer = Diagnostics.Start(); // DIAG
		ExportResultsToPdf(list5, P_0, text8, out List<DrawingIndexItem> list9, flag3, num25, num27, list, P_4, P_5, num24, num28, num26, progressTracker, P_8, num23, P_11, flag4, P_13, P_15, P_16, P_18, num, P_19);
		Diagnostics.Split("export pdf", exportTimer); // DIAG
		return (outPath: text8, finalRows: list9);
	}

	private static string ToRomanNumeral(int P_0)
	{
		if (P_0 <= 0)
		{
			return P_0.ToString();
		}
		(int, string)[] array = new (int, string)[13]
		{
			(1000, "M"),
			(900, "CM"),
			(500, "D"),
			(400, "CD"),
			(100, "C"),
			(90, "XC"),
			(50, "L"),
			(40, "XL"),
			(10, "X"),
			(9, "IX"),
			(5, "V"),
			(4, "IV"),
			(1, "I")
		};
		StringBuilder stringBuilder = new StringBuilder();
		(int, string)[] array2 = array;
		for (int i = 0; i < array2.Length; i++)
		{
			var (num, text) = array2[i];
			while (P_0 >= num)
			{
				stringBuilder.Append(text);
				P_0 -= num;
			}
		}
		return stringBuilder.ToString();
	}

	private static string GeneralDrawingNumber(int P_0, bool P_1)
	{
		if (P_1)
		{
			return $"G–{P_0:00}";
		}
		return $"00–G–{P_0:00}";
	}

	private static List<DrawingIndexItem> BuildRowsWithIndexPageEntries(List<DrawingIndexItem> P_0, int P_1, bool P_2 = false)
	{
		List<DrawingIndexItem> list = new List<DrawingIndexItem>();
		if (P_0.Count > 0)
		{
			list.Add(P_0[0]);
		}
		for (int i = 1; i <= P_1; i++)
		{
			string title = ((P_1 == 1) ? "DRAWING INDEX" : ("DRAWING INDEX - " + ToRomanNumeral(i)));
			list.Add(new DrawingIndexItem
			{
				DrawingNumber = GeneralDrawingNumber(i, P_2),
				Discipline = "GENERAL",
				Title = title,
				SourcePdfPath = string.Empty
			});
		}
		for (int j = 1; j < P_0.Count; j++)
		{
			list.Add(P_0[j]);
		}
		int num = 1;
		foreach (DrawingIndexItem item in list)
		{
			item.SeqNo = num++;
		}
		return list;
	}

	private static string? NormalizeTemplateRotationIfNeeded(string P_0, string P_1)
	{
		try
		{
			List<(string, double, double, double, double)> list = new List<(string, double, double, double, double)>();
			int num;
			double num2;
			double num3;
			using (PdfSharp.Pdf.PdfDocument pdfDocument = PdfReader.Open(P_0, PdfDocumentOpenMode.Import))
			{
				if (pdfDocument.PageCount == 0)
				{
					return null;
				}
				PdfSharp.Pdf.PdfPage pdfPage = pdfDocument.Pages[0];
				num = (pdfPage.Elements.GetInteger("/Rotate") % 360 + 360) % 360;
				if (num == 0)
				{
					return null;
				}
				PdfSharp.Pdf.PdfRectangle mediaBox = pdfPage.MediaBox;
				num2 = mediaBox.X2 - mediaBox.X1;
				num3 = mediaBox.Y2 - mediaBox.Y1;
				PdfItem pdfItem = pdfPage.Elements["/Annots"];
				PdfArray pdfArray = (pdfItem as PdfArray) ?? ((pdfItem as PdfReference)?.Value as PdfArray);
				if (pdfArray != null)
				{
					foreach (PdfItem element in pdfArray.Elements)
					{
						PdfDictionary pdfDictionary = ((element as PdfReference)?.Value as PdfDictionary) ?? (element as PdfDictionary);
						if (pdfDictionary == null || !string.Equals(pdfDictionary.Elements.GetName("/Subtype"), "/Widget", StringComparison.OrdinalIgnoreCase))
						{
							continue;
						}
						string text = pdfDictionary.Elements.GetString("/T");
						if (!string.IsNullOrWhiteSpace(text))
						{
							PdfItem pdfItem2 = pdfDictionary.Elements["/Rect"];
							PdfArray pdfArray2 = (pdfItem2 as PdfArray) ?? ((pdfItem2 as PdfReference)?.Value as PdfArray);
							if (pdfArray2 != null && pdfArray2.Elements.Count >= 4)
							{
								list.Add((text, pdfArray2.Elements.GetReal(0), pdfArray2.Elements.GetReal(1), pdfArray2.Elements.GetReal(2), pdfArray2.Elements.GetReal(3)));
							}
						}
					}
				}
			}
			double a;
			double b;
			double c;
			double d;
			double e;
			double f;
			double num4;
			double num5;
			switch (num)
			{
			case 90:
				a = 0.0;
				b = -1.0;
				c = 1.0;
				d = 0.0;
				e = 0.0;
				f = num2;
				num4 = num3;
				num5 = num2;
				break;
			case 180:
				a = -1.0;
				b = 0.0;
				c = 0.0;
				d = -1.0;
				e = num2;
				f = num3;
				num4 = num2;
				num5 = num3;
				break;
			case 270:
				a = 0.0;
				b = 1.0;
				c = -1.0;
				d = 0.0;
				e = num3;
				f = 0.0;
				num4 = num3;
				num5 = num2;
				break;
			default:
				File.AppendAllText(P_1, $"{DateTime.Now}: Template page has unexpected /Rotate={num} (not a multiple of 90); leaving it as-is.\n");
				return null;
			}
			(double x, double y) TransformPoint(double x, double y)
			{
				return (a * x + c * y + e, b * x + d * y + f);
			}
			XPdfForm image = XPdfForm.FromFile(P_0);
			PdfSharp.Pdf.PdfDocument pdfDocument2 = new PdfSharp.Pdf.PdfDocument();
			PdfSharp.Pdf.PdfPage pdfPage2 = pdfDocument2.AddPage();
			pdfPage2.Width = XUnit.FromPoint(num4);
			pdfPage2.Height = XUnit.FromPoint(num5);
			using (XGraphics xGraphics = XGraphics.FromPdfPage(pdfPage2))
			{
				xGraphics.DrawImage(image, 0.0, 0.0, num4, num5);
			}
			if (list.Count > 0)
			{
				PdfArray pdfArray3 = new PdfArray(pdfDocument2);
				foreach (var item5 in list)
				{
					(double x, double y) tuple = TransformPoint(item5.Item2, item5.Item3);
					double item = tuple.x;
					double item2 = tuple.y;
					(double x, double y) tuple2 = TransformPoint(item5.Item4, item5.Item5);
					double item3 = tuple2.x;
					double item4 = tuple2.y;
					PdfDictionary pdfDictionary2 = new PdfDictionary(pdfDocument2);
					pdfDocument2.Internals.AddObject(pdfDictionary2);
					pdfDictionary2.Elements.SetName("/Type", "/Annot");
					pdfDictionary2.Elements.SetName("/Subtype", "/Widget");
					pdfDictionary2.Elements.SetString("/T", item5.Item1);
					PdfArray pdfArray4 = new PdfArray(pdfDocument2);
					pdfArray4.Elements.Add(new PdfReal(Math.Min(item, item3)));
					pdfArray4.Elements.Add(new PdfReal(Math.Min(item2, item4)));
					pdfArray4.Elements.Add(new PdfReal(Math.Max(item, item3)));
					pdfArray4.Elements.Add(new PdfReal(Math.Max(item2, item4)));
					pdfDictionary2.Elements["/Rect"] = pdfArray4;
					pdfArray3.Elements.Add(pdfDictionary2.Reference);
				}
				pdfPage2.Elements["/Annots"] = pdfArray3;
			}
			string text2 = Path.Combine(Path.GetTempPath(), $"AutoDIS_TemplateNormalized_{Guid.NewGuid():N}.pdf");
			pdfDocument2.Save(text2);
			File.AppendAllText(P_1, $"{DateTime.Now}: Template page had /Rotate={num}; normalized into a rotation-baked temp copy at {text2} ({list.Count} form field widget(s) carried over)\n");
			return text2;
		}
		catch (Exception ex)
		{
			try
			{
				File.AppendAllText(P_1, $"{DateTime.Now}: Failed to normalize template rotation (continuing with the original template as-is): {ex.Message}\n");
			}
			catch
			{
			}
			return null;
		}
	}

	private static (PdfSharp.Pdf.PdfDocument doc, string tempCopyToDelete) OpenTemplateForImport(string P_0, string P_1)
	{
		try
		{
			return (doc: PdfReader.Open(P_0, PdfDocumentOpenMode.Import), tempCopyToDelete: string.Empty);
		}
		catch (Exception ex) when (ex.Message.Contains("Cannot retrieve stream length"))
		{
			string text = Path.Combine(Path.GetTempPath(), $"AutoDIS_Template_{Guid.NewGuid():N}.pdf");
			File.Copy(P_0, text, true);
			if (TryRepairMissingStreamLengths(text, P_1))
			{
				return (doc: PdfReader.Open(text, PdfDocumentOpenMode.Import), tempCopyToDelete: text);
			}
			try
			{
				File.Delete(text);
			}
			catch
			{
			}
			throw;
		}
	}

	private static List<BookmarkNode> ReadBookmarkTree(PdfSharp.Pdf.PdfDocument P_0, string P_1, int P_2 = 20000)
	{
		List<BookmarkNode> list = new List<BookmarkNode>();
		try
		{
			PdfItem pdfItem = P_0.Internals.Catalog.Elements["/Outlines"];
			PdfDictionary pdfDictionary = (pdfItem as PdfDictionary) ?? ((pdfItem as PdfReference)?.Value as PdfDictionary);
			if (pdfDictionary == null)
			{
				return list;
			}
			Dictionary<string, int> pageIndexByObjId = new Dictionary<string, int>();
			for (int i = 0; i < P_0.PageCount; i++)
			{
				PdfReference reference = P_0.Pages[i].Reference;
				if (reference != null)
				{
					pageIndexByObjId[reference.ObjectID.ToString()] = i;
				}
			}
			int nodeCount = 0;
			PdfItem pdfItem2 = pdfDictionary.Elements["/First"];
			PdfDictionary pdfDictionary2 = (pdfItem2 as PdfDictionary) ?? ((pdfItem2 as PdfReference)?.Value as PdfDictionary);
			while (pdfDictionary2 != null && nodeCount <= P_2)
			{
				BookmarkNode bookmarkNode = ReadItem(pdfDictionary2);
				if (bookmarkNode != null)
				{
					list.Add(bookmarkNode);
				}
				PdfItem pdfItem3 = pdfDictionary2.Elements["/Next"];
				pdfDictionary2 = (pdfItem3 as PdfDictionary) ?? ((pdfItem3 as PdfReference)?.Value as PdfDictionary);
			}
			File.AppendAllText(P_1, $"Read {nodeCount} bookmark node(s) from combined set's outline tree\n");
			BookmarkNode? ReadItem(PdfDictionary pdfDictionary3)
			{
				if (nodeCount++ > P_2)
				{
					return null;
				}
				BookmarkNode bookmarkNode2 = new BookmarkNode
				{
					Title = (pdfDictionary3.Elements.GetString("/Title") ?? string.Empty),
					SourcePageIndex = ResolvePageIndex(pdfDictionary3)
				};
				PdfItem pdfItem4 = pdfDictionary3.Elements["/First"];
				PdfDictionary pdfDictionary4 = (pdfItem4 as PdfDictionary) ?? ((pdfItem4 as PdfReference)?.Value as PdfDictionary);
				while (pdfDictionary4 != null && nodeCount <= P_2)
				{
					BookmarkNode bookmarkNode3 = ReadItem(pdfDictionary4);
					if (bookmarkNode3 != null)
					{
						bookmarkNode2.Children.Add(bookmarkNode3);
					}
					PdfItem pdfItem5 = pdfDictionary4.Elements["/Next"];
					pdfDictionary4 = (pdfItem5 as PdfDictionary) ?? ((pdfItem5 as PdfReference)?.Value as PdfDictionary);
				}
				return bookmarkNode2;
			}
			int? ResolvePageIndex(PdfDictionary pdfDictionary3)
			{
				PdfItem pdfItem4 = pdfDictionary3.Elements["/Dest"];
				if (pdfItem4 == null)
				{
					PdfItem pdfItem5 = pdfDictionary3.Elements["/A"];
					PdfDictionary pdfDictionary4 = (pdfItem5 as PdfDictionary) ?? ((pdfItem5 as PdfReference)?.Value as PdfDictionary);
					if (pdfDictionary4 != null)
					{
						pdfItem4 = pdfDictionary4.Elements["/D"];
					}
				}
				PdfArray pdfArray = (pdfItem4 as PdfArray) ?? ((pdfItem4 as PdfReference)?.Value as PdfArray);
				if (pdfArray == null || pdfArray.Elements.Count == 0)
				{
					return null;
				}
				if (!(pdfArray.Elements[0] is PdfReference pdfReference))
				{
					return null;
				}
				if (!pageIndexByObjId.TryGetValue(pdfReference.ObjectID.ToString(), out var value))
				{
					return null;
				}
				return value;
			}
		}
		catch (Exception ex)
		{
			try
			{
				File.AppendAllText(P_1, "Failed to read bookmarks from combined set: " + ex.Message + "\n");
			}
			catch
			{
			}
		}
		return list;
	}

	private static int WriteBookmarkTree(PdfSharp.Pdf.PdfDocument P_0, List<BookmarkNode> P_1, string P_2, Dictionary<(string Path, int PageIndex), int> P_3, int? P_4 = null, List<(string Title, int OutputPageIndex)>? P_5 = null)
	{
		int totalAdded = 0;
		List<BookmarkNode> list = ((P_5 != null) ? FilterOutStaleGeneralNodes(P_1) : P_1);
		List<BookmarkNode> freshNodes;
		if (P_5 != null && P_5.Count > 0)
		{
			freshNodes = P_5.Select(((string Title, int OutputPageIndex) e) => new BookmarkNode
			{
				Title = e.Title,
				ForcedOutputPageIndex = e.OutputPageIndex
			}).ToList();
			if (!TryInsertIntoGeneralFolder(list))
			{
				list.InsertRange(0, freshNodes);
			}
		}
		PdfDictionary pdfDictionary = new PdfDictionary(P_0);
		P_0.Internals.AddObject(pdfDictionary);
		PdfDictionary pdfDictionary2 = null;
		PdfDictionary pdfDictionary3 = null;
		PdfDictionary pdfDictionary4 = null;
		int num = 0;
		foreach (BookmarkNode item in list)
		{
			PdfDictionary pdfDictionary5 = BuildItem(item, pdfDictionary.Reference);
			if (pdfDictionary2 == null)
			{
				pdfDictionary2 = pdfDictionary5;
			}
			if (pdfDictionary4 != null)
			{
				pdfDictionary4.Elements["/Next"] = pdfDictionary5.Reference;
				pdfDictionary5.Elements["/Prev"] = pdfDictionary4.Reference;
			}
			pdfDictionary3 = pdfDictionary5;
			pdfDictionary4 = pdfDictionary5;
			num++;
			totalAdded++;
		}
		if (pdfDictionary2 == null)
		{
			return 0;
		}
		pdfDictionary.Elements["/First"] = pdfDictionary2.Reference;
		pdfDictionary.Elements["/Last"] = pdfDictionary3.Reference;
		pdfDictionary.Elements.SetInteger("/Count", num);
		P_0.Internals.Catalog.Elements["/Outlines"] = pdfDictionary.Reference;
		return totalAdded;
		PdfDictionary BuildItem(BookmarkNode bookmarkNode, PdfReference value)
		{
			PdfDictionary pdfDictionary6 = new PdfDictionary(P_0);
			P_0.Internals.AddObject(pdfDictionary6);
			pdfDictionary6.Elements.SetString("/Title", bookmarkNode.Title);
			pdfDictionary6.Elements["/Parent"] = value;
			int? num2 = ResolveOutputPageIndex(bookmarkNode);
			if (num2.HasValue && num2.Value < P_0.Pages.Count)
			{
				PdfArray value2 = new PdfArray(P_0)
				{
					Elements = 
					{
						(PdfItem)P_0.Pages[num2.Value].Reference,
						(PdfItem)new PdfName("/Fit")
					}
				};
				pdfDictionary6.Elements["/Dest"] = value2;
			}
			PdfDictionary pdfDictionary7 = null;
			PdfDictionary pdfDictionary8 = null;
			PdfDictionary pdfDictionary9 = null;
			int num3 = 0;
			foreach (BookmarkNode child in bookmarkNode.Children)
			{
				PdfDictionary pdfDictionary10 = BuildItem(child, pdfDictionary6.Reference);
				if (pdfDictionary7 == null)
				{
					pdfDictionary7 = pdfDictionary10;
				}
				if (pdfDictionary9 != null)
				{
					pdfDictionary9.Elements["/Next"] = pdfDictionary10.Reference;
					pdfDictionary10.Elements["/Prev"] = pdfDictionary9.Reference;
				}
				pdfDictionary8 = pdfDictionary10;
				pdfDictionary9 = pdfDictionary10;
				num3++;
				totalAdded++;
			}
			if (pdfDictionary7 != null)
			{
				pdfDictionary6.Elements["/First"] = pdfDictionary7.Reference;
				pdfDictionary6.Elements["/Last"] = pdfDictionary8.Reference;
				pdfDictionary6.Elements.SetInteger("/Count", num3);
			}
			return pdfDictionary6;
		}
		static List<BookmarkNode> FilterOutStaleGeneralNodes(List<BookmarkNode> list3)
		{
			List<BookmarkNode> list2 = new List<BookmarkNode>();
			foreach (BookmarkNode item2 in list3)
			{
				if (!IsCoverOrIndexTitle(item2.Title))
				{
					item2.Children = FilterOutStaleGeneralNodes(item2.Children);
					list2.Add(item2);
				}
			}
			return list2;
		}
		static bool IsCoverOrIndexTitle(string text2)
		{
			string text = text2.Trim();
			if (!string.Equals(text, "COVER SHEET", StringComparison.OrdinalIgnoreCase))
			{
				return text.Contains("DRAWING INDEX", StringComparison.OrdinalIgnoreCase);
			}
			return true;
		}
		int? ResolveOutputPageIndex(BookmarkNode bookmarkNode)
		{
			if (bookmarkNode.ForcedOutputPageIndex.HasValue)
			{
				return bookmarkNode.ForcedOutputPageIndex;
			}
			if (!bookmarkNode.SourcePageIndex.HasValue)
			{
				return null;
			}
			if (P_4.HasValue && bookmarkNode.SourcePageIndex.Value == P_4.Value)
			{
				return 0;
			}
			if (P_3.TryGetValue((P_2, bookmarkNode.SourcePageIndex.Value), out var value) && value >= 0)
			{
				return value;
			}
			return null;
		}
		bool TryInsertIntoGeneralFolder(List<BookmarkNode> list2)
		{
			foreach (BookmarkNode item3 in list2)
			{
				if (string.Equals(item3.Title.Trim(), "GENERAL", StringComparison.OrdinalIgnoreCase))
				{
					item3.Children.InsertRange(0, freshNodes);
					return true;
				}
				if (TryInsertIntoGeneralFolder(item3.Children))
				{
					return true;
				}
			}
			return false;
		}
	}

	private static List<BookmarkNode> BuildBookmarkTreeFromGeneratedIndex(List<DrawingIndexItem> P_0, Dictionary<(string Path, int PageIndex), int> P_1, List<DrawingIndexItem> P_2, bool P_3)
	{
		List<BookmarkNode> list = new List<BookmarkNode>();
		DrawingIndexItem drawingIndexItem = ((P_3 && P_0.Count > 0) ? P_0[0] : null);
		foreach (var item in ChunkConsecutiveByDiscipline(P_0))
		{
			BookmarkNode bookmarkNode = new BookmarkNode
			{
				Title = item.Key
			};
			foreach (DrawingIndexItem item2 in item.Items)
			{
				int? forcedOutputPageIndex;
				if (drawingIndexItem != null && item2 == drawingIndexItem)
				{
					forcedOutputPageIndex = 0;
				}
				else if (!(item2.Discipline == "GENERAL") || !(item2.SourcePdfPath == string.Empty) || !(item2.Title ?? string.Empty).StartsWith("DRAWING INDEX"))
				{
					forcedOutputPageIndex = ((string.IsNullOrEmpty(item2.SourcePdfPath) || !P_1.TryGetValue((item2.SourcePdfPath, item2.SourcePageIndex), out var value) || value < 0) ? ((int?)null) : new int?(value));
				}
				else
				{
					int num = P_2.IndexOf(item2);
					forcedOutputPageIndex = ((num >= 0) ? new int?(num + 1) : ((int?)null));
				}
				if (forcedOutputPageIndex.HasValue)
				{
					string title = (string.IsNullOrWhiteSpace(item2.Title) ? (item2.DrawingNumber ?? string.Empty) : (item2.DrawingNumber + " - " + item2.Title));
					bookmarkNode.Children.Add(new BookmarkNode
					{
						Title = title,
						ForcedOutputPageIndex = forcedOutputPageIndex
					});
				}
			}
			if (bookmarkNode.Children.Count > 0)
			{
				list.Add(bookmarkNode);
			}
		}
		return list;
	}

	private static void ExportResultsToPdf(IEnumerable<DrawingIndexItem> P_0, string P_1, string P_2, out List<DrawingIndexItem> P_3, bool P_4 = false, double P_5 = 48.0, double P_6 = 40.0, IEnumerable<(string Path, int PageIndex)>? P_7 = null, bool P_8 = false, bool P_9 = false, double P_10 = 20.0, double P_11 = 40.0, double? P_12 = null, ProgressTracker? P_13 = null, CancellationToken P_14 = default(CancellationToken), double? P_15 = null, string P_16 = "Arial", bool P_17 = false, string? P_18 = null, bool P_19 = false, bool P_20 = true, FieldRect? P_21 = null, int P_22 = 0, bool P_23 = false)
	{
		P_3 = (P_0 as List<DrawingIndexItem>) ?? P_0.ToList();
		var xTimer = Diagnostics.Start(); // DIAG
		string logPath = GetLogFilePath("export_log.txt");
		PdfSharp.Pdf.PdfDocument outDoc = null;
		try
		{
			File.AppendAllText(logPath, $"ExportResultsToPdf started at {DateTime.Now}\n");
			if (GlobalFontSettings.FontResolver == null)
			{
				GlobalFontSettings.FontResolver = new SystemFontResolver();
			}
			List<DrawingIndexItem> list = P_3;
			bool flag = P_17 && list.Count > 0 && !string.IsNullOrEmpty(list[0].SourcePdfPath) && File.Exists(list[0].SourcePdfPath);
			string coverSheetPath = (flag ? list[0].SourcePdfPath : null);
			int coverSheetPageIndex = (flag ? list[0].SourcePageIndex : 0);
			RunWithFileLockRetry(() =>
			{
				File.Copy(P_1, P_2, true);
			}, logPath, "Copy template to " + P_2);
			File.AppendAllText(logPath, "Copied template to " + P_2 + "\n");
			Diagnostics.Split("copy template", xTimer); // DIAG
			try
			{
				outDoc = PdfReader.Open(P_2, PdfDocumentOpenMode.Modify);
			}
			catch (Exception ex)
			{
				File.AppendAllText(logPath, "Failed to open copied output for modify: " + ex?.ToString() + "\n");
				bool flag2 = false;
				try
				{
					flag2 = TryRepairMissingStreamLengths(P_2, logPath);
				}
				catch (Exception ex2)
				{
					File.AppendAllText(logPath, "TryRepairMissingStreamLengths threw: " + ex2?.ToString() + "\n");
				}
				if (flag2)
				{
					try
					{
						outDoc = PdfReader.Open(P_2, PdfDocumentOpenMode.Modify);
						File.AppendAllText(logPath, "Reopened repaired output successfully\n");
					}
					catch (Exception ex3)
					{
						File.AppendAllText(logPath, "Reopen after repair failed: " + ex3?.ToString() + "\n");
					}
				}
			}
			if (outDoc == null)
			{
				try
				{
					File.AppendAllText(logPath, "Attempting fallback render (overlay-only)\n");
					CreateOverlayOnlyPdf(P_1, P_2, P_0, logPath, P_16);
					return;
				}
				catch (Exception ex4)
				{
					File.AppendAllText(logPath, "Fallback failed: " + ex4?.ToString() + "\n");
					return;
				}
			}
			File.AppendAllText(logPath, $"Opened output with {outDoc.Pages.Count} pages\n");
			Diagnostics.Split("open output (modify)", xTimer); // DIAG
			if (outDoc.Pages.Count == 0)
			{
				File.AppendAllText(logPath, "No pages to modify\n");
				RunWithFileLockRetry(() =>
				{
					outDoc.Save(P_2);
				}, logPath, "Save " + P_2 + " (no pages)");
				return;
			}
			if (flag)
			{
				try
				{
					var (pdfDocument, text) = OpenTemplateForImport(coverSheetPath, logPath);
					Diagnostics.Split("open cover template", xTimer); // DIAG
					try
					{
						int index = Math.Max(0, Math.Min(coverSheetPageIndex, pdfDocument.PageCount - 1));
						if (outDoc.Pages.Count > 0)
						{
							outDoc.Pages.RemoveAt(0);
						}
						outDoc.InsertPage(0, pdfDocument.Pages[index]);
						File.AppendAllText(logPath, "Inserted cover sheet as page 1\n");
					}
					finally
					{
						pdfDocument.Dispose();
						if (!string.IsNullOrEmpty(text))
						{
							try
							{
								File.Delete(text);
							}
							catch
							{
							}
						}
					}
				}
				catch (Exception ex5)
				{
					File.AppendAllText(logPath, "Failed to insert cover sheet: " + ex5.Message + "\n");
					flag = false;
				}
			}
			PdfSharp.Pdf.PdfDocument pdfDocument2 = null;
			Diagnostics.Split("cover sheet insert", xTimer); // DIAG
			string text2 = string.Empty;
			PdfSharp.Pdf.PdfPage pdfPage;
			if (flag)
			{
				(PdfSharp.Pdf.PdfDocument doc, string tempCopyToDelete) tuple2 = OpenTemplateForImport(P_1, logPath);
				pdfDocument2 = tuple2.doc;
				text2 = tuple2.tempCopyToDelete;
				Diagnostics.Split("open template for import", xTimer); // DIAG
				pdfPage = outDoc.AddPage(pdfDocument2.Pages[0]);
				File.AppendAllText(logPath, "Cover sheet placed as page 1; index table placed on imported template page 2\n");
			}
			else
			{
				pdfPage = outDoc.Pages[0];
			}
			double point = pdfPage.Width.Point;
			double point2 = pdfPage.Height.Point;
			double sizeScale = Math.Min(1.0, point / 2448.0);
			int num = Math.Max(6, (int)Math.Round(12.0 * sizeScale));
			int num2 = Math.Max(6, (int)Math.Round(12.0 * sizeScale));
			double tableAreaWidth = GetTableAreaWidth(point, point2, P_4, P_15);
			double num3 = 12.0 * sizeScale;
			double num4 = Math.Max(220.0 * sizeScale, (tableAreaWidth - P_6 - P_11 - num3) / 2.0);
			double num5 = 90.0 * sizeScale;
			double num6 = 110.0 * sizeScale;
			double num7 = Math.Max(160.0 * sizeScale, num4 - (num5 + num6));
			try
			{
				using PdfSharp.Pdf.PdfDocument pdfDocument3 = new PdfSharp.Pdf.PdfDocument();
				PdfSharp.Pdf.PdfPage pdfPage2 = pdfDocument3.AddPage();
				pdfPage2.Width = XUnit.FromPoint(point);
				pdfPage2.Height = XUnit.FromPoint(point2);
				using XGraphics xGraphics = XGraphics.FromPdfPage(pdfPage2);
				XFont font = new XFont(P_16, num);
				double num8 = xGraphics.MeasureString("SEQ.\u00a0NO.", font).Width + 8.0 * sizeScale;
				double num9 = xGraphics.MeasureString("DWG.\u00a0NO.", font).Width + 8.0 * sizeScale;
				if (num8 > num5)
				{
					num5 = num8;
				}
				if (num9 > num6)
				{
					num6 = num9;
				}
				if (num6 > 140.0 * sizeScale)
				{
					num6 = 140.0 * sizeScale;
				}
				if (num5 + num6 + 80.0 * sizeScale > num4)
				{
					num4 = num5 + num6 + 80.0 * sizeScale;
				}
				XFont font2 = new XFont(P_16, num2);
				double num10 = 0.0;
				foreach (DrawingIndexItem item in P_0)
				{
					double width = xGraphics.MeasureString(item.Title ?? string.Empty, font2).Width;
					if (width > num10)
					{
						num10 = width;
					}
				}
				double num11 = num10 + 16.0 * sizeScale;
				num7 = Math.Min(Math.Max(160.0 * sizeScale, num4 - (num5 + num6)), Math.Max(160.0 * sizeScale, num11));
			}
			catch
			{
			}
			double[] array = new double[3] { num5, num6, num7 };
			double num12 = P_10 + GetDrawingIndexTitleReservedHeight();
			double num13 = num5 + num6 + num7;
			double num14 = Math.Max(0.0, tableAreaWidth - P_6 - P_11);
			double num15 = Math.Max((P_4 ? 40.0 : 12.0) * sizeScale, num14 - 2.0 * num13);
			double value = P_12 ?? P_5;
			if (P_15.HasValue && P_6 + 2.0 * num13 + num15 <= P_15.Value)
			{
				value = P_5;
			}
			List<DrawingIndexItem> list2 = list;
			TablePagination tablePagination2;
			if (flag)
			{
				int num16 = 1;
				TablePagination tablePagination = null;
				List<DrawingIndexItem> list3 = null;
				for (int num17 = 0; num17 < 4; num17++)
				{
					list3 = BuildRowsWithIndexPageEntries(list, num16, P_23);
					tablePagination = BuildTablePagination(list3, point, point2, num12, P_5, array, num, num2, P_16, value);
					if (tablePagination.Pages.Count == num16)
					{
						break;
					}
					num16 = tablePagination.Pages.Count;
				}
				list2 = list3;
				tablePagination2 = tablePagination;
				File.AppendAllText(logPath, $"Added {num16} \"DRAWING INDEX\" GENERAL row(s) for the index page(s)\n");
			}
			else
			{
				tablePagination2 = BuildTablePagination(list2, point, point2, num12, P_5, array, num, num2, P_16, value);
			}
			P_3 = list2;
			Diagnostics.Split("table pagination", xTimer, $"{tablePagination2.Pages.Count} index page(s)"); // DIAG
			int sheetNumberTotalSheets = ((list2.Count > 0) ? list2.Max((DrawingIndexItem r) => r.SeqNo) : P_22);
			bool stampSheetNumbers = P_21.HasValue && P_22 > 0;
			List<DrawingIndexItem> indexPageEntries = (flag ? list2.Where((DrawingIndexItem r) => r.Discipline == "GENERAL" && r.SourcePdfPath == string.Empty && (r.Title ?? string.Empty).StartsWith("DRAWING INDEX")).ToList() : new List<DrawingIndexItem>());
			List<(string, int)> list4 = null;
			if (flag)
			{
				list4 = new List<(string, int)>();
				if (list2.Count > 0)
				{
					DrawingIndexItem drawingIndexItem = list2[0];
					list4.Add((drawingIndexItem.DrawingNumber + " - " + drawingIndexItem.Title, 0));
				}
				for (int num18 = 0; num18 < indexPageEntries.Count; num18++)
				{
					DrawingIndexItem drawingIndexItem2 = indexPageEntries[num18];
					list4.Add((drawingIndexItem2.DrawingNumber + " - " + drawingIndexItem2.Title, num18 + 1));
				}
			}
			ApplyIndexPageFormFields(pdfPage, 1);
			Dictionary<DrawingIndexItem, (PdfSharp.Pdf.PdfPage, PdfDictionary)> dictionary = new Dictionary<DrawingIndexItem, (PdfSharp.Pdf.PdfPage, PdfDictionary)>();
			PlaceTableColumnAnnotations(outDoc, pdfPage, tablePagination2.Pages[0], P_6, num12, array, tablePagination2, P_16, dictionary, P_9, num15);
			if (tablePagination2.Pages.Count > 1)
			{
				PdfSharp.Pdf.PdfDocument pdfDocument4 = null;
				bool flag3 = false;
				try
				{
					if (pdfDocument2 != null)
					{
						pdfDocument4 = pdfDocument2;
					}
					else
					{
						pdfDocument4 = PdfReader.Open(P_2, PdfDocumentOpenMode.Import);
						flag3 = true;
					}
					PdfSharp.Pdf.PdfPage page = pdfDocument4.Pages[0];
					for (int num19 = 1; num19 < tablePagination2.Pages.Count; num19++)
					{
						PdfSharp.Pdf.PdfPage pdfPage3 = outDoc.AddPage(page);
						ApplyIndexPageFormFields(pdfPage3, num19 + 1);
						PlaceTableColumnAnnotations(outDoc, pdfPage3, tablePagination2.Pages[num19], P_6, num12, array, tablePagination2, P_16, dictionary, P_9, num15);
					}
					File.AppendAllText(logPath, $"Added {tablePagination2.Pages.Count - 1} extra page(s) for overflow entries\n");
				}
				finally
				{
					if (flag3)
					{
						pdfDocument4?.Dispose();
					}
				}
			}
			Diagnostics.Split("index page layout", xTimer); // DIAG
			pdfDocument2?.Dispose();
			if (!string.IsNullOrEmpty(text2))
			{
				try
				{
					File.Delete(text2);
				}
				catch
				{
				}
			}
			if (P_9 & flag)
			{
				try
				{
					int num20 = 0;
					for (int num21 = 0; num21 <= tablePagination2.Pages.Count && num21 < list2.Count && num21 < outDoc.Pages.Count; num21++)
					{
						if (dictionary.TryGetValue(list2[num21], out var value2))
						{
							SetGoToAction(outDoc, value2.Item2, outDoc.Pages[num21]);
							num20++;
						}
					}
					File.AppendAllText(logPath, $"Added {num20} GENERAL-section self-link(s) (cover sheet + index page(s))\n");
				}
				catch (Exception ex6)
				{
					File.AppendAllText(logPath, "Failed to add GENERAL-section self-links: " + ex6.Message + "\n");
				}
			}
			if (P_8 && P_7 != null)
			{
				List<(string, int)> list5 = P_7.ToList();
				File.AppendAllText(logPath, $"Sheet import requested: includeSheetsInOutput={P_8}, sheetEntries count={list5.Count}\n");
				P_13?.Status("Importing sheets into the index...");
				int count = outDoc.Pages.Count;
				List<(string, int)> list6 = (flag ? list5.Where(((string Path, int PageIndex) en) => !(en.Path == coverSheetPath) || en.PageIndex != coverSheetPageIndex).ToList() : list5);
				File.AppendAllText(logPath, $"Starting sheet import: indexPageCount={count}, sheets to import={list6.Count}\n");
				Dictionary<(string, int), int> dictionary2 = new Dictionary<(string, int), int>();
				foreach (DrawingIndexItem item2 in list2)
				{
					if (!string.IsNullOrEmpty(item2.SourcePdfPath))
					{
						dictionary2[(item2.SourcePdfPath, item2.SourcePageIndex)] = item2.SeqNo;
					}
				}
				Dictionary<(string, int), int> dictionary3 = new Dictionary<(string, int), int>();
				int num22 = count;
				List<BookmarkNode> list7 = null;
				int num23 = 0;
				foreach (IGrouping<string, (string, int)> item3 in list6.GroupBy(((string Path, int PageIndex) en) => en.Path))
				{
					string key = item3.Key;
					var srcTimer = Diagnostics.Start(); // DIAG
					List<int> list8 = item3.Select(((string Path, int PageIndex) en) => en.PageIndex).ToList();
					P_14.ThrowIfCancellationRequested();
					if (!File.Exists(key))
					{
						foreach (int item4 in list8)
						{
							dictionary3[(key, item4)] = -1;
						}
						num23 += list8.Count;
						continue;
					}
					PdfSharp.Pdf.PdfDocument pdfDocument5 = null;
					try
					{
						try
						{
							pdfDocument5 = PdfReader.Open(key, PdfDocumentOpenMode.Import);
							File.AppendAllText(logPath, $"Opened sheet source {Path.GetFileName(key)} ({pdfDocument5.PageCount} page(s)) for import\n");
						}
						catch (Exception ex7) when (ex7.Message.Contains("Cannot retrieve stream length"))
						{
							File.AppendAllText(logPath, "Stream length issue in " + Path.GetFileName(key) + ", attempting repair\n");
							if (!TryRepairMissingStreamLengths(key, logPath))
							{
								File.AppendAllText(logPath, "Repair failed for " + Path.GetFileName(key) + "\n");
								throw;
							}
							File.AppendAllText(logPath, "Repair successful, retrying " + Path.GetFileName(key) + "\n");
							pdfDocument5 = PdfReader.Open(key, PdfDocumentOpenMode.Import);
						}
						if (P_19 && !P_20 && !string.IsNullOrEmpty(P_18) && string.Equals(key, P_18, StringComparison.OrdinalIgnoreCase))
						{
							list7 = ReadBookmarkTree(pdfDocument5, logPath);
						}
						foreach (int item5 in list8)
						{
							P_14.ThrowIfCancellationRequested();
							num23++;
							P_13?.Report($"Importing sheet {num23} of {list6.Count}: {Path.GetFileName(key)}");
							if (item5 < 0 || item5 >= pdfDocument5.PageCount)
							{
								dictionary3[(key, item5)] = -1;
								continue;
							}
							try
							{
								PdfSharp.Pdf.PdfPage pdfPage4 = pdfDocument5.Pages[item5];
								int integer = pdfPage4.Elements.GetInteger("/Rotate");
								double num24 = pdfPage4.MediaBox.X2 - pdfPage4.MediaBox.X1;
								double num25 = pdfPage4.MediaBox.Y2 - pdfPage4.MediaBox.Y1;
								PdfSharp.Pdf.PdfPage pdfPage5 = outDoc.AddPage(pdfPage4);
								dictionary3[(key, item5)] = num22;
								num22++;
								if (P_21.HasValue && P_22 > 0 && string.Equals(key, P_18, StringComparison.OrdinalIgnoreCase) && dictionary2.TryGetValue((key, item5), out var value3))
								{
									try
									{
										SetSheetNumberFormFieldValue(outDoc, pdfPage5, P_21.Value, value3, sheetNumberTotalSheets, P_16, integer, num24, num25);
									}
									catch (Exception ex8)
									{
										File.AppendAllText(logPath, $"Failed to stamp SHEET NO. on page {item5} of {key}: {ex8.Message}\n");
									}
								}
							}
							catch (Exception ex9)
							{
								File.AppendAllText(logPath, $"Failed to import page {item5} of {key}: {ex9.Message}\n");
								dictionary3[(key, item5)] = -1;
							}
						}
					}
					catch (Exception ex10)
					{
						File.AppendAllText(logPath, $"Failed to import sheet source {key}: {ex10.Message}\n");
						foreach (int item6 in list8)
						{
							dictionary3[(key, item6)] = -1;
						}
					}
					finally
					{
						pdfDocument5?.Dispose();
					}
					Diagnostics.Split("import sheet source", srcTimer, $"{Path.GetFileName(key)} ({list8.Count} page(s))"); // DIAG
				}
				int num26 = Enumerable.Count(dictionary3.Values, (int v) => v >= 0);
				File.AppendAllText(logPath, $"Imported {num26}/{list6.Count} sheet(s); total pages now: {outDoc.Pages.Count}\n");
				if (P_9 && count > 0 && dictionary3.Count > 0)
				{
					try
					{
						int num27 = 0;
						foreach (DrawingIndexItem item7 in list2)
						{
							if (dictionary.TryGetValue(item7, out var value4) && !string.IsNullOrEmpty(item7.SourcePdfPath) && dictionary3.TryGetValue((item7.SourcePdfPath, item7.SourcePageIndex), out var value5) && value5 >= 0 && value5 < outDoc.Pages.Count)
							{
								SetGoToAction(outDoc, value4.Item2, outDoc.Pages[value5]);
								num27++;
							}
						}
						File.AppendAllText(logPath, $"Added {num27} jump-to-sheet action(s) to drawing number markups\n");
					}
					catch (Exception ex11)
					{
						File.AppendAllText(logPath, "Failed to add hyperlinks: " + ex11.Message + "\n");
					}
				}
				if (P_19 & P_20)
				{
					try
					{
						List<BookmarkNode> list9 = BuildBookmarkTreeFromGeneratedIndex(list2, dictionary3, indexPageEntries, flag);
						if (list9.Count > 0)
						{
							int num28 = WriteBookmarkTree(outDoc, list9, P_18 ?? string.Empty, dictionary3);
							File.AppendAllText(logPath, $"Generated {num28} bookmark(s) from the drawing index\n");
						}
					}
					catch (Exception ex12)
					{
						File.AppendAllText(logPath, "Failed to generate bookmarks from the drawing index: " + ex12.Message + "\n");
					}
				}
				else if (P_19 && list7 != null && list7.Count > 0 && !string.IsNullOrEmpty(P_18))
				{
					try
					{
						int? num29 = ((flag && string.Equals(coverSheetPath, P_18, StringComparison.OrdinalIgnoreCase)) ? new int?(coverSheetPageIndex) : ((int?)null));
						int num30 = WriteBookmarkTree(outDoc, list7, P_18, dictionary3, num29, list4);
						File.AppendAllText(logPath, $"Transferred {num30} bookmark(s) from combined set\n");
					}
					catch (Exception ex13)
					{
						File.AppendAllText(logPath, "Failed to transfer bookmarks: " + ex13.Message + "\n");
					}
				}
			}
			Diagnostics.Split("bookmarks and links", xTimer); // DIAG
			EnsureAcroFormDefaultResources(outDoc, P_16);
			P_13?.Status("Saving the finished index and sheets (large sets take a moment)...");
			RunWithFileLockRetry(() =>
			{
				outDoc.Save(P_2);
			}, logPath, "Save " + P_2);
			File.AppendAllText(logPath, "Saved output to " + P_2 + "\n");
			Diagnostics.Split("acroform + save output", xTimer); // DIAG
			void ApplyIndexPageFormFields(PdfSharp.Pdf.PdfPage pdfPage6, int num31)
			{
				if (num31 - 1 < indexPageEntries.Count)
				{
					DrawingIndexItem drawingIndexItem3 = indexPageEntries[num31 - 1];
					SetIndexPageFormFieldValues(outDoc, pdfPage6, drawingIndexItem3.DrawingNumber, drawingIndexItem3.Title, P_16, sizeScale);
				}
				else
				{
					RemoveTemplateFormFieldWidgets(outDoc, pdfPage6);
				}
				if (stampSheetNumbers && num31 - 1 < indexPageEntries.Count)
				{
					try
					{
						SetSheetNumberFormFieldValue(outDoc, pdfPage6, P_21.Value, indexPageEntries[num31 - 1].SeqNo, sheetNumberTotalSheets, P_16, pdfPage6.Elements.GetInteger("/Rotate"), pdfPage6.MediaBox.X2 - pdfPage6.MediaBox.X1, pdfPage6.MediaBox.Y2 - pdfPage6.MediaBox.Y1);
					}
					catch (Exception ex15)
					{
						File.AppendAllText(logPath, $"Failed to stamp SHEET NO. on drawing index page {num31}: {ex15.Message}\n");
					}
				}
			}
		}
		catch (Exception ex14)
		{
			try
			{
				File.AppendAllText(GetLogFilePath("export_log.txt"), "Error: " + ex14?.ToString() + "\n");
			}
			catch
			{
			}
			throw;
		}
		finally
		{
			outDoc?.Dispose();
		}
	}

	private static void SetGoToAction(PdfSharp.Pdf.PdfDocument P_0, PdfDictionary P_1, PdfSharp.Pdf.PdfPage P_2)
	{
		if (P_2.Reference != null)
		{
			PdfArray pdfArray = new PdfArray(P_0);
			pdfArray.Elements.Add(P_2.Reference);
			pdfArray.Elements.Add(new PdfName("/Fit"));
			PdfDictionary pdfDictionary = new PdfDictionary(P_0);
			pdfDictionary.Elements.SetName("/S", "/GoTo");
			pdfDictionary.Elements["/D"] = pdfArray;
			P_1.Elements["/A"] = pdfDictionary;
		}
	}
}
