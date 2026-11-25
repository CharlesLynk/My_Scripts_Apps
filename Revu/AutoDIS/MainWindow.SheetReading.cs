using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Tesseract;

namespace MbrohDrawingIndexCreator;

// Reading the title block of every sheet: text layer first, OCR for whatever the text layer lacks, run on several worker threads.
public sealed partial class MainWindow
{
	private enum SheetOutcome
	{
		// Listed in the "skipped" log line (folder mode: not a single-page PDF, or unreadable).
		Skipped,
		// Not part of the output at all (the page could not be opened).
		Dropped,
		// Imported into the output but without an index row (an error while reading its fields).
		KeptWithoutRow,
		Cover,
		Row
	}

	private sealed class SheetRead
	{
		internal SheetOutcome Outcome = SheetOutcome.Row;
		internal string? SkippedName;
		internal int SkippedPages;
		internal string Label = string.Empty;
		internal string DrawingNumber = string.Empty;
		internal string Discipline = string.Empty;
		internal string Title = string.Empty;
		internal bool IdFromOcr;
		internal bool DisciplineFromOcr;
		internal bool TitleFromOcr;
		internal double IdConfidence = 100;
		internal double DisciplineConfidence = 100;
		internal double TitleConfidence = 100;
		internal bool NeededOcr;
		// The title block area is empty (a blank page in the set): nothing to read, and nothing to infer from the neighbours.
		internal bool Blank;
		// The OCR text exactly as read, kept so the index can be assembled again after a second look.
		internal string RawDrawingNumber = string.Empty;
		internal string RawDiscipline = string.Empty;
		internal string RawTitle = string.Empty;
		internal double PageWidth;
		internal double PageHeight;
		// Every distinct reading the passes produced, for the index analysis to choose between.
		internal List<ReadingCandidate> IdCandidates = new List<ReadingCandidate>();
		internal List<ReadingCandidate> DisciplineCandidates = new List<ReadingCandidate>();
		internal List<ReadingCandidate> TitleCandidates = new List<ReadingCandidate>();
	}

	// A sheet whose OCR-read fields are still in doubt after the index was checked, and which fields they are.
	private readonly record struct FlaggedSheet(SheetRead Read, int SourceIndex, bool Id, bool Discipline, bool Title);

	private sealed class SheetWorker : IDisposable
	{
		internal UglyToad.PdfPig.PdfDocument? CombinedDocument;
		internal OcrEngineSet? Engines;

		public void Dispose()
		{
			CombinedDocument?.Dispose();
			Engines?.Dispose();
		}
	}

	// Replaced at the start of each run from the resource plan (how many sheet images may be drawn at once).
	private static SemaphoreSlim RenderGate = new SemaphoreSlim(OcrPresets.MaxConcurrentRenders);

	// Helper processes that draw pages for this run (null when the plan has none); see RenderPool. PDFium is not thread-safe (one render at a
	// time per process), so concurrent renders need separate processes.
	private static RenderPool? _renderPool;

	// The resource plan of the run in progress.
	private static ResourcePlan _resourcePlan = ResourcePlanner.Plan(ResourcePlanner.Default, HardwareProbe.Probe());

	private const double OcrFieldMarginX = 4.0;
	private const double OcrFieldMarginY = 4.0;
	private const double OcrTitleMarginX = 6.0;
	private const double OcrTitleMarginY = 2.0;
	private const double LowConfidenceLimit = 60.0;

	// Share of the title block image that must be dark for the page not to count as blank (a real title block has frame lines and text).
	private const double BlankInkFraction = 0.0005;

	// The first pass always uses the fast model. For presets with a second look, the large model is loaded up front too
	// (it only takes a moment), so the second look cannot stall the run when it starts.
	private static OcrEngineSet? TryCreateOcrEngines(string logPath, OcrRunStats stats, out OcrEngineSet? largeEngines)
	{
		largeEngines = null;
		OcrEngineSet? engines = OcrSupport.TryCreateEngineSet(OcrSupport.DefaultTessdataDirectory, false, out string? reason);
		if (engines != null && CurrentOcrPreset.SecondLook)
		{
			largeEngines = OcrSupport.TryCreateEngineSet(OcrSupport.DefaultTessdataDirectory, true, out string? largeReason);
			if (largeEngines != null && largeEngines.Description != OcrSupport.BestModel)
			{
				largeEngines.Dispose();
				largeEngines = null;
			}
			if (largeEngines == null)
			{
				AppLog.TryAppend(logPath, $"{DateTime.Now}: The large OCR model is not available, so the second look is skipped: {largeReason}");
			}
		}
		stats.SecondLookAvailable = largeEngines != null;
		stats.EngineAvailable = engines != null;
		stats.ModelUsed = engines?.Description;
		stats.EngineMessage = reason;
		if (engines == null)
		{
			AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR unavailable (preset {CurrentOcrPreset.Preset}): {reason}");
		}
		else if (reason != null)
		{
			AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR running with '{engines.Description}' only (preset {CurrentOcrPreset.Preset}): {reason}");
		}
		else
		{
			OcrFieldRecognizer.ResetScaleStats();
			AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR ready (preset {CurrentOcrPreset.Preset}, model '{engines.Description}'{(largeEngines != null ? " + large model for a second look" : string.Empty)}, vector lettering drawn at {OcrPresets.Oversample(CurrentOcrPreset.Preset, false):0.##} x the working letter size (second look {OcrPresets.Oversample(CurrentOcrPreset.Preset, true):0.##} x), scans up to {OcrPresets.FirstPassDpi(CurrentOcrPreset.Preset):0} DPI (second look {OcrPresets.SecondLookDpi(CurrentOcrPreset.Preset):0}) but never above 1.5 x their own resolution)");
		}
		return engines;
	}

	// Copy with a status line, so a slow network share shows progress and Cancel still works during the copy. Uses the operating system's
	// copy routine, and a managed read/write loop if that fails. Returns a short description (size, speed, method) for the logs.
	private static string CopyFileWithProgress(string source, string destination, CancellationToken token, IProgress<(int current, int total, string status)>? progress, string activity)
	{
		Stopwatch timer = Stopwatch.StartNew();
		long total = new FileInfo(source).Length;
		string method = "system copy";
		int lastReportedMegabytes = -1;
		void Report(long copied)
		{
			int megabytes = (int)(copied / (1024 * 1024));
			if (megabytes != lastReportedMegabytes)
			{
				lastReportedMegabytes = megabytes;
				progress?.Report((0, 0, $"{activity} ({megabytes} of {total / (1024 * 1024)} MB)..."));
			}
		}
		try
		{
			FastCopy.Copy(source, destination, token, (copied, _) => Report(copied));
		}
		catch (IOException)
		{
			method = "managed copy";
			const int ChunkBytes = 4 * 1024 * 1024;
			lastReportedMegabytes = -1;
			using FileStream input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read, ChunkBytes, FileOptions.SequentialScan);
			using FileStream output = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None, ChunkBytes, FileOptions.SequentialScan);
			byte[] buffer = new byte[ChunkBytes];
			long copied = 0;
			int read;
			while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
			{
				token.ThrowIfCancellationRequested();
				output.Write(buffer, 0, read);
				copied += read;
				Report(copied);
			}
		}
		double seconds = Math.Max(0.001, timer.Elapsed.TotalSeconds);
		return $"{total / (1024.0 * 1024.0):0.0} MB in {seconds:0.0} s ({total / (1024.0 * 1024.0) / seconds:0.0} MB/s, {method})";
	}
	private static SheetRead[] ReadSheets(IReadOnlyList<(string path, int page)> sources, bool combined, byte[]? combinedBytes, FieldRect idRect, FieldRect disciplineRect, FieldRect titleRect, OcrEngineSet? probeEngines, OcrRunStats stats, string logPath, ProgressTracker progress, CancellationToken token)
	{
		SheetRead[] reads = new SheetRead[sources.Count];
		ConcurrentBag<OcrEngineSet> spareEngines = new ConcurrentBag<OcrEngineSet>();
		if (probeEngines != null)
		{
			spareEngines.Add(probeEngines);
		}
		ParallelOptions options = new ParallelOptions { MaxDegreeOfParallelism = _resourcePlan.Workers, CancellationToken = token };
		try
		{
			// Chunks of one, handed out in page order, so the sheets being read (and the progress shown) advance steadily down the set.
			Parallel.ForEach(Partitioner.Create(0, sources.Count, 1), options,
				() =>
				{
					SheetWorker worker = new SheetWorker();
					if (stats.EngineAvailable)
					{
						// Eager: each worker has its engines before its first sheet, so OCR never stalls the run mid-way.
						worker.Engines = spareEngines.TryTake(out OcrEngineSet? spare) ? spare : OcrSupport.TryCreateEngineSet(OcrSupport.DefaultTessdataDirectory, false, out _);
					}
					return worker;
				},
				(range, loopState, worker) =>
				{
					int index = range.Item1;
					reads[index] = ReadOneSheet(index, sources[index], combined, combinedBytes, idRect, disciplineRect, titleRect, worker, stats, logPath, progress, sources.Count, token);
					return worker;
				},
				worker => worker.Dispose());
		}
		finally
		{
			while (spareEngines.TryTake(out OcrEngineSet? unused))
			{
				unused.Dispose();
			}
		}
		return reads;
	}

	// Size of the template's page, which the field boxes are measured on (0 when unknown).
	private static (double width, double height) _templatePageSize;

	private static (double width, double height) ReadPageSize(string path)
	{
		try
		{
			using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(path);
			UglyToad.PdfPig.Content.Page? page = document.GetPages().FirstOrDefault();
			return page == null ? (0.0, 0.0) : (page.Width, page.Height);
		}
		catch
		{
			return (0.0, 0.0);
		}
	}

	// A sheet that is a reduced or enlarged print of the template's sheet (same shape, different size: a half-size copy of a drawing) has
	// its title block scaled with it, so the field boxes are scaled the same way. Other shapes are left alone.
	private static (FieldRect id, FieldRect discipline, FieldRect title) FieldsForPage(double pageWidth, double pageHeight, FieldRect id, FieldRect discipline, FieldRect title)
	{
		(double templateWidth, double templateHeight) = _templatePageSize;
		if (templateWidth <= 0.0 || templateHeight <= 0.0 || pageWidth <= 0.0 || pageHeight <= 0.0)
		{
			return (id, discipline, title);
		}
		double sx = pageWidth / templateWidth;
		double sy = pageHeight / templateHeight;
		if ((Math.Abs(sx - 1.0) < 0.02 && Math.Abs(sy - 1.0) < 0.02) || Math.Abs(sx - sy) > 0.03 * sx)
		{
			return (id, discipline, title);
		}
		FieldRect Scale(FieldRect r) => new FieldRect(r.X1 * sx, r.Y1 * sx, r.X2 * sx, r.Y2 * sx);
		return (Scale(id), Scale(discipline), Scale(title));
	}

	private static SheetRead ReadOneSheet(int index, (string path, int page) source, bool combined, byte[]? combinedBytes, FieldRect idRect, FieldRect disciplineRect, FieldRect titleRect, SheetWorker worker, OcrRunStats stats, string logPath, ProgressTracker progress, int totalSources, CancellationToken token)
	{
		token.ThrowIfCancellationRequested();
		string path = source.path;
		int pageIndex = source.page;
		string label = combined ? $"page {pageIndex + 1} of the combined set" : Path.GetFileName(path);
		SheetRead read = new SheetRead { Label = label };
		var sheetTimer = Diagnostics.Start(); // DIAG
		UglyToad.PdfPig.PdfDocument? ownDocument = null;
		try
		{
			UglyToad.PdfPig.Content.Page? page;
			if (combined)
			{
				try
				{
					worker.CombinedDocument ??= UglyToad.PdfPig.PdfDocument.Open(combinedBytes!);
					page = worker.CombinedDocument.GetPage(pageIndex + 1);
				}
				catch (Exception ex) when (ex is not OperationCanceledException)
				{
					AppLog.TryAppend(logPath, $"{DateTime.Now}: Sheet {label} skipped after an error: {ex.Message}");
					read.Outcome = SheetOutcome.Dropped;
					return read;
				}
			}
			else
			{
				try
				{
					ownDocument = UglyToad.PdfPig.PdfDocument.Open(path);
					int pageCount = ownDocument.NumberOfPages;
					if (pageCount != 1)
					{
						read.Outcome = SheetOutcome.Skipped;
						read.SkippedName = Path.GetFileName(path);
						read.SkippedPages = pageCount;
						return read;
					}
					page = ownDocument.GetPages().FirstOrDefault();
				}
				catch
				{
					read.Outcome = SheetOutcome.Skipped;
					read.SkippedName = Path.GetFileName(path);
					read.SkippedPages = -1;
					return read;
				}
			}
			progress.Report(count => $"Reading sheet {count} of {totalSources}");
			if (index == 0)
			{
				read.Outcome = SheetOutcome.Cover;
				return read;
			}
			Diagnostics.Split("open sheet", sheetTimer, label); // DIAG
			if (page == null)
			{
				read.Outcome = SheetOutcome.KeptWithoutRow;
				return read;
			}
			(idRect, disciplineRect, titleRect) = FieldsForPage(page.Width, page.Height, idRect, disciplineRect, titleRect);
			string id = FindTextInRect(page, idRect, out List<TextHit> idHits);
			string discipline = FindTextInRect(page, disciplineRect, out List<TextHit> disciplineHits);
			string title = FindTextInRect(page, titleRect, out List<TextHit> titleHits);
			if (idHits.Count >= 2)
			{
				id = (!HasOverlappingHits(idHits)) ? ClosestHitToRect(idHits, idRect).Text : string.Empty;
			}
			Diagnostics.Split("field lookup", sheetTimer, label); // DIAG
			bool titleFromText = !string.IsNullOrWhiteSpace(title);
			bool needId = string.IsNullOrWhiteSpace(id);
			bool needDiscipline = string.IsNullOrWhiteSpace(discipline);
			bool needTitle = string.IsNullOrWhiteSpace(title);
			read.NeededOcr = needId || needDiscipline || needTitle;
			if (read.NeededOcr)
			{
				stats.AddNeedingOcr();
				if (worker.Engines != null)
				{
					OcrFieldResult? ocrId = null;
					OcrFieldResult? ocrDiscipline = null;
					OcrFieldResult? ocrTitle = null;
					bool rendered = ReadFieldsByOcr(worker.Engines, CurrentOcrPreset, false, page, stats, path, pageIndex, page.Width, page.Height, idRect, disciplineRect, titleRect, needId, needDiscipline, needTitle, logPath, out bool blankPage, out ocrId, out ocrDiscipline, out ocrTitle, out List<OcrFieldResult> idAll, out List<OcrFieldResult> disciplineAll, out List<OcrFieldResult> titleAll);
					if (rendered && blankPage)
					{
						read.Blank = true;
						stats.AddBlank();
					}
					else if (rendered)
					{
						stats.AddAttempted();
						read.IdCandidates = idAll.Select(r => new ReadingCandidate(r.Text, r.Confidence, r.Large)).ToList();
						read.DisciplineCandidates = disciplineAll.Select(r => new ReadingCandidate(r.Text, r.Confidence, r.Large)).ToList();
						read.TitleCandidates = titleAll.Select(r => new ReadingCandidate(r.Text, r.Confidence, r.Large)).ToList();
						if (ocrId.HasValue)
						{
							id = ocrId.Value.Text;
							read.IdFromOcr = true;
							read.IdConfidence = ocrId.Value.Confidence;
						}
						if (ocrDiscipline.HasValue)
						{
							discipline = ocrDiscipline.Value.Text;
							read.DisciplineFromOcr = true;
							read.DisciplineConfidence = ocrDiscipline.Value.Confidence;
						}
						if (ocrTitle.HasValue)
						{
							title = ocrTitle.Value.Text;
							read.TitleFromOcr = true;
							read.TitleConfidence = ocrTitle.Value.Confidence;
						}
					}
					else
					{
						stats.AddRenderFailure();
					}
				}
			}
			Diagnostics.Split("ocr fallback", sheetTimer, label); // DIAG
			// Only text-layer values are tidied here; OCR text is cleaned later, once the whole set is known.
			if (!read.IdFromOcr)
			{
				id = FlattenEmbeddedNewlines(id);
			}
			if (!read.DisciplineFromOcr)
			{
				discipline = FlattenEmbeddedNewlines(discipline).Trim(' ', '\t', '\r', '\n', '|');
			}
			if (!read.TitleFromOcr)
			{
				title = FlattenEmbeddedNewlines(title);
				if (titleFromText)
				{
					title = StripLeadingDuplicateFieldByPosition(title, titleHits, discipline, disciplineHits);
					title = StripLeadingDuplicateFieldByPosition(title, titleHits, id, idHits);
				}
			}
			read.DrawingNumber = id;
			read.Discipline = discipline;
			read.Title = title;
			read.RawDrawingNumber = id;
			read.RawDiscipline = discipline;
			read.RawTitle = title;
			read.PageWidth = page.Width;
			read.PageHeight = page.Height;
			return read;
		}
		catch (Exception ex) when (ex is not OperationCanceledException)
		{
			AppLog.TryAppend(logPath, $"{DateTime.Now}: Sheet {Path.GetFileName(path)} (page {pageIndex + 1}) skipped after an error: {ex.Message}");
			read.Outcome = SheetOutcome.KeptWithoutRow;
			return read;
		}
		finally
		{
			ownDocument?.Dispose();
		}
	}

	// Second look (Thorough preset): the entries still in doubt after the whole index was checked for patterns are
	// read again with the large, more detailed model. The new reading replaces the old only if it scores at least as well,
	// and the index is then assembled and checked again.
	private static void SecondLook(List<FlaggedSheet> flagged, IReadOnlyList<(string path, int page)> sources, FieldRect idRect, FieldRect disciplineRect, FieldRect titleRect, OcrEngineSet? probeLarge, OcrRunStats stats, string logPath, ProgressTracker progress, CancellationToken token)
	{
		OcrPresetSettings preset = CurrentOcrPreset;
		ConcurrentBag<OcrEngineSet> spare = new ConcurrentBag<OcrEngineSet>();
		if (probeLarge != null)
		{
			spare.Add(probeLarge);
		}
		int done = 0;
		ParallelOptions options = new ParallelOptions { MaxDegreeOfParallelism = _resourcePlan.Workers, CancellationToken = token };
		try
		{
			Parallel.ForEach(Partitioner.Create(0, flagged.Count, 1), options,
				() => (engines: spare.TryTake(out OcrEngineSet? set) ? set : OcrSupport.TryCreateEngineSet(OcrSupport.DefaultTessdataDirectory, true, out _), pages: new Dictionary<string, UglyToad.PdfPig.PdfDocument>()),
				(range, loopState, local) =>
				{
					token.ThrowIfCancellationRequested();
					FlaggedSheet item = flagged[range.Item1];
					SheetRead read = item.Read;
					(string path, int page) source = sources[item.SourceIndex];
					progress.Status($"Taking a second look at doubtful entries ({Interlocked.Increment(ref done)} of {flagged.Count})...");
					if (local.engines == null || read.PageWidth <= 0.0)
					{
						return local;
					}
					// Lightweight first: the page is drawn again from its vector shapes at a higher resolution (cheap), and only then does the
					// large model read it. Pages that are pictures are drawn by PDFium as before.
					UglyToad.PdfPig.Content.Page? vectorPage = null;
					try
					{
						if (!local.pages.TryGetValue(source.path, out UglyToad.PdfPig.PdfDocument? opened))
						{
							opened = UglyToad.PdfPig.PdfDocument.Open(source.path);
							local.pages[source.path] = opened;
						}
						vectorPage = opened.GetPage(source.page + 1);
					}
					catch (Exception ex) when (ex is not OperationCanceledException)
					{
						AppLog.TryAppend(logPath, $"{DateTime.Now}: Second look could not reopen page {source.page + 1} for vector drawing: {ex.Message}");
					}
					(FieldRect pageId, FieldRect pageDiscipline, FieldRect pageTitle) = FieldsForPage(read.PageWidth, read.PageHeight, idRect, disciplineRect, titleRect);
					bool rendered = ReadFieldsByOcr(local.engines, preset, true, vectorPage, stats, source.path, source.page, read.PageWidth, read.PageHeight, pageId, pageDiscipline, pageTitle, item.Id, item.Discipline, item.Title, logPath, out bool _, out OcrFieldResult? id, out OcrFieldResult? discipline, out OcrFieldResult? title, out List<OcrFieldResult> idAll, out List<OcrFieldResult> disciplineAll, out List<OcrFieldResult> titleAll);
					if (!rendered)
					{
						return local;
					}
					// The large model's readings join the first pass's; the analysis picks whichever fits the set best.
					read.IdCandidates.AddRange(idAll.Select(r => new ReadingCandidate(r.Text, r.Confidence, true)));
					read.DisciplineCandidates.AddRange(disciplineAll.Select(r => new ReadingCandidate(r.Text, r.Confidence, true)));
					read.TitleCandidates.AddRange(titleAll.Select(r => new ReadingCandidate(r.Text, r.Confidence, true)));
					if (id.HasValue)
					{
						// A reading that has lost the hyphenated form of a sheet number the first pass read whole ("4NE" for "HH-H-05") is not better.
						bool keepsForm = !Regex.IsMatch(OcrTextCleanup.CleanSheetId(read.RawDrawingNumber ?? string.Empty, true), @"^[A-Z0-9]+(?:-[A-Z0-9]+)+$") || Regex.IsMatch(OcrTextCleanup.CleanSheetId(id.Value.Text, true), @"^[A-Z0-9]+(?:-[A-Z0-9]+)+$");
						bool better = keepsForm && !string.IsNullOrWhiteSpace(id.Value.Text) && (string.IsNullOrWhiteSpace(read.RawDrawingNumber) || id.Value.Confidence >= read.IdConfidence);
						stats.AddSecondLook(better && !string.Equals(id.Value.Text, read.RawDrawingNumber, StringComparison.Ordinal));
						if (better)
						{
							read.RawDrawingNumber = id.Value.Text;
							read.IdConfidence = id.Value.Confidence;
							read.IdFromOcr = true;
						}
					}
					if (discipline.HasValue)
					{
						bool better = !string.IsNullOrWhiteSpace(discipline.Value.Text) && (string.IsNullOrWhiteSpace(read.RawDiscipline) || discipline.Value.Confidence >= read.DisciplineConfidence);
						stats.AddSecondLook(better && !string.Equals(discipline.Value.Text, read.RawDiscipline, StringComparison.Ordinal));
						if (better)
						{
							read.RawDiscipline = discipline.Value.Text;
							read.DisciplineConfidence = discipline.Value.Confidence;
							read.DisciplineFromOcr = true;
						}
					}
					if (title.HasValue)
					{
						bool better = !string.IsNullOrWhiteSpace(title.Value.Text) && (string.IsNullOrWhiteSpace(read.RawTitle) || title.Value.Confidence >= read.TitleConfidence);
						stats.AddSecondLook(better && !string.Equals(title.Value.Text, read.RawTitle, StringComparison.Ordinal));
						if (better)
						{
							read.RawTitle = title.Value.Text;
							read.TitleConfidence = title.Value.Confidence;
							read.TitleFromOcr = true;
						}
					}
					return local;
				},
				local =>
				{
					local.engines?.Dispose();
					foreach (UglyToad.PdfPig.PdfDocument document in local.pages.Values)
					{
						document.Dispose();
					}
				});
		}
		finally
		{
			while (spare.TryTake(out OcrEngineSet? unused))
			{
				unused.Dispose();
			}
		}
	}
	// Renders the needed field regions of one sheet once and reads each field. Returns false if the sheet image could not be made.
	private sealed class FieldReadings
	{
		internal OcrFieldResult? Id;
		internal OcrFieldResult? Discipline;
		internal OcrFieldResult? Title;
		internal bool Blank;
		internal readonly List<OcrFieldResult> IdAll = new List<OcrFieldResult>();
		internal readonly List<OcrFieldResult> DisciplineAll = new List<OcrFieldResult>();
		internal readonly List<OcrFieldResult> TitleAll = new List<OcrFieldResult>();
	}
	private static bool ReadFieldsByOcr(OcrEngineSet engines, OcrPresetSettings preset, bool secondLook, UglyToad.PdfPig.Content.Page? vectorPage, OcrRunStats stats, string path, int pageIndex, double pageWidth, double pageHeight, FieldRect idRect, FieldRect disciplineRect, FieldRect titleRect, bool needId, bool needDiscipline, bool needTitle, string logPath, out bool blank, out OcrFieldResult? id, out OcrFieldResult? discipline, out OcrFieldResult? title, out List<OcrFieldResult> idAll, out List<OcrFieldResult> disciplineAll, out List<OcrFieldResult> titleAll)
	{
		FieldReadings result = new FieldReadings();
		bool ok = ReadFieldsCore(engines, preset, secondLook, vectorPage, stats, path, pageIndex, pageWidth, pageHeight, idRect, disciplineRect, titleRect, needId, needDiscipline, needTitle, logPath, result);
		blank = result.Blank;
		id = result.Id;
		discipline = result.Discipline;
		title = result.Title;
		idAll = result.IdAll;
		disciplineAll = result.DisciplineAll;
		titleAll = result.TitleAll;
		return ok;
	}

	private static bool ReadFieldsCore(OcrEngineSet engines, OcrPresetSettings preset, bool secondLook, UglyToad.PdfPig.Content.Page? vectorPage, OcrRunStats stats, string path, int pageIndex, double pageWidth, double pageHeight, FieldRect idRect, FieldRect disciplineRect, FieldRect titleRect, bool needId, bool needDiscipline, bool needTitle, string logPath, FieldReadings result)
	{
		FieldRect Padded(FieldRect r, double padX, double padY) => new FieldRect(Math.Max(0.0, r.X1 - padX), Math.Max(0.0, r.Y1 - padY), Math.Min(pageWidth, r.X2 + padX), Math.Min(pageHeight, r.Y2 + padY));
		// Single-line fields are cut from a generous area (a sixth of the box's width and 60% of its height beyond each side): sheets drawn
		// by hand or by different people place the lettering a little high, low or to one side, and the line nearest the expected place is read.
		FieldRect paddedId = Padded(idRect, Math.Max(OcrFieldMarginX, (idRect.X2 - idRect.X1) / 6.0), Math.Max(OcrFieldMarginY, 0.6 * (idRect.Y2 - idRect.Y1)));
		FieldRect paddedDiscipline = Padded(disciplineRect, OcrFieldMarginX, Math.Max(OcrFieldMarginY, 0.6 * (disciplineRect.Y2 - disciplineRect.Y1)));
		FieldRect paddedTitle = Padded(titleRect, OcrTitleMarginX, OcrTitleMarginY);
		List<FieldRect> wanted = new List<FieldRect>();
		if (needId)
		{
			wanted.Add(paddedId);
		}
		if (needDiscipline)
		{
			wanted.Add(paddedDiscipline);
		}
		if (needTitle)
		{
			wanted.Add(paddedTitle);
		}
		if (wanted.Any(r => r.X2 - r.X1 < 1.0 || r.Y2 - r.Y1 < 1.0))
		{
			AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR skipped for {Path.GetFileName(path)} page {pageIndex + 1}: the page ({pageWidth:0} x {pageHeight:0} pt) is smaller than the template's title block area, so it is probably a different sheet size.");
			return false;
		}
		FieldRect region = new FieldRect(wanted.Min(r => r.X1), wanted.Min(r => r.Y1), wanted.Max(r => r.X2), wanted.Max(r => r.Y2));
		string sheetName = Path.GetFileName(path);
		var timer = Diagnostics.Start(); // DIAG
		bool idDone = !needId;
		bool disciplineDone = !needDiscipline;
		bool titleDone = !needTitle;
		bool anyRead = false;
		bool vectorTried = false;

		// Where the field box centre sits within its padded cut, from the top (0) to the bottom (1).
		double AnchorOf(FieldRect field, FieldRect cut) => Math.Clamp((cut.Y2 - (field.Y1 + field.Y2) / 2.0) / Math.Max(1e-6, cut.Y2 - cut.Y1), 0.0, 1.0);

		GrayImage Crop(GrayImage image, FieldRect r)
		{
			double scaleX = image.Width / (region.X2 - region.X1);
			double scaleY = image.Height / (region.Y2 - region.Y1);
			return image.Crop((int)Math.Floor((r.X1 - region.X1) * scaleX), (int)Math.Floor((region.Y2 - r.Y2) * scaleY), (int)Math.Ceiling((r.X2 - r.X1) * scaleX), (int)Math.Ceiling((r.Y2 - r.Y1) * scaleY));
		}

		// Reads the wanted fields that have no usable reading yet from pictures of the title block (one picture per field box, supplied by
		// `picture`); a reading that is blank or cannot be a sheet number / discipline / title leaves the field open for the next route.
		bool ReadFrom(Func<FieldRect, GrayImage> picture, bool final)
		{
			Dictionary<FieldRect, GrayImage> cache = new Dictionary<FieldRect, GrayImage>();
			GrayImage Picture(FieldRect r)
			{
				if (!cache.TryGetValue(r, out GrayImage? image))
				{
					image = picture(r);
					cache[r] = image;
				}
				return image;
			}
			// A page with nothing in any of the field boxes (a blank sheet in the scan) has nothing to read; reading it would only invent text.
			if (final && wanted.All(r => VectorRaster.InkFraction(Picture(r)) < BlankInkFraction))
			{
				result.Blank = true;
				idDone = disciplineDone = titleDone = true;
				return true;
			}
			try
			{
				if (!idDone)
				{
					OcrFieldResult read = OcrFieldRecognizer.ReadAll(engines.Engine, Picture(paddedId), OcrFieldKind.SheetId, preset.OptionsFor(OcrFieldKind.SheetId, secondLook), out List<OcrFieldResult> all, AnchorOf(idRect, paddedId));
					result.IdAll.AddRange(all);
					result.Id = Preferred(result.Id, read, OcrFieldKind.SheetId);
					idDone = final || OcrFieldRecognizer.IsPlausible(result.Id.Value.Text, OcrFieldKind.SheetId);
				}
				if (!disciplineDone)
				{
					OcrFieldResult read = OcrFieldRecognizer.ReadAll(engines.Engine, Picture(paddedDiscipline), OcrFieldKind.Discipline, preset.OptionsFor(OcrFieldKind.Discipline, secondLook), out List<OcrFieldResult> all, AnchorOf(disciplineRect, paddedDiscipline));
					result.DisciplineAll.AddRange(all);
					result.Discipline = Preferred(result.Discipline, read, OcrFieldKind.Discipline);
					disciplineDone = final || OcrFieldRecognizer.IsPlausible(result.Discipline.Value.Text, OcrFieldKind.Discipline);
				}
				if (!titleDone)
				{
					OcrFieldResult read = OcrFieldRecognizer.ReadAll(engines.Engine, Picture(paddedTitle), OcrFieldKind.Title, preset.OptionsFor(OcrFieldKind.Title, secondLook), out List<OcrFieldResult> all);
					result.TitleAll.AddRange(all);
					result.Title = Preferred(result.Title, read, OcrFieldKind.Title);
					titleDone = final || OcrFieldRecognizer.IsPlausible(result.Title.Value.Text, OcrFieldKind.Title);
				}
				return true;
			}
			catch (Exception ex) when (ex is not OperationCanceledException)
			{
				AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR failed on {sheetName} page {pageIndex + 1}: {ex.Message}");
				return false;
			}
		}

		// Only the field boxes are ever drawn. When they lie far apart (so the box around all of them would take in a lot of the
		// drawing itself) each is drawn on its own rather than drawing the whole area between them.
		double unionArea = (region.X2 - region.X1) * (region.Y2 - region.Y1);
		bool separateRegions = unionArea > 3.0 * wanted.Sum(r => (r.X2 - r.X1) * (r.Y2 - r.Y1));

		// First choice: draw the title block from the page's own vector shapes (CAD plots draw their lettering as strokes). That needs no
		// page rendering and is several times quicker. It is only used when it gives usable readings; otherwise the page is drawn as a picture.
		if (vectorPage != null)
		{
			vectorTried = true;
			Func<FieldRect, GrayImage>? drawn = null;
			string? declined = null;
			if (PdfPigVectorSource.TryCollect(vectorPage, wanted.Select(r => (r.X1, r.Y1, r.X2, r.Y2)).ToList(), out List<VectorShape> shapes, out declined))
			{
				double ink;
				// Every field is drawn on its own, fine enough that its letters come out at the preset's pixel height.
				double oversample = OcrPresets.Oversample(preset.Preset, secondLook);
				Dictionary<FieldRect, GrayImage> fields = wanted.Distinct().ToDictionary(r => r, r => DrawVectorField(shapes, r, oversample));
				ink = fields.Values.Max(VectorRaster.InkFraction);
				Func<FieldRect, GrayImage> make = r => fields.TryGetValue(r, out GrayImage? drawnField) ? drawnField : DrawVectorField(shapes, r, oversample);
				if (ink >= 0.002)
				{
					drawn = make;
				}
				else
				{
					declined = "nothing visible in the title block";
				}
			}
			Diagnostics.Split("ocr vector render", timer, sheetName); // DIAG
			if (drawn != null && ReadFrom(drawn, false))
			{
				anyRead = true;
				stats.AddVectorRead();
			}
			else if (drawn == null && stats.ShouldLogVectorDecline())
			{
				AppLog.TryAppend(logPath, $"{DateTime.Now}: Title block of {sheetName} page {pageIndex + 1} is drawn as a picture instead of read from vector shapes: {declined}");
			}
		}
		if (!(idDone && disciplineDone && titleDone))
		{
			timer = Diagnostics.Start(); // DIAG
			double drawDpi = secondLook ? OcrPresets.SecondLookDpi(preset.Preset) : OcrPresets.FirstPassDpi(preset.Preset);
			// A scan has a resolution of its own (the Bachman scans are 300 DPI bilevel pictures); drawing it much finer only enlarges the
			// scanner's noise and the reading gets worse (measured: good at 300-450 DPI, garbage from 600).
			if (NativePictureDpi(vectorPage, region) is double nativeDpi)
			{
				drawDpi = Math.Min(drawDpi, Math.Max(200.0, nativeDpi * 1.5));
			}
			// Whatever the setting, one drawing is kept to a sane size (a weak machine, or a very large title block area).
			double drawnPixels = unionArea * drawDpi * drawDpi / (72.0 * 72.0);
			if (drawnPixels > 2.0 * MaxFieldPixels)
			{
				drawDpi *= Math.Sqrt(2.0 * MaxFieldPixels / drawnPixels);
			}
			// Drawing takes about the same time per page whatever its size, so separate drawings only pay when the single one would be huge.
			bool drawSeparately = separateRegions && unionArea * drawDpi * drawDpi / (72.0 * 72.0) > 20.0e6;
			Func<FieldRect, GrayImage>? source = null;
			if (drawSeparately)
			{
				Dictionary<FieldRect, GrayImage?> drawnByField = wanted.ToDictionary(r => r, r => RenderRegion(path, pageIndex, r, drawDpi, logPath));
				if (drawnByField.Values.All(image => image != null))
				{
					source = r => drawnByField[r]!;
				}
			}
			else
			{
				GrayImage? whole = RenderRegion(path, pageIndex, region, drawDpi, logPath);
				if (whole != null)
				{
					source = r => Crop(whole, r);
				}
			}
			Diagnostics.Split("ocr render", timer, sheetName); // DIAG
			if (source == null || !ReadFrom(source, true))
			{
				return anyRead;
			}
			anyRead = true;
			stats.AddImageRead(vectorTried);
		}
		Diagnostics.Split("ocr recognize", timer, sheetName); // DIAG
		return anyRead;
	}
	// The better of two readings of the same field: a usable one beats an unusable one, then higher confidence wins.
	private static OcrFieldResult Preferred(OcrFieldResult? current, OcrFieldResult candidate, OcrFieldKind kind)
	{
		if (!current.HasValue)
		{
			return candidate;
		}
		bool currentUsable = OcrFieldRecognizer.IsPlausible(current.Value.Text, kind);
		bool candidateUsable = OcrFieldRecognizer.IsPlausible(candidate.Text, kind);
		if (currentUsable != candidateUsable)
		{
			return candidateUsable ? candidate : current.Value;
		}
		return candidate.Confidence > current.Value.Confidence ? candidate : current.Value;
	}
	// Cap on the size of one drawn field, so a huge field or a tiny letter cannot ask for a gigabyte.
	private const double MaxFieldPixels = 24.0e6;

	// Draws one field of a vector sheet at the resolution its own lettering needs: a probe drawing measures the letter height, then the
	// field is drawn so the letters are TargetLetterHeight x oversample pixels tall (the reader shrinks them to the target anyway, so more
	// than that is only a finer-grained shrink). A sheet printed at half size has half-size lettering and so gets twice the DPI.
	private static GrayImage DrawVectorField(IReadOnlyList<VectorShape> shapes, FieldRect r, double oversample)
	{
		const double probeDpi = 150.0;
		double pen = VectorRaster.ChoosePenPoints(shapes, r.X1, r.Y1, r.X2, r.Y2);
		int measured = VectorRaster.Render(shapes, r.X1, r.Y1, r.X2, r.Y2, probeDpi, pen).TypicalLetterHeight();
		int target = OcrFieldRecognizer.TargetLetterHeight > 0 ? OcrFieldRecognizer.TargetLetterHeight : 44;
		double dpi = measured <= 0 ? 300.0 : Math.Clamp(probeDpi * target * oversample / measured, 100.0, 1200.0);
		double pixels = (r.X2 - r.X1) * (r.Y2 - r.Y1) * dpi * dpi / (72.0 * 72.0);
		if (pixels > MaxFieldPixels)
		{
			dpi *= Math.Sqrt(MaxFieldPixels / pixels);
		}
		return VectorRaster.Render(shapes, r.X1, r.Y1, r.X2, r.Y2, dpi, pen);
	}

	// The resolution of the picture(s) that cover the region of a scanned page (pixels per inch of the page), or null when there is none.
	private static double? NativePictureDpi(UglyToad.PdfPig.Content.Page? page, FieldRect region)
	{
		if (page == null)
		{
			return null;
		}
		double best = 0.0;
		try
		{
			foreach (UglyToad.PdfPig.Content.IPdfImage image in page.GetImages())
			{
				UglyToad.PdfPig.Core.PdfRectangle box = image.BoundingBox;
				if (box.Width > 1.0 && box.Left < region.X2 && box.Right > region.X1 && box.Bottom < region.Y2 && box.Top > region.Y1)
				{
					best = Math.Max(best, image.WidthInSamples / (box.Width / 72.0));
				}
			}
		}
		catch (Exception)
		{
			return null;
		}
		return best > 0.0 ? best : null;
	}
	// A helper process draws the page when there is a pool (several at once); otherwise, or if the helpers fail, this process does,
	// a few at a time.
	private static GrayImage? RenderRegion(string path, int pageIndex, FieldRect region, double dpi, string logPath)
	{
		RenderPool? pool = _renderPool;
		if (pool != null)
		{
			GrayImage? drawn = pool.TryRender(new RenderRequest(path, pageIndex, region.X1, region.Y1, region.X2, region.Y2, dpi));
			if (drawn != null)
			{
				return drawn;
			}
		}
		RenderGate.Wait();
		try
		{
			return RenderRegionToGray(path, pageIndex, region, dpi, logPath);
		}
		finally
		{
			RenderGate.Release();
		}
	}

	private static GrayImage? LogRenderProblem(string logPath, string path, int pageIndex, string? problem)
	{
		AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR render skipped for {Path.GetFileName(path)} page {pageIndex + 1}: {problem ?? "the page could not be drawn"}");
		return null;
	}

	private static GrayImage? RenderRegionToGray(string path, int pageIndex, FieldRect region, double dpi, string logPath)
	{
		try
		{
			return PdfiumRender.TryRegion(path, pageIndex, region.X1, region.Y1, region.X2, region.Y2, dpi, out string? problem) ?? LogRenderProblem(logPath, path, pageIndex, problem);

		}
		catch (Exception ex) when (ex is not OperationCanceledException)
		{
			AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR render failed for {Path.GetFileName(path)} page {pageIndex + 1}: {ex.Message}");
			return null;
		}
	}
	// Turns the per-sheet reads into index rows, in page order. OCR text is cleaned here, with knowledge of the whole set:
	// disciplines and title words are matched against those seen on text-layer sheets, and sheet IDs that break the numbering
	// pattern are repaired from their neighbours.
	private static (List<DrawingIndexItem> rows, List<(string, int)> sheets, List<(string, int)> skipped, List<FlaggedSheet> flagged) AssembleSheets(SheetRead[] reads, IReadOnlyList<(string path, int page)> sources, bool shortGeneralNumbers, OcrRunStats stats, string logPath, bool finalPass)
	{
		List<DrawingIndexItem> rows = new List<DrawingIndexItem>();
		List<(string, int)> sheets = new List<(string, int)>();
		List<(string, int)> skipped = new List<(string, int)>();
		List<FlaggedSheet> flagged = new List<FlaggedSheet>();
		List<(SheetRead read, int sourceIndex)> data = new List<(SheetRead, int)>();
		for (int i = 0; i < reads.Length; i++)
		{
			SheetRead read = reads[i];
			switch (read.Outcome)
			{
				case SheetOutcome.Skipped:
					skipped.Add((read.SkippedName ?? string.Empty, read.SkippedPages));
					break;
				case SheetOutcome.Dropped:
					break;
				case SheetOutcome.KeptWithoutRow:
					sheets.Add(sources[i]);
					break;
				case SheetOutcome.Cover:
					sheets.Add(sources[i]);
					break;
				default:
					sheets.Add(sources[i]);
					data.Add((read, i));
					break;
			}
		}
		// In folder mode the first kept sheet is the cover even when earlier files were skipped, so it was read like any other.
		int coverSourceIndex = -1;
		if (sheets.Count > 0)
		{
			coverSourceIndex = sources.Select((s, i) => (s, i)).First(x => x.s == sheets[0]).i;
			data.RemoveAll(d => d.sourceIndex == coverSourceIndex);
		}

		OcrVocabulary vocabulary = OcrVocabulary.CreateDefault();
		Dictionary<string, int> ocrDisciplineCounts = new Dictionary<string, int>();
		int hyphenated = 0;
		int textIds = 0;
		foreach ((SheetRead read, int _) in data)
		{
			if (!read.DisciplineFromOcr && !string.IsNullOrWhiteSpace(read.Discipline))
			{
				vocabulary.AddDiscipline(read.Discipline);
			}
			else if (read.DisciplineFromOcr && read.DisciplineConfidence >= 80.0)
			{
				string cleaned = vocabulary.RefineDiscipline(OcrTextCleanup.CleanDiscipline(read.RawDiscipline));
				ocrDisciplineCounts[cleaned] = ocrDisciplineCounts.GetValueOrDefault(cleaned) + 1;
			}
			if (!read.TitleFromOcr && !string.IsNullOrWhiteSpace(read.Title))
			{
				vocabulary.AddTitle(read.Title);
			}
			if (!read.IdFromOcr && !string.IsNullOrWhiteSpace(read.DrawingNumber))
			{
				textIds++;
				if (read.DrawingNumber.IndexOfAny(new[] { '-', '\u2013', '\u2014' }) >= 0)
				{
					hyphenated++;
				}
			}
		}
		vocabulary.LearnDisciplines(ocrDisciplineCounts);
		bool expectHyphen = textIds == 0 || hyphenated * 2 >= textIds;

		foreach ((SheetRead read, int _) in data)
		{
			if (read.IdFromOcr)
			{
				read.DrawingNumber = OcrTextCleanup.CleanSheetId(read.RawDrawingNumber, expectHyphen);
			}
			if (read.DisciplineFromOcr)
			{
				string snapped = vocabulary.SnapDiscipline(vocabulary.RefineDiscipline(OcrTextCleanup.CleanDiscipline(read.RawDiscipline)));
				read.Discipline = read.DisciplineConfidence < 35.0 && !vocabulary.ContainsDiscipline(snapped) ? string.Empty : snapped;
			}
			if (read.TitleFromOcr)
			{
				read.Title = vocabulary.SnapTitle(OcrTextCleanup.CleanTitle(read.RawTitle));
			}
		}
		List<IndexRow> index = data.Select(d => new IndexRow
		{
			Label = d.read.Label,
			Id = d.read.DrawingNumber,
			Discipline = d.read.Discipline,
			Title = d.read.Title,
			IdOcr = d.read.IdFromOcr,
			DisciplineOcr = d.read.DisciplineFromOcr,
			TitleOcr = d.read.TitleFromOcr,
			IdConfidence = d.read.IdConfidence,
			DisciplineConfidence = d.read.DisciplineConfidence,
			TitleConfidence = d.read.TitleConfidence,
			IdCandidates = d.read.IdFromOcr ? d.read.IdCandidates : new List<ReadingCandidate>(),
			DisciplineCandidates = d.read.DisciplineFromOcr ? d.read.DisciplineCandidates : new List<ReadingCandidate>(),
			TitleCandidates = d.read.TitleFromOcr ? d.read.TitleCandidates : new List<ReadingCandidate>(),
			Blank = d.read.Blank
		}).ToList();
		IndexAnalysisSummary analysis = IndexAnalyzer.Analyze(index, vocabulary, IdExampleSetting);
		for (int i = 0; finalPass && i < analysis.IdsRepaired; i++)
		{
			stats.AddIdRepaired();
		}
		if (finalPass)
		AppLog.TryAppend(logPath, $"{DateTime.Now}: Index analysis: {analysis.IdsRepaired} sheet number(s) repaired from numbering patterns, {analysis.TitlesRefined} title(s) refined, {analysis.DisciplinesFilled} discipline(s) inferred, {analysis.FlaggedRows} row(s) flagged for review");

		int sequence = 1;
		if (sheets.Count > 0)
		{
			rows.Add(new DrawingIndexItem
			{
				SeqNo = sequence++,
				DrawingNumber = GeneralDrawingNumber(0, shortGeneralNumbers),
				Discipline = "GENERAL",
				Title = "COVER SHEET",
				SourcePdfPath = sheets[0].Item1,
				SourcePageIndex = sheets[0].Item2
			});
		}
		for (int k = 0; k < data.Count; k++)
		{
			SheetRead read = data[k].read;
			int sourceIndex = data[k].sourceIndex;
			IndexRow analyzed = index[k];
			string id = analyzed.Id;
			string discipline = analyzed.Discipline;
			string title = analyzed.Title;
			id = id.Replace('£', 'E');
			id = id.Replace("/", string.Empty);
			id = id.Replace("=", string.Empty).Replace("%", string.Empty);
			id = CollapseConsecutiveDashes(NormalizeDashesToEnDash(id));
			title = NormalizeDashesToEnDash(title);
			if (read.TitleFromOcr)
			{
				title = CollapseConsecutiveDashes(title);
			}
			if (string.IsNullOrWhiteSpace(id))
			{
				id = "[NOT FOUND]";
			}
			if (string.IsNullOrWhiteSpace(title))
			{
				title = read.Blank ? "[BLANK PAGE]" : "[DESCRIPTION NOT FOUND]";
			}
			bool idUncertain = analyzed.IdUncertain && id != "[NOT FOUND]";
			bool disciplineUncertain = analyzed.DisciplineUncertain;
			bool titleUncertain = analyzed.TitleUncertain && title != "[DESCRIPTION NOT FOUND]";
			bool idBlank = id == "[NOT FOUND]";
			bool titleBlank = title == "[DESCRIPTION NOT FOUND]";
			bool disciplineBlank = string.IsNullOrWhiteSpace(discipline);
			bool wantId = !read.Blank && read.NeededOcr && (idUncertain || idBlank) && (read.IdFromOcr || idBlank);
			bool wantDiscipline = !read.Blank && read.NeededOcr && (disciplineUncertain || disciplineBlank) && (read.DisciplineFromOcr || disciplineBlank);
			bool wantTitle = !read.Blank && read.NeededOcr && (titleUncertain || titleBlank) && (read.TitleFromOcr || titleBlank);
			if (wantId || wantDiscipline || wantTitle)
			{
				flagged.Add(new FlaggedSheet(read, sourceIndex, wantId, wantDiscipline, wantTitle));
			}
			if (read.Blank && finalPass)
			{
				AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR {read.Label}: the title block area is empty, so this page looks blank and nothing was read from it");
			}
			if (read.NeededOcr && !read.Blank)
			{
				bool missing = id == "[NOT FOUND]" || title == "[DESCRIPTION NOT FOUND]" || string.IsNullOrWhiteSpace(discipline);
				if (missing && finalPass)
				{
					stats.AddStillMissing();
				}
				bool doubtful = idUncertain || disciplineUncertain || titleUncertain;
				if ((doubtful || missing) && finalPass)
				{
					string description = $"{read.Label}: {id} | {discipline} | {title}";
					if (doubtful)
					{
						stats.AddLowConfidence(description);
					}
					if (idUncertain)
					{
						AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR candidates for sheet number of {read.Label}: " + string.Join(" | ", read.IdCandidates.Select(c => $"{c.Text.Replace('\n', ' ')}@{c.Confidence:0}{(c.Large ? "L" : string.Empty)}")));
					}
					string why = analyzed.Reasons.Count > 0 ? string.Join("; ", analyzed.Reasons) : (missing ? "field missing" : "low confidence");
					AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR sheet needs review ({why}) - {description}");
				}
				if (finalPass)
				AppLog.TryAppend(logPath, $"{DateTime.Now}: OCR {read.Label}: id=[{id}] (c{read.IdConfidence:0}) discipline=[{discipline}] (c{read.DisciplineConfidence:0}) title=[{title}] (c{read.TitleConfidence:0})");
			}
			rows.Add(new DrawingIndexItem
			{
				SeqNo = sequence++,
				DrawingNumber = id,
				Discipline = discipline,
				Title = title,
				SourcePdfPath = sources[sourceIndex].path,
				SourcePageIndex = sources[sourceIndex].page,
				DrawingNumberUncertain = idUncertain,
				DisciplineUncertain = disciplineUncertain,
				TitleUncertain = titleUncertain,
				Note = string.Join("; ", analyzed.Reasons),
				PageWidth = read.PageWidth,
				PageHeight = read.PageHeight
			});
		}
		return (rows, sheets, skipped, flagged);
	}
}