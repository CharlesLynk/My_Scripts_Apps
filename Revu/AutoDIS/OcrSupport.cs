using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Tesseract;

namespace MbrohDrawingIndexCreator
{
	internal sealed class OcrEngineResult
	{
		internal TesseractEngine? Engine { get; init; }

		internal string? ModelUsed { get; init; }

		// True when the preferred model could not be loaded and the other one was used instead.
		internal bool UsedFallbackModel { get; init; }

		// Why OCR is unavailable (Engine is null) or why the preferred model was skipped.
		internal string? Message { get; init; }
	}

	// The engine one worker thread uses. Tesseract engines must not be shared between threads, so every worker owns one.
	internal sealed class OcrEngineSet : IDisposable
	{
		internal TesseractEngine Engine { get; }

		internal string Description { get; }

		// Set when the requested model could not be loaded and the other one is used instead.
		internal string? Message { get; }

		internal OcrEngineSet(TesseractEngine engine, string description, string? message)
		{
			Engine = engine;
			Description = description;
			Message = message;
		}

		public void Dispose()
		{
			Engine.Dispose();
		}
	}
	// Per-run OCR counters. Thread-safe because the sheet loop runs on several threads.
	internal sealed class OcrRunStats
	{
		private const int MaxLowConfidenceExamples = 25;

		private readonly object _gate = new object();
		private readonly List<string> _lowConfidenceSheets = new List<string>();
		private int _sheetsNeedingOcr;
		private int _sheetsOcrAttempted;
		private int _sheetsStillMissing;
		private int _renderFailures;
		private int _lowConfidenceCount;
		private int _idsRepaired;
		private int _secondLookFields;
		private int _secondLookChanged;
		private int _blankPages;
		private int _vectorReads;
		private int _imageReads;
		private int _vectorFallbacks;

		internal bool EngineAvailable { get; set; }

		internal string? ModelUsed { get; set; }

		internal string? EngineMessage { get; set; }

		// True when the large model loaded, so doubtful entries get a second look (Thorough preset).
		internal bool SecondLookAvailable { get; set; }

		internal int SecondLookFields => Volatile.Read(ref _secondLookFields);

		internal int SecondLookChanged => Volatile.Read(ref _secondLookChanged);

		internal void AddSecondLook(bool changed)
		{
			Interlocked.Increment(ref _secondLookFields);
			if (changed)
			{
				Interlocked.Increment(ref _secondLookChanged);
			}
		}

		// Sheets whose title block was read from the page's vector shapes, sheets drawn as pictures instead, and of those the ones where the vector reading was not usable.
		// Pages whose title block area is empty (blank sheets in a scan); they are neither read nor counted as failures.
		internal int BlankPages => Volatile.Read(ref _blankPages);

		internal void AddBlank() => Interlocked.Increment(ref _blankPages);

		internal int VectorReads => Volatile.Read(ref _vectorReads);

		internal int ImageReads => Volatile.Read(ref _imageReads);

		internal int VectorFallbacks => Volatile.Read(ref _vectorFallbacks);

		internal void AddVectorRead() => Interlocked.Increment(ref _vectorReads);

		private int _vectorDeclinesLogged;

		// Only the first few declined sheets are written to the log; a scanned set would otherwise add a line per sheet.
		internal bool ShouldLogVectorDecline() => Interlocked.Increment(ref _vectorDeclinesLogged) <= 5;

		internal void AddImageRead(bool afterVector)
		{
			Interlocked.Increment(ref _imageReads);
			if (afterVector)
			{
				Interlocked.Increment(ref _vectorFallbacks);
			}
		}

		internal int SheetsNeedingOcr => Volatile.Read(ref _sheetsNeedingOcr);

		internal int SheetsOcrAttempted => Volatile.Read(ref _sheetsOcrAttempted);

		internal int SheetsStillMissing => Volatile.Read(ref _sheetsStillMissing);

		internal int RenderFailures => Volatile.Read(ref _renderFailures);

		internal int LowConfidenceCount => Volatile.Read(ref _lowConfidenceCount);

		internal int IdsRepaired => Volatile.Read(ref _idsRepaired);

		internal IReadOnlyList<string> LowConfidenceSheets
		{
			get
			{
				lock (_gate)
				{
					return _lowConfidenceSheets.ToArray();
				}
			}
		}

		internal void AddNeedingOcr() => Interlocked.Increment(ref _sheetsNeedingOcr);

		internal void AddAttempted() => Interlocked.Increment(ref _sheetsOcrAttempted);

		internal void AddStillMissing() => Interlocked.Increment(ref _sheetsStillMissing);

		internal void AddRenderFailure() => Interlocked.Increment(ref _renderFailures);

		internal void AddIdRepaired() => Interlocked.Increment(ref _idsRepaired);

		internal void AddLowConfidence(string sheetDescription)
		{
			Interlocked.Increment(ref _lowConfidenceCount);
			lock (_gate)
			{
				if (_lowConfidenceSheets.Count < MaxLowConfidenceExamples)
				{
					_lowConfidenceSheets.Add(sheetDescription);
				}
			}
		}

		// A run needs the user's attention when OCR was required but did not fully deliver.
		internal bool HasProblem => SheetsNeedingOcr > 0 && (!EngineAvailable || SheetsStillMissing > 0 || RenderFailures > 0 || LowConfidenceCount > 0);

		internal string Describe()
		{
			string engine = EngineAvailable ? $"OCR engines '{ModelUsed}'" : $"OCR engine unavailable ({EngineMessage})";
			return $"{engine}; sheets needing OCR: {SheetsNeedingOcr}, OCR attempted: {SheetsOcrAttempted}, still missing fields afterwards: {SheetsStillMissing}, low confidence: {LowConfidenceCount}, sheet IDs repaired from sequence: {IdsRepaired}, second look: {SecondLookFields} field(s) re-read, {SecondLookChanged} changed, blank pages: {BlankPages}, read from vector shapes: {VectorReads}, drawn as pictures: {ImageReads} ({VectorFallbacks} after a vector attempt), render failures: {RenderFailures}";
		}
	}

	internal static class OcrSupport
	{
		internal const string BestModel = "eng_best";

		internal const string FastModel = "eng";

		internal static string DefaultTessdataDirectory => Path.Combine(AppContext.BaseDirectory, "tessdata");

		// Returns an engine for the preferred model, or the other model if that one is unusable.
		// Never throws; a failure is reported in Message so callers can tell the user instead of silently skipping OCR.
		internal static OcrEngineResult TryCreateEngine(string tessdataDirectory, string preferredModel)
		{
			if (!Directory.Exists(tessdataDirectory))
			{
				return new OcrEngineResult { Message = $"the OCR language-data folder is missing: {tessdataDirectory}" };
			}
			string otherModel = string.Equals(preferredModel, BestModel, StringComparison.OrdinalIgnoreCase) ? FastModel : BestModel;
			List<string> problems = new List<string>();
			foreach (string model in new[] { preferredModel, otherModel })
			{
				string file = Path.Combine(tessdataDirectory, model + ".traineddata");
				if (!File.Exists(file))
				{
					problems.Add($"{model}.traineddata is missing from {tessdataDirectory}");
					continue;
				}
				try
				{
					TesseractEngine engine = new TesseractEngine(tessdataDirectory, model, EngineMode.LstmOnly);
					bool fallback = !string.Equals(model, preferredModel, StringComparison.OrdinalIgnoreCase);
					return new OcrEngineResult
					{
						Engine = engine,
						ModelUsed = model,
						UsedFallbackModel = fallback,
						Message = fallback ? string.Join("; ", problems) : null
					};
				}
				catch (Exception ex)
				{
					problems.Add($"model '{model}' failed to load ({new FileInfo(file).Length:N0} bytes): {ex.Message}");
				}
			}
			return new OcrEngineResult { Message = string.Join("; ", problems) };
		}

		// useBestModel: the larger, more accurate model (the Thorough preset) instead of the fast one.
		// Returns null when no model can be loaded; reason then says why (or why the other model was used).
		internal static OcrEngineSet? TryCreateEngineSet(string tessdataDirectory, bool useBestModel, out string? reason)
		{
			OcrEngineResult result = TryCreateEngine(tessdataDirectory, useBestModel ? BestModel : FastModel);
			reason = result.Message;
			return result.Engine == null ? null : new OcrEngineSet(result.Engine, result.ModelUsed ?? string.Empty, result.Message);
		}	}
}
