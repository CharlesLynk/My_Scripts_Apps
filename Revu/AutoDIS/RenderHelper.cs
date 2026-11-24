using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MbrohDrawingIndexCreator
{
	// The connection to one render helper: another copy of this program, started with a switch that makes it do nothing but draw pages on
	// request. Helpers inherit the package identity, read the same local copy of the sheet set, and end by themselves if the program
	// goes away (their input closes).
	internal sealed class ProcessRenderConnection : IRenderConnection
	{
		private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(90);

		private readonly Process _process;
		private readonly Stream _toHelper;
		private readonly Stream _fromHelper;
		private readonly Action<string>? _log;
		private volatile bool _alive = true;

		internal ProcessRenderConnection(Action<string>? log)
		{
			_log = log;
			ProcessStartInfo start = new ProcessStartInfo(Environment.ProcessPath ?? throw new InvalidOperationException("The program's own path is unknown."))
			{
				UseShellExecute = false,
				CreateNoWindow = true,
				RedirectStandardInput = true,
				RedirectStandardOutput = true
			};
			start.ArgumentList.Add(RenderHelper.Switch);
			// The helper only draws pages: a single small heap is enough, and the server collector would reserve a heap per processor.
			start.Environment["DOTNET_gcServer"] = "0";
			start.Environment["DOTNET_gcConcurrent"] = "0";
			_process = Process.Start(start) ?? throw new InvalidOperationException("The render helper did not start.");
			_toHelper = _process.StandardInput.BaseStream;
			_fromHelper = _process.StandardOutput.BaseStream;
		}

		public bool IsAlive
		{
			get
			{
				if (!_alive)
				{
					return false;
				}
				try
				{
					return !_process.HasExited;
				}
				catch (InvalidOperationException)
				{
					return false;
				}
			}
		}

		public GrayImage? Render(RenderRequest request)
		{
			try
			{
				Task<(GrayImage? image, string? error)> exchange = Task.Run(() =>
				{
					RenderProtocol.WriteRequest(_toHelper, request);
					return RenderProtocol.ReadResponse(_fromHelper);
				});
				if (!exchange.Wait(RequestTimeout))
				{
					_log?.Invoke($"A render helper did not answer within {RequestTimeout.TotalSeconds:0} s for page {request.PageIndex + 1}; it was stopped.");
					Stop();
					return null;
				}
				(GrayImage? image, string? error) = exchange.Result;
				if (image == null)
				{
					_log?.Invoke($"A render helper could not draw page {request.PageIndex + 1}: {error}");
				}
				return image;
			}
			catch (Exception ex) when (ex is AggregateException || ex is IOException || ex is InvalidDataException || ex is ObjectDisposedException || ex is InvalidOperationException)
			{
				_log?.Invoke($"A render helper stopped responding: {(ex as AggregateException)?.InnerException?.Message ?? ex.Message}");
				Stop();
				return null;
			}
		}

		private void Stop()
		{
			_alive = false;
			try
			{
				if (!_process.HasExited)
				{
					_process.Kill(true);
				}
			}
			catch (Exception ex) when (ex is InvalidOperationException || ex is System.ComponentModel.Win32Exception)
			{
			}
		}

		public void Dispose()
		{
			if (_alive)
			{
				try
				{
					RenderProtocol.WriteQuit(_toHelper);
					_process.WaitForExit(1000);
				}
				catch (Exception ex) when (ex is IOException || ex is ObjectDisposedException || ex is InvalidOperationException)
				{
				}
			}
			Stop();
			_process.Dispose();
		}
	}

	internal static class RenderHelper
	{
		internal const string Switch = "--render-helper";
		internal const string SelfTestSwitch = "--render-selftest";

		// Entry point of a helper process: answer render requests on standard input/output until told to quit or the input closes.
		internal static int RunHelper()
		{
			int code = 0;
			Thread worker = new Thread(() => code = Serve());
			worker.SetApartmentState(ApartmentState.MTA);
			worker.Start();
			worker.Join();
			return code;
		}

		private static int Serve()
		{
			using Stream input = Console.OpenStandardInput();
			using Stream output = Console.OpenStandardOutput();
			while (true)
			{
				RenderRequest? request;
				try
				{
					request = RenderProtocol.ReadRequest(input);
				}
				catch (Exception ex) when (ex is IOException || ex is InvalidDataException || ex is ObjectDisposedException)
				{
					return 2;
				}
				if (request == null)
				{
					return 0;
				}
				try
				{
					RenderRequest r = request.Value;
					GrayImage? picture = PdfiumRender.TryRegion(r.Path, r.PageIndex, r.X1, r.Y1, r.X2, r.Y2, r.Dpi, out string? problem);
					if (picture == null)
					{
						throw new InvalidOperationException(problem ?? "the page could not be drawn");
					}
					RenderProtocol.WriteImage(output, picture);
				}
				catch (Exception ex) when (ex is not OutOfMemoryException)
				{
					try
					{
						RenderProtocol.WriteFailure(output, ex.Message);
					}
					catch (Exception writeFailure) when (writeFailure is IOException || writeFailure is ObjectDisposedException)
					{
						return 3;
					}
				}
			}
		}

		// Diagnostic, run inside the installed package:  MbrohDrawingIndexCreator.exe --render-selftest <pdf> <resultFile> <pageHeight (unused)> <x1> <y1> <x2> <y2> [pages] [processes]
		// Draws pages through a pool of helper processes and in this process, compares the pictures and writes the findings to resultFile.
		internal static int RunSelfTest(string[] args)
		{
			if (args.Length < 8)
			{
				return 64;
			}
			StringBuilder report = new StringBuilder();
			int code = 0;
			try
			{
				string pdf = args[1];
				string resultFile = args[2];
				double[] v = args.Skip(3).Take(5).Select(a => double.Parse(a, CultureInfo.InvariantCulture)).ToArray();
				int pages = args.Length > 8 ? int.Parse(args[8], CultureInfo.InvariantCulture) : 8;
				int processes = args.Length > 9 ? int.Parse(args[9], CultureInfo.InvariantCulture) : 2;
				report.AppendLine($"self-test of {pages} page(s) through up to {processes} helper process(es); this program is {Environment.ProcessPath}");
				Thread worker = new Thread(() =>
				{
					try
					{
						code = SelfTestBody(pdf, v[1], v[2], v[3], v[4], pages, processes, report);
					}
					catch (Exception ex)
					{
						report.AppendLine("FAILED: " + ex);
						code = 1;
					}
				});
				worker.SetApartmentState(ApartmentState.MTA);
				worker.Start();
				worker.Join();
				File.WriteAllText(resultFile, report.ToString());
			}
			catch (Exception ex)
			{
				return Fail(args.Length > 2 ? args[2] : null, report, ex);
			}
			return code;
		}

		private static int Fail(string? resultFile, StringBuilder report, Exception ex)
		{
			report.AppendLine("FAILED: " + ex);
			if (resultFile != null)
			{
				try
				{
					File.WriteAllText(resultFile, report.ToString());
				}
				catch (IOException)
				{
				}
			}
			return 1;
		}

		private static int SelfTestBody(string pdf, double x1, double y1, double x2, double y2, int pages, int processes, StringBuilder report)
		{
			GrayImage?[] expected = new GrayImage?[pages];
			Stopwatch timer = Stopwatch.StartNew();
			for (int i = 0; i < pages; i++)
			{
				expected[i] = PdfiumRender.TryRegion(pdf, i, x1, y1, x2, y2, 300.0, out _);
			}
			report.AppendLine($"in this process, one after another: {timer.Elapsed.TotalSeconds:0.0} s");
			object gate = new object();
			using RenderPool pool = new RenderPool(processes, () => new ProcessRenderConnection(m => { lock (gate) { report.AppendLine("  helper: " + m); } }), m => { lock (gate) { report.AppendLine("  pool: " + m); } });
			GrayImage?[] viaHelpers = new GrayImage?[pages];
			timer.Restart();
			Parallel.For(0, pages, new ParallelOptions { MaxDegreeOfParallelism = processes }, i => viaHelpers[i] = pool.TryRender(new RenderRequest(pdf, i, x1, y1, x2, y2, 300.0)));
			report.AppendLine($"through the helper pool, {processes} at a time: {timer.Elapsed.TotalSeconds:0.0} s; helper processes started: {pool.ProcessesStarted}, rendered by helpers: {pool.Rendered}, declined: {pool.Declined}");
			int identical = 0;
			for (int i = 0; i < pages; i++)
			{
				GrayImage? local = expected[i];
				GrayImage? other = viaHelpers[i];
				if (local != null && other != null && local.Width == other.Width && local.Height == other.Height && local.Pixels.AsSpan().SequenceEqual(other.Pixels))
				{
					identical++;
				}
			}
			report.AppendLine($"pictures from helpers identical to the in-process pictures: {identical} of {pages}");
			bool ok = pool.Rendered == pages && identical == pages;
			report.AppendLine(ok ? "RESULT: PASS" : "RESULT: FAIL");
			return ok ? 0 : 1;
		}
	}
}