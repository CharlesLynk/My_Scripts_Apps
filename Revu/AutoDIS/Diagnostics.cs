using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace MbrohDrawingIndexCreator
{
	// Opt-in timing log. On when the AUTODIS_DIAGNOSTICS environment variable is 1, or when the marker file
	// %TEMP%\AutoDIS\diagnostics.on exists (the packaged app cannot see environment variables set in a terminal).
	// Logs go to %TEMP%\AutoDIS\diagnostics only; the newest 20 runs are kept.
	// Remove: delete this file and every line marked "// DIAG".
	internal static class Diagnostics
	{
		private const int KeepRuns = 20;
		private static readonly object Gate = new object();
		private static readonly Dictionary<string, (int Count, double TotalMs)> Totals = new Dictionary<string, (int Count, double TotalMs)>();
		private static StreamWriter? _writer;

		internal static bool Enabled { get; set; } = string.Equals(Environment.GetEnvironmentVariable("AUTODIS_DIAGNOSTICS"), "1", StringComparison.Ordinal)
			|| File.Exists(Path.Combine(Path.GetTempPath(), "AutoDIS", "diagnostics.on"));

		internal static string RootFolder { get; set; } = Path.Combine(Path.GetTempPath(), "AutoDIS", "diagnostics");

		internal static string? CurrentLogPath { get; private set; }

		internal static Stopwatch? Start() => Enabled ? Stopwatch.StartNew() : null;

		internal static void BeginRun(string runName)
		{
			if (!Enabled)
			{
				return;
			}
			try
			{
				Directory.CreateDirectory(RootFolder);
				PruneOldRuns();
				string path = Path.Combine(RootFolder, $"{DateTime.Now:yyyyMMdd_HHmmss_fff}_{runName}.log");
				lock (Gate)
				{
					_writer?.Dispose();
					_writer = new StreamWriter(path, false, Encoding.UTF8) { AutoFlush = true };
					Totals.Clear();
					CurrentLogPath = path;
					WriteLocked($"Run '{runName}' started");
				}
			}
			catch (Exception)
			{
				// A diagnostics failure must never stop generation, so diagnostics switch themselves off.
				Enabled = false;
			}
		}

		// Records the time since the timer was started or last split, then restarts it.
		internal static void Split(string label, Stopwatch? timer, string? detail = null)
		{
			if (timer == null)
			{
				return;
			}
			double ms = timer.Elapsed.TotalMilliseconds;
			timer.Restart();
			lock (Gate)
			{
				Totals.TryGetValue(label, out (int Count, double TotalMs) t);
				Totals[label] = (t.Count + 1, t.TotalMs + ms);
				WriteLocked($"{label,-24} {ms,10:0.0} ms" + (detail == null ? string.Empty : "  " + detail));
			}
		}

		internal static void EndRun()
		{
			lock (Gate)
			{
				if (_writer == null)
				{
					return;
				}
				WriteLocked("Summary (label, total ms, count, average ms):");
				foreach (KeyValuePair<string, (int Count, double TotalMs)> kv in Totals.OrderByDescending(k => k.Value.TotalMs))
				{
					WriteLocked($"  {kv.Key,-24} {kv.Value.TotalMs,10:0.0} ms  x{kv.Value.Count,-6} avg {kv.Value.TotalMs / kv.Value.Count,8:0.0} ms");
				}
				try
				{
					_writer.Dispose();
				}
				catch (IOException)
				{
				}
				_writer = null;
			}
		}

		// Caller must hold Gate.
		private static void WriteLocked(string line)
		{
			try
			{
				_writer?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} {line}");
			}
			catch (IOException)
			{
			}
		}

		private static void PruneOldRuns()
		{
			// Keep one slot free for the log that BeginRun is about to create, so the total stays at KeepRuns.
			IEnumerable<string> old = Directory.GetFiles(RootFolder, "*.log").OrderByDescending(f => f, StringComparer.Ordinal).Skip(KeepRuns - 1);
			foreach (string file in old)
			{
				try
				{
					File.Delete(file);
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
}
