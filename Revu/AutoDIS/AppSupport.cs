using System;
using System.IO;

namespace MbrohDrawingIndexCreator
{
	internal static class AppLog
	{
		private static readonly object Gate = new object();

		// Log writes must never cause a generation to fail, so write errors are ignored.
		// Serialized because the sheet-reading workers log from several threads.
		internal static void TryAppend(string logPath, string message)
		{
			try
			{
				lock (Gate)
				{
					File.AppendAllText(logPath, message.EndsWith("\n", StringComparison.Ordinal) ? message : message + "\n");
				}
			}
			catch (IOException)
			{
			}
			catch (UnauthorizedAccessException)
			{
			}
		}
	}

	internal static class FileHygiene
	{
		// Output is written beside the destination under a unique name, then moved into place.
		// A crash or cancel therefore never leaves a half-written file at the chosen path.
		internal static string PartialPathFor(string finalPath)
		{
			string directory = Path.GetDirectoryName(finalPath) ?? string.Empty;
			string name = Path.GetFileName(finalPath);
			string token = Guid.NewGuid().ToString("N").Substring(0, 8);
			return Path.Combine(directory, $"{name}.{token}.partial");
		}

		// Throws IOException if the destination is open in another program.
		internal static void Commit(string partialPath, string finalPath)
		{
			File.Move(partialPath, finalPath, overwrite: true);
		}

		internal static void DeleteQuietly(string path)
		{
			try
			{
				if (File.Exists(path))
				{
					File.Delete(path);
				}
			}
			catch (IOException)
			{
			}
			catch (UnauthorizedAccessException)
			{
			}
		}

		// Deletes files matching the pattern whose last write is older than maxAge.
		internal static int SweepStale(string folder, string pattern, TimeSpan maxAge)
		{
			int removed = 0;
			try
			{
				if (!Directory.Exists(folder))
				{
					return 0;
				}
				DateTime cutoff = DateTime.Now - maxAge;
				foreach (string file in Directory.EnumerateFiles(folder, pattern))
				{
					try
					{
						if (File.GetLastWriteTime(file) < cutoff)
						{
							File.Delete(file);
							removed++;
						}
					}
					catch (IOException)
					{
					}
					catch (UnauthorizedAccessException)
					{
					}
				}
			}
			catch (IOException)
			{
			}
			catch (UnauthorizedAccessException)
			{
			}
			return removed;
		}
	}
}
