using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace MbrohDrawingIndexCreator
{
	// Copies a file with the operating system's own copy routine (what Explorer uses), which moves data in large, well-pipelined
	// requests. On a network share this was several times quicker than a managed read/write loop. Progress and cancellation work
	// through the routine's callback.
	internal static class FastCopy
	{
		private const uint ProgressContinue = 0;
		private const uint ProgressCancel = 1;

		private delegate uint CopyProgressRoutine(long totalFileSize, long totalBytesTransferred, long streamSize, long streamBytesTransferred, uint streamNumber, uint callbackReason, IntPtr sourceFile, IntPtr destinationFile, IntPtr data);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CopyFileExW")]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool CopyFileEx(string existingFileName, string newFileName, CopyProgressRoutine? progressRoutine, IntPtr data, ref int cancel, uint copyFlags);

		// Throws OperationCanceledException when cancelled (the partial destination is removed) and IOException if the copy fails.
		internal static void Copy(string source, string destination, CancellationToken token, Action<long, long>? progress)
		{
			int cancel = 0;
			CopyProgressRoutine routine = (total, transferred, _, _, _, _, _, _, _) =>
			{
				progress?.Invoke(transferred, total);
				return token.IsCancellationRequested ? ProgressCancel : ProgressContinue;
			};
			bool ok;
			int error = 0;
			try
			{
				ok = CopyFileEx(source, destination, routine, IntPtr.Zero, ref cancel, 0);
				if (!ok)
				{
					error = Marshal.GetLastWin32Error();
				}
			}
			finally
			{
				GC.KeepAlive(routine);
			}
			if (ok)
			{
				return;
			}
			if (token.IsCancellationRequested)
			{
				try
				{
					File.Delete(destination);
				}
				catch (IOException)
				{
				}
				catch (UnauthorizedAccessException)
				{
				}
				throw new OperationCanceledException(token);
			}
			throw new IOException($"Copying '{Path.GetFileName(source)}' failed: {new Win32Exception(error).Message}", new Win32Exception(error));
		}
	}
}
