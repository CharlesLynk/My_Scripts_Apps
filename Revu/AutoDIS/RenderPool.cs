using System;
using System.Collections.Generic;
using System.Threading;

namespace MbrohDrawingIndexCreator
{
	// One connection to a helper that draws page regions. The pool does not care whether it is a separate process or a test double.
	internal interface IRenderConnection : IDisposable
	{
		// False once the helper has died or stopped answering; the pool then discards the connection.
		bool IsAlive { get; }

		// Null when the region could not be drawn (a bad page, or the helper failed). A transport failure also clears IsAlive.
		GrayImage? Render(RenderRequest request);
	}

	// A small pool of render helper processes. PDFium is not thread-safe, so a process draws one page at a time, so extra processes are the
	// way to draw several pages at once. Helpers start the first time a render is wanted (a run that never needs a picture of a page
	// starts none), are reused for the whole run, and are shut down by Dispose. Any failure returns null so the caller can render the
	// page itself, and a pool that keeps failing switches itself off for the rest of the run.
	internal sealed class RenderPool : IDisposable
	{
		private const int MaximumFailures = 4;

		private readonly int _maximum;
		private readonly Func<IRenderConnection> _factory;
		private readonly Action<string>? _log;
		private readonly object _gate = new object();
		private readonly Stack<IRenderConnection> _idle = new Stack<IRenderConnection>();
		private readonly List<IRenderConnection> _all = new List<IRenderConnection>();
		private int _creating;
		private int _failures;
		private bool _disabled;
		private bool _disposed;
		private int _rendered;
		private int _declined;
		private int _startedTotal;

		internal RenderPool(int maximumProcesses, Func<IRenderConnection> factory, Action<string>? log = null)
		{
			_maximum = Math.Max(1, maximumProcesses);
			_factory = factory;
			_log = log;
		}

		internal int Rendered => Volatile.Read(ref _rendered);

		// Requests answered with null (the caller rendered those itself).
		internal int Declined => Volatile.Read(ref _declined);

		// How many helpers were started over the pool's life (replacements of failed helpers count too).
		internal int ProcessesStarted
		{
			get
			{
				lock (_gate)
				{
					return _startedTotal;
				}
			}
		}

		internal bool IsDisabled
		{
			get
			{
				lock (_gate)
				{
					return _disabled || _disposed;
				}
			}
		}

		internal GrayImage? TryRender(RenderRequest request)
		{
			IRenderConnection? connection = Acquire();
			if (connection == null)
			{
				Interlocked.Increment(ref _declined);
				return null;
			}
			GrayImage? image;
			try
			{
				image = connection.Render(request);
			}
			catch (Exception ex) when (ex is not OutOfMemoryException)
			{
				_log?.Invoke($"A render helper failed on page {request.PageIndex + 1}: {ex.Message}");
				Discard(connection);
				Interlocked.Increment(ref _declined);
				return null;
			}
			Release(connection);
			if (image != null)
			{
				Interlocked.Increment(ref _rendered);
			}
			else
			{
				Interlocked.Increment(ref _declined);
			}
			return image;
		}

		private IRenderConnection? Acquire()
		{
			lock (_gate)
			{
				while (true)
				{
					if (_disabled || _disposed)
					{
						return null;
					}
					if (_idle.Count > 0)
					{
						return _idle.Pop();
					}
					if (_all.Count + _creating < _maximum)
					{
						_creating++;
						break;
					}
					Monitor.Wait(_gate);
				}
			}
			IRenderConnection? created = null;
			try
			{
				created = _factory();
			}
			catch (Exception ex) when (ex is not OutOfMemoryException)
			{
				_log?.Invoke($"A render helper could not be started: {ex.Message}");
			}
			lock (_gate)
			{
				_creating--;
				if (created == null)
				{
					NoteFailureLocked();
					Monitor.PulseAll(_gate);
					return null;
				}
				if (_disposed)
				{
					created.Dispose();
					return null;
				}
				_all.Add(created);
				_startedTotal++;
				return created;
			}
		}

		private void Release(IRenderConnection connection)
		{
			lock (_gate)
			{
				if (connection.IsAlive && !_disposed && !_disabled)
				{
					_idle.Push(connection);
				}
				else
				{
					if (!connection.IsAlive && !_disposed)
					{
						NoteFailureLocked();
					}
					_all.Remove(connection);
					connection.Dispose();
				}
				Monitor.PulseAll(_gate);
			}
		}

		// A connection that threw is closed and forgotten; a later render starts a replacement (until too many have failed).
		private void Discard(IRenderConnection connection)
		{
			lock (_gate)
			{
				_all.Remove(connection);
				NoteFailureLocked();
				Monitor.PulseAll(_gate);
			}
			connection.Dispose();
		}

		private void NoteFailureLocked()
		{
			_failures++;
			if (_failures >= MaximumFailures && !_disabled)
			{
				_disabled = true;
				_log?.Invoke("Render helpers failed repeatedly; pages are drawn in the main program for the rest of this run.");
			}
		}

		public void Dispose()
		{
			List<IRenderConnection> toClose;
			lock (_gate)
			{
				if (_disposed)
				{
					return;
				}
				_disposed = true;
				toClose = new List<IRenderConnection>(_all);
				_all.Clear();
				_idle.Clear();
				Monitor.PulseAll(_gate);
			}
			foreach (IRenderConnection connection in toClose)
			{
				try
				{
					connection.Dispose();
				}
				catch (Exception ex) when (ex is not OutOfMemoryException)
				{
				}
			}
		}
	}
}
