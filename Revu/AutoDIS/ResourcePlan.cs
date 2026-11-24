using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MbrohDrawingIndexCreator
{
	internal enum ResourceLevel
	{
		Background,
		Normal,
		High,
		Maximum
	}

	internal readonly record struct HardwareSnapshot(int LogicalProcessors, int PhysicalCores, long TotalMemoryBytes, long AvailableMemoryBytes);

	// How much of the computer one run may use: worker threads reading sheets, how many sheets may be drawn to images at once,
	// a memory budget the worker count must fit inside, and the scheduling priority of the process while it runs.
	internal sealed record ResourcePlan(ResourceLevel Level, int Workers, int RenderConcurrency, int RenderProcesses, long MemoryBudgetBytes, ProcessPriorityClass Priority, bool LimitedByMemory);

	internal static class HardwareProbe
	{
		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
		private sealed class MemoryStatus
		{
			public uint Length = (uint)Marshal.SizeOf<MemoryStatus>();
			public uint MemoryLoad;
			public ulong TotalPhys;
			public ulong AvailPhys;
			public ulong TotalPageFile;
			public ulong AvailPageFile;
			public ulong TotalVirtual;
			public ulong AvailVirtual;
			public ulong AvailExtendedVirtual;
		}

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool GlobalMemoryStatusEx([In, Out] MemoryStatus status);

		[DllImport("kernel32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool GetLogicalProcessorInformationEx(int relationship, IntPtr buffer, ref uint length);

		// Cores that actually compute (hyper-threads count once). Extra threads beyond this share a core and give little on this kind of work.
		// Falls back to half the logical processors, never less than one.
		private static int CountPhysicalCores(int logical)
		{
			try
			{
				uint length = 0;
				GetLogicalProcessorInformationEx(0, IntPtr.Zero, ref length);
				if (length == 0)
				{
					return Math.Max(1, logical / 2);
				}
				IntPtr buffer = Marshal.AllocHGlobal((int)length);
				try
				{
					if (!GetLogicalProcessorInformationEx(0, buffer, ref length))
					{
						return Math.Max(1, logical / 2);
					}
					int cores = 0;
					long offset = 0;
					while (offset + 8 <= length)
					{
						// Each record starts with its relationship (int) and its size (uint); relationship 0 is one processor core.
						int relationship = Marshal.ReadInt32(buffer, (int)offset);
						int size = Marshal.ReadInt32(buffer, (int)offset + 4);
						if (size <= 0)
						{
							break;
						}
						if (relationship == 0)
						{
							cores++;
						}
						offset += size;
					}
					return cores > 0 ? Math.Min(cores, logical) : Math.Max(1, logical / 2);
				}
				finally
				{
					Marshal.FreeHGlobal(buffer);
				}
			}
			catch (Exception ex) when (ex is DllNotFoundException || ex is EntryPointNotFoundException || ex is MarshalDirectiveException)
			{
				return Math.Max(1, logical / 2);
			}
		}

		// Never throws: if memory cannot be read, the figures fall back to a modest assumption so the plan stays conservative.
		internal static HardwareSnapshot Probe()
		{
			int processors = Math.Max(1, Environment.ProcessorCount);
			try
			{
				MemoryStatus status = new MemoryStatus();
				if (GlobalMemoryStatusEx(status))
				{
					return new HardwareSnapshot(processors, CountPhysicalCores(processors), (long)status.TotalPhys, (long)status.AvailPhys);
				}
			}
			catch (Exception ex) when (ex is DllNotFoundException || ex is EntryPointNotFoundException || ex is MarshalDirectiveException)
			{
			}
			return new HardwareSnapshot(processors, CountPhysicalCores(processors), 8L * 1024 * 1024 * 1024, 3L * 1024 * 1024 * 1024);
		}
	}

	// The one resource setting the user sees. It is a priority ("how much may this program take"), applied to what this particular
	// computer has, so the same choice means 2 threads on a laptop and more on a workstation.
	internal static class ResourcePlanner
	{
		internal const ResourceLevel Default = ResourceLevel.Normal;

		// What one worker thread needs while it reads sheets (its own PDF reader state and OCR engine), measured with margin.
		private const long BytesPerWorker = 300L * 1024 * 1024;

		// Page drawing stops speeding up beyond a few simultaneous renders, so the plan never asks for more.
		private const int RenderPlateau = 5;

		internal static string DisplayName(ResourceLevel level) => level switch
		{
			ResourceLevel.Background => "Background (stay out of the way)",
			ResourceLevel.Normal => "Normal (recommended)",
			ResourceLevel.High => "High",
			_ => "Maximum"
		};

		internal static string Description(ResourceLevel level) => level switch
		{
			ResourceLevel.Background => "Uses a small share of the processor and memory at low priority, so you can keep working. Slowest.",
			ResourceLevel.Normal => "Uses about half the processor and a moderate share of free memory. A good balance for most work.",
			ResourceLevel.High => "Uses most of the processor and free memory. Other programs may slow down while it runs.",
			_ => "Uses every processor core at raised priority. The computer may feel sluggish until it finishes; on most computers this is only slightly faster than High."
		};

		internal static ResourceLevel ParseOrDefault(string? text) => Enum.TryParse(text, ignoreCase: true, out ResourceLevel level) && Enum.IsDefined(level) ? level : Default;

		internal static ResourcePlan Plan(ResourceLevel level, HardwareSnapshot hardware, long workSetBytes = 0)
		{
			// Measured on an 8-core / 16-thread machine: throughput stops improving at about four workers and does not improve at all beyond the
			// physical core count, so the levels scale with physical cores, not logical processors.
			int cores = Math.Max(1, Math.Min(hardware.PhysicalCores > 0 ? hardware.PhysicalCores : hardware.LogicalProcessors, hardware.LogicalProcessors));
			(int workers, double memoryShare, ProcessPriorityClass priority, int renders) = level switch
			{
				ResourceLevel.Background => (Math.Clamp(cores / 4, 1, 2), 0.20, ProcessPriorityClass.BelowNormal, 1),
				ResourceLevel.Normal => (Math.Clamp(cores / 2, 2, 4), 0.40, ProcessPriorityClass.Normal, 4),
				ResourceLevel.High => (Math.Clamp(cores * 3 / 4, 3, 6), 0.65, ProcessPriorityClass.Normal, 4),
				_ => (Math.Clamp(cores, 2, 8), 0.85, ProcessPriorityClass.AboveNormal, RenderPlateau)
			};
			workers = Math.Min(workers, hardware.LogicalProcessors);
			long budget = (long)(Math.Max(0, hardware.AvailableMemoryBytes) * memoryShare);
			// The memory budget caps the workers: what is left after holding the sheet set, divided by what a worker needs.
			int affordable = (int)Math.Max(1, (budget - Math.Max(0, workSetBytes)) / BytesPerWorker);
			bool limited = affordable < workers;
			workers = Math.Min(workers, affordable);
			// Helper processes that draw pages with PDFium (one process renders about 4 sheets/s; measured with render + OCR, 4
			// processes gave 2.6x, 6 gave 3.4x, 8 gave 3.7x). Conservative: no more than there are workers to use them.
			int helperTarget = level switch
			{
				ResourceLevel.Background => 1,
				ResourceLevel.Normal => Math.Clamp(cores / 2, 1, 4),
				ResourceLevel.High => Math.Clamp(cores * 3 / 4, 2, 6),
				_ => Math.Clamp(cores, 2, 8)
			};
			int helpers = Math.Max(1, Math.Min(helperTarget, workers));
			return new ResourcePlan(level, workers, Math.Clamp(Math.Min(renders, workers), 1, RenderPlateau), helpers, budget, priority, limited);
		}

		internal static string Summary(ResourcePlan plan, HardwareSnapshot hardware)
		{
			double gb = 1024.0 * 1024 * 1024;
			return $"{plan.Level}: {plan.Workers} worker thread(s), up to {plan.RenderProcesses} page-drawing helper process(es), about {plan.MemoryBudgetBytes / gb:0.0} GB memory budget, {plan.Priority} priority"
				+ (plan.LimitedByMemory ? " (worker count limited by available memory)" : string.Empty)
				+ $". This computer: {hardware.PhysicalCores} processor cores ({hardware.LogicalProcessors} threads), {hardware.TotalMemoryBytes / gb:0.0} GB memory, {hardware.AvailableMemoryBytes / gb:0.0} GB free.";
		}
	}
}
