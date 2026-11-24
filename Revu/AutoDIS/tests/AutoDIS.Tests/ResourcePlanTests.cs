using System;
using System.Linq;
using System.Diagnostics;
using MbrohDrawingIndexCreator;
using Xunit;

namespace AutoDIS.Tests;

public class ResourcePlanTests
{
	private static readonly long Gb = 1024L * 1024 * 1024;
	// Hyper-threaded: two logical processors per core.
	private static HardwareSnapshot Machine(int threads, long totalGb, long freeGb) => new(threads, Math.Max(1, threads / 2), totalGb * Gb, freeGb * Gb);

	[Fact]
	public void Levels_AreOrderedByWorkers_OnTypicalLaptop()
	{
		HardwareSnapshot hw = Machine(16, 32, 20);
		int[] workers = Enum.GetValues<ResourceLevel>().Select(l => ResourcePlanner.Plan(l, hw).Workers).ToArray();
		Assert.Equal(workers.OrderBy(w => w), workers);
		Assert.True(workers[0] < workers[^1]);
	}

	[Fact]
	public void Normal_MatchesTheTunedDefault()
	{
		ResourcePlan plan = ResourcePlanner.Plan(ResourceLevel.Normal, Machine(16, 32, 20));
		Assert.Equal(4, plan.Workers);
		Assert.Equal(4, plan.RenderConcurrency);
		Assert.Equal(ProcessPriorityClass.Normal, plan.Priority);
	}

	[Fact]
	public void Background_IsGentle()
	{
		ResourcePlan plan = ResourcePlanner.Plan(ResourceLevel.Background, Machine(16, 32, 20));
		Assert.True(plan.Workers <= 2);
		Assert.Equal(1, plan.RenderConcurrency);
		Assert.Equal(ProcessPriorityClass.BelowNormal, plan.Priority);
	}

	[Fact]
	public void Maximum_UsesEveryPhysicalCore_AndRaisesPriority()
	{
		ResourcePlan plan = ResourcePlanner.Plan(ResourceLevel.Maximum, Machine(16, 32, 20));
		Assert.Equal(8, plan.Workers);
		Assert.Equal(ProcessPriorityClass.AboveNormal, plan.Priority);
	}

	[Fact]
	public void SmallMachine_NeverExceedsItsThreads()
	{
		foreach (ResourceLevel level in Enum.GetValues<ResourceLevel>())
		{
			ResourcePlan plan = ResourcePlanner.Plan(level, Machine(2, 8, 6));
			Assert.InRange(plan.Workers, 1, 2);
			Assert.InRange(plan.RenderConcurrency, 1, plan.Workers);
		}
	}

	[Fact]
	public void LowFreeMemory_LimitsWorkers_ButNeverBelowOne()
	{
		ResourcePlan plan = ResourcePlanner.Plan(ResourceLevel.Maximum, Machine(16, 8, 1));
		Assert.True(plan.LimitedByMemory);
		Assert.InRange(plan.Workers, 1, 3);
		ResourcePlan starved = ResourcePlanner.Plan(ResourceLevel.Maximum, new HardwareSnapshot(16, 8, 8 * Gb, 0));
		Assert.Equal(1, starved.Workers);
	}

	[Fact]
	public void LargeSheetSet_ReducesAffordableWorkers()
	{
		HardwareSnapshot hw = Machine(16, 16, 4);
		int small = ResourcePlanner.Plan(ResourceLevel.High, hw, 50L * 1024 * 1024).Workers;
		int large = ResourcePlanner.Plan(ResourceLevel.High, hw, 2L * Gb).Workers;
		Assert.True(large < small);
	}

	[Theory]
	[InlineData("High", "High")]
	[InlineData("maximum", "Maximum")]
	[InlineData("", "Normal")]
	[InlineData("garbage", "Normal")]
	[InlineData("42", "Normal")]
	[InlineData(null, "Normal")]
	public void Parse_FallsBackToNormal(string? text, string expectedName)
	{
		Assert.Equal(expectedName, ResourcePlanner.ParseOrDefault(text).ToString());
	}

	[Fact]
	public void Probe_ReturnsPlausibleHardware()
	{
		HardwareSnapshot hw = HardwareProbe.Probe();
		Assert.True(hw.LogicalProcessors >= 1);
		Assert.True(hw.TotalMemoryBytes > 0);
		Assert.InRange(hw.AvailableMemoryBytes, 1, hw.TotalMemoryBytes);
	}

	[Fact]
	public void RenderHelpersFollowTheLevelAndNeverExceedTheWorkers()
	{
		HardwareSnapshot hw = Machine(16, 32, 20);
		int[] helpers = Enum.GetValues<ResourceLevel>().Select(l => ResourcePlanner.Plan(l, hw).RenderProcesses).ToArray();

		Assert.Equal(new[] { 1, 4, 6, 8 }, helpers);
		foreach (ResourceLevel level in Enum.GetValues<ResourceLevel>())
		{
			ResourcePlan plan = ResourcePlanner.Plan(level, Machine(4, 8, 6));
			Assert.InRange(plan.RenderProcesses, 1, plan.Workers);
		}
	}
}