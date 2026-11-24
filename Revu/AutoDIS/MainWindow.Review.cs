using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Storage.Streams;

namespace MbrohDrawingIndexCreator;

// The review step between reading the sheets and writing the index: the user can rename, remove or split discipline headings. Headings and
// page breaks of the index follow from the disciplines of the rows, so changing them re-flows everything after when the index is built.
public sealed partial class MainWindow
{
	// The window that shows the review dialog for the run in progress (set while a run is active).
	private static MainWindow? _reviewHost;

	private static bool ReviewBeforeSaveSetting = LoadReviewBeforeSave();

	private static bool LoadReviewBeforeSave()
	{
		return !LoadSettingsFile().TryGetValue("ReviewBeforeSave", out string value) || !string.Equals(value, "False", StringComparison.OrdinalIgnoreCase);
	}

	private static void SaveReviewBeforeSave(bool value)
	{
		Dictionary<string, string> dictionary = LoadSettingsFile();
		dictionary["ReviewBeforeSave"] = value.ToString();
		SaveSettingsFile(dictionary);
	}

	// Called from the worker thread that builds the index: shows the dialog on the window's thread and waits for the answer. The rows are
	// edited in place. Throws OperationCanceledException when the user cancels.
	private static void ReviewEntries(List<DrawingIndexItem> rows, FieldRect disciplineRect, FieldRect titleRect, string logPath, IProgress<(int current, int total, string status)> progress, CancellationToken token)
	{
		MainWindow? host = _reviewHost;
		if (host == null || !ReviewBeforeSaveSetting || rows.Count == 0)
		{
			return;
		}
		progress.Report((0, 0, "Waiting for your review of the discipline headings..."));
		var waitTimer = Diagnostics.Start(); // DIAG
		using ManualResetEventSlim finished = new ManualResetEventSlim(false);
		bool proceed = true;
		Exception? failure = null;
		bool queued = host.DispatcherQueue.TryEnqueue(async () =>
		{
			try
			{
				proceed = await host.ShowReviewDialogAsync(rows, disciplineRect, titleRect, logPath);
			}
			catch (Exception ex)
			{
				failure = ex;
			}
			finally
			{
				finished.Set();
			}
		});
		if (!queued)
		{
			return;
		}
		finished.Wait(token);
		Diagnostics.Split("review dialog (waiting for the user)", waitTimer); // DIAG
		if (failure != null)
		{
			AppLog.TryAppend(logPath, $"{DateTime.Now}: The review dialog could not be shown: {failure.Message}");
			return;
		}
		if (!proceed)
		{
			throw new OperationCanceledException(token);
		}
		AppLog.TryAppend(logPath, $"{DateTime.Now}: Discipline headings reviewed by the user; the index continues with {rows.Count} row(s) in {GroupByDiscipline(rows).Count} heading group(s).");
	}

	private static readonly SolidColorBrush Orange = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 255, 204, 128));
	private static readonly SolidColorBrush Black = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 0, 0, 0));

	private static void MarkUncertain(TextBox box, bool uncertain)
	{
		if (uncertain)
		{
			box.Background = Orange;
			box.Foreground = Black;
		}
		else
		{
			box.ClearValue(Control.BackgroundProperty);
			box.ClearValue(Control.ForegroundProperty);
		}
	}

	// Runs of consecutive rows with the same discipline: each run gets a heading in the index.
	private static List<List<DrawingIndexItem>> GroupByDiscipline(List<DrawingIndexItem> rows)
	{
		List<List<DrawingIndexItem>> groups = new List<List<DrawingIndexItem>>();
		foreach (DrawingIndexItem row in rows)
		{
			if (groups.Count == 0 || !string.Equals((groups[^1][0].Discipline ?? string.Empty).Trim(), (row.Discipline ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase))
			{
				groups.Add(new List<DrawingIndexItem>());
			}
			groups[^1].Add(row);
		}
		return groups;
	}

	// Up to three sheets of a heading group worth looking at: the first doubtful one, then the first and the last.
	private static List<DrawingIndexItem> SamplesOf(List<DrawingIndexItem> group)
	{
		List<DrawingIndexItem> samples = new List<DrawingIndexItem>();
		DrawingIndexItem? doubtful = group.FirstOrDefault(r => r.DisciplineUncertain);
		foreach (DrawingIndexItem candidate in new[] { doubtful, group[0], group[^1] })
		{
			if (candidate != null && !samples.Contains(candidate))
			{
				samples.Add(candidate);
			}
		}
		return samples;
	}

	// Small pictures of every sheet's discipline and title area, drawn in the background while the dialog is open (by the render helper
	// processes, a few at a time) and kept in memory as PNG: about 10-20 KB each, so a few MB for a whole set, nothing on disk. A picture
	// that is asked for before its turn is drawn at once.
	private sealed class ReviewPreviews : IDisposable
	{
		private readonly ConcurrentDictionary<DrawingIndexItem, Lazy<Task<byte[]?>>> _pictures = new ConcurrentDictionary<DrawingIndexItem, Lazy<Task<byte[]?>>>();
		private readonly Func<DrawingIndexItem, byte[]?> _draw;
		private readonly SemaphoreSlim _turns;
		private readonly CancellationTokenSource _stop = new CancellationTokenSource();

		internal ReviewPreviews(Func<DrawingIndexItem, byte[]?> draw, int parallel)
		{
			_draw = draw;
			_turns = new SemaphoreSlim(Math.Max(1, parallel));
		}

		internal Task<byte[]?> GetAsync(DrawingIndexItem row)
		{
			return _pictures.GetOrAdd(row, r => new Lazy<Task<byte[]?>>(() => Task.Run(async () =>
			{
				try
				{
					await _turns.WaitAsync(_stop.Token);
					try
					{
						return _draw(r);
					}
					finally
					{
						_turns.Release();
					}
				}
				catch (Exception ex) when (ex is OperationCanceledException || ex is ObjectDisposedException)
				{
					return null;
				}
			}))).Value;
		}

		// Queues pictures in the given order: the sheets people look at first (samples of each heading), then everything else.
		internal void Prefetch(IEnumerable<DrawingIndexItem> rows)
		{
			foreach (DrawingIndexItem row in rows)
			{
				_ = GetAsync(row);
			}
		}

		public void Dispose()
		{
			_stop.Cancel();
		}
	}

	// The discipline and title areas of a sheet, with room around them, as a small PNG (the helper processes draw it with PDFium, so scans work).
	private static byte[]? DrawFieldsPreview(DrawingIndexItem row, FieldRect disciplineRect, FieldRect titleRect, string logPath)
	{
		(_, disciplineRect, titleRect) = FieldsForPage(row.PageWidth, row.PageHeight, disciplineRect, disciplineRect, titleRect);
		double x1 = Math.Max(0.0, Math.Min(disciplineRect.X1, titleRect.X1) - 12.0);
		double y1 = Math.Max(0.0, Math.Min(disciplineRect.Y1, titleRect.Y1) - 8.0);
		double x2 = Math.Max(disciplineRect.X2, titleRect.X2) + 12.0;
		double y2 = Math.Max(disciplineRect.Y2, titleRect.Y2) + 8.0;
		double dpi = Math.Clamp(760.0 / Math.Max(1.0, x2 - x1) * 72.0, 72.0, 200.0);
		GrayImage? picture = RenderRegion(row.SourcePdfPath, row.SourcePageIndex, new FieldRect(x1, y1, x2, y2), dpi, logPath);
		if (picture == null)
		{
			return null;
		}
#pragma warning disable CA1416 // Windows-only System.Drawing is intentional.
		using MemoryStream bmp = new MemoryStream(picture.ToBmp());
		using System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(bmp);
		using MemoryStream png = new MemoryStream();
		bitmap.Save(png, ImageFormat.Png);
		return png.ToArray();
#pragma warning restore CA1416
	}

	private static async Task<BitmapImage?> LoadBitmapAsync(byte[]? png)
	{
		if (png == null)
		{
			return null;
		}
		BitmapImage bitmap = new BitmapImage();
		using InMemoryRandomAccessStream stream = new InMemoryRandomAccessStream();
		using (DataWriter writer = new DataWriter(stream))
		{
			writer.WriteBytes(png);
			await writer.StoreAsync();
			writer.DetachStream();
		}
		stream.Seek(0UL);
		await bitmap.SetSourceAsync(stream);
		return bitmap;
	}

	// An image that fills itself in when the picture of the sheet is ready (immediately when it already is).
	private static Grid PreviewImage(ReviewPreviews previews, DrawingIndexItem row, double width)
	{
		Grid cell = new Grid { MaxWidth = width, MinHeight = 60.0, HorizontalAlignment = HorizontalAlignment.Stretch };
		TextBlock waiting = new TextBlock { Text = "drawing the title block...", Opacity = 0.6, VerticalAlignment = VerticalAlignment.Center, FontSize = 12.0 };
		cell.Children.Add(waiting);
		_ = FillAsync();
		return cell;

		async Task FillAsync()
		{
			BitmapImage? bitmap = await LoadBitmapAsync(await previews.GetAsync(row));
			cell.Children.Clear();
			if (bitmap == null)
			{
				cell.Children.Add(new TextBlock { Text = "(this sheet could not be drawn)", Opacity = 0.6, VerticalAlignment = VerticalAlignment.Center, FontSize = 12.0 });
				return;
			}
			cell.Children.Add(new Image { Source = bitmap, Stretch = Stretch.Uniform, HorizontalAlignment = HorizontalAlignment.Left, MaxHeight = 110.0 });
		}
	}

	private static readonly Windows.UI.Color HoverFill = Windows.UI.Color.FromArgb(34, 0, 120, 215);
	private static readonly Windows.UI.Color AccentBar = Windows.UI.Color.FromArgb(255, 0, 120, 215);
	private static readonly Windows.UI.Color Clear = Windows.UI.Color.FromArgb(0, 0, 0, 0);

	// A heading row: it lights up under the pointer (a bar at its left edge and a tint across it) so it is plain which row the buttons at its
	// right belong to; its own brush is also what flashes when it absorbs other rows.
	private sealed class HeadingRow
	{
		internal readonly Border Host;
		internal readonly SolidColorBrush Fill = new SolidColorBrush(Clear);
		internal readonly SolidColorBrush Bar = new SolidColorBrush(Clear);

		internal HeadingRow(UIElement content)
		{
			Host = new Border
			{
				Child = content,
				Background = Fill,
				BorderBrush = Bar,
				BorderThickness = new Thickness(4, 0, 0, 0),
				CornerRadius = new CornerRadius(4),
				Padding = new Thickness(6, 5, 6, 5)
			};
			Host.PointerEntered += (object s, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs a) =>
			{
				Fill.Color = HoverFill;
				Bar.Color = AccentBar;
			};
			Host.PointerExited += (object s, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs a) =>
			{
				Fill.Color = Clear;
				Bar.Color = Clear;
			};
			// Clip to the row's own (possibly animated) size so a row folding away does not spill over its neighbours.
			Microsoft.UI.Composition.Visual visual = Microsoft.UI.Xaml.Hosting.ElementCompositionPreview.GetElementVisual(Host);
			visual.Clip = visual.Compositor.CreateInsetClip();
		}
	}

	private const double AnimationSeconds = 0.28;

	// Folds a row away (height and opacity to nothing); `done` runs when it has gone.
	private static void Collapse(FrameworkElement host, Action done)
	{
		Microsoft.UI.Xaml.Media.Animation.Storyboard story = new Microsoft.UI.Xaml.Media.Animation.Storyboard();
		host.Height = host.ActualHeight;
		AddAnimation(story, host, "Height", host.ActualHeight, 0.0, true);
		AddAnimation(story, host, "Opacity", 1.0, 0.0, false);
		story.Completed += (object? s, object a) => done();
		story.Begin();
	}

	// Unfolds a new row (height and opacity from nothing to its natural size).
	private static void Expand(FrameworkElement host, double width)
	{
		host.Measure(new Windows.Foundation.Size(Math.Max(1.0, width), double.PositiveInfinity));
		double natural = host.DesiredSize.Height;
		host.Height = 0.0;
		host.Opacity = 0.0;
		Microsoft.UI.Xaml.Media.Animation.Storyboard story = new Microsoft.UI.Xaml.Media.Animation.Storyboard();
		AddAnimation(story, host, "Height", 0.0, natural, true);
		AddAnimation(story, host, "Opacity", 0.0, 1.0, false);
		story.Completed += (object? s, object a) =>
		{
			host.ClearValue(FrameworkElement.HeightProperty);
			host.Opacity = 1.0;
		};
		story.Begin();
	}

	private static void AddAnimation(Microsoft.UI.Xaml.Media.Animation.Storyboard story, FrameworkElement target, string property, double from, double to, bool dependent)
	{
		Microsoft.UI.Xaml.Media.Animation.DoubleAnimation animation = new Microsoft.UI.Xaml.Media.Animation.DoubleAnimation
		{
			From = from,
			To = to,
			Duration = new Duration(TimeSpan.FromSeconds(AnimationSeconds)),
			EnableDependentAnimation = dependent,
			EasingFunction = new Microsoft.UI.Xaml.Media.Animation.CubicEase { EasingMode = Microsoft.UI.Xaml.Media.Animation.EasingMode.EaseInOut }
		};
		Microsoft.UI.Xaml.Media.Animation.Storyboard.SetTarget(animation, target);
		Microsoft.UI.Xaml.Media.Animation.Storyboard.SetTargetProperty(animation, property);
		story.Children.Add(animation);
	}

	// A short glow on a row that has just taken in others.
	private static void Flash(SolidColorBrush brush)
	{
		Microsoft.UI.Xaml.Media.Animation.Storyboard story = new Microsoft.UI.Xaml.Media.Animation.Storyboard();
		Microsoft.UI.Xaml.Media.Animation.ColorAnimation animation = new Microsoft.UI.Xaml.Media.Animation.ColorAnimation
		{
			From = Windows.UI.Color.FromArgb(120, 0, 120, 215),
			To = Clear,
			Duration = new Duration(TimeSpan.FromSeconds(1.1))
		};
		Microsoft.UI.Xaml.Media.Animation.Storyboard.SetTarget(animation, brush);
		Microsoft.UI.Xaml.Media.Animation.Storyboard.SetTargetProperty(animation, "Color");
		story.Children.Add(animation);
		story.Begin();
	}

	private async Task<bool> ShowReviewDialogAsync(List<DrawingIndexItem> rows, FieldRect disciplineRect, FieldRect titleRect, string logPath)
	{
		using ReviewPreviews previews = new ReviewPreviews(r => DrawFieldsPreview(r, disciplineRect, titleRect, logPath), Math.Max(2, _resourcePlan.Workers));
		previews.Prefetch(GroupByDiscipline(rows).SelectMany(SamplesOf).Concat(rows).Distinct());
		Window? reviewWindow = null;

		StackPanel groupsPanel = new StackPanel { Spacing = 2.0 };
		string[] original = rows.Select(r => r.Discipline ?? string.Empty).ToArray();
		bool[] originalUncertain = rows.Select(r => r.DisciplineUncertain).ToArray();
		Stack<string[]> history = new Stack<string[]>();
		Button undo = new Button { Content = "Undo last change", IsEnabled = false };
		Button reset = new Button { Content = "Reset all headings to the original", IsEnabled = false };
		// The rows now on screen, by the first sheet of their heading, so a change can animate rows that go away or come in.
		Dictionary<DrawingIndexItem, HeadingRow> shown = new Dictionary<DrawingIndexItem, HeadingRow>();
		bool animating = false;

		void Snapshot() => history.Push(rows.Select(r => r.Discipline ?? string.Empty).ToArray());

		void UpdateButtons()
		{
			undo.IsEnabled = history.Count > 0;
			reset.IsEnabled = !rows.Select(r => r.Discipline ?? string.Empty).SequenceEqual(original);
		}

		void Restore(string[] disciplines)
		{
			for (int i = 0; i < rows.Count; i++)
			{
				rows[i].Discipline = disciplines[i];
				rows[i].DisciplineUncertain = string.Equals(disciplines[i], original[i], StringComparison.Ordinal) && originalUncertain[i];
			}
		}

		void RebuildGroups()
		{
			groupsPanel.Children.Clear();
			shown.Clear();
			UpdateButtons();
			List<List<DrawingIndexItem>> groups = GroupByDiscipline(rows);
			for (int g = 0; g < groups.Count; g++)
			{
				List<DrawingIndexItem> group = groups[g];
				int index = g;
				Grid line = new Grid { ColumnSpacing = 8.0 };
				line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(36.0) });
				line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2.0, GridUnitType.Star), MinWidth = 150.0 });
				line.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(3.0, GridUnitType.Star), MinWidth = 120.0 });
				line.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
				line.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
				line.Children.Add(BuildPreviewIcon(group, previews));
				TextBox name = new TextBox { Text = group[0].Discipline ?? string.Empty, PlaceholderText = "(no heading)" };
				MarkUncertain(name, group.Any(r => r.DisciplineUncertain) || group.Count <= 2);

				// The typed heading is saved as soon as the box loses focus, without redrawing the list: a redraw at that moment would remove
				// the button the user is in the middle of clicking (Remove heading, Split...) before the click arrives.
				void CommitName()
				{
					string typed = (name.Text ?? string.Empty).Trim();
					if (!string.Equals(typed, (group[0].Discipline ?? string.Empty).Trim(), StringComparison.Ordinal))
					{
						Snapshot();
						foreach (DrawingIndexItem row in group)
						{
							row.Discipline = typed;
							row.DisciplineUncertain = false;
						}
						MarkUncertain(name, false);
						UpdateButtons();
					}
				}

				name.LostFocus += (object s, RoutedEventArgs a) => CommitName();
				name.KeyDown += (object s, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs a) =>
				{
					if (a.Key == Windows.System.VirtualKey.Enter)
					{
						ApplyChange(CommitName);
					}
				};
				string first = group[0].DrawingNumber;
				string last = group[^1].DrawingNumber;
				TextBlock info = new TextBlock
				{
					Text = group.Count == 1 ? $"1 sheet: {first}" : $"{group.Count} sheets: {first} to {last}",
					VerticalAlignment = VerticalAlignment.Center,
					TextWrapping = TextWrapping.Wrap
				};
				Button split = new Button { Content = "Split...", IsEnabled = group.Count >= 2 };
				split.Click += async (object s, RoutedEventArgs a) =>
				{
					CommitName();
					await ShowSplitWindowAsync(group, rows, previews, reviewWindow, Snapshot, () => ApplyChange(() => { }));
				};
				Button merge = new Button { Content = "Remove heading", IsEnabled = index > 0 };
				ToolTipService.SetToolTip(merge, "The sheets of this heading join the heading above it");
				merge.Click += (object s, RoutedEventArgs a) =>
				{
					CommitName();
					ApplyChange(() =>
					{
						// Regrouped after the box was saved: renaming the heading to the one above may already have joined the two.
						List<List<DrawingIndexItem>> fresh = GroupByDiscipline(rows);
						int position = fresh.FindIndex(candidate => candidate.Contains(group[0]));
						if (position > 0 && ReferenceEquals(fresh[position][0], group[0]))
						{
							Snapshot();
							string above = fresh[position - 1][0].Discipline ?? string.Empty;
							foreach (DrawingIndexItem row in fresh[position])
							{
								row.Discipline = above;
								row.DisciplineUncertain = false;
							}
						}
					});
				};
				Grid.SetColumn(name, 1);
				Grid.SetColumn(info, 2);
				Grid.SetColumn(split, 3);
				Grid.SetColumn(merge, 4);
				line.Children.Add(name);
				line.Children.Add(info);
				line.Children.Add(split);
				line.Children.Add(merge);
				HeadingRow row2 = new HeadingRow(line);
				shown[group[0]] = row2;
				groupsPanel.Children.Add(row2.Host);
			}
			// Room below the last row, so that when rows at the end fold away the rows above stay where they are instead of the list
			// shrinking under the scroll position and sliding everything down.
			groupsPanel.Children.Add(new Border { Height = 360.0 });
		}

		// Runs a change to the headings with the rows moving visibly: rows that the change absorbs fold away (and the row that takes them in
		// glows) before the list is redrawn; rows that the change adds unfold.
		void ApplyChange(Action change)
		{
			if (animating)
			{
				return;
			}
			List<List<DrawingIndexItem>> before = GroupByDiscipline(rows);
			Dictionary<DrawingIndexItem, HeadingRow> beforeRows = new Dictionary<DrawingIndexItem, HeadingRow>(shown);
			change();
			List<List<DrawingIndexItem>> after = GroupByDiscipline(rows);
			HashSet<DrawingIndexItem> afterFirsts = after.Select(g => g[0]).ToHashSet();
			List<DrawingIndexItem> vanishing = before.Select(g => g[0]).Where(first => !afterFirsts.Contains(first) && beforeRows.ContainsKey(first)).ToList();

			void Finish()
			{
				RebuildGroups();
				double width = groupsPanel.ActualWidth;
				foreach (List<DrawingIndexItem> group in after)
				{
					if (!beforeRows.ContainsKey(group[0]) && shown.TryGetValue(group[0], out HeadingRow? added))
					{
						Expand(added.Host, width);
					}
				}
				animating = false;
			}

			if (vanishing.Count == 0)
			{
				animating = true;
				Finish();
				return;
			}
			animating = true;
			foreach (List<DrawingIndexItem> group in after)
			{
				if (before.Any(old => old.Any(row => group.Contains(row)) && !afterFirsts.Contains(old[0])) && beforeRows.TryGetValue(group[0], out HeadingRow? absorber))
				{
					Flash(absorber.Fill);
				}
			}
			int remaining = vanishing.Count;
			foreach (DrawingIndexItem first in vanishing)
			{
				Collapse(beforeRows[first].Host, () =>
				{
					if (--remaining == 0)
					{
						Finish();
					}
				});
			}
		}

		undo.Click += (object s, RoutedEventArgs a) => ApplyChange(() =>
		{
			if (history.Count > 0)
			{
				Restore(history.Pop());
			}
		});
		reset.Click += (object s, RoutedEventArgs a) => ApplyChange(() =>
		{
			history.Clear();
			Restore(original);
		});
		RebuildGroups();
		StackPanel top = new StackPanel { Spacing = 10.0 };
		top.Children.Add(new TextBlock
		{
			Text = "Each change of discipline starts a new heading in the index. If a heading comes from a misread sheet, remove it (its sheets join the heading above) or rename it; if sheets sit under a heading they do not belong to, use Split to start a new heading part way through. The headings, columns and pages after a change re-flow when the index is made. Point at the eye icon to see the title blocks as they were drawn for reading. Orange marks a heading with one or two sheets, or one the program was unsure of.",
			TextWrapping = TextWrapping.Wrap
		});
		StackPanel buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8.0 };
		buttons.Children.Add(undo);
		buttons.Children.Add(reset);
		top.Children.Add(buttons);
		return await ShowPopupWindowAsync("Review discipline headings", top, groupsPanel, "Create the index", "Cancel", 1150.0, 800.0, windowCreated: w => reviewWindow = w);
	}
	// The Split window: every sheet after the first with a picture of its title block, so the sheet where the new heading starts can be
	// picked by looking at it, and the name of the new heading. It can be moved and resized, and closes when another window is clicked.
	private Task ShowSplitWindowAsync(List<DrawingIndexItem> group, List<DrawingIndexItem> rows, ReviewPreviews previews, Window? owner, Action snapshot, Action rebuild)
	{
		Window? splitWindow = null;
		ListView list = new ListView { SelectionMode = ListViewSelectionMode.Single, HorizontalAlignment = HorizontalAlignment.Stretch };
		Style stretch = new Style(typeof(ListViewItem));
		stretch.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch));
		stretch.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(8, 0, 8, 0)));
		list.ItemContainerStyle = stretch;
		foreach (DrawingIndexItem member in group.Skip(1))
		{
			Grid item = new Grid { ColumnSpacing = 14.0, Padding = new Thickness(0, 10, 0, 10) };
			item.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(230.0) });
			item.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.0, GridUnitType.Star) });
			StackPanel text = new StackPanel { Spacing = 2.0, VerticalAlignment = VerticalAlignment.Center };
			text.Children.Add(new TextBlock { Text = member.DrawingNumber, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
			text.Children.Add(new TextBlock { Text = member.Title, FontSize = 12.0, Opacity = 0.75, TextWrapping = TextWrapping.Wrap });
			Grid picture = PreviewImage(previews, member, 560.0);
			picture.VerticalAlignment = VerticalAlignment.Center;
			Grid.SetColumn(picture, 1);
			item.Children.Add(text);
			item.Children.Add(picture);
			Border row = new Border
			{
				Child = item,
				BorderThickness = new Thickness(0, 0, 0, 1),
				BorderBrush = (Brush)Application.Current.Resources["ControlStrokeColorDefaultBrush"]
			};
			list.Items.Add(new ListViewItem { Content = row, Tag = member });
		}
		ComboBox newName = new ComboBox { Header = "Heading for the sheets from the selected one to the end of this group", HorizontalAlignment = HorizontalAlignment.Stretch, IsEditable = true };
		foreach (string known in rows.Select(r => (r.Discipline ?? string.Empty).Trim()).Where(d => d.Length > 0).Distinct().OrderBy(d => d))
		{
			newName.Items.Add(known);
		}
		TextBlock summary = new TextBlock { TextWrapping = TextWrapping.Wrap, Opacity = 0.8 };
		Button apply = new Button { Content = "Split the group", IsEnabled = false, Style = (Style)Application.Current.Resources["AccentButtonStyle"] };

		void Refresh()
		{
			string typed = (newName.Text ?? string.Empty).Trim();
			if (list.SelectedItem is ListViewItem { Tag: DrawingIndexItem start })
			{
				int from = group.IndexOf(start);
				summary.Text = $"{group.Count - from} sheet(s), {start.DrawingNumber} to {group[^1].DrawingNumber}, will go under \"{(typed.Length > 0 ? typed : "(type a heading)")}\"; the first {from} sheet(s) keep \"{group[0].Discipline}\".";
			}
			else
			{
				summary.Text = "Select the first sheet that belongs under the new heading.";
			}
			apply.IsEnabled = list.SelectedItem != null && typed.Length > 0;
		}

		list.SelectionChanged += (object s, SelectionChangedEventArgs a) => Refresh();
		newName.SelectionChanged += (object s, SelectionChangedEventArgs a) => Refresh();
		newName.LostFocus += (object s, RoutedEventArgs a) => Refresh();
		newName.KeyUp += (object s, Microsoft.UI.Xaml.Input.KeyRoutedEventArgs a) => Refresh();
		apply.Click += (object s, RoutedEventArgs a) =>
		{
			if (list.SelectedItem is ListViewItem { Tag: DrawingIndexItem start } && (newName.Text ?? string.Empty).Trim() is { Length: > 0 } typed)
			{
				snapshot();
				for (int k = group.IndexOf(start); k < group.Count; k++)
				{
					group[k].Discipline = typed;
					group[k].DisciplineUncertain = false;
				}
				splitWindow?.Close();
				rebuild();
			}
		};
		Refresh();
		StackPanel top = new StackPanel { Spacing = 6.0 };
		top.Children.Add(new TextBlock { Text = "Where does the new heading start?", FontSize = 18.0, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
		top.Children.Add(new TextBlock { Text = "Select the first sheet that belongs under the new heading. The pictures show each sheet's discipline and title as drawn.", TextWrapping = TextWrapping.Wrap });
		StackPanel body = new StackPanel { Spacing = 10.0 };
		body.Children.Add(list);
		body.Children.Add(newName);
		body.Children.Add(summary);
		body.Children.Add(apply);
		return ShowPopupWindowAsync("Split heading", top, body, null, "Cancel", 940.0, 720.0, owner: owner, closeWhenDeactivated: true, modal: false, windowCreated: w => splitWindow = w);
	}

	// An eye icon that shows the title blocks of a few of the group's sheets the moment the pointer reaches it. It is a Popup placed by hand,
	// not a tooltip (those wait half a second) and not a Flyout (which closes by itself as the pointer moves between target and flyout).
	// The pictures come from the background cache and are already in the popup when it opens.
	private static UIElement BuildPreviewIcon(List<DrawingIndexItem> group, ReviewPreviews previews)
	{
		FontIcon icon = new FontIcon { Glyph = "\uE7B3", FontSize = 20.0, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center };
		List<DrawingIndexItem> samples = SamplesOf(group);
		StackPanel pictures = new StackPanel { Spacing = 6.0, Width = 780.0 };
		foreach (DrawingIndexItem row in samples)
		{
			pictures.Children.Add(new TextBlock { Text = $"{row.DrawingNumber}{(row.DisciplineUncertain ? " (unsure)" : string.Empty)}: read as \"{row.Discipline}\"", FontSize = 12.0 });
			pictures.Children.Add(PreviewImage(previews, row, 760.0));
		}
		Border card = new Border
		{
			Child = pictures,
			Padding = new Thickness(12),
			CornerRadius = new CornerRadius(6),
			BorderThickness = new Thickness(1),
			BorderBrush = new SolidColorBrush(Windows.UI.Color.FromArgb(160, 128, 128, 128)),
			Background = (Brush)Application.Current.Resources["ApplicationPageBackgroundThemeBrush"]
		};
		Microsoft.UI.Xaml.Controls.Primitives.Popup popup = new Microsoft.UI.Xaml.Controls.Primitives.Popup { Child = card, IsLightDismissEnabled = false };
		Border target = new Border { Background = new SolidColorBrush(Windows.UI.Color.FromArgb(1, 0, 0, 0)), Child = icon, Width = 36.0, Height = 32.0 };
		target.PointerEntered += (object s, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs a) =>
		{
			if (target.XamlRoot == null)
			{
				return;
			}
			popup.XamlRoot = target.XamlRoot;
			card.Measure(new Windows.Foundation.Size(double.PositiveInfinity, double.PositiveInfinity));
			// The pictures arrive a moment after the frame, so leave room for them when keeping the card inside the window.
			double expected = Math.Max(card.DesiredSize.Height, samples.Count * 120.0 + 24.0);
			Windows.Foundation.Point at = target.TransformToVisual(null).TransformPoint(new Windows.Foundation.Point(target.ActualWidth + 4.0, 0.0));
			double rootWidth = target.XamlRoot.Size.Width;
			double rootHeight = target.XamlRoot.Size.Height;
			popup.HorizontalOffset = Math.Max(0.0, Math.Min(at.X, rootWidth - card.DesiredSize.Width - 8.0));
			popup.VerticalOffset = Math.Max(0.0, Math.Min(at.Y - 8.0, rootHeight - expected - 8.0));
			popup.IsOpen = true;
		};
		target.PointerExited += (object s, Microsoft.UI.Xaml.Input.PointerRoutedEventArgs a) => popup.IsOpen = false;
		target.Unloaded += (object s, RoutedEventArgs a) => popup.IsOpen = false;
		return target;
	}
}