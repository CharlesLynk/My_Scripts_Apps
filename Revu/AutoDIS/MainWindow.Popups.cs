using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Windows.Graphics;

namespace MbrohDrawingIndexCreator;

// Pop-ups that the user can move and resize (a ContentDialog can do neither): each is a window of its own, owned by the window it
// belongs to, and a modal one disables its owner until it closes.
public sealed partial class MainWindow
{
	private const int OwnerWindowIndex = -8;

	[DllImport("user32.dll")]
	private static extern bool EnableWindow(IntPtr window, bool enable);

	[DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
	private static extern IntPtr SetWindowLongPtr(IntPtr window, int index, IntPtr value);

	private static IntPtr HandleOf(Window window) => WinRT.Interop.WindowNative.GetWindowHandle(window);

	// Shows `body` in a window titled `title` with an optional fixed `top` part above it (always visible) and a Close button plus an optional
	// primary button below. Completes with true when the primary button was pressed (onPrimary runs first), false otherwise.
	// closeWhenDeactivated: the window closes as soon as another window is clicked (a light pop-up).
	private Task<bool> ShowPopupWindowAsync(string title, UIElement? top, UIElement body, string? primaryText, string closeText, double width, double height, Window? owner = null, Action? onPrimary = null, bool closeWhenDeactivated = false, bool modal = true, Action<Window>? windowCreated = null)
	{
		Window ownerWindow = owner ?? this;
		TaskCompletionSource<bool> finished = new TaskCompletionSource<bool>();
		Window popup = new Window { Title = title };
		windowCreated?.Invoke(popup);
		bool primaryPressed = false;

		Grid root = new Grid
		{
			Padding = new Thickness(16),
			RowSpacing = 10.0,
			Background = (Brush)Application.Current.Resources["ApplicationPageBackgroundThemeBrush"]
		};
		root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
		root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1.0, GridUnitType.Star) });
		root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
		if (top != null)
		{
			Grid.SetRow((FrameworkElement)top, 0);
			root.Children.Add(top);
		}
		ScrollViewer scroller = new ScrollViewer { Content = body, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled, Padding = new Thickness(0, 0, 16, 0) };
		Grid.SetRow(scroller, 1);
		root.Children.Add(scroller);
		StackPanel buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8.0, HorizontalAlignment = HorizontalAlignment.Right };
		Button close = new Button { Content = closeText };
		close.Click += (object s, RoutedEventArgs a) => popup.Close();
		buttons.Children.Add(close);
		if (primaryText != null)
		{
			Button primary = new Button { Content = primaryText, Style = (Style)Application.Current.Resources["AccentButtonStyle"] };
			primary.Click += (object s, RoutedEventArgs a) =>
			{
				onPrimary?.Invoke();
				primaryPressed = true;
				popup.Close();
			};
			buttons.Children.Add(primary);
		}
		Grid.SetRow(buttons, 2);
		root.Children.Add(buttons);
		root.KeyDown += (object s, KeyRoutedEventArgs a) =>
		{
			if (a.Key == Windows.System.VirtualKey.Escape)
			{
				popup.Close();
			}
		};
		popup.Content = root;
		try
		{
			string icon = Path.Combine(GetPackageContentDirectory(), "Assets", "AppIcon.ico");
			if (File.Exists(icon))
			{
				popup.AppWindow.SetIcon(icon);
			}
		}
		catch (Exception)
		{
		}
		double scale = ownerWindow.Content?.XamlRoot?.RasterizationScale ?? 1.0;
		SizeInt32 size = new SizeInt32((int)(width * scale), (int)(height * scale));
		popup.AppWindow.Resize(size);
		try
		{
			PointInt32 ownerPosition = ownerWindow.AppWindow.Position;
			SizeInt32 ownerSize = ownerWindow.AppWindow.Size;
			popup.AppWindow.Move(new PointInt32(Math.Max(0, ownerPosition.X + (ownerSize.Width - size.Width) / 2), Math.Max(0, ownerPosition.Y + (ownerSize.Height - size.Height) / 2)));
		}
		catch (Exception)
		{
		}
		if (popup.AppWindow.Presenter is OverlappedPresenter presenter)
		{
			presenter.IsMinimizable = false;
		}
		IntPtr popupHandle = HandleOf(popup);
		IntPtr ownerHandle = HandleOf(ownerWindow);
		SetWindowLongPtr(popupHandle, OwnerWindowIndex, ownerHandle);
		if (modal)
		{
			EnableWindow(ownerHandle, false);
		}
		bool activated = false;
		if (closeWhenDeactivated)
		{
			popup.Activated += (object s, WindowActivatedEventArgs a) =>
			{
				if (a.WindowActivationState == WindowActivationState.Deactivated)
				{
					if (activated)
					{
						popup.Close();
					}
				}
				else
				{
					activated = true;
				}
			};
		}
		popup.Closed += (object s, WindowEventArgs a) =>
		{
			if (modal)
			{
				EnableWindow(ownerHandle, true);
			}
			ownerWindow.Activate();
			finished.TrySetResult(primaryPressed);
		};
		popup.Activate();
		return finished.Task;
	}
}
