using System.Windows;
using System.Windows.Threading;

namespace DatabaseManager.Wpf;

public partial class App : Application
{
	public App()
	{
		DispatcherUnhandledException += App_DispatcherUnhandledException;
	}

	/// <summary>
	/// Without this, any unhandled exception on the UI thread (e.g. the CommandPaletteWindow
	/// double-Close reentrancy bug) tears down the whole app with no warning beyond a brief
	/// freeze. Show what happened and keep the app running instead - losing unsaved work to a
	/// hard crash is worse than a warning dialog for a bug we didn't anticipate.
	/// </summary>
	private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
	{
		MessageBox.Show(
			$"An unexpected error occurred:\n\n{e.Exception.Message}\n\nThe application will keep running, but please save your work and consider restarting if you see repeated errors.",
			"Unexpected Error",
			MessageBoxButton.OK,
			MessageBoxImage.Warning);
		e.Handled = true;
	}

	public static void ApplyTheme(bool darkMode)
	{
		var app = Current;
		if (app is null)
		{
			return;
		}

		var themeUri = new Uri(
			darkMode ? "Themes/DarkTheme.xaml" : "Themes/LightTheme.xaml",
			UriKind.Relative);

		if (app.Resources.MergedDictionaries.Count == 0)
		{
			app.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = themeUri });
			return;
		}

		app.Resources.MergedDictionaries[0] = new ResourceDictionary { Source = themeUri };
	}
}

