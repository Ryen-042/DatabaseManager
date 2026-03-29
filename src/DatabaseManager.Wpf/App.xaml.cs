using System.Windows;

namespace DatabaseManager.Wpf;

public partial class App : Application
{
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

