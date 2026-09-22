using System.Collections.ObjectModel;
using System.Windows;
using DatabaseManager.Core.Services;
using DatabaseManager.Core.Services.Schema;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Wpf.Windows;

public partial class ConnectionPickerWindow : Window
{
    private static readonly HashSet<string> SystemDatabaseNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "master",
        "model",
        "msdb",
        "tempdb"
    };

    private readonly IConnectionProfileStoreService _store;
    private readonly IDatabaseSchemaService _databaseSchemaService;
    private readonly ObservableCollection<ConnectionProfileViewModel> _profiles = new();
    private Guid? _editingProfileId;
    private List<string> _allDatabases = new();
    private string? _selectedDatabaseName;
    private bool _isUpdatingDatabaseSelection;

    public ConnectionPickerWindow(IConnectionProfileStoreService store, IDatabaseSchemaService databaseSchemaService, string currentConnectionString)
    {
        InitializeComponent();
        _store = store;
        _databaseSchemaService = databaseSchemaService;
        ProfilesListBox.ItemsSource = _profiles;
        RawConnectionStringTextBox.Text = currentConnectionString;

        // Default to the Raw tab if there's already a connection string in play and no
        // saved profiles exist yet - avoids landing on an empty "Saved Connections" list
        // when the user hasn't created any profiles.
        Loaded += async (_, _) => await LoadProfilesAsync();
    }

    /// <summary>Set once the dialog is confirmed (DialogResult == true); null connection string means canceled.</summary>
    public string? SelectedConnectionString { get; private set; }

    /// <summary>Null when the user connected via the Raw tab instead of a saved profile.</summary>
    public string? SelectedProfileName { get; private set; }

    /// <summary>Null when the user connected via the Raw tab. Used by the caller to record last-used time only once the connection actually succeeds.</summary>
    public Guid? SelectedProfileId { get; private set; }

    private async Task LoadProfilesAsync()
    {
        var previouslySelectedId = (ProfilesListBox.SelectedItem as ConnectionProfileViewModel)?.Id;

        _profiles.Clear();
        var all = await _store.GetAllAsync(CancellationToken.None);
        foreach (var profile in all.OrderByDescending(p => p.IsDefault).ThenBy(p => p.Name, StringComparer.OrdinalIgnoreCase))
        {
            _profiles.Add(new ConnectionProfileViewModel(profile));
        }

        ProfilesListBox.SelectedItem = previouslySelectedId.HasValue
            ? _profiles.FirstOrDefault(p => p.Id == previouslySelectedId.Value)
            : null;

        ProfilesListBox.SelectedItem ??= _profiles.FirstOrDefault(p => p.IsDefault) ?? _profiles.FirstOrDefault();

        if (_profiles.Count == 0)
        {
            PickerTabControl.SelectedIndex = 1;
        }

        UpdateButtonStates();
    }

    private void ProfilesListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        UpdateButtonStates();
    }

    private void UpdateButtonStates()
    {
        var hasSelection = ProfilesListBox.SelectedItem is not null;
        EditButton.IsEnabled = hasSelection;
        DeleteButton.IsEnabled = hasSelection;
        SetDefaultButton.IsEnabled = hasSelection && ProfilesListBox.SelectedItem is ConnectionProfileViewModel { IsDefault: false };
    }

    private void ProfilesListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (ProfilesListBox.SelectedItem is ConnectionProfileViewModel)
        {
            ConnectButton_Click(sender, new RoutedEventArgs());
        }
    }

    private void NewButton_Click(object sender, RoutedEventArgs e) => ShowEditForm(null);

    private void EditButton_Click(object sender, RoutedEventArgs e)
    {
        if (ProfilesListBox.SelectedItem is ConnectionProfileViewModel vm)
        {
            ShowEditForm(vm);
        }
    }

    private void ShowEditForm(ConnectionProfileViewModel? existing)
    {
        _editingProfileId = existing?.Id;
        EditNameTextBox.Text = existing?.Name ?? string.Empty;
        EditConnectionStringTextBox.Text = existing is not null ? _store.Decrypt(existing.Profile) : string.Empty;
        EditDefaultCheckBox.IsChecked = existing?.IsDefault ?? false;

        ListPanel.Visibility = Visibility.Collapsed;
        EditPanel.Visibility = Visibility.Visible;
        EditNameTextBox.Focus();
    }

    private async void SaveProfileButton_Click(object sender, RoutedEventArgs e)
    {
        var name = EditNameTextBox.Text.Trim();
        var connectionString = EditConnectionStringTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(connectionString))
        {
            MessageBox.Show(
                this,
                "Both a name and a connection string are required.",
                "Missing Information",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        await _store.SaveAsync(_editingProfileId, name, connectionString, EditDefaultCheckBox.IsChecked == true, CancellationToken.None);

        EditPanel.Visibility = Visibility.Collapsed;
        ListPanel.Visibility = Visibility.Visible;
        await LoadProfilesAsync();
    }

    private void CancelEditButton_Click(object sender, RoutedEventArgs e)
    {
        EditPanel.Visibility = Visibility.Collapsed;
        ListPanel.Visibility = Visibility.Visible;
    }

    private async void SetDefaultButton_Click(object sender, RoutedEventArgs e)
    {
        if (ProfilesListBox.SelectedItem is not ConnectionProfileViewModel vm)
        {
            return;
        }

        await _store.SaveAsync(vm.Id, vm.Name, _store.Decrypt(vm.Profile), isDefault: true, CancellationToken.None);
        await LoadProfilesAsync();
    }

    private async void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (ProfilesListBox.SelectedItem is not ConnectionProfileViewModel vm)
        {
            return;
        }

        var result = MessageBox.Show(
            this,
            $"Delete the saved connection \"{vm.Name}\"? This cannot be undone.",
            "Delete Connection",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        await _store.DeleteAsync(vm.Id, CancellationToken.None);
        await LoadProfilesAsync();
    }

    private void ConnectButton_Click(object sender, RoutedEventArgs e)
    {
        if (PickerTabControl.SelectedIndex == 0)
        {
            if (ProfilesListBox.SelectedItem is not ConnectionProfileViewModel vm)
            {
                MessageBox.Show(this, "Select a saved connection first.", "No Connection Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedConnectionString = _store.Decrypt(vm.Profile);
            SelectedProfileName = vm.Name;
            SelectedProfileId = vm.Id;
        }
        else if (PickerTabControl.SelectedIndex == 1)
        {
            var text = RawConnectionStringTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                MessageBox.Show(this, "Enter a connection string first.", "Missing Connection String", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedConnectionString = text;
            SelectedProfileName = null;
            SelectedProfileId = null;
        }
        else
        {
            if (!TryGetBaseConnectionString(out var baseConnectionString))
            {
                return;
            }

            SelectedConnectionString = string.IsNullOrWhiteSpace(_selectedDatabaseName)
                ? baseConnectionString
                : ConnectionStringHelper.WithDatabase(baseConnectionString, _selectedDatabaseName);

            var selectedProfile = ProfilesListBox.SelectedItem as ConnectionProfileViewModel;
            SelectedProfileName = selectedProfile?.Name;
            SelectedProfileId = selectedProfile?.Id;
        }

        DialogResult = true;
    }

    /// <summary>
    /// The connection string to browse/connect with from the Databases tab: whichever of the
    /// Saved Connections selection or the Raw text currently has a value, regardless of which
    /// tab happens to be active (the user picks one there, then switches to Databases).
    /// </summary>
    private bool TryGetBaseConnectionString(out string connectionString)
    {
        if (ProfilesListBox.SelectedItem is ConnectionProfileViewModel vm)
        {
            connectionString = _store.Decrypt(vm.Profile);
            return true;
        }

        var text = RawConnectionStringTextBox.Text.Trim();
        if (!string.IsNullOrWhiteSpace(text))
        {
            connectionString = text;
            return true;
        }

        connectionString = string.Empty;
        MessageBox.Show(this, "Select a saved connection or enter a connection string first.", "No Connection Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
        return false;
    }

    private async void BrowseDatabasesButton_Click(object sender, RoutedEventArgs e)
    {
        if (!TryGetBaseConnectionString(out var baseConnectionString))
        {
            return;
        }

        var listConnectionString = ConnectionStringHelper.WithFallbackDatabase(baseConnectionString, "master");

        BrowseDatabasesButton.IsEnabled = false;
        try
        {
            var databases = await _databaseSchemaService.GetDatabasesAsync(listConnectionString, CancellationToken.None);
            _allDatabases = databases
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            ApplyDatabaseFilter();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Failed to list databases: {ex.Message}", "Browse Databases", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BrowseDatabasesButton.IsEnabled = true;
        }
    }

    private void DatabaseSearchTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) => ApplyDatabaseFilter();

    private void ShowSystemDatabasesCheckBox_Changed(object sender, RoutedEventArgs e) => ApplyDatabaseFilter();

    private void ApplyDatabaseFilter()
    {
        var showSystemDatabases = ShowSystemDatabasesCheckBox.IsChecked == true;
        var query = DatabaseSearchTextBox.Text.Trim();

        var filtered = _allDatabases
            .Where(name => showSystemDatabases
                || !SystemDatabaseNames.Contains(name)
                || string.Equals(name, _selectedDatabaseName, StringComparison.OrdinalIgnoreCase))
            .Where(name => string.IsNullOrWhiteSpace(query) || name.Contains(query, StringComparison.OrdinalIgnoreCase))
            .ToList();

        DatabasesListBox.ItemsSource = filtered;

        if (!string.IsNullOrWhiteSpace(_selectedDatabaseName))
        {
            _isUpdatingDatabaseSelection = true;
            DatabasesListBox.SelectedItem = filtered.FirstOrDefault(x => string.Equals(x, _selectedDatabaseName, StringComparison.OrdinalIgnoreCase));
            _isUpdatingDatabaseSelection = false;
        }
    }

    private void DatabasesListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (_isUpdatingDatabaseSelection)
        {
            return;
        }

        _selectedDatabaseName = DatabasesListBox.SelectedItem as string;
    }
}
