using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Threading;
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

    private const string NoConnectionStatusText = "Select a saved connection or enter a connection string above to browse its databases.";

    private readonly IConnectionProfileStoreService _store;
    private readonly IDatabaseSchemaService _databaseSchemaService;
    private readonly ObservableCollection<ConnectionProfileViewModel> _profiles = new();
    private readonly DispatcherTimer _databaseFetchDebounceTimer = new() { Interval = TimeSpan.FromMilliseconds(500) };
    private Guid? _editingProfileId;
    private List<string> _allDatabases = new();
    private string? _selectedDatabaseName;
    private bool _isUpdatingDatabaseSelection;
    private bool _suppressDatabaseAutoFetchOnSelectionChange;
    private CancellationTokenSource? _databasesFetchCts;

    public ConnectionPickerWindow(IConnectionProfileStoreService store, IDatabaseSchemaService databaseSchemaService, string currentConnectionString)
    {
        _databaseFetchDebounceTimer.Tick += DatabaseFetchDebounceTimer_Tick;

        InitializeComponent();
        _store = store;
        _databaseSchemaService = databaseSchemaService;
        ProfilesListBox.ItemsSource = _profiles;
        RawConnectionStringTextBox.Text = currentConnectionString;

        Loaded += async (_, _) => await LoadProfilesAsync();
    }

    /// <summary>Set once the dialog is confirmed (DialogResult == true); null connection string means canceled.</summary>
    public string? SelectedConnectionString { get; private set; }

    /// <summary>Null when the user connected via the raw text box instead of a saved profile.</summary>
    public string? SelectedProfileName { get; private set; }

    /// <summary>Null when the user connected via the raw text box. Used by the caller to record last-used time only once the connection actually succeeds.</summary>
    public Guid? SelectedProfileId { get; private set; }

    private async Task LoadProfilesAsync()
    {
        var previouslySelectedId = (ProfilesListBox.SelectedItem as ConnectionProfileViewModel)?.Id;

        _profiles.Clear();
        var all = await _store.GetAllAsync(CancellationToken.None);
        foreach (var profile in all.OrderByDescending(p => p.IsDefault).ThenBy(p => p.Name, StringComparer.OrdinalIgnoreCase))
        {
            string preview;
            try
            {
                preview = _store.Decrypt(profile);
            }
            catch (Exception)
            {
                preview = "(unable to decrypt connection string)";
            }

            _profiles.Add(new ConnectionProfileViewModel(profile, preview));
        }

        ProfilesListBox.SelectedItem = previouslySelectedId.HasValue
            ? _profiles.FirstOrDefault(p => p.Id == previouslySelectedId.Value)
            : null;

        ProfilesListBox.SelectedItem ??= _profiles.FirstOrDefault(p => p.IsDefault) ?? _profiles.FirstOrDefault();

        UpdateButtonStates();
    }

    private void ProfilesListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        UpdateButtonStates();

        if (_suppressDatabaseAutoFetchOnSelectionChange)
        {
            return;
        }

        _databaseFetchDebounceTimer.Stop();
        _ = RefreshDatabasesAsync();
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
        EditConnectionStringTextBox.Text = existing?.ConnectionStringPreview ?? string.Empty;
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

        await _store.SaveAsync(vm.Id, vm.Name, vm.ConnectionStringPreview, isDefault: true, CancellationToken.None);
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
        if (!TryGetBaseConnectionString(out var baseConnectionString, showWarningIfMissing: true))
        {
            return;
        }

        SelectedConnectionString = string.IsNullOrWhiteSpace(_selectedDatabaseName)
            ? baseConnectionString
            : ConnectionStringHelper.WithDatabase(baseConnectionString, _selectedDatabaseName);

        var selectedProfile = ProfilesListBox.SelectedItem as ConnectionProfileViewModel;
        SelectedProfileName = selectedProfile?.Name;
        SelectedProfileId = selectedProfile?.Id;

        DialogResult = true;
    }

    /// <summary>
    /// The connection string to connect/browse-databases with: a selected saved profile takes
    /// priority over the raw text box (they're mutually exclusive in practice - typing into the
    /// raw box clears the saved selection, see RawConnectionStringTextBox_TextChanged), matching
    /// the whole-window "one active source at a time" model now that both are always visible
    /// instead of living on separate tabs. showWarningIfMissing is false for the auto-fetch path
    /// (typing/browsing shouldn't pop a modal), true for the explicit Connect click.
    /// </summary>
    private bool TryGetBaseConnectionString(out string connectionString, bool showWarningIfMissing)
    {
        if (ProfilesListBox.SelectedItem is ConnectionProfileViewModel vm)
        {
            connectionString = vm.ConnectionStringPreview;
            return true;
        }

        var text = RawConnectionStringTextBox.Text.Trim();
        if (!string.IsNullOrWhiteSpace(text))
        {
            connectionString = text;
            return true;
        }

        connectionString = string.Empty;
        if (showWarningIfMissing)
        {
            MessageBox.Show(this, "Select a saved connection or enter a connection string first.", "No Connection Selected", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        return false;
    }

    /// <summary>
    /// Typing into the raw box means "use this instead of whatever's selected" - clearing the
    /// saved-profile selection keeps TryGetBaseConnectionString's precedence unambiguous. The
    /// clear itself would otherwise trigger ProfilesListBox_SelectionChanged's immediate fetch on
    /// every keystroke, defeating the debounce below, hence the suppress flag.
    /// </summary>
    private void RawConnectionStringTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (ProfilesListBox.SelectedItem is not null)
        {
            _suppressDatabaseAutoFetchOnSelectionChange = true;
            ProfilesListBox.SelectedItem = null;
            _suppressDatabaseAutoFetchOnSelectionChange = false;
        }

        _databaseFetchDebounceTimer.Stop();
        _databaseFetchDebounceTimer.Start();
    }

    private void DatabaseFetchDebounceTimer_Tick(object? sender, EventArgs e)
    {
        _databaseFetchDebounceTimer.Stop();
        _ = RefreshDatabasesAsync();
    }

    private void RefreshDatabasesButton_Click(object sender, RoutedEventArgs e)
    {
        _databaseFetchDebounceTimer.Stop();
        _ = RefreshDatabasesAsync();
    }

    /// <summary>
    /// Auto-fetches the database list for whichever connection is currently active - on load (if
    /// a default profile is pre-selected), whenever the saved-profile selection changes, and
    /// (debounced) whenever the raw text box changes. Cancels any still-in-flight fetch first so a
    /// slow stale request can't overwrite a newer, faster one's results. Failures show inline in
    /// DatabaseStatusTextBlock rather than a MessageBox, since this now fires from typing/browsing,
    /// not just an explicit button click.
    /// </summary>
    private async Task RefreshDatabasesAsync()
    {
        _databasesFetchCts?.Cancel();
        var cts = new CancellationTokenSource();
        _databasesFetchCts = cts;

        if (!TryGetBaseConnectionString(out var baseConnectionString, showWarningIfMissing: false))
        {
            _allDatabases = new List<string>();
            ApplyDatabaseFilter();
            DatabaseStatusTextBlock.Text = NoConnectionStatusText;
            return;
        }

        DatabaseStatusTextBlock.Text = "Loading databases...";

        var listConnectionString = ConnectionStringHelper.WithFallbackDatabase(baseConnectionString, "master");

        try
        {
            var databases = await _databaseSchemaService.GetDatabasesAsync(listConnectionString, cts.Token);
            if (cts.IsCancellationRequested)
            {
                return;
            }

            _allDatabases = databases
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            ApplyDatabaseFilter();
            DatabaseStatusTextBlock.Text = string.Empty;
        }
        catch (OperationCanceledException)
        {
            // Superseded by a newer fetch; leave whatever that one produces alone.
        }
        catch (Exception ex)
        {
            if (cts.IsCancellationRequested)
            {
                return;
            }

            _allDatabases = new List<string>();
            ApplyDatabaseFilter();
            DatabaseStatusTextBlock.Text = $"Could not load databases: {ex.Message}";
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
