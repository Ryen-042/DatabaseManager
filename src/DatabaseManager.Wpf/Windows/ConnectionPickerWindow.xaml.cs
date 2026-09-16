using System.Collections.ObjectModel;
using System.Windows;
using DatabaseManager.Core.Services;
using DatabaseManager.Wpf.ViewModels;

namespace DatabaseManager.Wpf.Windows;

public partial class ConnectionPickerWindow : Window
{
    private readonly IConnectionProfileStoreService _store;
    private readonly ObservableCollection<ConnectionProfileViewModel> _profiles = new();
    private Guid? _editingProfileId;

    public ConnectionPickerWindow(IConnectionProfileStoreService store, string currentConnectionString)
    {
        InitializeComponent();
        _store = store;
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
        else
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

        DialogResult = true;
    }
}
