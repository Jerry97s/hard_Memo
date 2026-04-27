using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using HdLabs.Memo.ViewModels;

namespace HdLabs.Memo;

public partial class ReminderSetupWindow : Window
{
    private readonly MainViewModel _vm;
    private bool _syncUi;

    public ReminderSetupWindow(MainViewModel vm)
    {
        InitializeComponent();
        _vm = vm;
        Loaded += OnLoaded;
        Closing += (_, _) => ApplyDisableOnCloseIfNeeded();
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        Loaded -= OnLoaded;
        Dispatcher.BeginInvoke(
            new Action(InitializeAll),
            DispatcherPriority.Loaded);
    }

    private void InitializeAll()
    {
        _syncUi = true;
        try
        {
            ApplyFromEditor();
            RefreshHistoryList();
        }
        finally
        {
            _syncUi = false;
        }
    }

    private void ApplyFromEditor()
    {
        if (_vm.EditorReminderAt is { } r)
        {
            var local = r.LocalDateTime;
            DatePick.SelectedDate = local.Date;
            DatePick.DisplayDate = local;
            TimeText.Text = local.ToString("HH:mm", CultureInfo.InvariantCulture);
            UseReminderCheck.IsChecked = true;
        }
        else
        {
            var t = DateTime.Now.AddMinutes(30);
            DatePick.SelectedDate = t.Date;
            DatePick.DisplayDate = t;
            TimeText.Text = t.ToString("HH:mm", CultureInfo.InvariantCulture);
            UseReminderCheck.IsChecked = false;
        }

        SyncTimeInputsEnabled();
    }

    private void SyncTimeInputsEnabled()
    {
        var on = UseReminderCheck.IsChecked == true;
        DatePick.IsEnabled = on;
        TimeText.IsEnabled = on;
    }

    private void UseReminderCheck_OnChanged(object sender, RoutedEventArgs e)
    {
        if (_syncUi)
            return;

        // 사용자가 "끄기"로 바꿨으면 즉시 꺼지게(창을 X로 닫아도 울리지 않도록)
        if (UseReminderCheck.IsChecked != true)
        {
            if (_vm.EditorReminderAt is not null)
            {
                _vm.EditorReminderAt = null;
                _vm.AddReminderDisabledHistory();
            }
        }

        SyncTimeInputsEnabled();
    }

    private void ApplyDisableOnCloseIfNeeded()
    {
        // 체크를 꺼둔 채로 닫으면, 알림이 남아있지 않게 보정
        if (UseReminderCheck.IsChecked == true)
            return;
        if (_vm.EditorReminderAt is null)
            return;
        _vm.EditorReminderAt = null;
        _vm.AddReminderDisabledHistory();
    }

    private void RefreshHistoryList()
    {
        var rows = _vm
            .GetReminderHistory()
            .Select(h => new ReminderHistoryItem(h.Label, h.Enabled, h.TargetAtLocal))
            .ToList();
        HistoryList.ItemsSource = rows;
    }

    private void HistoryList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_syncUi)
            return;
        if (HistoryList.SelectedItem is not ReminderHistoryItem row)
            return;

        _syncUi = true;
        try
        {
            UseReminderCheck.IsChecked = row.Enabled;
            if (row.Enabled && row.TargetAtLocal is { } t)
            {
                var local = t.LocalDateTime;
                DatePick.SelectedDate = local.Date;
                DatePick.DisplayDate = local;
                TimeText.Text = local.ToString("HH:mm", CultureInfo.InvariantCulture);
            }
            SyncTimeInputsEnabled();
        }
        finally
        {
            _syncUi = false;
        }
    }

    private void Clear_Click(object sender, RoutedEventArgs e)
    {
        _vm.EditorReminderAt = null;
        _vm.AddReminderDisabledHistory();
        DialogResult = true;
        Close();
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        if (UseReminderCheck.IsChecked != true)
        {
            _vm.EditorReminderAt = null;
            _vm.AddReminderDisabledHistory();
            DialogResult = true;
            Close();
            return;
        }

        if (DatePick.SelectedDate is not { } datePart)
        {
            MessageBox.Show("날짜를 선택하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var timeRaw = TimeText.Text.Trim();
        var parts = timeRaw.Split(':', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2
            || !int.TryParse(parts[0], NumberStyles.None, CultureInfo.CurrentCulture, out var h)
            || !int.TryParse(parts[1], NumberStyles.None, CultureInfo.CurrentCulture, out var min))
        {
            MessageBox.Show("시간은 HH:mm 형식(예: 14:30)으로 입력하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (h is < 0 or > 23 || min is < 0 or > 59)
        {
            MessageBox.Show("시간은 00:00~23:59 사이로 입력하세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var local = datePart.Date + new TimeSpan(h, min, 0);
        var dto = new DateTimeOffset(local, TimeZoneInfo.Local.GetUtcOffset(local));
        _vm.EditorReminderAt = dto;
        _vm.AddReminderEnabledHistory(dto);
        DialogResult = true;
        Close();
    }

    private sealed class ReminderHistoryItem
    {
        public ReminderHistoryItem(string label, bool enabled, DateTimeOffset? targetAtLocal)
        {
            Label = label;
            Enabled = enabled;
            TargetAtLocal = targetAtLocal;
        }

        public string Label { get; }
        public bool Enabled { get; }
        public DateTimeOffset? TargetAtLocal { get; }
    }
}
