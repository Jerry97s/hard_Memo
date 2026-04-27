using System.Collections.ObjectModel;
using System.Linq;
using System.Media;
using System.Windows;
using System.Windows.Threading;
using HdLabs.Common.Mvvm;
using HdLabs.Memo;
using HdLabs.Memo.Helpers;
using HdLabs.Memo.Models;
using HdLabs.Memo.Services;

namespace HdLabs.Memo.ViewModels;

public sealed class MainViewModel : ViewModelBase
{
    private readonly MemoDataService _data = new();
    private MemoDataRoot _root = new();
    private string _newTitle = "";
    private string _newTitleXaml = "";
    private string _newBody = "";
    private MemoItem? _selected;
    private string _currentView = "Editor";
    private Guid? _editingId;
    private bool _isDirty;
    private bool _suppressDirty;
    private readonly DispatcherTimer _autoSave = new() { Interval = TimeSpan.FromMilliseconds(900) };
    private readonly DispatcherTimer _reminderTimer = new() { Interval = TimeSpan.FromSeconds(1) };
    private DateTime? _editorScheduleDate;
    private DateTimeOffset? _editorReminderAt;

    public string Title => "HdLabs Memo";

    public ObservableCollection<MemoItem> Items { get; } = new();

    public string CurrentView
    {
        get => _currentView;
        set
        {
            if (!SetProperty(ref _currentView, value))
                return;
            if (string.Equals(_currentView, "List", StringComparison.OrdinalIgnoreCase))
                FlushEditorToItem();

            OnPropertyChanged(nameof(IsEditorView));
            OnPropertyChanged(nameof(IsListView));
        }
    }

    public bool IsEditorView => string.Equals(CurrentView, "Editor", StringComparison.OrdinalIgnoreCase);
    public bool IsListView => string.Equals(CurrentView, "List", StringComparison.OrdinalIgnoreCase);

    public string NewTitle
    {
        get => _newTitle;
        set
        {
            if (!SetProperty(ref _newTitle, value))
                return;
            if (!_suppressDirty)
                _isDirty = true;
            RestartAutoSave();
        }
    }

    public string NewTitleXaml
    {
        get => _newTitleXaml;
        set
        {
            if (!SetProperty(ref _newTitleXaml, value))
                return;
            if (!_suppressDirty)
                _isDirty = true;
            RestartAutoSave();
        }
    }

    public string NewBody
    {
        get => _newBody;
        set
        {
            if (!SetProperty(ref _newBody, value))
                return;
            if (!_suppressDirty)
                _isDirty = true;
            RestartAutoSave();
        }
    }

    public MemoItem? Selected
    {
        get => _selected;
        set
        {
            if (!SetProperty(ref _selected, value))
                return;
            OpenSelectedCommand.RaiseCanExecuteChanged();
            DeleteSelectedCommand.RaiseCanExecuteChanged();
        }
    }

    public bool WindowTopmost
    {
        get => _root.Settings.WindowTopmost;
        set
        {
            if (_root.Settings.WindowTopmost == value)
                return;
            _root.Settings.WindowTopmost = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsPinned));
            Persist();
        }
    }

    public bool IsPinned => WindowTopmost;

    public string? CardTintHex
    {
        get => _root.Settings.CardTintHex;
        set
        {
            if (_root.Settings.CardTintHex == value)
                return;
            _root.Settings.CardTintHex = value;
            OnPropertyChanged();
            Persist();
        }
    }

    public string BringToFrontHotkey
    {
        get => string.IsNullOrWhiteSpace(_root.Settings.BringToFrontHotkey) ? "Ctrl+Alt+M" : _root.Settings.BringToFrontHotkey.Trim();
        set
        {
            var v = string.IsNullOrWhiteSpace(value) ? "Ctrl+Alt+M" : value.Trim();
            if (string.Equals(_root.Settings.BringToFrontHotkey, v, StringComparison.OrdinalIgnoreCase))
                return;
            _root.Settings.BringToFrontHotkey = v;
            OnPropertyChanged();
            Persist();
        }
    }

    /// <summary>편집 중인 메모에 캘린더로 붙이는 날짜(시간은 무시).</summary>
    public DateTime? EditorScheduleDate
    {
        get => _editorScheduleDate;
        set
        {
            if (!SetProperty(ref _editorScheduleDate, value))
                return;
            if (!_suppressDirty)
            {
                _isDirty = true;
                RestartAutoSave();
            }
        }
    }

    /// <summary>편집 중인 메모의 미리 알림 시각.</summary>
    public DateTimeOffset? EditorReminderAt
    {
        get => _editorReminderAt;
        set
        {
            if (!SetProperty(ref _editorReminderAt, value))
                return;
            if (!_suppressDirty)
            {
                _isDirty = true;
                RestartAutoSave();
            }
        }
    }

    public RelayCommand NewCommand { get; }
    public RelayCommand SaveCommand { get; }
    public RelayCommand ShowEditorCommand { get; }
    public RelayCommand ShowListCommand { get; }
    public RelayCommand OpenSelectedCommand { get; }
    public RelayCommand DeleteSelectedCommand { get; }
    public RelayCommand TogglePinCommand { get; }

    public Guid? EditingId => _editingId;

    public MainViewModel()
    {
        NewCommand = new RelayCommand(New, () => true);
        SaveCommand = new RelayCommand(CommitAndPersist);
        ShowEditorCommand = new RelayCommand(() => CurrentView = "Editor");
        ShowListCommand = new RelayCommand(GoList);
        OpenSelectedCommand = new RelayCommand(OpenSelected, () => Selected is not null);
        DeleteSelectedCommand = new RelayCommand(DeleteSelected, () => Selected is not null);
        TogglePinCommand = new RelayCommand(() => WindowTopmost = !WindowTopmost);

        _autoSave.Tick += (_, _) =>
        {
            _autoSave.Stop();
            FlushEditorToItem();
            Persist();
        };

        _reminderTimer.Tick += (_, _) => OnReminderTimerTick();
        _reminderTimer.Start();

        _root = _data.Load();
        if (_root.Items.Count == 0)
            Seed();
        else
        {
            foreach (var m in _root.Items.OrderByDescending(x => x.ModifiedAt))
                Items.Add(m);
        }
    }

    public void OnWindowClosing()
    {
        _reminderTimer.Stop();
        _autoSave.Stop();
        FlushEditorToItem();
        Persist();
    }

    private void OnReminderTimerTick()
    {
        var wasDirty = _isDirty;
        FlushEditorToItem();
        if (wasDirty)
            Persist();

        var now = DateTimeOffset.Now;
        var anyFired = false;
        var owner = Application.Current?.MainWindow;

        foreach (var m in Items)
        {
            if (m.ReminderAt is not { } t || t > now)
                continue;
            anyFired = true;
            m.ReminderAt = null;
            if (_editingId == m.Id)
            {
                _editorReminderAt = null;
                OnPropertyChanged(nameof(EditorReminderAt));
            }

            var showTitle = string.IsNullOrWhiteSpace(m.Title) ? "제목 없음" : m.Title.Trim();
            if (owner is Window w)
            {
                if (w.WindowState == WindowState.Minimized)
                    w.WindowState = WindowState.Normal;
                w.Activate();
            }

            SystemSounds.Asterisk.Play();
            new MemoReminderNotificationWindow(showTitle).ShowDialog();
        }

        if (anyFired)
            Persist();
    }

    public void SetCardTintFromUi(string? hex8)
    {
        if (string.IsNullOrWhiteSpace(hex8))
            return;
        _root.Settings.CardTintHex = hex8.Trim();
        OnPropertyChanged(nameof(CardTintHex));
        Persist();
    }

    public void SetTopmostFromUi(bool topmost)
    {
        if (_root.Settings.WindowTopmost == topmost)
            return;
        _root.Settings.WindowTopmost = topmost;
        OnPropertyChanged(nameof(WindowTopmost));
        OnPropertyChanged(nameof(IsPinned));
        Persist();
    }

    public sealed record ReminderHistoryUiItem(
        string Label,
        bool Enabled,
        DateTimeOffset ChangedAtLocal,
        DateTimeOffset? TargetAtLocal);

    private void EnsureReminderHistoryMigrated()
    {
        _root.Settings.ReminderHistoryV2 ??= new();
        _root.Settings.ReminderTimeHistory ??= new();

        if (_root.Settings.ReminderHistoryV2.Count > 0)
            return;

        // v1(시간만) → v2(상태 포함): 기존 항목은 "켜짐"으로 간주
        foreach (var s in _root.Settings.ReminderTimeHistory)
        {
            if (!DateTimeOffset.TryParse(s, null, out var p))
                continue;
            _root.Settings.ReminderHistoryV2.Add(new ReminderHistoryEntry
            {
                Enabled = true,
                TargetAtIso = p.ToString("O"),
                ChangedAtIso = DateTimeOffset.Now.ToString("O"),
            });
            if (_root.Settings.ReminderHistoryV2.Count >= 5)
                break;
        }
    }

    /// <summary>알림 창 '최근 설정' 리스트(상태 포함), 최대 5, 최신이 앞.</summary>
    public IReadOnlyList<ReminderHistoryUiItem> GetReminderHistory()
    {
        EnsureReminderHistoryMigrated();
        var items = new List<ReminderHistoryUiItem>();
        foreach (var e in _root.Settings.ReminderHistoryV2)
        {
            if (!DateTimeOffset.TryParse(e.ChangedAtIso, null, out var changed))
                continue;
            DateTimeOffset? target = null;
            if (!string.IsNullOrWhiteSpace(e.TargetAtIso) && DateTimeOffset.TryParse(e.TargetAtIso, null, out var t))
                target = t;

            var changedLocal = changed.ToLocalTime();
            var targetLocal = target?.ToLocalTime();
            var label = e.Enabled
                ? $"켜짐 · {targetLocal:yyyy.MM.dd HH:mm} (설정: {changedLocal:MM.dd HH:mm})"
                : $"꺼짐 · {changedLocal:yyyy.MM.dd HH:mm}";

            items.Add(new ReminderHistoryUiItem(label, e.Enabled, changedLocal, targetLocal));
        }

        return items;
    }

    /// <summary>알림을 켠 이력을 추가(최대 5), 저장합니다.</summary>
    public void AddReminderEnabledHistory(DateTimeOffset targetAtLocal)
    {
        EnsureReminderHistoryMigrated();
        var dto = targetAtLocal.ToLocalTime();
        var entry = new ReminderHistoryEntry
        {
            Enabled = true,
            TargetAtIso = dto.ToString("O"),
            ChangedAtIso = DateTimeOffset.Now.ToString("O"),
        };
        PushReminderHistoryEntry(entry);
    }

    /// <summary>알림을 끈 이력을 추가(최대 5), 저장합니다.</summary>
    public void AddReminderDisabledHistory()
    {
        EnsureReminderHistoryMigrated();
        var entry = new ReminderHistoryEntry
        {
            Enabled = false,
            TargetAtIso = null,
            ChangedAtIso = DateTimeOffset.Now.ToString("O"),
        };
        PushReminderHistoryEntry(entry);
    }

    private void PushReminderHistoryEntry(ReminderHistoryEntry entry)
    {
        var next = new List<ReminderHistoryEntry> { entry };

        foreach (var e in _root.Settings.ReminderHistoryV2)
        {
            if (next.Count >= 5)
                break;
            if (e.Enabled == entry.Enabled && string.Equals(e.TargetAtIso, entry.TargetAtIso, StringComparison.Ordinal))
                continue;
            next.Add(e);
        }

        _root.Settings.ReminderHistoryV2 = next;
        Persist();
    }

    private void GoList()
    {
        FlushEditorToItem();
        Persist();
        CurrentView = "List";
    }

    private void New()
    {
        if (_isDirty && (NewTitle.Length > 0 || !MemoBodyDocumentHelper.IsBodyVisuallyEmpty(NewBody)
                         || _editorScheduleDate is not null || _editorReminderAt is not null))
        {
            var r = MessageBox.Show(
                "편집 중인 내용이 있습니다. 저장하지 않고 새 메모를 시작할까요?\n(아니요: 취소, 예: 지우고 새로)",
                "HdLabs Memo",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (r == MessageBoxResult.No)
                return;
        }

        _editingId = null;
        _suppressDirty = true;
        try
        {
            _newTitle = "";
            _newTitleXaml = "";
            _newBody = "";
            _editorScheduleDate = null;
            _editorReminderAt = null;
            OnPropertyChanged(nameof(NewTitle));
            OnPropertyChanged(nameof(NewTitleXaml));
            OnPropertyChanged(nameof(NewBody));
            OnPropertyChanged(nameof(EditorScheduleDate));
            OnPropertyChanged(nameof(EditorReminderAt));
            _isDirty = false;
        }
        finally
        {
            _suppressDirty = false;
        }

        _autoSave.Stop();
        CurrentView = "Editor";
    }

    public void OpenSelected()
    {
        if (Selected is null)
            return;
        if (_isDirty)
            FlushEditorToItem();
        _editingId = Selected.Id;
        _suppressDirty = true;
        try
        {
            _newTitle = Selected.Title;
            _newTitleXaml = Selected.TitleXaml ?? "";
            _newBody = Selected.Body;
            _editorScheduleDate = Selected.ScheduleDate;
            _editorReminderAt = Selected.ReminderAt;
            OnPropertyChanged(nameof(NewTitle));
            OnPropertyChanged(nameof(NewTitleXaml));
            OnPropertyChanged(nameof(NewBody));
            OnPropertyChanged(nameof(EditorScheduleDate));
            OnPropertyChanged(nameof(EditorReminderAt));
            _isDirty = false;
        }
        finally
        {
            _suppressDirty = false;
        }

        _autoSave.Stop();
        CurrentView = "Editor";
    }

    private void DeleteSelected()
    {
        if (Selected is null)
            return;
        var id = Selected.Id;
        if (_editingId == id)
        {
            _editingId = null;
            _suppressDirty = true;
            try
            {
                _newTitle = "";
                _newTitleXaml = "";
                _newBody = "";
                _editorScheduleDate = null;
                _editorReminderAt = null;
                OnPropertyChanged(nameof(NewTitle));
                OnPropertyChanged(nameof(NewTitleXaml));
                OnPropertyChanged(nameof(NewBody));
                OnPropertyChanged(nameof(EditorScheduleDate));
                OnPropertyChanged(nameof(EditorReminderAt));
                _isDirty = false;
            }
            finally
            {
                _suppressDirty = false;
            }
        }

        for (var i = Items.Count - 1; i >= 0; i--)
        {
            if (Items[i].Id == id)
            {
                Items.RemoveAt(i);
                break;
            }
        }

        Selected = null;
        Persist();
        OpenSelectedCommand.RaiseCanExecuteChanged();
        DeleteSelectedCommand.RaiseCanExecuteChanged();
    }

    public bool DeleteByIds(IEnumerable<Guid> ids)
    {
        var idSet = new HashSet<Guid>(ids);
        if (idSet.Count == 0)
            return false;

        var removedAny = false;

        if (_editingId is Guid eid && idSet.Contains(eid))
        {
            _editingId = null;
            _suppressDirty = true;
            try
            {
                _newTitle = "";
                _newTitleXaml = "";
                _newBody = "";
                _editorScheduleDate = null;
                _editorReminderAt = null;
                OnPropertyChanged(nameof(NewTitle));
                OnPropertyChanged(nameof(NewTitleXaml));
                OnPropertyChanged(nameof(NewBody));
                OnPropertyChanged(nameof(EditorScheduleDate));
                OnPropertyChanged(nameof(EditorReminderAt));
                _isDirty = false;
            }
            finally
            {
                _suppressDirty = false;
            }
        }

        for (var i = Items.Count - 1; i >= 0; i--)
        {
            if (!idSet.Contains(Items[i].Id))
                continue;
            Items.RemoveAt(i);
            removedAny = true;
        }

        if (Selected is { } s && idSet.Contains(s.Id))
            Selected = null;

        if (removedAny)
        {
            Persist();
            OpenSelectedCommand.RaiseCanExecuteChanged();
            DeleteSelectedCommand.RaiseCanExecuteChanged();
        }

        return removedAny;
    }

    public bool DeleteEditingOrSelected()
    {
        if (_editingId is Guid eid)
            return DeleteByIds(new[] { eid });
        if (Selected is { } s)
            return DeleteByIds(new[] { s.Id });
        return false;
    }

    private void CommitAndPersist()
    {
        FlushEditorToItem();
        Persist();
    }

    private void FlushEditorToItem()
    {
        if (!_isDirty)
            return;
        var title = string.IsNullOrWhiteSpace(NewTitle) ? "제목 없음" : NewTitle.Trim();
        var body = NewBody ?? "";
        var now = DateTimeOffset.Now;

        if (_editingId is Guid eid)
        {
            var item = Items.FirstOrDefault(x => x.Id == eid);
            if (item is not null)
            {
                item.Title = title;
                item.TitleXaml = string.IsNullOrWhiteSpace(NewTitleXaml) ? null : NewTitleXaml;
                item.Body = body;
                item.ModifiedAt = now;
                item.ScheduleDate = _editorScheduleDate;
                item.ReminderAt = _editorReminderAt;
            }
        }
        else if (!string.IsNullOrEmpty(NewTitle) || !MemoBodyDocumentHelper.IsBodyVisuallyEmpty(NewBody)
                 || _editorScheduleDate is not null
                 || _editorReminderAt is not null)
        {
            var item = new MemoItem
            {
                Id = Guid.NewGuid(),
                Title = title,
                TitleXaml = string.IsNullOrWhiteSpace(NewTitleXaml) ? null : NewTitleXaml,
                Body = body,
                CreatedAt = now,
                ModifiedAt = now,
                ScheduleDate = _editorScheduleDate,
                ReminderAt = _editorReminderAt,
            };
            Items.Insert(0, item);
            _editingId = item.Id;
        }

        _isDirty = false;
    }

    private void RestartAutoSave()
    {
        if (CurrentView != "Editor" || _suppressDirty)
            return;
        if (_editingId is null && string.IsNullOrEmpty(NewTitle) && string.IsNullOrEmpty(NewTitleXaml)
            && MemoBodyDocumentHelper.IsBodyVisuallyEmpty(NewBody)
            && _editorScheduleDate is null
            && _editorReminderAt is null)
            return;
        _autoSave.Stop();
        _autoSave.Start();
    }

    private void Persist()
    {
        _root.Items = Items.ToList();
        _data.Save(_root);
    }

    private void Seed()
    {
        var now = DateTimeOffset.Now;
        Items.Add(new MemoItem
        {
            Title = "할일 목록",
            Body = "3월 결산" + Environment.NewLine + "보고 자료 작성" + Environment.NewLine + "업무 일지 작성",
            CreatedAt = now.AddDays(-2),
            ModifiedAt = now.AddHours(-3),
            Group = "업무"
        });
        Items.Add(new MemoItem
        {
            Title = "고객 미팅",
            Body = "오후 4시 화상 회의",
            CreatedAt = now.AddDays(-1),
            ModifiedAt = now.AddMinutes(-20),
            Group = "일정"
        });
        _root.Items = Items.ToList();
        _data.Save(_root);
    }
}
