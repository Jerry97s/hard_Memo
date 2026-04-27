using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Interop;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using HdLabs.Memo.Ai;
using HdLabs.Memo.Helpers;
using HdLabs.Memo.Models;
using HdLabs.Memo.ViewModels;
using Microsoft.Win32;
using Key = System.Windows.Input.Key;

namespace HdLabs.Memo;

public partial class MainWindow : Window
{
    private const string ClipboardFormatMemoImages = "HdLabs.Memo.Images.v1";
    private enum FormatTarget
    {
        Body = 0,
        Title = 1,
    }

    private AiAssistantWindow? _aiWindow;

    private bool _loadingBody;
    private string? _lastBodyXaml;
    private bool _updatingFontPickers;
    private bool _normalizingCaretAfterZwspClick;
    private bool _applyingPerParagraphTypingFormat;
    private FormatTarget _formatTarget = FormatTarget.Body;
    private ObservableCollection<FontFamily>? _fontFamilies;
    private ICollectionView? _fontFamiliesView;
    private string _fontFamilyFilterText = "";
    private readonly DispatcherTimer _fontFilterTimer = new() { Interval = TimeSpan.FromMilliseconds(120) };
    private bool _suppressFontFamilySelectionApply;
    private bool _openingFontDropdownFromTyping;
    /// <summary>RichTextBox 빈 selection의 GetPropertyValue(글꼴)가 "왼쪽 런"을 반환하는(WPF) 동작을 보완: 툴바로 삽입 서식을 적용한 캐럿(심볼 오프셋)에 캐시.</summary>
    private int? _insertionToolbarSymbolOffset;
    private FontFamily? _insertionToolbarFontFamily;
    private double? _insertionToolbarSizeDips;
    private string? _insertionToolbarPlainTextSnapshot;
    // "칸(문단)" 단위로 마지막으로 지정한 타이핑 서식을 유지
    private readonly Dictionary<Paragraph, (FontFamily? Family, double? SizeDips)> _typingFormatByParagraph = new();
    private readonly DispatcherTimer _bulletPopupCloseTimer;
    private GlobalHotkey? _hotkey;
    private HwndSource? _hwndSource;
    private bool _exiting;
    private static readonly int[] FontSizePresets =
        { 8, 9, 10, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 28, 32, 36, 40, 48, 56, 64, 72 };

    /// <summary>WPF <see cref="TextElement.FontSize"/> 는 DIP(1/96in). typographic pt(1/72in) = DIP×72/96.</summary>
    private static double PtsToDips(double pt) => pt * (96.0 / 72.0);

    private static double DipsToPts(double dips) => dips * (72.0 / 96.0);

    public MainWindow()
    {
        InitializeComponent();
        _bulletPopupCloseTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(240) };
        _bulletPopupCloseTimer.Tick += (_, _) =>
        {
            _bulletPopupCloseTimer.Stop();
            if (IsLoaded)
                BulletStylesPopup.IsOpen = false;
        };
        var vm = new MainViewModel();
        DataContext = vm;
        vm.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName is nameof(MainViewModel.WindowTopmost))
                Topmost = vm.WindowTopmost;
            else if (e.PropertyName is nameof(MainViewModel.NewTitleXaml))
                LoadTitleFromViewModel();
            else if (e.PropertyName is nameof(MainViewModel.NewBody))
                LoadBodyFromViewModel();
        };
        Topmost = vm.WindowTopmost;
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel vm)
            ApplyCardBackground(vm.CardTintHex);

        InitializeFontPickers();
        LoadTitleFromViewModel();
        MemoBody.Focus();
        // MemoBody.Loaded(→LoadBody)보다 먼저 올 수 있어, 한 프레임 뒤에 다시 맞춤
        Dispatcher.BeginInvoke(
            new Action(() =>
            {
                if (DataContext is not MainViewModel v || MemoBody.Document is not { } d)
                    return;
                SetCaretToPreferredOnBodyLoad(v.NewBody, d);
            }),
            DispatcherPriority.Input);

        EnsureTrayAndHotkey();
    }

    private void EnsureTrayAndHotkey()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero)
            return;

        _hwndSource ??= HwndSource.FromHwnd(hwnd);
        if (_hwndSource is not null)
        {
            _hwndSource.RemoveHook(WndProc);
            _hwndSource.AddHook(WndProc);
        }

        _hotkey ??= new GlobalHotkey(hwnd, id: 0xBEEF);
        if (DataContext is MainViewModel vm && TryParseHotkey(vm.BringToFrontHotkey, out var mods, out var key))
            _hotkey.Register(mods | GlobalHotkey.Modifiers.NoRepeat, key);
        else
            _hotkey.Register(GlobalHotkey.Modifiers.Control | GlobalHotkey.Modifiers.Alt | GlobalHotkey.Modifiers.NoRepeat, Key.M);
    }

    private static bool TryParseHotkey(string? text, out GlobalHotkey.Modifiers mods, out Key key)
    {
        mods = GlobalHotkey.Modifiers.None;
        key = Key.None;
        if (string.IsNullOrWhiteSpace(text))
            return false;
        var parts = text.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length < 2)
            return false;

        if (!Enum.TryParse(parts[^1], ignoreCase: true, out key))
            return false;

        for (var i = 0; i < parts.Length - 1; i++)
        {
            var p = parts[i];
            if (p.Equals("Ctrl", StringComparison.OrdinalIgnoreCase) || p.Equals("Control", StringComparison.OrdinalIgnoreCase))
                mods |= GlobalHotkey.Modifiers.Control;
            else if (p.Equals("Alt", StringComparison.OrdinalIgnoreCase))
                mods |= GlobalHotkey.Modifiers.Alt;
            else if (p.Equals("Shift", StringComparison.OrdinalIgnoreCase))
                mods |= GlobalHotkey.Modifiers.Shift;
            else if (p.Equals("Win", StringComparison.OrdinalIgnoreCase) || p.Equals("Windows", StringComparison.OrdinalIgnoreCase))
                mods |= GlobalHotkey.Modifiers.Win;
        }

        return mods != GlobalHotkey.Modifiers.None && key != Key.None;
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == GlobalHotkey.WmHotkey)
        {
            handled = true;
            BringToFrontFromAnywhere();
        }
        return IntPtr.Zero;
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_exiting)
        {
            if (DataContext is MainViewModel vm)
                vm.OnWindowClosing();
            _hotkey?.Dispose();
            return;
        }

        e.Cancel = true;
        HideToTray();
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            HideToTray();
            e.Handled = true;
            return;
        }
        if (e.Key != Key.Delete)
            return;
        if (Keyboard.FocusedElement is not ListView && Keyboard.FocusedElement is not ListViewItem)
            return;
        if (DataContext is not MainViewModel vm)
            return;
        if (vm.IsListView && vm.DeleteSelectedCommand.CanExecute(null))
        {
            vm.DeleteSelectedCommand.Execute(null);
            e.Handled = true;
        }
    }

    private void MemoList_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is MainViewModel vm)
            vm.OpenSelected();
    }

    private void MemoList_OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Delete)
            return;
        DeleteSelectedMemosFromList();
        e.Handled = true;
    }

    private void Chrome_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.LeftButton == MouseButtonState.Pressed)
            DragMove();
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    private void HideToTray_Click(object sender, RoutedEventArgs e) => HideToTray();

    private void ExitApp_Click(object sender, RoutedEventArgs e) => ExitApp();

    private void HotkeySetup_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not MainViewModel vm)
            return;
        new HotkeySetupWindow(vm) { Owner = this }.ShowDialog();
        EnsureTrayAndHotkey();
    }

    private void MemoList_DeleteSelected_Click(object sender, RoutedEventArgs e) => DeleteSelectedMemosFromList();

    private void DeleteCurrentMemo_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not MainViewModel vm)
            return;

        var title = vm.NewTitle;
        var display = string.IsNullOrWhiteSpace(title) ? "제목 없음" : title.Trim();
        if (!MemoConfirmDialog.ShowDialog(this, "삭제", $"\"{display}\" 메모를 삭제할까요?"))
            return;

        vm.DeleteEditingOrSelected();
        vm.ShowListCommand.Execute(null);
    }

    private void DeleteSelectedMemosFromList()
    {
        if (DataContext is not MainViewModel vm)
            return;

        var selected = MemoList.SelectedItems.Cast<object>().ToList();
        var memos = selected.OfType<MemoItem>().ToList();
        if (memos.Count == 0 && vm.Selected is not null)
            memos = new() { vm.Selected };

        if (memos.Count == 0)
            return;

        var msg = memos.Count == 1
            ? $"\"{memos[0].Title}\" 메모를 삭제할까요?"
            : $"선택한 {memos.Count}개 메모를 삭제할까요?";

        if (!MemoConfirmDialog.ShowDialog(this, "삭제", msg))
            return;

        vm.DeleteByIds(memos.Select(m => m.Id));
    }

    private void HideToTray()
    {
        EnsureTrayAndHotkey();
        if (WindowState != WindowState.Minimized)
            WindowState = WindowState.Minimized;
        Hide();
    }

    private void ExitApp()
    {
        _exiting = true;
        Close();
        Application.Current.Shutdown();
    }

    private void BringToFrontFromAnywhere()
    {
        EnsureTrayAndHotkey();
        Show();
        if (WindowState == WindowState.Minimized)
            WindowState = WindowState.Normal;

        // Topmost toggle trick to ensure foreground
        var wasTopmost = Topmost;
        Topmost = true;
        Activate();
        Topmost = wasTopmost;
        Focus();
    }

    private void AiAssistant_Click(object sender, RoutedEventArgs e)
    {
        if (_aiWindow is null || !_aiWindow.IsLoaded)
        {
            _aiWindow = new AiAssistantWindow(GetAiContextSnapshot);
            _aiWindow.Owner = this;
        }
        _aiWindow.Show();
        _aiWindow.Activate();
    }

    private AiMemoContext GetAiContextSnapshot()
    {
        var titlePlain = "";
        var titleXaml = "";
        try
        {
            if (MemoTitle?.Document is FlowDocument td)
            {
                titlePlain = MemoTitleDocumentHelper.ToPlainText(td);
                titleXaml = MemoTitleDocumentHelper.ToXamlString(td);
            }
        }
        catch
        {
            // ignore
        }

        var bodyPlain = "";
        var bodyXaml = "";
        try
        {
            if (MemoBody?.Document is FlowDocument bd)
            {
                bodyPlain = new TextRange(bd.ContentStart, bd.ContentEnd).Text;
                bodyXaml = MemoBodyDocumentHelper.ToXamlString(bd);
            }
        }
        catch
        {
            // ignore
        }

        return new AiMemoContext(
            TitlePlain: titlePlain,
            TitleXaml: titleXaml,
            BodyPlain: bodyPlain,
            BodyXaml: bodyXaml);
    }

    private void MoreOptions_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.ContextMenu is null)
            return;

        btn.ContextMenu.PlacementTarget = btn;
        btn.ContextMenu.Placement = PlacementMode.Bottom;
        btn.ContextMenu.DataContext = DataContext;
        btn.ContextMenu.IsOpen = true;
        e.Handled = true;
    }

    private void CalendarButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not MainViewModel vm)
            return;
        new CalendarPickerWindow(vm) { Owner = this }.ShowDialog();
    }

    private void AlarmButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not MainViewModel vm)
            return;
        new ReminderSetupWindow(vm) { Owner = this }.ShowDialog();
    }

    private void ColorPickerButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button b || b.ContextMenu is null)
            return;
        b.ContextMenu.PlacementTarget = b;
        b.ContextMenu.Placement = PlacementMode.Bottom;
        b.ContextMenu.DataContext = DataContext;
        b.ContextMenu.IsOpen = true;
    }

    private void ColorTintClick(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem { Tag: string hex } || DataContext is not MainViewModel vm)
            return;
        vm.SetCardTintFromUi(hex);
        ApplyCardBackground(hex);
    }

    private void ApplyCardBackground(string? hex8)
    {
        if (string.IsNullOrWhiteSpace(hex8))
        {
            RootMemoCard.Background = (Brush)FindResource("MemoBodyBrush");
            return;
        }

        try
        {
            var c = (Color)ColorConverter.ConvertFromString(hex8);
            var top = Lighter(c, 0.12);
            var bottom = Darker(c, 0.11);
            var brush = new LinearGradientBrush
            {
                StartPoint = new Point(0, 0),
                EndPoint = new Point(0, 1),
            };
            brush.GradientStops.Add(new GradientStop(top, 0));
            brush.GradientStops.Add(new GradientStop(Color.FromArgb(c.A, c.R, c.G, c.B), 0.55));
            brush.GradientStops.Add(new GradientStop(bottom, 1));
            RootMemoCard.Background = brush;
        }
        catch
        {
            RootMemoCard.Background = (Brush)FindResource("MemoBodyBrush");
        }
    }

    private static Color Lighter(Color c, double amount) =>
        Color.FromArgb(c.A,
            (byte)Math.Min(255, (int)(c.R + 255 * amount)),
            (byte)Math.Min(255, (int)(c.G + 255 * amount)),
            (byte)Math.Min(255, (int)(c.B + 255 * amount)));

    private static Color Darker(Color c, double amount) =>
        Color.FromArgb(c.A,
            (byte)Math.Max(0, (int)(c.R * (1 - amount))),
            (byte)Math.Max(0, (int)(c.G * (1 - amount))),
            (byte)Math.Max(0, (int)(c.B * (1 - amount))));

    #region RichTextBox 본문·서식

    private void MemoBodyContext_Cut_Click(object sender, RoutedEventArgs e)
    {
        MemoBody.Focus();
        ApplicationCommands.Cut.Execute(null, MemoBody);
        OnMemoBodyUserEdit();
    }

    private void MemoBodyContext_Copy_Click(object sender, RoutedEventArgs e)
    {
        MemoBody.Focus();

        var paths = CollectImagePathsFromBodySelection();
        if (paths.Count > 0)
        {
            try
            {
                var json = JsonSerializer.Serialize(paths);
                var data = new DataObject();
                data.SetData(ClipboardFormatMemoImages, json);
                data.SetText(string.Join(Environment.NewLine, paths));
                Clipboard.SetDataObject(data, copy: true);
            }
            catch
            {
                // fallback
                ApplicationCommands.Copy.Execute(null, MemoBody);
            }
            return;
        }

        ApplicationCommands.Copy.Execute(null, MemoBody);
    }

    private void MemoBodyContext_Paste_Click(object sender, RoutedEventArgs e)
    {
        MemoBody.Focus();
        try
        {
            if (TryPasteImagesFromClipboard())
            {
                OnMemoBodyUserEdit();
                return;
            }
        }
        catch
        {
            // ignore and fallback
        }

        ApplicationCommands.Paste.Execute(null, MemoBody);
        OnMemoBodyUserEdit();
    }

    private bool TryPasteImagesFromClipboard()
    {
        if (Clipboard.ContainsData(ClipboardFormatMemoImages)
            && Clipboard.GetData(ClipboardFormatMemoImages) is string json)
        {
            var paths = JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
            var any = false;
            foreach (var p in paths.Where(File.Exists))
            {
                if (!IsImageFilePath(p))
                    continue;
                MemoRtbInserter.InsertImageFromFilePath(MemoBody, p, TryFindResource("MemoTextBrush") as Brush, OnMemoBodyUserEdit);
                any = true;
            }
            return any;
        }

        if (Clipboard.ContainsFileDropList())
        {
            var list = Clipboard.GetFileDropList();
            var any = false;
            foreach (var p in list.Cast<string>().Where(File.Exists))
            {
                if (!IsImageFilePath(p))
                    continue;
                MemoRtbInserter.InsertImageFromFilePath(MemoBody, p, TryFindResource("MemoTextBrush") as Brush, OnMemoBodyUserEdit);
                any = true;
            }
            return any;
        }

        if (Clipboard.ContainsImage())
        {
            var bmp = Clipboard.GetImage();
            if (bmp is null)
                return false;
            var tempPath = MemoRtbInserter.StoreBitmapToAppDataAsPng(bmp);
            if (string.IsNullOrEmpty(tempPath) || !File.Exists(tempPath))
                return false;
            MemoRtbInserter.InsertImageFromFilePath(MemoBody, tempPath, TryFindResource("MemoTextBrush") as Brush, OnMemoBodyUserEdit);
            return true;
        }

        return false;
    }

    private List<string> CollectImagePathsFromBodySelection()
    {
        var res = new List<string>();
        if (MemoBody.Document is null)
            return res;

        var start = MemoBody.Selection.Start;
        var end = MemoBody.Selection.End;
        if (start.CompareTo(end) == 0)
        {
            // 커서만 있는 경우: 앞/뒤 요소를 확인
            var back = start.GetAdjacentElement(LogicalDirection.Backward);
            var fwd = start.GetAdjacentElement(LogicalDirection.Forward);
            TryAddImagePathFromElement(back, res);
            TryAddImagePathFromElement(fwd, res);
            return res.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        var p = start;
        while (p is not null && p.CompareTo(end) < 0)
        {
            if (p.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.ElementStart)
            {
                var el = p.GetAdjacentElement(LogicalDirection.Forward);
                TryAddImagePathFromElement(el, res);
            }
            p = p.GetNextContextPosition(LogicalDirection.Forward);
        }

        return res.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static void TryAddImagePathFromElement(object? el, List<string> acc)
    {
        if (el is not InlineUIContainer iuc)
            return;
        if (iuc.Child is not DependencyObject root)
            return;
        var img = FindDescendant<Image>(root);
        if (img?.Source is BitmapImage bi
            && bi.UriSource is { IsFile: true } uri
            && File.Exists(uri.LocalPath))
        {
            acc.Add(uri.LocalPath);
        }
    }

    private static T? FindDescendant<T>(DependencyObject root) where T : DependencyObject
    {
        var n = VisualTreeHelper.GetChildrenCount(root);
        for (var i = 0; i < n; i++)
        {
            var c = VisualTreeHelper.GetChild(root, i);
            if (c is T t)
                return t;
            var sub = FindDescendant<T>(c);
            if (sub is not null)
                return sub;
        }
        return null;
    }

    private bool _loadingTitle;
    private string? _lastTitleXaml;

    private void MemoTitle_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_loadingTitle)
            return;
        if (DataContext is not MainViewModel vm)
            return;
        if (MemoTitle.Document is null)
            return;

        var x = MemoTitleDocumentHelper.ToXamlString(MemoTitle.Document);
        if (string.Equals(x, _lastTitleXaml, StringComparison.Ordinal))
        {
            SyncFontPickersFromSelection();
            return;
        }

        _lastTitleXaml = x;
        vm.NewTitleXaml = x;
        vm.NewTitle = MemoTitleDocumentHelper.ToPlainText(MemoTitle.Document);
        SyncFontPickersFromSelection();
    }

    private void LoadTitleFromViewModel()
    {
        if (DataContext is not MainViewModel vm)
            return;
        var incoming = vm.NewTitleXaml ?? "";
        if (string.IsNullOrEmpty(incoming) && !string.IsNullOrEmpty(vm.NewTitle))
            incoming = vm.NewTitle; // 평문 데이터 호환
        if (string.Equals(incoming, _lastTitleXaml, StringComparison.Ordinal))
            return;

        _loadingTitle = true;
        try
        {
            var doc = MemoTitleDocumentHelper.FromStorageString(incoming);
            MemoTitle.Document = doc;
            _lastTitleXaml = MemoTitleDocumentHelper.ToXamlString(doc);
        }
        finally
        {
            _loadingTitle = false;
        }
        SyncFontPickersFromSelection();
    }

    private void MemoTitle_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;
        e.Handled = true;
        MemoBody.Focus();
    }

    private void MemoBody_Loaded(object sender, RoutedEventArgs e) => LoadBodyFromViewModel();

    private void MemoBody_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_loadingBody)
            return;
        if (DataContext is not MainViewModel vm)
            return;
        if (MemoBody.Document is null)
            return;
        if (_insertionToolbarPlainTextSnapshot is { } snap)
        {
            var plain = new TextRange(MemoBody.Document.ContentStart, MemoBody.Document.ContentEnd).Text;
            if (!string.Equals(plain, snap, StringComparison.Ordinal))
                ClearToolbarInsertionFormatCache();
        }
        var s = MemoBodyDocumentHelper.ToXamlString(MemoBody.Document);
        if (string.Equals(s, _lastBodyXaml, StringComparison.Ordinal))
        {
            SyncFontPickersFromSelection();
            return;
        }
        _lastBodyXaml = s;
        vm.NewBody = s;
        SyncFontPickersFromSelection();
    }

    private void MemoBody_SelectionChanged(object sender, RoutedEventArgs e)
    {
        if (_loadingBody)
            return;
        NormalizeCaretIfOnZwspBoundary();
        ApplyPerParagraphTypingFormatIfNeeded();
        InvalidateToolbarInsertionFormatCacheOnCaretOrSelection();
        SyncFontPickersFromSelection();
    }

    /// <summary>
    /// 글꼴 변경 시 삽입한 U+200B Run 경계로 다시 클릭하면, 캐럿이 U+200B "앞"에 걸려 왼쪽 Run 서식으로 되돌아가는 경우가 있다.
    /// Forward가 정확히 U+200B 한 글자면 캐럿을 그 뒤로 보정한다.
    /// </summary>
    private void NormalizeCaretIfOnZwspBoundary()
    {
        if (_normalizingCaretAfterZwspClick)
            return;
        if (MemoBody.Document is null)
            return;
        if (!MemoBody.Selection.IsEmpty)
            return;

        var p = MemoBody.CaretPosition;
        var trF = GetOneCharForwardAtPointer(p);
        if (trF is null)
            return;
        if (!string.Equals(trF.Text, "\u200b", StringComparison.Ordinal))
            return;

        try
        {
            _normalizingCaretAfterZwspClick = true;
            MemoBody.CaretPosition = trF.End;
        }
        finally
        {
            _normalizingCaretAfterZwspClick = false;
        }
    }

    private void ApplyPerParagraphTypingFormatIfNeeded()
    {
        if (_applyingPerParagraphTypingFormat)
            return;
        if (MemoBody.Document is null)
            return;
        if (!MemoBody.Selection.IsEmpty)
            return;

        var par = MemoRtbInserter.GetParagraphContaining(MemoBody);
        if (par is null)
            return;
        if (!_typingFormatByParagraph.TryGetValue(par, out var fmt))
            return;

        try
        {
            _applyingPerParagraphTypingFormat = true;
            MemoBody.Focus();
            if (fmt.Family is { } ff)
                MemoBody.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, ff);
            if (fmt.SizeDips is { } sz && sz > 0 && !double.IsNaN(sz) && !double.IsInfinity(sz))
                MemoBody.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, sz);
            UpdateToolbarInsertionFormatCacheAfterEmptyApply(fmt.Family, fmt.SizeDips);
        }
        finally
        {
            _applyingPerParagraphTypingFormat = false;
        }
    }

    private void LoadBodyFromViewModel()
    {
        if (DataContext is not MainViewModel vm)
            return;
        var incoming = vm.NewBody ?? "";
        if (string.Equals(incoming, _lastBodyXaml, StringComparison.Ordinal))
            return;
        _loadingBody = true;
        try
        {
            var fg = (Brush)FindResource("MemoTextBrush");
            var doc = MemoBodyDocumentHelper.FromStorageString(incoming, fg);
            MemoBody.Document = doc;
            _lastBodyXaml = MemoBodyDocumentHelper.ToXamlString(doc);
        }
        finally
        {
            _loadingBody = false;
        }
        _typingFormatByParagraph.Clear();
        ClearToolbarInsertionFormatCache();
        if (MemoBody.Document is { } d)
            SetCaretToPreferredOnBodyLoad(incoming, d);
        SyncFontPickersFromSelection();
    }

    /// <summary>빈·거의 빈 본문은 맨 윗칸(첫 문단 시작), 그 외엔 끝에서 이어쓰기.</summary>
    private void SetCaretToPreferredOnBodyLoad(string? body, FlowDocument d)
    {
        var empty = string.IsNullOrEmpty(body) || MemoBodyDocumentHelper.IsBodyVisuallyEmpty(body);
        if (empty)
        {
            NormalizeDocumentToSingleEmptyParagraph(d);
            if (d.Blocks.FirstBlock is Paragraph p)
                MemoBody.CaretPosition = p.ContentStart;
            else
                MemoBody.CaretPosition = d.ContentStart;
            return;
        }

        MemoBody.CaretPosition = d.ContentEnd;
    }

    /// <summary>빈 문서인데도 저장 XAML에 줄바꿈/빈 문단이 남아 "2번째 줄"처럼 시작하는 것을 방지.</summary>
    private static void NormalizeDocumentToSingleEmptyParagraph(FlowDocument d)
    {
        // 기존 블록/인라인을 비워서 1줄(첫 문단)에서 시작하게 만든다.
        d.Blocks.Clear();
        var p = new Paragraph
        {
            Margin = new Thickness(0),
            Padding = new Thickness(0),
        };
        d.Blocks.Add(p);
    }

    private void MemoBody_FontContext_MouseUp(object sender, MouseButtonEventArgs e) =>
        SyncFontPickersFromSelection();

    private void MemoBody_FontContext_KeyUp(object sender, KeyEventArgs e)
    {
        if (_loadingBody || e.IsRepeat)
            return;
        if (e.Key is Key.Left or Key.Right or Key.Up or Key.Down or Key.Home or Key.End
            or Key.PageUp or Key.PageDown)
        {
            SyncFontPickersFromSelection();
            return;
        }
        if (e.Key is Key.A && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            SyncFontPickersFromSelection();
    }

    private void PushXamlToViewModel()
    {
        if (DataContext is not MainViewModel vm || MemoBody.Document is null)
            return;
        if (_loadingBody)
            return;
        var s = MemoBodyDocumentHelper.ToXamlString(MemoBody.Document);
        if (string.Equals(s, _lastBodyXaml, StringComparison.Ordinal))
            return;
        _lastBodyXaml = s;
        vm.NewBody = s;
    }

    private void FormatBold_Click(object sender, RoutedEventArgs e) =>
        RunEditing(EditingCommands.ToggleBold);

    private void FormatItalic_Click(object sender, RoutedEventArgs e) =>
        RunEditing(EditingCommands.ToggleItalic);

    private void FormatUnderline_Click(object sender, RoutedEventArgs e) =>
        RunEditing(EditingCommands.ToggleUnderline);

    private void FormatAlignLeft_Click(object sender, RoutedEventArgs e) =>
        RunEditing(EditingCommands.AlignLeft);

    private void FormatAlignCenter_Click(object sender, RoutedEventArgs e) =>
        RunEditing(EditingCommands.AlignCenter);

    private void FormatAlignRight_Click(object sender, RoutedEventArgs e) =>
        RunEditing(EditingCommands.AlignRight);

    private void FormatStrike_Click(object sender, RoutedEventArgs e)
    {
        MemoBody.Focus();
        // 선택이 비어 있어도(커서만 있어도) 이후 입력에 적용되도록 Selection에 설정
        var tr = new TextRange(MemoBody.Selection.Start, MemoBody.Selection.End);
        var dec = (tr.GetPropertyValue(Inline.TextDecorationsProperty) as TextDecorationCollection) ?? new TextDecorationCollection();
        var hasStrike = dec.Any(d => d.Location == TextDecorationLocation.Strikethrough);
        var next = new TextDecorationCollection();
        foreach (var t in dec)
        {
            if (t.Location != TextDecorationLocation.Strikethrough)
                next.Add(t);
        }
        if (!hasStrike)
        {
            foreach (var t in TextDecorations.Strikethrough)
                next.Add(t);
        }
        MemoBody.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, next);
        PushXamlToViewModel();
    }

    private void RunEditing(RoutedUICommand command)
    {
        MemoBody.Focus();
        command.Execute(null, MemoBody);
        PushXamlToViewModel();
    }

    private void OnMemoBodyUserEdit() => SyncMemoBodyToViewModelIfNotLoading();

    private void SyncMemoBodyToViewModelIfNotLoading()
    {
        if (_loadingBody)
            return;
        PushXamlToViewModel();
    }

    private void InsertChecklistLine_Click(object sender, RoutedEventArgs e)
    {
        var fg = TryFindResource("MemoTextBrush") as Brush;
        MemoRtbInserter.InsertChecklistLine(MemoBody, fg, OnMemoBodyUserEdit);
    }

    private void InsertImageFromFile_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "이미지|*.png;*.jpg;*.jpeg;*.jpe;*.gif;*.bmp;*.tif;*.tiff;*.ico;*.wdp;*.hdp;*.jxr|모든 파일|*.*",
            Title = "이미지 파일 선택",
        };
        if (dlg.ShowDialog() != true)
            return;
        if (string.IsNullOrWhiteSpace(dlg.FileName) || !File.Exists(dlg.FileName))
            return;
        if (!IsImageFilePath(dlg.FileName))
        {
            MessageBox.Show("지원하는 이미지 파일만 넣을 수 있어요.", "이미지", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        try
        {
            MemoRtbInserter.InsertImageFromFilePath(MemoBody, dlg.FileName, TryFindResource("MemoTextBrush") as Brush, OnMemoBodyUserEdit);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"이미지를 넣을 수 없습니다.\n{ex.Message}", "이미지", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void MemoBody_PreviewDragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop)
            && e.Data.GetData(DataFormats.FileDrop) is string[] paths
            && paths.Any(static p => !string.IsNullOrWhiteSpace(p) && IsImageFilePath(p) && File.Exists(p)))
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }
    }

    private void MemoBody_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is not string[] paths)
            return;
        e.Handled = true;
        foreach (var p in paths)
        {
            if (string.IsNullOrWhiteSpace(p) || !File.Exists(p) || !IsImageFilePath(p))
                continue;
            try
            {
                MemoRtbInserter.InsertImageFromFilePath(MemoBody, p, TryFindResource("MemoTextBrush") as Brush, OnMemoBodyUserEdit);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이미지를 넣을 수 없습니다.\n{p}\n{ex.Message}", "이미지", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }

    private static bool IsImageFilePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return false;
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext is ".png" or ".jpg" or ".jpeg" or ".jpe" or ".gif" or ".bmp" or ".tif" or ".tiff" or ".ico"
            or ".wdp" or ".hdp" or ".jxr";
    }

    private void BulletAnchorButton_MouseEnter(object sender, MouseEventArgs e)
    {
        _bulletPopupCloseTimer.Stop();
        BulletStylesPopup.IsOpen = true;
    }

    private void BulletAnchorButton_MouseLeave(object sender, MouseEventArgs e) =>
        ScheduleCloseBulletStylesPopup();

    private void BulletPopupContent_MouseEnter(object sender, MouseEventArgs e) =>
        _bulletPopupCloseTimer.Stop();

    private void BulletPopupContent_MouseLeave(object sender, MouseEventArgs e) =>
        ScheduleCloseBulletStylesPopup();

    private void ScheduleCloseBulletStylesPopup()
    {
        _bulletPopupCloseTimer.Stop();
        _bulletPopupCloseTimer.Start();
    }

    private const string DefaultBulletGaulJumTag = "· ";

    private void BulletAnchorButton_Click(object sender, RoutedEventArgs e) =>
        ApplyBulletTag(DefaultBulletGaulJumTag);

    private void InsertBulletWithTag_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button b || b.Tag is not string tag)
            return;
        ApplyBulletTag(tag);
    }

    private void ApplyBulletTag(string tag)
    {
        _bulletPopupCloseTimer.Stop();
        BulletStylesPopup.IsOpen = false;
        InsertLinePrefixRtb(tag);
        PushXamlToViewModel();
        MemoBody.Focus();
    }

    private void InsertLinePrefixRtb(string prefix)
    {
        if (MemoBody.Document is null)
            return;
        var a = new TextRange(MemoBody.Document.ContentStart, MemoBody.CaretPosition);
        var text = a.Text;
        var i = text.Length;
        var insert = (i == 0 || text[i - 1] is '\n' or '\r') ? prefix : "\r" + prefix;
        InsertTextAtCaretRtb(insert);
    }

    private void InsertTextAtCaretRtb(string text)
    {
        var p = !MemoBody.Selection.IsEmpty
            ? MemoBody.Selection.Start
            : MemoBody.CaretPosition;
        new TextRange(p, p) { Text = text };
        MemoBody.Focus();
    }

    #region Font family / size (editable + pt/px)

    private void InitializeFontPickers()
    {
        _fontFamilies = new ObservableCollection<FontFamily>(Fonts.SystemFontFamilies
            .OrderBy(f => f.Source, StringComparer.CurrentCultureIgnoreCase));
        FontFamilyPicker.ItemsSource = _fontFamilies;
        _fontFamiliesView = CollectionViewSource.GetDefaultView(FontFamilyPicker.ItemsSource);
        _fontFamiliesView.Filter = o =>
        {
            if (o is not FontFamily f)
                return false;
            if (string.IsNullOrWhiteSpace(_fontFamilyFilterText))
                return true;
            return f.Source?.IndexOf(_fontFamilyFilterText, StringComparison.CurrentCultureIgnoreCase) >= 0;
        };

        // editable ComboBox의 내부 TextBox 입력 이벤트를 받아서 "타이핑 → 후보 드롭다운"을 구현
        FontFamilyPicker.AddHandler(TextBoxBase.TextChangedEvent, new TextChangedEventHandler(FontFamilyPicker_TextChanged));
        _fontFilterTimer.Tick += (_, _) =>
        {
            _fontFilterTimer.Stop();
            if (_fontFamiliesView is null)
                return;
            _suppressFontFamilySelectionApply = true;
            try
            {
                _fontFamiliesView.Refresh();
            }
            finally
            {
                _suppressFontFamilySelectionApply = false;
            }
            if (FontFamilyPicker.IsKeyboardFocusWithin && !FontFamilyPicker.IsDropDownOpen)
            {
                _openingFontDropdownFromTyping = true;
                FontFamilyPicker.IsDropDownOpen = true;
                _openingFontDropdownFromTyping = false;
            }
        };

        FontSizePicker.ItemsSource = FontSizePresets;
        FontSizeUnitPicker.ItemsSource = new[] { "pt", "px" };
        _updatingFontPickers = true;
        try
        {
            FontSizeUnitPicker.SelectedIndex = 0;
        }
        finally
        {
            _updatingFontPickers = false;
        }
        SyncFontPickersFromSelection();
    }

    private void FontFamilyPicker_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_updatingFontPickers)
            return;
        if (!FontFamilyPicker.IsKeyboardFocusWithin)
            return;

        _fontFamilyFilterText = (FontFamilyPicker.Text ?? "").Trim();
        _fontFilterTimer.Stop();
        _fontFilterTimer.Start();
    }

    private void FontFamilyPicker_DropDownOpened(object sender, EventArgs e)
    {
        if (sender is ComboBox cbb)
            SetFooterComboBoxDropdownOpensUpward(cbb);

        // 버튼/마우스로 리스트를 열 때는 전체 목록이 보여야 함 (검색은 "타이핑 중"에만)
        if (_openingFontDropdownFromTyping)
            return;
        if (string.IsNullOrWhiteSpace(_fontFamilyFilterText))
            return;
        _fontFamilyFilterText = "";
        _fontFamiliesView?.Refresh();
    }

    private void FooterCombo_DropDownOpened(object sender, EventArgs e)
    {
        if (sender is ComboBox cbb)
            SetFooterComboBoxDropdownOpensUpward(cbb);
    }

    /// <summary>하단 툴바 콤보: 드롭이 아래로 길게 빠지며 "별도 창"처럼 보이는 것을 막기 위해, 목록을 위(본문 쪽)로 펼친다.</summary>
    private static void SetFooterComboBoxDropdownOpensUpward(ComboBox cb)
    {
        try
        {
            cb.ApplyTemplate();
            if (cb.Template.FindName("PART_Popup", cb) is Popup { } pop)
            {
                pop.Placement = PlacementMode.Top;
                pop.PlacementTarget = cb;
            }
        }
        catch
        {
            // 템플릿 커스텀 시 PART_Popup 없을 수 있음
        }
    }

    private void FontFamilyPicker_DropDownClosed(object sender, EventArgs e)
    {
        // 닫힌 뒤에는 다음에 열면 전체가 나오도록 필터를 초기화
        if (string.IsNullOrWhiteSpace(_fontFamilyFilterText))
            return;
        _fontFamilyFilterText = "";
        _fontFamiliesView?.Refresh();
    }

    private string CurrentFontSizeUnit() =>
        FontSizeUnitPicker.SelectedItem is string s && s is "pt" or "px" ? s : "pt";

    private void SyncFontPickersFromSelection()
    {
        if (_fontFamilies is null)
            return;
        if (_formatTarget == FormatTarget.Title && MemoTitle is not null)
        {
            // 제목 편집 중이면 제목 글꼴/크기를 콤보에 반영
            _updatingFontPickers = true;
            try
            {
                var ff = MemoTitle.Selection.GetPropertyValue(TextElement.FontFamilyProperty) as FontFamily
                         ?? MemoTitle.FontFamily
                         ?? new FontFamily("Malgun Gothic");
                EnsureFontInList(ff, out var match);
                FontFamilyPicker.Text = match.Source;
                FontFamilyPicker.SelectedItem = match;

                var sizeObj = MemoTitle.Selection.GetPropertyValue(TextElement.FontSizeProperty);
                var sizeDips = sizeObj is double sz && sz > 0 && !double.IsNaN(sz) && !double.IsInfinity(sz)
                    ? sz
                    : MemoTitle.FontSize;
                if (sizeDips <= 0 || double.IsNaN(sizeDips) || double.IsInfinity(sizeDips))
                    sizeDips = 15.0;
                if (string.Equals(FontSizeUnitPicker.SelectedItem as string, "pt", StringComparison.Ordinal))
                    FontSizePicker.Text = DipsToPts(sizeDips).ToString("0.##", CultureInfo.CurrentCulture);
                else
                    FontSizePicker.Text = sizeDips.ToString("0.##", CultureInfo.CurrentCulture);
                var n = CoerceToFontSizePreset(FontSizePicker.Text);
                FontSizePicker.SelectedItem = n;
            }
            finally
            {
                _updatingFontPickers = false;
            }
            return;
        }
        var cur = GetCurrentFontAtCaret();
        _updatingFontPickers = true;
        try
        {
            if (cur.IsFontFamilyMixed)
            {
                FontFamilyPicker.Text = "";
                FontFamilyPicker.SelectedItem = null;
            }
            else
            {
                var ff = cur.Family ?? (MemoBody.Document is FlowDocument d
                    ? d.FontFamily
                    : new FontFamily("Malgun Gothic"));
                EnsureFontInList(ff, out var match);
                FontFamilyPicker.Text = match.Source;
                FontFamilyPicker.SelectedItem = match;
            }
            if (string.Equals(FontSizeUnitPicker.SelectedItem as string, "pt", StringComparison.Ordinal))
                FontSizePicker.Text = DipsToPts(cur.Size).ToString("0.##", CultureInfo.CurrentCulture);
            else
                FontSizePicker.Text = cur.Size.ToString("0.##", CultureInfo.CurrentCulture);
            var n = CoerceToFontSizePreset(FontSizePicker.Text);
            FontSizePicker.SelectedItem = n;
        }
        finally
        {
            _updatingFontPickers = false;
        }
    }

    private static int? CoerceToFontSizePreset(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return null;
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out var d)
            && Math.Abs(d - Math.Round(d)) < 0.0001
            && int.TryParse(d.ToString("F0", CultureInfo.InvariantCulture), out var w))
        {
            return FontSizePresets.Any(p => p == w) ? w : null;
        }
        return null;
    }

    private void UpdateFontSizeDisplayForUnitOnly(double sizeDips)
    {
        _updatingFontPickers = true;
        try
        {
            if (CurrentFontSizeUnit() is "pt")
                FontSizePicker.Text = DipsToPts(sizeDips).ToString("0.##", CultureInfo.CurrentCulture);
            else
                FontSizePicker.Text = sizeDips.ToString("0.##", CultureInfo.CurrentCulture);
            var n = CoerceToFontSizePreset(FontSizePicker.Text);
            FontSizePicker.SelectedItem = n;
        }
        finally
        {
            _updatingFontPickers = false;
        }
    }

    private int GetCaretSymbolOffsetInDocument()
    {
        if (MemoBody.Document is not FlowDocument d)
            return 0;
        return d.ContentStart.GetOffsetToPosition(MemoBody.CaretPosition);
    }

    private void ClearToolbarInsertionFormatCache()
    {
        _insertionToolbarSymbolOffset = null;
        _insertionToolbarFontFamily = null;
        _insertionToolbarSizeDips = null;
        _insertionToolbarPlainTextSnapshot = null;
    }

    private void InvalidateToolbarInsertionFormatCacheOnCaretOrSelection()
    {
        if (MemoBody.Document is null)
            return;
        if (!MemoBody.Selection.IsEmpty)
        {
            ClearToolbarInsertionFormatCache();
            return;
        }
        if (_insertionToolbarSymbolOffset is { } o && GetCaretSymbolOffsetInDocument() != o)
            ClearToolbarInsertionFormatCache();
    }

    /// <summary>빈 선택에 툴바/컨텍스트로 FontFamily/FontSize 를 적용한 뒤, 콤보가 D 옆 런만 읽는(WPF) 문제를 막기 위해 캐시.</summary>
    private void UpdateToolbarInsertionFormatCacheAfterEmptyApply(FontFamily? newFamily, double? newSizeDips)
    {
        if (MemoBody.Document is not FlowDocument d || !MemoBody.Selection.IsEmpty)
            return;
        var off = GetCaretSymbolOffsetInDocument();
        var defF = d.FontFamily ?? new FontFamily("Malgun Gothic");
        var defS = 15.0;
        if (d.FontSize > 0 && !double.IsNaN(d.FontSize) && !double.IsInfinity(d.FontSize))
            defS = d.FontSize;

        FontFamily fam;
        if (newFamily is { } ff)
        {
            fam = ff;
        }
        else if (_insertionToolbarSymbolOffset == off && _insertionToolbarFontFamily is { } cff)
        {
            fam = cff;
        }
        else if (MemoBody.Selection.GetPropertyValue(TextElement.FontFamilyProperty) is FontFamily sff)
        {
            fam = sff;
        }
        else
        {
            fam = defF;
        }

        double sizeD;
        if (newSizeDips is { } nd and > 0 && !double.IsNaN(nd) && !double.IsInfinity(nd))
        {
            sizeD = nd;
        }
        else if (_insertionToolbarSymbolOffset == off && _insertionToolbarSizeDips is { } cs
                 && cs > 0 && !double.IsNaN(cs) && !double.IsInfinity(cs))
        {
            sizeD = cs;
        }
        else if (MemoBody.Selection.GetPropertyValue(TextElement.FontSizeProperty) is double sd
                 && sd > 0 && !double.IsNaN(sd) && !double.IsInfinity(sd))
        {
            sizeD = sd;
        }
        else
        {
            sizeD = defS;
        }

        _insertionToolbarSymbolOffset = off;
        _insertionToolbarFontFamily = fam;
        _insertionToolbarSizeDips = sizeD;
        _insertionToolbarPlainTextSnapshot = new TextRange(d.ContentStart, d.ContentEnd).Text;
    }

    private readonly record struct CaretFontInfo(FontFamily? Family, double Size, bool IsFontFamilyMixed);

    private CaretFontInfo GetCurrentFontAtCaret()
    {
        var d = MemoBody.Document;
        var defF = d?.FontFamily ?? new FontFamily("Malgun Gothic");
        var defS = 15.0;
        if (d is not null)
        {
            var ds = d.FontSize;
            if (ds > 0 && !double.IsNaN(ds) && !double.IsInfinity(ds))
                defS = ds;
        }
        if (MemoBody.Document is null)
            return new CaretFontInfo(defF, defS, false);

        if (MemoBody.Selection.IsEmpty
            && _insertionToolbarSymbolOffset is { } insOff
            && GetCaretSymbolOffsetInDocument() == insOff
            && _insertionToolbarFontFamily is { } cFam
            && _insertionToolbarSizeDips is { } cSz
            && cSz > 0 && !double.IsNaN(cSz) && !double.IsInfinity(cSz))
        {
            return new CaretFontInfo(cFam, cSz, false);
        }

        // 선택이 있으면 선택 범위 스타일. 커서만: 먼저 Selection(=다음 입력/삽입 기본) → 1자는 옆 'D'만 읽는 경우를 보완.
        object fo;
        object so;
        if (!MemoBody.Selection.IsEmpty)
        {
            var tr = new TextRange(MemoBody.Selection.Start, MemoBody.Selection.End);
            fo = tr.GetPropertyValue(TextElement.FontFamilyProperty);
            so = tr.GetPropertyValue(TextElement.FontSizeProperty);
            if (object.ReferenceEquals(fo, DependencyProperty.UnsetValue))
            {
                var s1 = defS;
                if (so is double sD0 && sD0 > 0 && !double.IsNaN(sD0) && !double.IsInfinity(sD0))
                    s1 = sD0;
                return new CaretFontInfo(null, s1, true);
            }
        }
        else
        {
            var p = MemoBody.CaretPosition;
            fo = MemoBody.Selection.GetPropertyValue(TextElement.FontFamilyProperty);
            so = MemoBody.Selection.GetPropertyValue(TextElement.FontSizeProperty);
            if (!(fo is FontFamily)
                && GetTextRangeForFontPickerAtCaret(p, defF) is { } tr1)
            {
                fo = tr1.GetPropertyValue(TextElement.FontFamilyProperty);
                so = tr1.GetPropertyValue(TextElement.FontSizeProperty);
            }
            else if (!(so is double sz0) || sz0 <= 0 || double.IsNaN(sz0) || double.IsInfinity(sz0))
            {
                if (GetTextRangeForFontPickerAtCaret(p, defF) is { } tr2)
                    so = tr2.GetPropertyValue(TextElement.FontSizeProperty);
            }
        }

        var f = defF;
        if (fo is FontFamily ffam)
            f = ffam;
        var s = defS;
        if (so is double sD && sD > 0 && !double.IsNaN(sD) && !double.IsInfinity(sD))
            s = sD;
        return new CaretFontInfo(f, s, false);
    }

    private static TextRange? GetOneCharForwardAtPointer(TextPointer p)
    {
        if (p.GetTextRunLength(LogicalDirection.Forward) > 0)
        {
            var end = p.GetPositionAtOffset(1, LogicalDirection.Forward);
            if (end is not null && p.CompareTo(end) < 0)
                return new TextRange(p, end);
        }
        return null;
    }

    private static TextRange? GetOneCharBackwardAtPointer(TextPointer p)
    {
        if (p.GetTextRunLength(LogicalDirection.Backward) > 0)
        {
            var start = p.GetPositionAtOffset(1, LogicalDirection.Backward);
            if (start is not null && start.CompareTo(p) < 0)
                return new TextRange(start, p);
        }
        return null;
    }

    /// <summary>콤보박스에 "지금 캐럿" 글꼴을 뿌릴 때 쓰는 1자 범위(앞/뒤 둘 있으면 \r·줄끊김/문서기본(앞) vs 실제(뒤) 정리).</summary>
    private static TextRange? GetTextRangeForFontPickerAtCaret(TextPointer p, FontFamily documentDefaultFont)
    {
        var trF = GetOneCharForwardAtPointer(p);
        var trB = GetOneCharBackwardAtPointer(p);
        if (trB is null && trF is null)
            return null;
        if (trB is not null && trF is null)
            return trB;
        if (trF is not null && trB is null)
            return trF;

        // 우리가 삽입 서식 고정용으로 넣는 U+200B는 0폭이라 경계 클릭 시 쉽게 앞으로 읽혀야 한다.
        // (기존에는 U+200B를 "약한 break"로 처리해 뒤(왼쪽) 글꼴로 되돌아갔음)
        if (string.Equals(trF!.Text, "\u200b", StringComparison.Ordinal))
            return trF!;

        var ffB = trB!.GetPropertyValue(TextElement.FontFamilyProperty);
        var ffF = trF!.GetPropertyValue(TextElement.FontFamilyProperty);
        if (ffB is FontFamily a && ffF is FontFamily b
            && string.Equals(a.Source, b.Source, StringComparison.OrdinalIgnoreCase))
            return trF;

        if (IsWeakBreakCharRangeForFontRead(trF!))
            return trB!;

        // 앞(Forward)이 문서 기본(상속)이고 뒤는 Webdings: 앞 1자가 “줄/제어”가 아니면(?, 공백 등) 그 글꼴이 맞다.
        var defS = documentDefaultFont.Source;
        if (ffB is FontFamily backF && ffF is FontFamily fwdF
            && !string.Equals(backF.Source, defS, StringComparison.OrdinalIgnoreCase)
            && string.Equals(fwdF.Source, defS, StringComparison.OrdinalIgnoreCase))
        {
            if (trF.Text is { Length: 1 } s0 && !IsSingleCharLineOrFormatBreak(s0[0]))
                return trF!;
            return trB!;
        }

        // "abc|Webdings" : 앞=일반, 뒤=심볼 — 오른쪽(다음) 글리프
        return trF!;
    }

    private static bool IsWeakBreakCharRangeForFontRead(TextRange trF)
    {
        var t = trF.Text;
        if (t.Length == 0)
            return true;
        if (t is "\r\n" or "\n\r")
            return true;
        if (t.Length == 1)
            return IsSingleCharLineOrFormatBreak(t[0]);
        return false;
    }

    private static bool IsSingleCharLineOrFormatBreak(char c) =>
        c is '\r' or '\n' or '\f' or '\u200b' or '\u00ad' or '\u200c' or '\u200d' or '\u2028' or '\u2029' or '\u00a0'
        || (c < ' ' && c != '\t') // C0(탭 제외) 제어
        || char.GetUnicodeCategory(c) is UnicodeCategory.LineSeparator
            or UnicodeCategory.ParagraphSeparator
            or UnicodeCategory.Format;

    private void EnsureFontInList(FontFamily f, out FontFamily selectedInstance)
    {
        selectedInstance = f;
        if (_fontFamilies is null)
            return;
        var match = _fontFamilies
            .FirstOrDefault(x => string.Equals(x.Source, f.Source, StringComparison.OrdinalIgnoreCase));
        if (match is not null)
        {
            selectedInstance = match;
            return;
        }
        _fontFamilies.Insert(0, f);
        selectedInstance = f;
    }

    private void FontFamilyPicker_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingFontPickers)
            return;
        if (_suppressFontFamilySelectionApply)
            return;
        if (FontFamilyPicker.SelectedItem is not FontFamily fam)
            return;
        if (MemoBody.Document is null)
            return;
        FontFamilyPicker.Text = fam.Source;
        ApplyCurrentFontFamily(fam);
    }

    private void FontFamilyPicker_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (_updatingFontPickers)
            return;
        // 드롭다운에서 항목을 클릭할 때는 포커스가 흔들려도 강제 적용/포커스 이동을 하지 않음
        if (FontFamilyPicker.IsDropDownOpen)
            return;
        TryApplyFontFamilyFromText(focusToMemo: false);
    }

    private void FontFamilyPicker_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;
        e.Handled = true;
        FontFamilyPicker.IsDropDownOpen = false;
        TryApplyFontFamilyFromText(focusToMemo: true);
    }

    private void TryApplyFontFamilyFromText(bool focusToMemo)
    {
        if (MemoBody.Document is null)
            return;
        var t = FontFamilyPicker.Text?.Trim() ?? "";
        if (string.IsNullOrEmpty(t))
            return;
        if (FontFamilyPicker.SelectedItem is FontFamily f
            && string.Equals(f.Source, t, StringComparison.OrdinalIgnoreCase))
        {
            if (focusToMemo)
                MemoBody.Focus();
            return;
        }
        FontFamily? use = _fontFamilies?.FirstOrDefault(x => string.Equals(x.Source, t, StringComparison.OrdinalIgnoreCase));
        try
        {
            use ??= new FontFamily(t);
        }
        catch
        {
            return;
        }
        if (use is not null)
        {
            if (_fontFamilies is not null
                && _fontFamilies.All(x => !string.Equals(x.Source, use.Source, StringComparison.OrdinalIgnoreCase)))
                _fontFamilies.Insert(0, use);
            FontFamilyPicker.SelectedItem = _fontFamilies?.FirstOrDefault(x =>
                string.Equals(x.Source, use.Source, StringComparison.OrdinalIgnoreCase)) ?? use;
            FontFamilyPicker.Text = use.Source;
            ApplyCurrentFontFamily(use);
        }
        if (focusToMemo)
            MemoBody.Focus();
    }

    private void ApplyCurrentFontFamily(FontFamily fam)
    {
        if (_formatTarget == FormatTarget.Title && MemoTitle is not null)
        {
            MemoTitle.Focus();
            MemoTitle.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, fam);
            SyncFontPickersFromSelection();
            return;
        }
        // 선택이 없을 때: '옆 한 글자' Run에 먹이지 말고 Selection(삽입 기본)에만 — 안 그러면 D 스냅 ITC로 다시 읽힘
        MemoBody.Focus();
        if (!MemoBody.Selection.IsEmpty)
        {
            new TextRange(MemoBody.Selection.Start, MemoBody.Selection.End)
                .ApplyPropertyValue(TextElement.FontFamilyProperty, fam);
        }
        else
        {
            if (!MemoRtbInserter.TryApplyEmptySelectionFontToDocument(MemoBody, fam, null))
                MemoBody.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, fam);
            if (MemoRtbInserter.GetParagraphContaining(MemoBody) is { } par)
            {
                _typingFormatByParagraph.TryGetValue(par, out var prev);
                _typingFormatByParagraph[par] = (fam, prev.SizeDips);
            }
            UpdateToolbarInsertionFormatCacheAfterEmptyApply(newFamily: fam, newSizeDips: null);
        }
        PushXamlToViewModel();
        SyncFontPickersFromSelection();
    }

    private void FontSizeUnitPicker_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingFontPickers)
            return;
        var d = GetCurrentFontAtCaret();
        UpdateFontSizeDisplayForUnitOnly(d.Size);
    }

    private void FontSizePicker_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingFontPickers)
            return;
        if (FontSizePicker.SelectedItem is int n)
            ApplyDipsForCurrentUnitValue(n);
    }

    private void FontSizePicker_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (_updatingFontPickers)
            return;
        if (FontSizePicker.IsDropDownOpen)
            return;
        TryApplyFontSizeFromText(focusToMemo: false);
    }

    private void FontSizePicker_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;
        e.Handled = true;
        FontSizePicker.IsDropDownOpen = false;
        TryApplyFontSizeFromText(focusToMemo: true);
    }

    private void TryApplyFontSizeFromText(bool focusToMemo)
    {
        if (MemoBody.Document is null)
            return;
        if (TextInputValueToDips(FontSizePicker.Text, out var dips))
        {
            ApplyDipsToSelection(dips);
            if (focusToMemo)
                MemoBody.Focus();
        }
    }

    private void ApplyDipsForCurrentUnitValue(double unitValue)
    {
        if (CurrentFontSizeUnit() is "pt")
            ApplyDipsToSelection(PtsToDips(unitValue));
        else
            ApplyDipsToSelection(unitValue);
    }

    private bool TextInputValueToDips(string? text, out double dips)
    {
        dips = 0;
        if (string.IsNullOrWhiteSpace(text))
            return false;
        if (!double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out var v) || v <= 0
            || double.IsNaN(v) || double.IsInfinity(v))
            return false;
        dips = CurrentFontSizeUnit() is "pt" ? PtsToDips(v) : v;
        return true;
    }

    private void ApplyDipsToSelection(double dips)
    {
        if (_formatTarget == FormatTarget.Title && MemoTitle is not null)
        {
            MemoTitle.Focus();
            MemoTitle.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, dips);
            SyncFontPickersFromSelection();
            return;
        }
        MemoBody.Focus();
        if (!MemoBody.Selection.IsEmpty)
        {
            new TextRange(MemoBody.Selection.Start, MemoBody.Selection.End)
                .ApplyPropertyValue(TextElement.FontSizeProperty, dips);
        }
        else
        {
            if (!MemoRtbInserter.TryApplyEmptySelectionFontToDocument(MemoBody, null, dips))
                MemoBody.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, dips);
            if (MemoRtbInserter.GetParagraphContaining(MemoBody) is { } par)
            {
                _typingFormatByParagraph.TryGetValue(par, out var prev);
                _typingFormatByParagraph[par] = (prev.Family, dips);
            }
            UpdateToolbarInsertionFormatCacheAfterEmptyApply(newFamily: null, newSizeDips: dips);
        }
        PushXamlToViewModel();
        SyncFontPickersFromSelection();
    }

    private void MemoBodyContext_FontFamily_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem { Tag: string src } || string.IsNullOrWhiteSpace(src) || MemoBody.Document is null)
            return;
        try
        {
            ApplyCurrentFontFamily(new FontFamily(src));
        }
        catch
        {
            // 글꼴 없음/이름만 다른 환경
        }

        MemoBody.Focus();
    }

    private void MemoTitle_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        _formatTarget = FormatTarget.Title;
        SyncFontPickersFromSelection();
    }

    private void MemoBody_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        _formatTarget = FormatTarget.Body;
        SyncFontPickersFromSelection();
    }

    private void MemoBodyContext_FontSizePreset_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem { Tag: string tag }
            || !int.TryParse(tag, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)
            || n <= 0
            || MemoBody.Document is null)
            return;
        ApplyDipsForCurrentUnitValue(n);
        MemoBody.Focus();
    }

    private void MemoBodyContext_FontSizeStep_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem { Tag: string tag }
            || !int.TryParse(tag, NumberStyles.Integer, CultureInfo.InvariantCulture, out var delta))
            return;
        if (MemoBody.Document is null)
            return;
        var d = GetCurrentFontAtCaret().Size;
        if (string.Equals(FontSizeUnitPicker.SelectedItem as string, "pt", StringComparison.Ordinal))
        {
            var nextPt = DipsToPts(d) + delta;
            if (nextPt is < 4 or > 200)
                return;
            ApplyDipsToSelection(PtsToDips(nextPt));
        }
        else
        {
            // px(DIP) 단위: 1씩
            var nextDip = d + delta;
            if (nextDip is < 4 or > 200)
                return;
            ApplyDipsToSelection(nextDip);
        }

        SyncFontPickersFromSelection();
        MemoBody.Focus();
    }

    private void MemoBodyContext_FontSizeUnitPt_Click(object sender, RoutedEventArgs e)
    {
        if (string.Equals(FontSizeUnitPicker.SelectedItem as string, "pt", StringComparison.Ordinal))
        {
            MemoBody.Focus();
            return;
        }

        _updatingFontPickers = true;
        try
        {
            FontSizeUnitPicker.SelectedItem = "pt";
        }
        finally
        {
            _updatingFontPickers = false;
        }

        UpdateFontSizeDisplayForUnitOnly(GetCurrentFontAtCaret().Size);
        MemoBody.Focus();
    }

    private void Button_Click(object sender, RoutedEventArgs e)
    {

    }

    private void MemoBodyContext_FontSizeUnitPx_Click(object sender, RoutedEventArgs e)
    {
        if (string.Equals(FontSizeUnitPicker.SelectedItem as string, "px", StringComparison.Ordinal))
        {
            MemoBody.Focus();
            return;
        }

        _updatingFontPickers = true;
        try
        {
            FontSizeUnitPicker.SelectedItem = "px";
        }
        finally
        {
            _updatingFontPickers = false;
        }

        UpdateFontSizeDisplayForUnitOnly(GetCurrentFontAtCaret().Size);
        MemoBody.Focus();
    }

    #endregion
    #endregion
}
