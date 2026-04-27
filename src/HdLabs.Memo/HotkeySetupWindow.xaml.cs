using System.Text;
using System.Windows;
using System.Windows.Input;
using HdLabs.Memo.ViewModels;

namespace HdLabs.Memo;

public partial class HotkeySetupWindow : Window
{
    private readonly MainViewModel _vm;

    public HotkeySetupWindow(MainViewModel vm)
    {
        InitializeComponent();
        _vm = vm;
        CurrentHotkeyText.Text = _vm.BringToFrontHotkey;
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            Close();
            return;
        }

        var key = (e.Key == Key.System) ? e.SystemKey : e.Key;
        if (key is Key.LeftCtrl or Key.RightCtrl or Key.LeftAlt or Key.RightAlt or Key.LeftShift or Key.RightShift
            or Key.LWin or Key.RWin)
            return;

        var mods = Keyboard.Modifiers;
        if (mods == ModifierKeys.None)
        {
            HintText.Text = "Ctrl/Alt/Shift/Win 중 하나 이상을 포함해주세요.";
            e.Handled = true;
            return;
        }

        var s = BuildHotkeyString(mods, key);
        _vm.BringToFrontHotkey = s;
        CurrentHotkeyText.Text = s;
        HintText.Text = "적용되었습니다. (바로 사용 가능)";
        e.Handled = true;
    }

    private static string BuildHotkeyString(ModifierKeys mods, Key key)
    {
        var sb = new StringBuilder();
        void Add(string t)
        {
            if (sb.Length > 0)
                sb.Append('+');
            sb.Append(t);
        }

        if (mods.HasFlag(ModifierKeys.Control)) Add("Ctrl");
        if (mods.HasFlag(ModifierKeys.Alt)) Add("Alt");
        if (mods.HasFlag(ModifierKeys.Shift)) Add("Shift");
        if (mods.HasFlag(ModifierKeys.Windows)) Add("Win");
        Add(key.ToString());
        return sb.ToString();
    }

    private void Reset_Click(object sender, RoutedEventArgs e)
    {
        _vm.BringToFrontHotkey = "Ctrl+Alt+M";
        CurrentHotkeyText.Text = _vm.BringToFrontHotkey;
        HintText.Text = "기본값으로 되돌렸습니다.";
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}

