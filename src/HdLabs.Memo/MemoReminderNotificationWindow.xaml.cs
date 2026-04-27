using System.Windows;
using System.Windows.Input;

namespace HdLabs.Memo;

public partial class MemoReminderNotificationWindow : Window
{
    public MemoReminderNotificationWindow(string memoTitleForDisplay)
    {
        InitializeComponent();
        var t = string.IsNullOrWhiteSpace(memoTitleForDisplay) ? "제목 없음" : memoTitleForDisplay.Trim();
        TitleRun.Text = t;
    }

    private void Ok_Click(object sender, RoutedEventArgs e) => Close();

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            e.Handled = true;
            Close();
        }
    }

    private void DragArea_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left)
            return;
        try
        {
            DragMove();
        }
        catch
        {
            // ignore
        }
    }
}
