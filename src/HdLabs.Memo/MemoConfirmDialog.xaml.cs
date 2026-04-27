using System.Windows;
using System.Windows.Input;

namespace HdLabs.Memo;

public partial class MemoConfirmDialog : Window
{
    public MemoConfirmDialog(string title, string message, string okText = "확인", string cancelText = "취소")
    {
        InitializeComponent();
        TitleText.Text = string.IsNullOrWhiteSpace(title) ? "확인" : title.Trim();
        MessageText.Text = message ?? "";
        OkButton.Content = okText;
        CancelButton.Content = cancelText;
    }

    public static bool ShowDialog(Window owner, string title, string message, string okText = "삭제", string cancelText = "취소")
    {
        var d = new MemoConfirmDialog(title, message, okText, cancelText) { Owner = owner };
        return d.ShowDialog() == true;
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            e.Handled = true;
            DialogResult = false;
            Close();
        }
    }

    private void DragArea_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left)
            return;
        try { DragMove(); } catch { /* ignore */ }
    }
}

