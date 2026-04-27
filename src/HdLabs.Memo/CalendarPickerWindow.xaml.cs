using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using HdLabs.Memo.ViewModels;

namespace HdLabs.Memo;

public partial class CalendarPickerWindow : Window
{
    private readonly MainViewModel _vm;

    public CalendarPickerWindow(MainViewModel vm)
    {
        InitializeComponent();
        _vm = vm;
        Loaded += (_, _) =>
        {
            if (_vm.EditorScheduleDate is { } s)
            {
                DayCalendar.SelectedDate = s.Date;
                DayCalendar.DisplayDate = s;
            }
            else
            {
                DayCalendar.SelectedDate = DateTime.Today;
                DayCalendar.DisplayDate = DateTime.Today;
            }

            UpdateStatus();
        };
    }

    private void UpdateStatus()
    {
        if (DayCalendar.SelectedDate is { } d)
            StatusText.Text = "이 메모에 붙은 날짜: " + d.ToString("yyyy년 M월 d일 (ddd)");
        else
            StatusText.Text = "날짜를 선택하거나 '날짜 제거'로 지울 수 있습니다.";
    }

    private void DayCalendar_OnSelectedDatesChanged(object? sender, SelectionChangedEventArgs e) => UpdateStatus();

    private void OpenWindowsCalendar_Click(object sender, RoutedEventArgs e) => TryOpenWindowsCalendar();

    private void ClearDate_Click(object sender, RoutedEventArgs e)
    {
        _vm.EditorScheduleDate = null;
        DayCalendar.SelectedDate = null;
        UpdateStatus();
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        if (DayCalendar.SelectedDate is { } d)
            _vm.EditorScheduleDate = d.Date;
        else
            _vm.EditorScheduleDate = null;
        DialogResult = true;
        Close();
    }

    private static void TryOpenWindowsCalendar()
    {
        try
        {
            Process.Start(new ProcessStartInfo { FileName = "outlookcal:", UseShellExecute = true });
            return;
        }
        catch
        {
            // ignore
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = "shell:AppsFolder\\Microsoft.Windows.Calendar_8wekyb3d8bbwe!App",
                UseShellExecute = true
            });
        }
        catch
        {
            // ignore
        }
    }
}
