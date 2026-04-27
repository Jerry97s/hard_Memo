using System.Runtime.InteropServices;
using System.Windows.Input;

namespace HdLabs.Memo.Helpers;

public sealed class GlobalHotkey : IDisposable
{
    private readonly IntPtr _hwnd;
    private readonly int _id;
    private bool _registered;

    public const int WmHotkey = 0x0312;

    public GlobalHotkey(IntPtr hwnd, int id)
    {
        _hwnd = hwnd;
        _id = id;
    }

    [Flags]
    public enum Modifiers : uint
    {
        None = 0,
        Alt = 0x0001,
        Control = 0x0002,
        Shift = 0x0004,
        Win = 0x0008,
        NoRepeat = 0x4000,
    }

    public bool Register(Modifiers mods, Key key)
    {
        Unregister();

        var vk = KeyInterop.VirtualKeyFromKey(key);
        _registered = RegisterHotKey(_hwnd, _id, (uint)mods, (uint)vk);
        return _registered;
    }

    public void Unregister()
    {
        if (!_registered)
            return;
        UnregisterHotKey(_hwnd, _id);
        _registered = false;
    }

    public void Dispose() => Unregister();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}

