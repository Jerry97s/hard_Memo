namespace HdLabs.Memo.Models;

public sealed class MemoDataRoot
{
    public List<MemoItem> Items { get; set; } = new();

    public MemoUserSettings Settings { get; set; } = new();
}

public sealed class MemoUserSettings
{
    public bool WindowTopmost { get; set; }

    /// <summary>카드 본문 기본 틴트 (ARGB #RRGGBB).</summary>
    public string? CardTintHex { get; set; }

    /// <summary>전역 단축키: 창을 앞으로 가져오기. 예) Ctrl+Alt+M</summary>
    public string BringToFrontHotkey { get; set; } = "Ctrl+Alt+M";

    /// <summary>알림 창에서 확인한 시각(로컬) 이력, 최대 5. ISO 8601 offset 문자열, 최신이 앞.</summary>
    public List<string> ReminderTimeHistory { get; set; } = new();

    /// <summary>
    /// 알림 이력(상태 포함) v2. 최대 5, 최신이 앞.
    /// - Enabled=true: TargetAtIso에 "울릴 시각" 저장
    /// - Enabled=false: TargetAtIso는 null
    /// </summary>
    public List<ReminderHistoryEntry> ReminderHistoryV2 { get; set; } = new();
}

public sealed class ReminderHistoryEntry
{
    public bool Enabled { get; set; }

    /// <summary>사용자가 설정한 알림 시각(ISO 8601 offset). 끔이면 null.</summary>
    public string? TargetAtIso { get; set; }

    /// <summary>사용자가 켜거나 끈 시각(ISO 8601 offset).</summary>
    public string ChangedAtIso { get; set; } = DateTimeOffset.Now.ToString("O");
}
