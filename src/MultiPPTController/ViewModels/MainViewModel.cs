using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MultiPPTController.Models;
using MultiPPTController.Services;

namespace MultiPPTController.ViewModels;

/// <summary>
/// 主控制台状态与命令。
/// </summary>
public sealed partial class MainViewModel : ObservableObject, IDisposable
{
    private readonly MonitorService _monitors = new();
    private readonly PlaybackOrchestrator _orchestrator = new();
    private readonly ThumbnailService _thumbnails = new();
    private bool _syncingScreenOptions;
    /// <summary>
    /// 初始化屏幕列表与 WPS 探测。
    /// </summary>
    public MainViewModel()
    {
        RefreshScreens();
        WpsReady = PlaybackOrchestrator.IsWpsAvailable();
        EngineText = WpsReady
            ? "PPTX：主屏 Web · 副屏 WPS　PDF：双屏 Web"
            : "未检测到 WPS（PPTX 仅主屏；PDF 可双屏）";
        Decks.CollectionChanged += (_, _) =>
        {
            HasDecks = Decks.Count > 0;
            StartCommand.NotifyCanExecuteChanged();
        };
    }

    /// <summary>已导入的演示文稿。</summary>
    public ObservableCollection<DeckItem> Decks { get; } = [];

    /// <summary>当前已连接屏幕。</summary>
    public ObservableCollection<DisplayScreen> Screens { get; } = [];

    /// <summary>右侧屏幕编排卡片。</summary>
    public ObservableCollection<ScreenCardViewModel> ScreenCards { get; } = [];

    /// <summary>WPS COM 是否可创建。</summary>
    [ObservableProperty]
    private bool _wpsReady;

    /// <summary>引擎状态短文案。</summary>
    [ObservableProperty]
    private string _engineText = string.Empty;

    /// <summary>底部提示。</summary>
    [ObservableProperty]
    private string _banner = "导入 PPT 或 PDF，并为每份指定一块或多块屏幕。";

    /// <summary>是否正在放映。</summary>
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(StartCommand))]
    [NotifyCanExecuteChangedFor(nameof(StopCommand))]
    [NotifyCanExecuteChangedFor(nameof(NextCommand))]
    [NotifyCanExecuteChangedFor(nameof(PreviousCommand))]
    [NotifyCanExecuteChangedFor(nameof(ImportCommand))]
    [NotifyCanExecuteChangedFor(nameof(TogglePauseCommand))]
    private bool _isPlaying;

    /// <summary>放映中是否暂停翻页。</summary>
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(NextCommand))]
    [NotifyCanExecuteChangedFor(nameof(PreviousCommand))]
    private bool _isPaused;

    /// <summary>片库是否已有文稿。</summary>
    [ObservableProperty]
    private bool _hasDecks;

    /// <summary>命令是否忙碌（开播/停播过程）。</summary>
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(StartCommand))]
    [NotifyCanExecuteChangedFor(nameof(StopCommand))]
    [NotifyCanExecuteChangedFor(nameof(NextCommand))]
    [NotifyCanExecuteChangedFor(nameof(PreviousCommand))]
    private bool _isBusy;

    /// <summary>屏幕数量说明。</summary>
    public string ScreenSummary => Screens.Count <= 1
        ? $"当前 {Screens.Count} 块屏幕 · 请用「扩展」模式连接更多显示器"
        : $"当前 {Screens.Count} 块屏幕 · 扩展桌面";

    /// <summary>
    /// 从文件对话框结果导入多份文稿。
    /// </summary>
    /// <param name="paths">绝对路径。</param>
    [RelayCommand(CanExecute = nameof(CanEditLibrary))]
    private async Task Import(IEnumerable<string>? paths)
    {
        if (paths is null)
        {
            return;
        }

        foreach (var path in paths.Where(File.Exists))
        {
            if (Decks.Any(item => string.Equals(item.FilePath, path, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var deck = new DeckItem(path);
            deck.PropertyChanged += OnDeckPropertyChanged;
            AutoAssignScreen(deck);
            ApplyScreenOptions(deck);
            Decks.Add(deck);
            await LoadCoverAsync(deck).ConfigureAwait(true);
        }

        SyncScreenOptions();
        Banner = Decks.Count == 0 ? "尚未导入演示文稿。" : $"已导入 {Decks.Count} 份文稿。";
        StartCommand.NotifyCanExecuteChanged();
    }

    /// <summary>
    /// 从片库移除一份文稿。
    /// </summary>
    /// <param name="deck">目标条目。</param>
    [RelayCommand(CanExecute = nameof(CanEditLibrary))]
    private void Remove(DeckItem? deck)
    {
        if (deck is null)
        {
            return;
        }

        deck.PropertyChanged -= OnDeckPropertyChanged;
        Decks.Remove(deck);
        SyncScreenOptions();
        StartCommand.NotifyCanExecuteChanged();
    }

    /// <summary>
    /// 重新枚举显示器。
    /// </summary>
    [RelayCommand]
    private void RefreshScreens()
    {
        Screens.Clear();
        foreach (var screen in _monitors.GetScreens())
        {
            Screens.Add(screen);
        }

        OnPropertyChanged(nameof(ScreenSummary));
        foreach (var deck in Decks)
        {
            var removed = false;
            for (var i = deck.AssignedDeviceNames.Count - 1; i >= 0; i--)
            {
                if (Screens.All(screen => screen.DeviceName != deck.AssignedDeviceNames[i]))
                {
                    deck.AssignedDeviceNames.RemoveAt(i);
                    removed = true;
                }
            }

            if (removed && !deck.HasAssignment)
            {
                deck.Status = "屏幕已断开，请重新分配";
            }
        }

        SyncScreenOptions();
    }

    /// <summary>
    /// 开始各屏同步放映。
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanStart))]
    private async Task StartAsync()
    {
        var assignments = BuildAssignments();
        if (assignments.Count == 0)
        {
            Banner = "请至少为一份文稿分配屏幕。";
            return;
        }

        if (assignments.GroupBy(item => item.Screen.DeviceName).Any(group => group.Count() > 1))
        {
            Banner = "同一块屏幕只能绑定一份文稿。";
            return;
        }

        IsBusy = true;
        Banner = "正在启动放映…";
        try
        {
            await _orchestrator.StartAsync(assignments).ConfigureAwait(true);
            IsPlaying = true;
            IsPaused = false;
            Banner = "放映中 · 方向键 / 空格同步下一步，Esc 结束。";
            _ = RelayoutAfterChromeSettlesAsync();
        }
        catch (Exception ex)
        {
            IsPlaying = false;
            Banner = $"开播失败：{ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>
    /// 结束全部放映。
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanStop))]
    private async Task StopAsync()
    {
        IsBusy = true;
        try
        {
            await _orchestrator.StopAsync().ConfigureAwait(true);
            Banner = "已结束放映。";
        }
        catch (Exception ex)
        {
            Banner = $"结束时出现问题：{ex.Message}";
        }
        finally
        {
            IsPlaying = false;
            IsPaused = false;
            IsBusy = false;
        }
    }

    /// <summary>
    /// 暂停或继续翻页（不结束放映）。
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanTogglePause))]
    private async Task TogglePause()
    {
        if (IsPaused)
        {
            IsPaused = false;
            Banner = "放映中 · 方向键 / 空格同步下一步，Esc 结束。";
            return;
        }

        var note = await _orchestrator.NextAsync().ConfigureAwait(true);
        Banner = string.IsNullOrWhiteSpace(note)
            ? "放映中 · 方向键 / 空格同步下一步，Esc 结束。"
            : note;
    }

    /// <summary>
    /// 各屏同步下一步动画。
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanControlPlayback))]
    private async Task NextAsync()
    {
        var note = await _orchestrator.NextAsync().ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(note))
        {
            Banner = note;
        }
    }

    /// <summary>
    /// 各屏同步上一步动画。
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanControlPlayback))]
    private async Task PreviousAsync()
    {
        var note = await _orchestrator.PreviousAsync().ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(note))
        {
            Banner = note;
        }
    }

    /// <summary>
    /// 释放编排器。
    /// </summary>
    public void Dispose()
    {
        if (IsPlaying)
        {
            try
            {
                _orchestrator.StopAsync().GetAwaiter().GetResult();
            }
            catch
            {
                // 退出时忽略
            }
        }

        _orchestrator.Dispose();
    }

    /// <summary>
    /// 主窗口隐藏后分几次重铺放映窗，盖住副屏任务栏并纠正主屏分辨率残留。
    /// </summary>
    /// <returns>完成任务。</returns>
    private async Task RelayoutAfterChromeSettlesAsync()
    {
        foreach (var delayMs in new[] { 800 })
        {
            await Task.Delay(delayMs).ConfigureAwait(true);
            if (!IsPlaying)
            {
                return;
            }

            try
            {
                await _orchestrator.RelayoutAllAsync().ConfigureAwait(true);
            }
            catch
            {
                // 放映已结束时忽略
            }
        }
    }

    /// <summary>
    /// 未放映时才允许改片库。
    /// </summary>
    /// <returns>可编辑则为 true。</returns>
    private bool CanEditLibrary() => !IsPlaying && !IsBusy;

    /// <summary>
    /// 是否可以开播：未在放映，且至少一份 PPT 已指定屏幕；副屏或 .dps 仍需 WPS。
    /// </summary>
    /// <returns>可开播则为 true。</returns>
    private bool CanStart() =>
        !IsPlaying && !IsBusy && HasAssignedPlayback() && (!NeedsWpsEngine() || WpsReady);

    /// <summary>
    /// 是否必须启动 WPS：副屏 PPTX，或主屏文稿需 WPS 另存。PDF 不走 WPS。
    /// </summary>
    /// <returns>需要 WPS 则为 true。</returns>
    private bool NeedsWpsEngine()
    {
        foreach (var deck in Decks.Where(item => item.HasAssignment))
        {
            if (DeckConvertService.IsPdf(deck.FilePath))
            {
                continue;
            }

            if (!DeckConvertService.IsWebNative(deck.FilePath))
            {
                return true;
            }

            foreach (var name in deck.AssignedDeviceNames)
            {
                var screen = Screens.FirstOrDefault(item => item.DeviceName == name);
                if (screen is { IsPrimary: false })
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// 是否已有文稿绑定到屏幕。
    /// </summary>
    /// <returns>至少一份已指定屏幕则为 true。</returns>
    private bool HasAssignedPlayback() =>
        Decks.Any(item => item.HasAssignment);

    /// <summary>
    /// 是否可以翻页或结束。
    /// </summary>
    /// <returns>放映中且不忙碌则为 true。</returns>
    private bool CanControlPlayback() => IsPlaying && !IsBusy && !IsPaused;

    /// <summary>
    /// 是否可以结束放映。
    /// </summary>
    /// <returns>放映中则为 true。</returns>
    private bool CanStop() => IsPlaying && !IsBusy;

    /// <summary>
    /// 是否可以暂停或继续。
    /// </summary>
    /// <returns>放映中则为 true。</returns>
    private bool CanTogglePause() => IsPlaying && !IsBusy;

    /// <summary>
    /// 为新导入文稿自动分到第一块空闲屏幕。
    /// </summary>
    /// <param name="deck">新条目。</param>
    private void AutoAssignScreen(DeckItem deck)
    {
        var used = Decks
            .SelectMany(item => item.AssignedDeviceNames)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var free = Screens.FirstOrDefault(screen => !used.Contains(screen.DeviceName));
        if (free is null)
        {
            deck.Status = "待分配屏幕";
            deck.RefreshAssignmentText(Screens);
            return;
        }

        deck.AssignedDeviceNames.Add(free.DeviceName);
        deck.RefreshAssignmentText(Screens);
    }

    /// <summary>
    /// 后台提取封面并刷新屏幕卡片。
    /// </summary>
    /// <param name="deck">片库条目。</param>
    /// <returns>完成任务。</returns>
    private async Task LoadCoverAsync(DeckItem deck)
    {
        var cover = await Task.Run(() => _thumbnails.Create(deck.FilePath)).ConfigureAwait(true);
        deck.Cover = cover;
        RebuildScreenCards();
    }

    /// <summary>
    /// 文稿绑定或封面变化时同步右侧卡片。
    /// </summary>
    /// <param name="sender">片库条目。</param>
    /// <param name="e">属性参数。</param>
    private void OnDeckPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(DeckItem.Cover))
        {
            RebuildScreenCards();
        }
    }

    /// <summary>
    /// 判断该设备名是否已被其他文稿占用。
    /// </summary>
    /// <param name="owner">当前文稿。</param>
    /// <param name="deviceName">拟绑定的设备名。</param>
    /// <returns>被其他文稿占用则为 true。</returns>
    private bool IsScreenTakenByOther(DeckItem owner, string? deviceName)
    {
        if (string.IsNullOrWhiteSpace(deviceName))
        {
            return false;
        }

        return Decks.Any(other =>
            !ReferenceEquals(other, owner) &&
            other.IsAssignedTo(deviceName));
    }

    /// <summary>
    /// 用户勾选或取消某块屏幕时更新绑定，并拒绝占用他人的屏幕。
    /// </summary>
    /// <param name="deck">当前文稿。</param>
    /// <param name="option">被操作的屏幕选项。</param>
    /// <param name="selected">是否勾选。</param>
    private void OnScreenOptionToggled(DeckItem deck, ScreenOption option, bool selected)
    {
        if (_syncingScreenOptions)
        {
            return;
        }

        if (selected)
        {
            if (IsScreenTakenByOther(deck, option.DeviceName))
            {
                option.SetSelectedSilent(false);
                return;
            }

            if (!deck.IsAssignedTo(option.DeviceName))
            {
                deck.AssignedDeviceNames.Add(option.DeviceName);
            }
        }
        else
        {
            for (var i = deck.AssignedDeviceNames.Count - 1; i >= 0; i--)
            {
                if (string.Equals(deck.AssignedDeviceNames[i], option.DeviceName, StringComparison.OrdinalIgnoreCase))
                {
                    deck.AssignedDeviceNames.RemoveAt(i);
                }
            }
        }

        SyncScreenOptions();
    }

    /// <summary>
    /// 去掉重复绑定，并按占用情况刷新每份文稿的下拉选项。
    /// </summary>
    private void SyncScreenOptions()
    {
        if (_syncingScreenOptions)
        {
            return;
        }

        _syncingScreenOptions = true;
        try
        {
            ResolveDuplicateAssignments();
            RefreshScreenOptions();
            RebuildScreenCards();
            StartCommand.NotifyCanExecuteChanged();
        }
        finally
        {
            _syncingScreenOptions = false;
        }
    }

    /// <summary>
    /// 同一屏幕若被多份文稿绑定，只保留片库中最先的一份。
    /// </summary>
    private void ResolveDuplicateAssignments()
    {
        var claimed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var deck in Decks)
        {
            for (var i = deck.AssignedDeviceNames.Count - 1; i >= 0; i--)
            {
                if (!claimed.Add(deck.AssignedDeviceNames[i]))
                {
                    deck.AssignedDeviceNames.RemoveAt(i);
                }
            }
        }
    }

    /// <summary>
    /// 同步每份文稿的屏幕选项列表，并标记被其他文稿占用的项为不可选。
    /// </summary>
    private void RefreshScreenOptions()
    {
        foreach (var deck in Decks)
        {
            ApplyScreenOptions(deck);
        }
    }

    /// <summary>
    /// 按当前显示器列表填充一份文稿的下拉选项，并写入占用禁用状态。
    /// </summary>
    /// <param name="deck">目标文稿。</param>
    private void ApplyScreenOptions(DeckItem deck)
    {
        for (var i = deck.ScreenOptions.Count - 1; i >= 0; i--)
        {
            if (Screens.All(screen => screen.DeviceName != deck.ScreenOptions[i].DeviceName))
            {
                deck.ScreenOptions.RemoveAt(i);
            }
        }

        for (var index = 0; index < Screens.Count; index++)
        {
            var screen = Screens[index];
            var option = deck.ScreenOptions.FirstOrDefault(item => item.DeviceName == screen.DeviceName);
            if (option is null)
            {
                option = new ScreenOption(screen);
                deck.ScreenOptions.Insert(Math.Min(index, deck.ScreenOptions.Count), option);
            }

            option.SelectionChanged = (item, selected) => OnScreenOptionToggled(deck, item, selected);
            option.IsAvailable = !IsScreenTakenByOther(deck, screen.DeviceName);
            option.SetSelectedSilent(deck.IsAssignedTo(screen.DeviceName));
        }

        deck.RefreshAssignmentText(Screens);
    }

    /// <summary>
    /// 按当前屏幕与片库绑定重绘编排卡片。
    /// </summary>
    private void RebuildScreenCards()
    {
        foreach (var screen in Screens)
        {
            var card = ScreenCards.FirstOrDefault(item => item.Screen.DeviceName == screen.DeviceName);
            if (card is null)
            {
                card = new ScreenCardViewModel(screen);
                ScreenCards.Add(card);
            }

            var deck = Decks.FirstOrDefault(item => item.IsAssignedTo(screen.DeviceName));
            card.Cover = deck?.Cover;
        }

        for (var i = ScreenCards.Count - 1; i >= 0; i--)
        {
            if (Screens.All(screen => screen.DeviceName != ScreenCards[i].Screen.DeviceName))
            {
                ScreenCards.RemoveAt(i);
            }
        }
    }

    /// <summary>
    /// 收集已绑定且文件仍存在的配对。
    /// </summary>
    /// <returns>开播清单。</returns>
    private List<(DeckItem Deck, DisplayScreen Screen)> BuildAssignments()
    {
        var result = new List<(DeckItem, DisplayScreen)>();
        foreach (var deck in Decks)
        {
            if (!deck.HasAssignment)
            {
                continue;
            }

            if (!File.Exists(deck.FilePath))
            {
                deck.Status = "文件已丢失";
                continue;
            }

            foreach (var deviceName in deck.AssignedDeviceNames)
            {
                var screen = Screens.FirstOrDefault(item => item.DeviceName == deviceName);
                if (screen is null)
                {
                    continue;
                }

                result.Add((deck, screen));
            }
        }

        return result;
    }
}
