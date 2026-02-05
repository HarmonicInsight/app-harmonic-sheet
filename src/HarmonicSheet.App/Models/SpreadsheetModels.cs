using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace HarmonicSheet.Models;

/// <summary>
/// スプレッドシートの1行を表します
/// </summary>
public partial class SpreadsheetRow : ObservableObject
{
    [ObservableProperty]
    private int _rowNumber;

    [ObservableProperty]
    private ObservableCollection<SpreadsheetCell> _cells = new();
}

/// <summary>
/// スプレッドシートの1セルを表します
/// </summary>
public partial class SpreadsheetCell : ObservableObject
{
    [ObservableProperty]
    private char _column;

    [ObservableProperty]
    private int _row;

    [ObservableProperty]
    private string _value = string.Empty;

    [ObservableProperty]
    private string? _formula;

    [ObservableProperty]
    private string? _displayValue;

    /// <summary>
    /// セルのアドレス（例: A1, B2）
    /// </summary>
    public string Address => $"{Column}{Row}";
}

/// <summary>
/// セルの変更内容
/// </summary>
public class CellChange
{
    public char Column { get; set; }
    public int Row { get; set; }
    public string Value { get; set; } = string.Empty;
    public string? Formula { get; set; }
}

/// <summary>
/// Claude APIからのスプレッドシート操作結果
/// </summary>
public class SpreadsheetCommandResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<CellChange> Changes { get; set; } = new();
}
