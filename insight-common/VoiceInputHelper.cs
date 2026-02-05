using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace InsightCommon.AI;

/// <summary>
/// Windows音声入力（Win+H）を起動するヘルパー
/// </summary>
public static class VoiceInputHelper
{
    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    private const byte VK_LWIN = 0x5B;  // Left Windows key
    private const byte VK_H = 0x48;      // H key
    private const uint KEYEVENTF_KEYUP = 0x0002;

    /// <summary>
    /// Windows音声入力を起動（Win+H キーストロークをシミュレート）
    /// </summary>
    public static void ActivateWindowsVoiceTyping()
    {
        // Press Win+H
        keybd_event(VK_LWIN, 0, 0, UIntPtr.Zero);
        keybd_event(VK_H, 0, 0, UIntPtr.Zero);

        // Release H, then Win
        keybd_event(VK_H, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
