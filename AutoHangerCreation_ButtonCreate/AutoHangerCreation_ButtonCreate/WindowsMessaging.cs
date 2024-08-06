using System.Runtime.InteropServices;
using Microsoft.VisualBasic.CompilerServices;


[StandardModule]
internal sealed class WindowsMessaging
{
    public const int WM_KEYDOWN = 256;
	[DllImport("User32.dll")]
	public static extern int SendMessage(int hWnd, int Msg, int wParam, int lParam);

	[DllImport("User32.dll")]
	public static extern int PostMessage(int hWnd, int Msg, int wParam, int lParam);

	public static int SendWindowsMessage(int hWnd, int Msg, int wParam, int lParam)
	{
		int result = 0;
		if (hWnd > 0)
		{
			result = SendMessage(hWnd, Msg, wParam, lParam);
		}
		return result;
	}

	public static int PostWindowsMessage(int hWnd, int Msg, int wParam, int lParam)
	{
		int result = 0;
		if (hWnd > 0)
		{
			result = PostMessage(hWnd, Msg, wParam, lParam);
		}
		return result;
	}
}
