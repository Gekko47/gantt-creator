using System.Runtime.InteropServices;

namespace GanttCreator.AddIn;

/// <summary>
/// P/Invoke wrappers for the Windows TaskDialog API (comctl32.dll),
/// used to display diagnostic information with a clickable hyperlink
/// to the active log file.
/// </summary>
/// <remarks>
/// TaskDialog is available on Windows Vista and later. The hyperlink
/// flag (TDF_ENABLE_HYPERLINKS) enables the TDN_HYPERLINK notification
/// when the user clicks a link in the dialog content.
/// </remarks>
internal static partial class TaskDialogApi
{
    private const string _comCtl32 = "comctl32.dll";

    /// <summary>
    /// Indicates that the dialog content contains hyperlinks that the
    /// user can click. When a link is clicked, the TDN_HYPERLINK
    /// notification is sent to the callback.
    /// </summary>
    private const int _tdfEnableHyperlinks = 0x00000020;

    /// <summary>
    /// Notification code for a hyperlink click within the dialog content.
    /// </summary>
    private const int _tdnHyperlink = unchecked((int)0xFFFFFD9F);

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate int TaskDialogCallback(IntPtr hwndDlg, int msg, IntPtr wParam, IntPtr lParam, IntPtr referenceData);

    // IDE1006 does not apply to the native TaskDialog entry point: an unused
    // private P/Invoke declaration still needs the exact native name.
#pragma warning disable IDE1006
    [DllImport(_comCtl32, SetLastError = true, CharSet = CharSet.Unicode)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
#pragma warning disable SYSLIB1054 // Keep the explicit SetLastError/CharSet P/Invoke shape used by existing interop guards.
    private static extern int TaskDialog(
        IntPtr hwndOwner,
        IntPtr hInstance,
        string title,
        string content,
        string? mainInstruction,
        int flags,
        string? radioButton1,
        string? verificationText,
        out int buttonId
    );
#pragma warning restore SYSLIB1054 // Keep the explicit SetLastError/CharSet P/Invoke shape used by existing interop guards.
#pragma warning restore IDE1006

    [DllImport(_comCtl32, SetLastError = true, CharSet = CharSet.Unicode)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern int TaskDialogIndirect(
        ref TASKDIALOGCONFIG config,
        out int buttonId,
        out int checkboxState,
        out int verificationState
    );

    [DllImport(_comCtl32, SetLastError = true)]
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    private static extern IntPtr GetDesktopWindow();

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct TASKDIALOGCONFIG
    {
        public int cbSize;
        public IntPtr hwndParent;
        public IntPtr hInstance;
        public int dwFlags;
        public int dwCommonButtons;
        public IntPtr pszWindowTitle;
        public IntPtr pszMainIcon;
        public IntPtr pszMainInstruction;
        public IntPtr pszContent;
        public uint cButtons;
        public IntPtr pButtons;
        public int nDefaultButton;
        public uint cRadioButtons;
        public IntPtr pRadioButtons;
        public int nDefaultRadioButton;
        public IntPtr pszVerificationText;
        public IntPtr pszExpandedInformation;
        public IntPtr pszExpandedControlText;
        public IntPtr pszCollapsedControlText;
        public IntPtr pszFooterIcon;
        public IntPtr pszFooter;
        public IntPtr pCallback;
        public IntPtr lpCallbackData;
        public uint cxWidth;
    }

    private sealed class CallbackHolder
    {
        public required TaskDialogCallback Callback { get; set; }
    }

    /// <summary>
    /// Shows a TaskDialog with the given title, main instruction, content,
    /// and a clickable hyperlink. The hyperlink is detected by the callback
    /// and the provided <paramref name="openFileAction"/> is invoked with
    /// the hyperlink text (expected to be a file path).
    /// </summary>
    /// <param name="title">Dialog title.</param>
    /// <param name="mainInstruction">Main instruction text.</param>
    /// <param name="content">Dialog content; may contain a hyperlink.</param>
    /// <param name="openFileAction">Action invoked with the hyperlink file path when the user clicks it.</param>
    /// <returns>The ID of the button the user clicked, or 0 on failure.</returns>
    public static int ShowWithHyperlink(string title, string mainInstruction, string content, Action<string> openFileAction)
    {
        var owner = GetDesktopWindow();

        var callbackHolder = new CallbackHolder
        {
            Callback = (hwndDlg, msg, wParam, lParam, referenceData) =>
            {
                if (msg == _tdnHyperlink)
                {
                    // The hyperlink text is passed as lParam (pointer to string)
                    var linkText = Marshal.PtrToStringUni(lParam);
                    if (!string.IsNullOrEmpty(linkText))
                    {
                        try
                        {
                            openFileAction(linkText);
                        }
#pragma warning disable CA1031
                        catch
#pragma warning restore CA1031
                        {
                            // Non-fatal: the dialog already displayed; a failed
                            // open attempt does not need to reach the user again.
                        }
                    }
                }
                return 0;
            },
        };

        var config = new TASKDIALOGCONFIG
        {
            cbSize = Marshal.SizeOf<TASKDIALOGCONFIG>(),
            hwndParent = owner,
            dwFlags = _tdfEnableHyperlinks,
            pszWindowTitle = Marshal.StringToCoTaskMemUni(title),
            pszMainInstruction = Marshal.StringToCoTaskMemUni(mainInstruction),
            pszContent = Marshal.StringToCoTaskMemUni(content),
            pCallback = Marshal.GetFunctionPointerForDelegate(callbackHolder.Callback),
        };

        try
        {
            var hresult = TaskDialogIndirect(ref config, out int buttonId, out _, out _);
            if (hresult < 0)
            {
                // HRESULT failure -- the dialog did not display. Don't report
                // the HRESULT as a button ID.
                return 0;
            }
            return buttonId;
        }
        finally
        {
            Marshal.FreeCoTaskMem(config.pszWindowTitle);
            Marshal.FreeCoTaskMem(config.pszMainInstruction);
            Marshal.FreeCoTaskMem(config.pszContent);
        }
    }
}
