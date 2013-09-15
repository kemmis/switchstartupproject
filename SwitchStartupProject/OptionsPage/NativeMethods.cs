using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace LucidConcepts.SwitchStartupProject.OptionsPage
{
    /// <remarks>
    /// This is source code from the Visual Studio 2012 SDK. It is included here to use the same functionality with the Visual Studio 2010 SDK.
    /// The source code was provided by Ryan Molden from Microsoft in the following Visual Studio development forum discussion page: (no license terms were indicated)
    /// http://social.msdn.microsoft.com/Forums/vstudio/en-US/6af9718e-8778-4233-875d-b38c03e9f4ba/vs-plugin-unable-to-access-wpf-user-control-in-option-dialog
    /// </remarks>
    internal static class NativeMethods
    {
        public const int WM_GETDLGCODE = 0x0087;
        public const int WM_SETFOCUS = 0x0007;

        public const int DLGC_WANTARROWS = 0x0001;
        public const int DLGC_WANTTAB = 0x0002;
        public const int DLGC_WANTCHARS = 0x0080;
        public const int DLGC_WANTALLKEYS = 0x0004;

        public const int GA_ROOT = 2;

        [DllImport("user32.dll", ExactSpelling = true)]
        internal static extern IntPtr GetAncestor(IntPtr hWnd, int flags);

        [DllImport("user32.dll", ExactSpelling = true)]
        internal static extern IntPtr GetNextDlgTabItem(IntPtr hDlg, IntPtr hCtl, [MarshalAs(UnmanagedType.Bool)] bool bPrevious);

        [DllImport("user32.dll")]
        public static extern void SetFocus(IntPtr hwnd);
    }
}
