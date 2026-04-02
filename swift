Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
public class LargeIcon {
    [DllImport("Shell32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);
    
    [DllImport("shell32.dll", CharSet=CharSet.Auto)]
    public static extern int ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, int nIcons);
    
    [DllImport("user32.dll")]
    public static extern bool DestroyIcon(IntPtr hIcon);
    
    [DllImport("Shell32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr SHExtractIcons(string pszFileName, int nIconIndex, int cxIcon, int cyIcon, IntPtr[] phicon, IntPtr[] piconid, int nIcons, int flags);
    
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
    public struct SHFILEINFO {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst=80)]
        public string szTypeName;
    }
}
"@

$iconIndex = 54
$size = 256
$hIcons = New-Object IntPtr[] 1
$hIds = New-Object IntPtr[] 1

[LargeIcon]::SHExtractIcons(
    "C:\Windows\System32\imageres.dll",
    $iconIndex, $size, $size,
    $hIcons, $hIds, 1, 0
)

$icon = [System.Drawing.Icon]::FromHandle($hIcons[0])
$bmp = New-Object System.Drawing.Bitmap($icon.ToBitmap(), 256, 256)
$bmp.Save("C:\temp\lock-icon-256.png", [System.Drawing.Imaging.ImageFormat]::Png)
[LargeIcon]::DestroyIcon($hIcons[0])
