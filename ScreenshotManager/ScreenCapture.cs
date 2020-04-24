using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace ScreenshotManager
{
    public class ScreenCapture
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetDesktopWindow();

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);

        public static void CaptureDesktop(string path)
        {
            int screenLeft = int.MaxValue;
            int screenTop = int.MaxValue;
            int screenRight = int.MinValue;
            int screenBottom = int.MinValue;

            int deviceIndex = 0;
            while (true)
            {
                NativeUtilities.DisplayDevice deviceData = new NativeUtilities.DisplayDevice { cb = Marshal.SizeOf(typeof(NativeUtilities.DisplayDevice)) };
                if (NativeUtilities.EnumDisplayDevices(null, deviceIndex, ref deviceData, 0) != 0)
                {
                    NativeUtilities.DEVMODE devMode = new NativeUtilities.DEVMODE();
                    if (NativeUtilities.EnumDisplaySettings(deviceData.DeviceName, NativeUtilities.ENUM_CURRENT_SETTINGS, ref devMode))
                    {
                        screenLeft = Math.Min(screenLeft, devMode.dmPositionX);
                        screenTop = Math.Min(screenTop, devMode.dmPositionY);
                        screenRight = Math.Max(screenRight, devMode.dmPositionX + devMode.dmPelsWidth);
                        screenBottom = Math.Max(screenBottom, devMode.dmPositionY + devMode.dmPelsHeight);
                    }
                    deviceIndex++;
                }
                else
                    break;
            }

            using (Bitmap bmp = new Bitmap(screenRight - screenLeft, screenBottom - screenTop))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                    g.CopyFromScreen(screenLeft, screenTop, 0, 0, bmp.Size);

                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            }
        }

        public static void CaptureActiveWindow(string path)
        {
            CaptureWindow(GetForegroundWindow(), path);
        }

        
        public static void CaptureWindow(IntPtr handle, string path)
        {
            Rect rect = new Rect();
            GetWindowRect(handle, ref rect);
            Rectangle bounds = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
            Bitmap bmp = new Bitmap(bounds.Width, bounds.Height);

            using (Graphics graphics = Graphics.FromImage(bmp))
            {
                graphics.CopyFromScreen(new Point(bounds.Left, bounds.Top), Point.Empty, bounds.Size);
            }

            bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        }

        class NativeUtilities
        {
            [Flags()]
            public enum DisplayDeviceStateFlags : int
            {
                AttachedToDesktop = 0x1,
                MultiDriver = 0x2,
                PrimaryDevice = 0x4,
                MirroringDriver = 0x8,
                VGACompatible = 0x16,
                Removable = 0x20,
                ModesPruned = 0x8000000,
                Remote = 0x4000000,
                Disconnect = 0x2000000
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
            public struct DisplayDevice
            {
                [MarshalAs(UnmanagedType.U4)]
                public int cb;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
                public string DeviceName;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
                public string DeviceString;
                [MarshalAs(UnmanagedType.U4)]
                public DisplayDeviceStateFlags StateFlags;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
                public string DeviceID;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
                public string DeviceKey;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct DEVMODE
            {
                private const int CCHDEVICENAME = 0x20;
                private const int CCHFORMNAME = 0x20;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
                public string dmDeviceName;
                public short dmSpecVersion;
                public short dmDriverVersion;
                public short dmSize;
                public short dmDriverExtra;
                public int dmFields;
                public int dmPositionX;
                public int dmPositionY;
                public ScreenOrientation dmDisplayOrientation;
                public int dmDisplayFixedOutput;
                public short dmColor;
                public short dmDuplex;
                public short dmYResolution;
                public short dmTTOption;
                public short dmCollate;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
                public string dmFormName;
                public short dmLogPixels;
                public int dmBitsPerPel;
                public int dmPelsWidth;
                public int dmPelsHeight;
                public int dmDisplayFlags;
                public int dmDisplayFrequency;
                public int dmICMMethod;
                public int dmICMIntent;
                public int dmMediaType;
                public int dmDitherType;
                public int dmReserved1;
                public int dmReserved2;
                public int dmPanningWidth;
                public int dmPanningHeight;
            }

            [DllImport("user32.dll")]
            public static extern bool EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);

            public const int ENUM_CURRENT_SETTINGS = -1;
            const int ENUM_REGISTRY_SETTINGS = -2;

            [DllImport("User32.dll")]
            public static extern int EnumDisplayDevices(string lpDevice, int iDevNum, ref DisplayDevice lpDisplayDevice, int dwFlags);
        }

    }
}
