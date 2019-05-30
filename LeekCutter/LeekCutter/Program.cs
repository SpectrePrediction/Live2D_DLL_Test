using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Drawing;
using OpenCvSharp;
using System.Threading;
using Tesseract;

namespace LeekCutter
{
    public struct POINT
    {
        public int x;
        public int y;
    }
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
    public static class APIMethod
    {
        [DllImport("user32.dll")]
        public static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        static extern IntPtr WindowFromPoint(POINT Point);

        [DllImport("user32.dll")]
        static extern IntPtr WindowFromPoint(int xPoint, int yPoint);

        [DllImport("user32.dll")]
        static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern bool SetWindowText(IntPtr hWnd, string lpString);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool PrintWindow(
        IntPtr hwnd, // Window to copy,Handle to the window that will be copied. 
        IntPtr hdcBlt, // HDC to print into,Handle to the device context. 
        UInt32 nFlags // Optional flags,Specifies the drawing options. It can be one of the following values. 
        );

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleBitmap(
        IntPtr hdc, // handle to DC
        int nWidth, // width of bitmap, in pixels
        int nHeight // height of bitmap, in pixels
        );

        [DllImport("gdi32.dll")]
        public static extern IntPtr SelectObject(
        IntPtr hdc, // handle to DC
        IntPtr hgdiobj // handle to object
        );

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr hwnd);

        [DllImport("gdi32.dll")]
        public static extern int DeleteDC(IntPtr hdc);   // handle to DC

        [DllImport("gdi32.dll")]
        public static extern IntPtr DeleteObject(IntPtr hObject);

        [DllImport("user32.dll")]
        public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public static POINT GetCursorPos()
        {
            POINT p=new LeekCutter.POINT();
            if (GetCursorPos(out p))
            {
                return p;
            }
            throw new Exception();
        }

        public static IntPtr WindowFromPoint()
        {
            POINT p = GetCursorPos();
            return WindowFromPoint(p);
        }

    }
    class Program
    {
        #region bVk参数 常量定义
        public const byte vbKeyLButton = 0x1;    // 鼠标左键
        public const byte vbKeyRButton = 0x2;    // 鼠标右键
        public const byte vbKeyCancel = 0x3;     // CANCEL 键
        public const byte vbKeyMButton = 0x4;    // 鼠标中键
        public const byte vbKeyBack = 0x8;       // BACKSPACE 键
        public const byte vbKeyTab = 0x9;        // TAB 键
        public const byte vbKeyClear = 0xC;      // CLEAR 键
        public const byte vbKeyReturn = 0xD;     // ENTER 键
        public const byte vbKeyShift = 0x10;     // SHIFT 键
        public const byte vbKeyControl = 0x11;   // CTRL 键
        public const byte vbKeyAlt = 18;         // Alt 键  (键码18)
        public const byte vbKeyMenu = 0x12;      // MENU 键
        public const byte vbKeyPause = 0x13;     // PAUSE 键
        public const byte vbKeyCapital = 0x14;   // CAPS LOCK 键
        public const byte vbKeyEscape = 0x1B;    // ESC 键
        public const byte vbKeySpace = 0x20;     // SPACEBAR 键
        public const byte vbKeyPageUp = 0x21;    // PAGE UP 键
        public const byte vbKeyEnd = 0x23;       // End 键
        public const byte vbKeyHome = 0x24;      // HOME 键
        public const byte vbKeyLeft = 0x25;      // LEFT ARROW 键
        public const byte vbKeyUp = 0x26;        // UP ARROW 键
        public const byte vbKeyRight = 0x27;     // RIGHT ARROW 键
        public const byte vbKeyDown = 0x28;      // DOWN ARROW 键
        public const byte vbKeySelect = 0x29;    // Select 键
        public const byte vbKeyPrint = 0x2A;     // PRINT SCREEN 键
        public const byte vbKeyExecute = 0x2B;   // EXECUTE 键
        public const byte vbKeySnapshot = 0x2C;  // SNAPSHOT 键
        public const byte vbKeyDelete = 0x2E;    // Delete 键
        public const byte vbKeyHelp = 0x2F;      // HELP 键
        public const byte vbKeyNumlock = 0x90;   // NUM LOCK 键

        //常用键 字母键A到Z
        public const byte vbKeyA = 65;
        public const byte vbKeyB = 66;
        public const byte vbKeyC = 67;
        public const byte vbKeyD = 68;
        public const byte vbKeyE = 69;
        public const byte vbKeyF = 70;
        public const byte vbKeyG = 71;
        public const byte vbKeyH = 72;
        public const byte vbKeyI = 73;
        public const byte vbKeyJ = 74;
        public const byte vbKeyK = 75;
        public const byte vbKeyL = 76;
        public const byte vbKeyM = 77;
        public const byte vbKeyN = 78;
        public const byte vbKeyO = 79;
        public const byte vbKeyP = 80;
        public const byte vbKeyQ = 81;
        public const byte vbKeyR = 82;
        public const byte vbKeyS = 83;
        public const byte vbKeyT = 84;
        public const byte vbKeyU = 85;
        public const byte vbKeyV = 86;
        public const byte vbKeyW = 87;
        public const byte vbKeyX = 88;
        public const byte vbKeyY = 89;
        public const byte vbKeyZ = 90;

        //数字键盘0到9
        public const byte vbKey0 = 48;    // 0 键
        public const byte vbKey1 = 49;    // 1 键
        public const byte vbKey2 = 50;    // 2 键
        public const byte vbKey3 = 51;    // 3 键
        public const byte vbKey4 = 52;    // 4 键
        public const byte vbKey5 = 53;    // 5 键
        public const byte vbKey6 = 54;    // 6 键
        public const byte vbKey7 = 55;    // 7 键
        public const byte vbKey8 = 56;    // 8 键
        public const byte vbKey9 = 57;    // 9 键


        public const byte vbKeyNumpad0 = 0x60;    //0 键
        public const byte vbKeyNumpad1 = 0x61;    //1 键
        public const byte vbKeyNumpad2 = 0x62;    //2 键
        public const byte vbKeyNumpad3 = 0x63;    //3 键
        public const byte vbKeyNumpad4 = 0x64;    //4 键
        public const byte vbKeyNumpad5 = 0x65;    //5 键
        public const byte vbKeyNumpad6 = 0x66;    //6 键
        public const byte vbKeyNumpad7 = 0x67;    //7 键
        public const byte vbKeyNumpad8 = 0x68;    //8 键
        public const byte vbKeyNumpad9 = 0x69;    //9 键
        public const byte vbKeyMultiply = 0x6A;   // MULTIPLICATIONSIGN(*)键
        public const byte vbKeyAdd = 0x6B;        // PLUS SIGN(+) 键
        public const byte vbKeySeparator = 0x6C;  // ENTER 键
        public const byte vbKeySubtract = 0x6D;   // MINUS SIGN(-) 键
        public const byte vbKeyDecimal = 0x6E;    // DECIMAL POINT(.) 键
        public const byte vbKeyDivide = 0x6F;     // DIVISION SIGN(/) 键


        //F1到F12按键
        public const byte vbKeyF1 = 0x70;   //F1 键
        public const byte vbKeyF2 = 0x71;   //F2 键
        public const byte vbKeyF3 = 0x72;   //F3 键
        public const byte vbKeyF4 = 0x73;   //F4 键
        public const byte vbKeyF5 = 0x74;   //F5 键
        public const byte vbKeyF6 = 0x75;   //F6 键
        public const byte vbKeyF7 = 0x76;   //F7 键
        public const byte vbKeyF8 = 0x77;   //F8 键
        public const byte vbKeyF9 = 0x78;   //F9 键
        public const byte vbKeyF10 = 0x79;  //F10 键
        public const byte vbKeyF11 = 0x7A;  //F11 键
        public const byte vbKeyF12 = 0x7B;  //F12 键

        //移动鼠标 
        public const int MOUSEEVENTF_MOVE = 0x0001;
        //模拟鼠标左键按下 
        public const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        //模拟鼠标左键抬起 
        public const int MOUSEEVENTF_LEFTUP = 0x0004;
        //模拟鼠标右键按下 
        public const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
        //模拟鼠标右键抬起 
        public const int MOUSEEVENTF_RIGHTUP = 0x0010;
        //模拟鼠标中键按下 
        public const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        //模拟鼠标中键抬起 
        public const int MOUSEEVENTF_MIDDLEUP = 0x0040;
        //标示是否采用绝对坐标 
        public const int MOUSEEVENTF_ABSOLUTE = 0x8000;

        #endregion

        Bitmap GetWindowFromTitle(String IN)
        {
            Mat m = new Mat();
            IntPtr Ptr = new IntPtr();
            RECT R = new RECT();
            Process[] localByName = Process.GetProcessesByName(IN);
            Bitmap b = new Bitmap(10,10);
            if (localByName.Length > 0)
            {
                Ptr = localByName[0].MainWindowHandle;
                #region getBitmap
                APIMethod.GetWindowRect(Ptr, out R);
                if(R.Right == R.Left|| R.Bottom == R.Top)
                {
                    return b;
                }
                b = new Bitmap(R.Right - R.Left, R.Bottom - R.Top);
                IntPtr hscrdc = APIMethod.GetWindowDC(Ptr);
                IntPtr hmemdc = APIMethod.CreateCompatibleDC(hscrdc);
                IntPtr hbitmap = APIMethod.CreateCompatibleBitmap(hscrdc, R.Right - R.Left, R.Bottom - R.Top);
                APIMethod.SelectObject(hmemdc, hbitmap);
                APIMethod.PrintWindow(Ptr, hmemdc, 0);
                b = Bitmap.FromHbitmap(hbitmap);
                APIMethod.DeleteObject(hbitmap);
                APIMethod.DeleteDC(hmemdc);
                APIMethod.ReleaseDC(Ptr, hscrdc);
                #endregion
                m = OpenCvSharp.Extensions.BitmapConverter.ToMat(b);
            }
            return b;
        }

        void TypingTest(String IN)
        {
            IntPtr Ptr = new IntPtr();
            Process[] localByName = Process.GetProcessesByName(IN);
            if (localByName.Length > 0)
            {
                Ptr = localByName[0].MainWindowHandle;
                APIMethod.SetForegroundWindow(Ptr);
                APIMethod.keybd_event(vbKeyA, 0, 0, 0);
                Thread.Sleep(1000);
                APIMethod.keybd_event(vbKeyA, 0, 2, 0);
                APIMethod.keybd_event(vbKeyBack, 0, 0, 0);
                Thread.Sleep(1000);
                APIMethod.keybd_event(vbKeyBack, 0, 2, 0);
            }
        }
        static void CursorTest()
        {
            POINT P = APIMethod.GetCursorPos();
            while (true)
            {
                IntPtr Ptr = APIMethod.WindowFromPoint();
                P = APIMethod.GetCursorPos();
                Console.WriteLine(Ptr.ToString("X")+" "+P.x.ToString()+" "+P.y.ToString());
            }
        }

        void OCR()
        {
            TesseractEngine ocr = new TesseractEngine("./tessdata", "chi_sim");
            while (true)
            {
                Bitmap b = GetWindowFromTitle("mspaint");
                if (b.Width > 0 && b.Height > 0)
                {
                    Mat element = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(7, 7));
                    Mat m = OpenCvSharp.Extensions.BitmapConverter.ToMat(b);
                    Cv2.Erode(m, m, element);
                    Cv2.Dilate(m, m, element);
                    b = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(m);
                    Bitmap bitmap24 = new Bitmap(b.Width, b.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    Graphics g = Graphics.FromImage(bitmap24);
                    g.DrawImageUnscaled(b, 0, 0);
                    Page p = ocr.Process(bitmap24);
                    using (new Window("Output", OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap24)))
                    {
                        Cv2.WaitKey(100);
                    }
                    String t = p.GetText();
                    Console.WriteLine(t);
                    p.Dispose();
                }
            }
        }

        static void MoveAndClick(int x,int y)
        {
            APIMethod.SetCursorPos(x, y);
            APIMethod.mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }
        static void Arknights(int x,int y)
        {
            Thread.Sleep(10000);
            Program.MoveAndClick(x, y);
            x -= 372;
            y += 260;
            Thread.Sleep(10000);
            Program.MoveAndClick(x, y);
            x -= 3;
            y -= 167;
            Thread.Sleep(10000);
            Program.MoveAndClick(x, y);
            x += 133;
            y += 97;
            Thread.Sleep(10000);
            for (int i = 0; i < 3; i++)
            {
                Program.MoveAndClick(x, y);
                x += 337;
                y += 74;
                Thread.Sleep(10000);
                //for applying designated mission
                //Program.MoveAndClick(985, 471);
                //Thread.Sleep(10000);
                Program.MoveAndClick(x, y);
                x += 3;
                y -= 86;
                Thread.Sleep(10000);
                Program.MoveAndClick(x, y);
                x -= 98;
                y -= 178;
                Thread.Sleep(2 * 60 * 1000);
                Program.MoveAndClick(x, y);
                x -= 242;
                y += 190;
            }
        }
        static void Main(string[] args)
        {
            Console.WriteLine("x:");
            int.TryParse(Console.ReadLine(), out int x);
            Console.WriteLine("y:");
            int.TryParse(Console.ReadLine(), out int y);
            Program.Arknights(x,y);
        }
    }
}
