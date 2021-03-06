﻿using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using IWshRuntimeLibrary;

namespace SnipIt
{
    public partial class Form1 : Form
    {

        int selectX;
        int selectY;
        int selectWidth;
        int selectHeight;
        Random rnd = new Random();
        SolidBrush myBrush = new SolidBrush(Color.FromArgb(50, 255, 255, 255));
        bool start = false;
        [System.Runtime.InteropServices.DllImport("Shell32.dll")]
        private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
        public Form1()
        {
            appShortcutToDesktop();
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Top = 0;
            Left = 0;
            Hide();
            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(printscreen as Image);
            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);
            using (MemoryStream s = new MemoryStream())
            {
                printscreen.Save(s, ImageFormat.Bmp);
                pictureBox1.Size = new System.Drawing.Size(Width, Height);
                pictureBox1.Image = Image.FromStream(s);
            }
            Show();
            Cursor = Cursors.Cross;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            if (pictureBox1.Image == null)
                return;
            if (start)
            {
                pictureBox1.Refresh();
                selectWidth = e.X - selectX;
                selectHeight = e.Y - selectY;
                pictureBox1.CreateGraphics().FillRectangle(myBrush, selectX, selectY, selectWidth, selectHeight);
            }
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                selectX = e.X;
                selectY = e.Y;
            }
            pictureBox1.Refresh();
            start = true;
        }

        private void pictureBox1_MouseRelease(object esneder, MouseEventArgs e)
        {
            if (pictureBox1.Image == null)
                return;
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                pictureBox1.Refresh();
                selectWidth = e.X - selectX;
                selectHeight = e.Y - selectY;
                pictureBox1.CreateGraphics().FillRectangle(myBrush, selectX,
                         selectY, selectWidth, selectHeight);

            }
            start = false;
            SaveToClipboardAsync();
        }
        private void appShortcutToDesktop()
        {
            string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string app = System.Reflection.Assembly.GetExecutingAssembly().Location;

            object shDesktop = (object)"Desktop";

            WshShell shell = new WshShell();

            string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + @"\SnipIt.lnk";
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
            shortcut.Description = "SnipIt Shortcut";
            shortcut.Hotkey = "Ctrl+Shift+C";
            shortcut.TargetPath = app;
            shortcut.Save();
        }
        public Byte[] BitmapToArray(Bitmap bitmap)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Bmp);
                return stream.ToArray();
            }
        }
        private void SaveToClipboardAsync()
        {
            if (selectWidth > 0)
            {

                Rectangle rect = new Rectangle(selectX, selectY, selectWidth, selectHeight);
                Bitmap OriginalImage = new Bitmap(pictureBox1.Image, pictureBox1.Width, pictureBox1.Height);
                Bitmap _img = new Bitmap(selectWidth, selectHeight);
                Graphics g = Graphics.FromImage(_img);
                g.InterpolationMode = InterpolationMode.High;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.DrawImage(OriginalImage, 0, 0, rect, GraphicsUnit.Pixel);
                int ourId = rnd.Next(1, 1000);
                var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                _img.Save(@desktop + @"\SnipIt_" + ourId + ".png", ImageFormat.Png);
                Clipboard.SetImage(_img);
            }
            SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);
            Application.Exit();
        }
    }
}
