/*
* Copyright 2019 Hannah's Hubby
*
*  Licensed under the Apache License, Version 2.0 (the "License");
*  you may not use this file except in compliance with the License.
*  You may obtain a copy of the License at
*
*      http://www.apache.org/licenses/LICENSE-2.0
*
*  Unless required by applicable law or agreed to in writing, software
*  distributed under the License is distributed on an "AS IS" BASIS,
*  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either expressed or implied.
*  See the License for the specific language governing permissions and
*  limitations under the License.
*/



using System;
using System.Collections;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;

using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;

using System.Runtime.InteropServices; //Needed
using System.Reflection;


using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Web;

using System.Threading.Tasks;


namespace CaptureMonitoreServer
{



    public partial class ScreenCapturePlugin : Form
    {
        public ScreenCapturePlugin()
        {
            InitializeComponent();
        }

        private void FrmCaptureMonitor_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                notifyIcon1.Visible = true;
                this.Hide();
                e.Cancel = true;
            }

            System.Windows.Forms.Application.ExitThread();
            Environment.Exit(0);
        }

        


    private bool firstVisibleIgnored = false;
        protected override void SetVisibleCore(bool value)
        {

            if (firstVisibleIgnored == false)
            {
                notifyIcon1.Visible = true;
                System.Diagnostics.Debug.Print("=================firstVisibleIgnored == false=======================");
                Init_Function();
                firstVisibleIgnored = true;
                base.SetVisibleCore(false);
                return;
            }

            base.SetVisibleCore(value);
        }


        private void FrmCaptureMonitor_Load(object sender, EventArgs e)
        {
        }

        private void Init_Function()
        { 
            System.Diagnostics.Debug.Print("=================Init_Function=======================");
            /* create temp directory  */
            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\tmp"))
            { }
            else
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\tmp");
            }
            /* create temp directory  */

            HttpServer httpServer;
            /*
            if (args.GetLength(0) > 0)
            {
                httpServer = new MyHttpServer(Convert.ToInt16(args[0]));
            }
            else
            {
                httpServer = new MyHttpServer(42395);
            }
             */
            httpServer = new MyHttpServer(42395);
            //httpServer.listen();

            Thread thread = new Thread(new ThreadStart(httpServer.listen));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        private void EndProgramToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.ExitThread();
            Environment.Exit(0);
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }
    }


    /* HttpProcessor class */
    public class HttpProcessor
    {
        public TcpClient socket;
        public HttpServer srv;

        private Stream inputStream;
        public StreamWriter outputStream;

        public String http_method;
        public String http_url;
        public String http_protocol_versionstring;
        public Hashtable httpHeaders = new Hashtable();


        private static int MAX_POST_SIZE = 10 * 1024 * 1024; // 10MB




        public HttpProcessor(TcpClient s, HttpServer srv)
        {
            this.socket = s;
            this.srv = srv;
        }


        private string streamReadLine(Stream inputStream)
        {
            int next_char;
            string data = "";
            while (true)
            {
                next_char = inputStream.ReadByte();
                if (next_char == '\n') { break; }
                if (next_char == '\r') { continue; }
                if (next_char == -1) { Thread.Sleep(1); continue; };
                data += Convert.ToChar(next_char);
            }
            return data;
        }
        public void process()
        {
            // we can't use a StreamReader for input, because it buffers up extra data on us inside it's
            // "processed" view of the world, and we want the data raw after the headers
            inputStream = new BufferedStream(socket.GetStream());

            // we probably shouldn't be using a streamwriter for all output from handlers either
            outputStream = new StreamWriter(new BufferedStream(socket.GetStream()));
            try
            {
                parseRequest();
                readHeaders();
                if (http_method.Equals("GET"))
                {
                    handleGETRequest();
                }
                else if (http_method.Equals("POST"))
                {
                    handlePOSTRequest();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.ToString());
                writeFailure();
            }
            outputStream.Flush();
            // bs.Flush(); // flush any remaining output
            inputStream = null; outputStream = null; // bs = null;            
            socket.Close();
        }

        public void parseRequest()
        {
            String request = streamReadLine(inputStream);
            string[] tokens = request.Split(' ');
            if (tokens.Length != 3)
            {
                throw new Exception("invalid http request line");
            }
            http_method = tokens[0].ToUpper();
            http_url = tokens[1];
            http_protocol_versionstring = tokens[2];

            Console.WriteLine("starting: " + request);
        }

        public void readHeaders()
        {
            Console.WriteLine("readHeaders()");
            String line;
            while ((line = streamReadLine(inputStream)) != null)
            {
                if (line.Equals(""))
                {
                    Console.WriteLine("got headers");
                    return;
                }

                int separator = line.IndexOf(':');
                if (separator == -1)
                {
                    throw new Exception("invalid http header line: " + line);
                }
                String name = line.Substring(0, separator);
                int pos = separator + 1;
                while ((pos < line.Length) && (line[pos] == ' '))
                {
                    pos++; // strip any spaces
                }

                string value = line.Substring(pos, line.Length - pos);
                Console.WriteLine("header: {0}:{1}", name, value);
                httpHeaders[name] = value;
            }
        }

        public void handleGETRequest()
        {
            srv.handleGETRequest(this);
        }

        private const int BUF_SIZE = 4096;
        public void handlePOSTRequest()
        {
            // this post data processing just reads everything into a memory stream.
            // this is fine for smallish things, but for large stuff we should really
            // hand an input stream to the request processor. However, the input stream 
            // we hand him needs to let him see the "end of the stream" at this content 
            // length, because otherwise he won't know when he's seen it all! 

            Console.WriteLine("get post data start");
            int content_len = 0;
            MemoryStream ms = new MemoryStream();
            if (this.httpHeaders.ContainsKey("Content-Length"))
            {
                content_len = Convert.ToInt32(this.httpHeaders["Content-Length"]);
                if (content_len > MAX_POST_SIZE)
                {
                    throw new Exception(
                        String.Format("POST Content-Length({0}) too big for this simple server",
                          content_len));
                }
                byte[] buf = new byte[BUF_SIZE];
                int to_read = content_len;
                while (to_read > 0)
                {
                    Console.WriteLine("starting Read, to_read={0}", to_read);

                    int numread = this.inputStream.Read(buf, 0, Math.Min(BUF_SIZE, to_read));
                    Console.WriteLine("read finished, numread={0}", numread);
                    if (numread == 0)
                    {
                        if (to_read == 0)
                        {
                            break;
                        }
                        else
                        {
                            throw new Exception("client disconnected during post");
                        }
                    }
                    to_read -= numread;
                    ms.Write(buf, 0, numread);
                }
                ms.Seek(0, SeekOrigin.Begin);
            }
            Console.WriteLine("get post data end");
            srv.handlePOSTRequest(this, new StreamReader(ms));

        }

        public void writeSuccess(string content_type = "text/html")
        {
            outputStream.WriteLine("HTTP/1.0 200 OK");
            outputStream.WriteLine("Content-Type: " + content_type);
            outputStream.WriteLine("Connection: close");
            outputStream.WriteLine("");
        }

        public void writeFailure()
        {
            outputStream.WriteLine("HTTP/1.0 404 File not found");
            outputStream.WriteLine("Connection: close");
            outputStream.WriteLine("");
        }
    }
    /* HttpProcessor class */


    /* HttpServer class */
    public abstract class HttpServer
    {

        protected int port;
        TcpListener listener;
        bool is_active = true;

        public HttpServer(int port)
        {
            this.port = port;
        }

        public void listen()
        {
            listener = new TcpListener(port);
            listener.Start();
            while (is_active)
            {
                TcpClient s = listener.AcceptTcpClient();
                HttpProcessor processor = new HttpProcessor(s, this);
                Thread thread = new Thread(new ThreadStart(processor.process));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                Thread.Sleep(1);
            }
        }

        public abstract void handleGETRequest(HttpProcessor p);
        public abstract void handlePOSTRequest(HttpProcessor p, StreamReader inputData);
    }
    /* HttpServer class */


    /* MyHttpServer class */
   
    public class MyHttpServer : HttpServer
    {

        /* excel valiable */
        static Microsoft.Office.Interop.Excel.Application objExcel = null;
        static Workbook workbook = null;
        static Worksheet worksheet = null;

        static int iCaptureCnt = 0;
        static int iExcelStartRowNum = 0;

        static string file_nm;
        static string file_path;


        /* excel valiable */


        public MyHttpServer(int port)
            : base(port)
        {
        }
        public override void handleGETRequest(HttpProcessor p)
        {

            if (p.http_url.Equals("/Test.png"))
            {
                Stream fs = File.Open("../../Test.png", FileMode.Open);

                p.writeSuccess("image/png");
                fs.CopyTo(p.outputStream.BaseStream);
                p.outputStream.BaseStream.Flush();
            }

            Console.WriteLine("request: {0}", p.http_url);

            if (p.http_url.Equals("/?type=start"))
            {
                startExcel();
            }
            else if (p.http_url.Equals("/?type=end"))
            {
                endExcel();
            }
            else if (p.http_url.Equals("/favicon.ico"))
            { }
            else
            {
                CaptureScreen(p.http_url);
            }

            p.writeSuccess();
            /*
            p.outputStream.WriteLine("<html><body><h1>test server</h1>");
            p.outputStream.WriteLine("Current Time: " + DateTime.Now.ToString());
            p.outputStream.WriteLine("url : {0}", p.http_url);

            p.outputStream.WriteLine("<form method=post action=/form>");
            p.outputStream.WriteLine("<input type=text name=foo value=foovalue>");
            p.outputStream.WriteLine("<input type=submit name=bar value=barvalue>");
            p.outputStream.WriteLine("</form>");
            */
        }

        public override void handlePOSTRequest(HttpProcessor p, StreamReader inputData)
        {
            string data = inputData.ReadToEnd();
            Console.WriteLine("POST request: {0}", p.http_url);
            Console.WriteLine("data request: {0}", data);
            
            Console.WriteLine("request: {0}", p.http_url);

            if (p.http_url.Equals("/?type=start"))
            {
                startExcel();
            }
            else if (p.http_url.Equals("/?type=end"))
            {
                endExcel();
            }
            else if (p.http_url.Equals("/favicon.ico"))
            { }
            else
            {
                CaptureScreen(data);
            }

            p.writeSuccess();



            p.writeSuccess();
            /*
            p.outputStream.WriteLine("<html><body><h1>test server</h1>");
            p.outputStream.WriteLine("<a href=/test>return</a><p>");
            p.outputStream.WriteLine("postbody: <pre>{0}</pre>", data);
            */

        }



        private static Bitmap CaptureCursor(ref int x, ref int y)
        {
            Bitmap bmp;
            IntPtr hicon;

            Win32Stuff.CURSORINFO ci = new Win32Stuff.CURSORINFO();
            Win32Stuff.ICONINFO icInfo;
            ci.cbSize = Marshal.SizeOf(ci);
            if (Win32Stuff.GetCursorInfo(out ci))
            {
                if (ci.flags == Win32Stuff.CURSOR_SHOWING)
                {
                    hicon = Win32Stuff.CopyIcon(ci.hCursor);
                    if (Win32Stuff.GetIconInfo(hicon, out icInfo))
                    {
                        x = ci.ptScreenPos.x - ((int)icInfo.xHotspot);
                        y = ci.ptScreenPos.y - ((int)icInfo.yHotspot);
                        System.Drawing.Icon ic = System.Drawing.Icon.FromHandle(hicon);
                        bmp = ic.ToBitmap();

                        return bmp;
                    }
                }
            }
            return null;
        }





        
        public static void CaptureScreen(string getData)
        {

            iCaptureCnt++;
            Range cellRange = null;
            Size sz = new Size(Screen.PrimaryScreen.Bounds.Width , Screen.PrimaryScreen.Bounds.Height );
            Bitmap bt = new Bitmap(Screen.PrimaryScreen.Bounds.Width , Screen.PrimaryScreen.Bounds.Height );

            Graphics g = Graphics.FromImage(bt);
            g.CopyFromScreen(0, 0, 0, 0, sz); // capture screen



            // mouse capture
            int iCurX = 0;
            int iCurY = 0;
            Bitmap Cur = CaptureCursor(ref iCurX, ref iCurY);
            System.Drawing.Rectangle curRec = new System.Drawing.Rectangle(iCurX, iCurY, Cur.Width, Cur.Height); // Rectangle for mouse point
            g.DrawImage(Cur, curRec); // create mouse point in g(capture screen)


            MemoryStream ms = new MemoryStream(); // memory stream for capture
            ms.Position = 0;

            // set ms buffer imager
            bt.Save(ms, ImageFormat.Jpeg);
            SaveMemoryStream(ms, file_path + @"\img_" + iCaptureCnt.ToString().PadLeft(3, '0') + ".jpg");
            string aa = HttpUtility.UrlDecode(getData);
            File.WriteAllText(file_path + @"\img_" + iCaptureCnt.ToString().PadLeft(3, '0') + ".dat", HttpUtility.UrlDecode(getData), Encoding.UTF8);


            /*
            bt.Size.Height = 500;
            bt.Size.Width = 500;

            int h = bt.Size.Height;
            int w = bt.Size.Width;

            Clipboard.SetImage(bt);
            int count = (int)Math.Ceiling((double)bt.Size.Height / 100);
            */


            //cellRange = worksheet.get_Range("A" + iExcelStartRowNum, "A" + iExcelStartRowNum);
            //cellRange.set_Value(Missing.Value, "Action #" + iCaptureCnt);
            //iExcelStartRowNum++;


            //cellRange = worksheet.get_Range("A" + iExcelStartRowNum, "A" + (iExcelStartRowNum + count - 1));
            //iExcelStartRowNum = iExcelStartRowNum + count;
            //cellRange.RowHeight = 100; //100

            //worksheet.Paste(cellRange, bt);
            //worksheet.Paste(cellRange);
            //workbook.Save();
            //objExcel.Quit();

            /*
            DataObject data = new DataObject();
            data.SetData("rawbinary", false, ms);
            Clipboard.SetDataObject(data, true);
            */

            //SaveMemoryStream(ms, @"d:\aaa.jpg");
        }


        public static void SaveMemoryStream(MemoryStream ms, string FileName)
        {
            FileStream outStream = File.OpenWrite(FileName);
            ms.WriteTo(outStream);
            outStream.Flush();
            outStream.Close();
        }

        private void startExcel()
        {
            file_nm = DateTime.Now.ToString("yyyyMMddHHmmss");



            file_path = System.Windows.Forms.Application.StartupPath + @"\tmp\" + file_nm;

            if (Directory.Exists(file_path))
            {}
            else
            {
                Directory.CreateDirectory(file_path);
            }
            
        }




        private void createExcel()
        {


            iExcelStartRowNum = 1;

            object missingParam = Type.Missing;
            objExcel = new Microsoft.Office.Interop.Excel.Application();
            workbook = objExcel.Workbooks.Add();
            worksheet = workbook.Worksheets.Add() as Worksheet;

            workbook.SaveAs(file_path + @"\" + file_nm + "_screenCapture.xlsx");
            
            //endExcel();
        }

        private void writeExcel()
        {
            Range cellRange = null;
            Bitmap sourceImage = null;
            string temp = null;

            String FolderName = file_path + "\\";
            //file sorting
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(FolderName);
            foreach (System.IO.FileInfo File in di.GetFiles())
            //foreach (String File in Directory.GetFiles(FolderName, "*").OrderBy(f => f))
            {

                if (File.Extension.ToLower().CompareTo(".jpg") == 0)
                {


                    //String FileNameOnly = File.Name.Substring(0, File.Name.Length - 4);
                    //String FullFileName = File.FullName;

                    //MessageBox.Show(FullFileName + " " + FileNameOnly);
                    if (System.IO.File.Exists(File.FullName.Replace(".jpg", ".dat")) == true)
                    {
                        temp = System.IO.File.ReadAllText(File.FullName.Replace(".jpg", ".dat"));
                        cellRange = worksheet.get_Range("A" + iExcelStartRowNum, "A" + iExcelStartRowNum);
                        cellRange.set_Value(Missing.Value, temp);
                    }

                    iExcelStartRowNum++;


                    sourceImage = new Bitmap(File.FullName);
                    int count = (int)Math.Ceiling((double)sourceImage.Height / 100);
                    Image picObject = Image.FromFile(File.FullName);
                    System.Windows.Forms.Clipboard.SetDataObject(picObject, true);
                    cellRange = worksheet.get_Range("A" + iExcelStartRowNum, "A" + (iExcelStartRowNum + count - 1));
                    cellRange.RowHeight = 100; //100
                    worksheet.Paste(cellRange, picObject);
                    System.Windows.Forms.Clipboard.Clear();

                    iExcelStartRowNum = iExcelStartRowNum + count;


                    //worksheet.Paste(cellRange, bt);
                    //worksheet.Paste(cellRange);


                }

            }
            workbook.Save();
        }


        private void endExcel()
        {

            createExcel();
            writeExcel();

            workbook.Close(0);
            int id;
            Win32Stuff.GetWindowThreadProcessId(objExcel.Hwnd, out id);
            Process processList = Process.GetProcessById(id);
            processList.Kill();
            //System.Diagnostics.Process.
        }

    }


    /* Win32Stuff class */
    class Win32Stuff
    {

        #region Class Variables

        public const int SM_CXSCREEN = 0;
        public const int SM_CYSCREEN = 1;

        public const Int32 CURSOR_SHOWING = 0x00000001;

        [StructLayout(LayoutKind.Sequential)]
        public struct ICONINFO
        {
            public bool fIcon;         // Specifies whether this structure defines an icon or a cursor. A value of TRUE specifies
            public Int32 xHotspot;     // Specifies the x-coordinate of a cursor's hot spot. If this structure defines an icon, the hot
            public Int32 yHotspot;     // Specifies the y-coordinate of the cursor's hot spot. If this structure defines an icon, the hot
            public IntPtr hbmMask;     // (HBITMAP) Specifies the icon bitmask bitmap. If this structure defines a black and white icon,
            public IntPtr hbmColor;    // (HBITMAP) Handle to the icon color bitmap. This member can be optional if this
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public Int32 x;
            public Int32 y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CURSORINFO
        {
            public Int32 cbSize;        // Specifies the size, in bytes, of the structure.
            public Int32 flags;         // Specifies the cursor state. This parameter can be one of the following values:
            public IntPtr hCursor;          // Handle to the cursor.
            public POINT ptScreenPos;       // A POINT structure that receives the screen coordinates of the cursor.
        }

        #endregion


        #region Class Functions


        [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")]
        public static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        [DllImport("user32.dll", EntryPoint = "GetDesktopWindow")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", EntryPoint = "GetDC")]
        public static extern IntPtr GetDC(IntPtr ptr);

        [DllImport("user32.dll", EntryPoint = "GetSystemMetrics")]
        public static extern int GetSystemMetrics(int abc);

        [DllImport("user32.dll", EntryPoint = "GetWindowDC")]
        public static extern IntPtr GetWindowDC(Int32 ptr);

        [DllImport("user32.dll", EntryPoint = "ReleaseDC")]
        public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDc);


        [DllImport("user32.dll", EntryPoint = "GetCursorInfo")]
        public static extern bool GetCursorInfo(out CURSORINFO pci);

        [DllImport("user32.dll", EntryPoint = "CopyIcon")]
        public static extern IntPtr CopyIcon(IntPtr hIcon);

        [DllImport("user32.dll", EntryPoint = "GetIconInfo")]
        public static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO piconinfo);


        #endregion
    }
    /* Win32Stuff class */






}
