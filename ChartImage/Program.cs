using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using NetOffice;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.OfficeApi.Enums;
using Application = NetOffice.ExcelApi.Application;
using Encoder = System.Drawing.Imaging.Encoder;

namespace ChartImage
{
    class Program
    {
        private static string _strOupPutPath;
        private static List<string[]> _listInfo;
        private static string _strInfoFile;

        [STAThread]
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("输入文件夹路径：");
                string strFolder = Console.ReadLine();
                if (!Directory.Exists(strFolder))
                {
                    Console.WriteLine("路径不正确");
                    continue;
                }

                string[] strsInfo = GetChartTypeInfo();
                if (strsInfo.Length == 0)
                {
                    Console.WriteLine("标签文件错误");
                    continue;
                }

                _listInfo = strsInfo.Select(s => s.Split('\t')).ToList();

                _strOupPutPath = $"{strFolder}\\output\\";
                if (!Directory.Exists(_strOupPutPath))
                    Directory.CreateDirectory(_strOupPutPath);

                _strInfoFile = _strOupPutPath + "info.txt";
                if (File.Exists(_strInfoFile))
                    File.Delete(_strInfoFile);

                var files = Directory.GetFiles(strFolder, "*.*", SearchOption.TopDirectoryOnly);
                files = files.Where(f => f.Contains(".xls") && !f.Contains("~$")).ToArray();
                foreach (string file in files)
                {
                    try
                    {
                        ExportChart(file);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
            }
            
        }

        static void ExportChart(string strFile)
        {
            Application app = GetApp();
            if (app == null)
                throw new Exception("先打开excel");

            Workbook workbook = app.Workbooks.Open(strFile);
            foreach (Worksheet sheet in workbook.Sheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoChart)
                        ExportChart(shape, strFile);
                }
            }
        }

        static void ExportChart(Shape shape, string strFileName)
        {
            Chart chart = shape.Chart;
            string strName = chart.ChartTitle.Caption;
            int iSeriesCnt = (chart.SeriesCollection() as SeriesCollection).Count;
            string strChartSaveName = _strOupPutPath;

            int i = 1;
            while (true)
            {
                strChartSaveName = $"{_strOupPutPath}{strName}.crtx";
                if (!File.Exists(strChartSaveName))
                    break;
                strName += i;
                i++;
            }

            chart.SaveChartTemplate(strChartSaveName);

            string strBigImg = $"{_strOupPutPath}\\{strName}_1.png";
            string strSmallImg = $"{_strOupPutPath}\\m_{strName}_1.png";
            shape.Copy();
            Object obj = Clipboard.GetData("PNG");
            Image imgTemp = new Bitmap((MemoryStream) obj);
            Clipboard.Clear();

            Image imgBig = new Bitmap(imgTemp, 680, 407);
            Image imgSmall = new Bitmap(imgTemp, 499, 299);
            imgBig.Save(strBigImg);
            imgSmall.Save(strSmallImg);

            imgTemp.Dispose();
            imgBig.Dispose();
            imgSmall.Dispose();


            string strChartType = chart.ChartType.ToString();
            var infos = _listInfo.Where(d => d[0] == strChartType);
            if (!infos.Any())
                throw new Exception($"找不到{strChartType}类型标签");
            else if (infos.Count() != 1)
                throw new Exception($"{strChartType}类型标签重复");

            List<string> listInfo = infos.Single().ToList();
            string strLine = strName + "\t";
            strLine += listInfo[1] + "\t";
            listInfo.RemoveAt(0);
            listInfo.RemoveAt(0);
            listInfo.Insert(0, iSeriesCnt.ToString());
            strLine += string.Join(" ", listInfo);
            strLine += "\r\n";

            File.AppendAllText(_strInfoFile, strLine, Encoding.GetEncoding("GB2312"));
        }

        static string[] GetChartTypeInfo()
        {
            string strFile = $"{Environment.CurrentDirectory}\\charttype.txt";
            return File.ReadAllLines(strFile, Encoding.GetEncoding("GB2312"));
        }

        static Application GetApp()
        {
            object obj = null;
            try
            {
                obj = Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {

            }

            if (obj != null)
            {
                return new NetOffice.ExcelApi.Application(new COMObject(obj));
                //return (NetOffice.PowerPointApi.Application)obj;
            }
            return new NetOffice.ExcelApi.Application("Excel.Application");
        }

        private static EncoderParameters encoderParams;
        private static ImageCodecInfo jpegImageCodecInfo;
        public static EncoderParameters GetEncoder(int iQuality)
        {
            encoderParams = new EncoderParameters();
            long[] quality = new long[1];
            quality[0] = iQuality;
            EncoderParameter encoderParam = new EncoderParameter(Encoder.Quality, quality);
            encoderParams.Param[0] = encoderParam;

            return encoderParams;
        }
    }
}
