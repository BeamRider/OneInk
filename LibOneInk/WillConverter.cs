using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace LibOneInk
{
    public class WillConverter
    {
        private static readonly log4net.ILog Logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public float ScaleFactor { get; set; } = 1.0f;
        public float PressureRatio { get; set; } = 0.25f;

        public List<OneInkStrokeGroup> Groups { get; } = new List<OneInkStrokeGroup>();

        public WillConverter(string path)
        {
            ReadWillFile(path);
        }

        private void ReadWillFile(string path)
        { 
            string baseDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

            try
            {
                ZipFile.ExtractToDirectory(path, baseDir);
                if (Directory.Exists(baseDir))
                {
                    List<string> strokeFiles = ReadStrokeFiles(baseDir);

                    foreach (string fileName in strokeFiles)
                    {
                        Logger.Info($"Importing {fileName}");

                        OneInkStrokeGroup ink = ImportStrokeFile(fileName);
                        Groups.Add(ink);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
            finally
            {
                if (Directory.Exists(baseDir))
                    Directory.Delete(baseDir, true);
            }
        }

        public List<string> ReadStrokeFiles(string baseDir)
        {
            List<string> strokeFiles = new List<string>();
            string sectionsPath = Path.Combine(baseDir, @"sections\_rels");
            int sid = 0;
            do
            {
                string sectionFile = Path.Combine(sectionsPath, $"section{sid}.svg.rels");
                Logger.Info($"Section {sid} found");
                if (!File.Exists(sectionFile))
                    break;
                XDocument secDoc = XDocument.Parse(File.ReadAllText(sectionFile));
                foreach (var rel in secDoc.Descendants().First().Descendants())
                {
                    XAttribute attr = rel.Attribute("Target");
                    if (attr != null)
                    {
                        Logger.Info($"Section {sid} target {attr.Value}");
                        strokeFiles.Add(Path.Combine(baseDir, attr.Value.Substring(1)));
                    }
                }
                sid++;
            } while (true);
            return strokeFiles;
        }

        private OneInkStrokeGroup ImportStrokeFile(string fileName)
        {
            OneInkStrokeGroup group = new OneInkStrokeGroup();
            using (var stream = File.OpenRead(fileName))
            {
                int len;
                while ((len = ReadLen(stream)) > 0)
                {
                    OneInkStroke stroke = new OneInkStroke();
                    byte[] data = new byte[len];
                    stream.Read(data, 0, len);
                    var willStroke = WacomInkFormat.Path.Parser.ParseFrom(data);
                    float x = 0;
                    float y = 0;
                    float z = 0;
                    long col = 0;
                    long lastcol = col;
                    float fpf = (float)Math.Pow(10, willStroke.DecimalPrecision) * ScaleFactor;
                    for (int idx = 0; idx < willStroke.Points.Count / 2; idx++)
                    {
                        float dx = willStroke.Points[2 * idx] / fpf;
                        float dy = willStroke.Points[2 * idx + 1] / fpf;
                        if (willStroke.StrokeWidths.Count > idx)
                        {
                            float dz = willStroke.StrokeWidths[idx] / fpf;
                            z += dz;
                        }
                        if (willStroke.StrokeColor.Count > idx)
                        {
                            int dcol = willStroke.StrokeColor[idx];
                            col += dcol;
                        }
                        x += dx;
                        y += dy;
                        float nz = z * PressureRatio;
                        if (stroke.Points.Count == 0)
                        {
                            stroke.Color = new OneInkColor
                            {
                                r = (byte)((col >> 24) & 0xff),
                                g = (byte)((col >> 16) & 0xff),
                                b = (byte)((col >> 8) & 0xff),
                                a = (byte)(col & 0xff)
                            };
                        }
                        else if (lastcol != col)
                        {
                            group.Strokes.Add(stroke);
                            stroke = new OneInkStroke();
                        }
                        stroke.Points.Add(new OneInkPoint { x = x, y = y, pressure = (nz <= 1.0f) ? nz : 1.0f });
                        lastcol = col;
                    }
                    if(stroke.Points.Count > 0)
                    {
                        group.Strokes.Add(stroke);
                    }
                }
            }
            return group;
        }

        private static int ReadLen(Stream s)
        {
            UInt32 res = 0;
            int i = 0;
            bool more;
            do
            {
                int v = s.ReadByte();
                if (v < 0)
                    return v;
                more = (v & 0x80) != 0;
                res = res | (((UInt32)(v & 0x7f)) << (i * 8 - i));
                i++;
            } while (more);
            return (int)res;
        }


    }
}
