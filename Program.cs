﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using FVSelectScenes;
using WinShell;
using System.Diagnostics;
using Interop;

namespace FVProduceScenes
{
    class Program
    {
        const string c_syntax =
@"Syntax: FVProduceScenes <filename> <destination folder> [options]
   Filenames should be .mp4. Each file should have an accompanying .csv
   containing segments. Filenames may contain wildcards.

Options:
  -t   Tolerate dates out of order.
";

        const string c_csvSuffix = " scenes.csv";

        static string s_srcFilename = null;
        static string s_destFolder = null;
        static bool s_tolerateDatesOutOfOrder = false;
        static bool s_showSyntax = false;

        static void Main(string[] args)
        {
            try
            {
                ParseCommandLine(args);
                if (s_showSyntax)
                {
                    Console.WriteLine(c_syntax);
                }
                else
                {
                    if (!Directory.Exists(s_destFolder))
                    {
                        throw new ArgumentException("Destination folder does not exist: " + s_destFolder);
                    }

                    foreach (var filename in CodeBit.PathEx.GetFilesByPattern(args[0]))
                    {
                        ProcessFile(filename);
                    }
                }
            }
            catch (Exception err)
            {
#if DEBUG
                Console.WriteLine(err.ToString());
#else
                Console.WriteLine(err.Message);
#endif
            }

#if DEBUG
            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine();
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
#endif
        }

        static void ParseCommandLine(string[] args)
        {
            foreach(string arg in args)
            {
                if (arg.Equals("-t", StringComparison.OrdinalIgnoreCase))
                {
                    s_tolerateDatesOutOfOrder = true;
                }
                else if (arg.Equals("-h", StringComparison.OrdinalIgnoreCase)
                    || arg.Equals("-?", StringComparison.Ordinal))
                {
                    s_showSyntax = true;
                }
                else if (s_srcFilename == null)
                {
                    s_srcFilename = arg;
                }
                else if (s_destFolder == null)
                {
                    s_destFolder = arg;
                }
                else
                {
                    throw new ApplicationException("Unexpected command-line argument: " + arg);
                }
            }
        }

        static void ProcessFile(string filename)
        {
            Console.WriteLine($"Processing: {filename}");

            string csvName = Path.Combine(Path.GetDirectoryName(filename),
                Path.GetFileNameWithoutExtension(filename) + c_csvSuffix);

            if (!File.Exists(csvName))
            {
                throw new ApplicationException($"Scene file not found: {csvName}");
            }

            var segments = SegmentModel.LoadFromCsv(csvName);

            // Check the segments
            var lastDate = DateTime.MinValue.AddSeconds(1);
            var lastPosition = TimeSpan.MinValue;
            foreach(var segment in segments)
            {
                if (segment.Disposition == SegmentDisposition.Keep)
                {

                    if (segment.Date < lastDate && !s_tolerateDatesOutOfOrder)
                    {
                        throw new ApplicationException($"Dates aren't ascending. {lastDate} > {segment.Date}.  Use -t option to tolerate this.");
                    }
                    if (segment.Position < lastPosition)
                    {
                        throw new ApplicationException("Positions aren't ascending.");
                    }
                    if (string.IsNullOrEmpty(segment.Subject))
                    {
                        throw new ApplicationException("Segment is missing subject.");
                    }
                    lastDate = segment.Date;
                    lastPosition = segment.Position;
                }
            }

            // Get the length of the entire video
            TimeSpan videoLength;
            using (var ps = PropertyStore.Open(filename))
            {
                videoLength = TimeSpan.FromTicks((long)(ulong)ps.GetValue(s_pkDuration));
            }

            // Produce the scenes
            lastDate = DateTime.MinValue;
            int ordinal = 0;
            for (int i=0; i<segments.Count; ++i)
            {
                var segment = segments[i];
                if (segment.Disposition == SegmentDisposition.Keep)
                {
                    // Find the end
                    TimeSpan end = videoLength;
                    for (int j=i+1; j<segments.Count; ++j)
                    {
                        if (segments[j].Disposition != SegmentDisposition.AddToPrevious)
                        {
                            end = segments[j].Position - TimeSpan.FromMilliseconds(34); // One NTSC frame short of the next segment
                            break;
                        }
                    }

                    if (segment.Date > lastDate)
                    {
                        ordinal = 0;
                    }
                    ++ordinal;
                    lastDate = segment.Date;

                    ProduceScene(filename, ordinal, segment.Position, end, segment.Date, segment.Subject, segment.Title);
                }
            }

        }

        static void ProduceScene(string filename, int ordinal, TimeSpan start, TimeSpan end, DateTime date, string subject, string title)
        {
            string dstFilename = GenerateFilename(date, ordinal, subject, title);
            Console.WriteLine(dstFilename);

            // Print the info
            Console.WriteLine($"  start={start} end={end} date={date:yyyy-MM-dd} ordinal={ordinal} subject=\"{subject}\" title=\"{title}\"");

            // Extract the video segment

            // This version is the fastest and maintains quality because it simply copies the video across.
            // However, it has to back up to the preceding key frame so cuts are less accurate.
            // string arguments = $"-hide_banner -ss {start.Ticks / 10}us -to {end.Ticks / 10}us -i \"{filename}\" -c:v copy -c:a copy \"{dstFilename}\"";

            // This version should give more precise cuts because it decompresses and re-compresses the video. Whether that results
            // in reduced quality is a question.
            //string arguments = $"-hide_banner -i \"{filename}\" -ss {start.Ticks / 10}us -to {end.Ticks / 10}us -c:v libx264 -profile:v main -level:v 3.1 -crf 20 -c:a aac \"{dstFilename}\"";

            // This version has precise cuts and uses the unsharp filter which actually sharpens the output.
            //string arguments = $"-hide_banner -i \"{filename}\" -ss {start.Ticks / 10}us -to {end.Ticks / 10}us -vf unsharp -c:v libx264 -profile:v main -level:v 3.1 -crf 20 -c:a aac \"{dstFilename}\"";

            // This version has precise cuts, uses the unsharp filter, and adds +faststart for quicker loading
            string arguments = $"-hide_banner -i \"{filename}\" -ss {start.Ticks / 10}us -to {end.Ticks / 10}us -vf unsharp -c:v libx264 -profile:v main -level:v 3.1 -crf 20 -c:a aac -movflags +faststart \"{dstFilename}\"";

            // This version is an attempt to speed things up by not having to read all of the video up to the cut point. But it doesn't work - it sums up both delays.
            //string arguments = $"-hide_banner -ss {start.Ticks / 10}us -i \"{filename}\" -ss {start.Ticks / 10}us -to {end.Ticks / 10}us -c:v libx264 -profile:v main -level:v 3.1 -crf 20 -c:a aac \"{dstFilename}\"";

            Console.WriteLine(arguments);

            var psi = new ProcessStartInfo("ffmpeg.exe", arguments);
            psi.UseShellExecute = false;
            psi.CreateNoWindow = false;

            using (var process = Process.Start(psi))
            {
                process.WaitForExit();
            }
            Console.WriteLine();

            // If only date, not time, translate ordinal into seconds
            bool dateOnly = (date.Hour == 0 && date.Minute == 0 && date.Second == 0);
            if (dateOnly)
            {
                date = new DateTime(date.Year, date.Month, date.Day, 12+(ordinal / (60 * 60)), (ordinal / 60) % 60, ordinal % 60, DateTimeKind.Utc);
            }

            // Apply date metadata
            using (var isom = new FileMeta.IsomCoreMetadata(dstFilename, true))
            {
                isom.CreationTime = date;
                isom.ModificationTime = date;
                isom.Commit();
            }

            // Apply the balance of the metadata
            using (var ps = WinShell.PropertyStore.Open(dstFilename, true))
            {
                string metaTitle = null;
                if (!string.IsNullOrEmpty(subject))
                {
                    if (!string.IsNullOrEmpty(title))
                    {
                        metaTitle = string.Concat(subject, " - ", title);
                    }
                    else
                    {
                        metaTitle = subject;
                    }
                }
                else if (!string.IsNullOrEmpty(title))
                {
                    metaTitle = title;
                }

                if (!string.IsNullOrEmpty(metaTitle))
                {
                    ps.SetValue(s_pkTitle, metaTitle);
                }

                ps.SetValue(s_pkTrackNum, (uint)ordinal);

                string comment = "&timezone=0";
                if (dateOnly)
                {
                    comment += " &datePrecision=8";
                }
                ps.SetValue(s_pkComment, comment);
                ps.Commit();
            }
        }

        static string GenerateFilename(DateTime date, int ordinal, string subject, string title)
        {
            for (char attempt=(char)('a'-1); attempt<'z'; ++attempt)
            {
                var attemptStr = (attempt >= 'a') ? attempt.ToString() : string.Empty;
                var filename = $"{date:yyyy-MM-dd} ({ordinal:d2}{attemptStr})";
                
                subject = subject.Trim();
                if (!string.IsNullOrEmpty(subject))
                {
                    filename = string.Concat(filename, " ", subject);
                }

                title = title.Trim();
                if (!string.IsNullOrEmpty(title))
                {
                    filename = string.Concat(filename, " - ", title);
                }

                filename += ".mp4";

                filename = Path.Combine(s_destFolder, filename);

                if (!File.Exists(filename))
                {
                    return filename;
                }
            }
            throw new ApplicationException("Too many duplicates.");
        }

        // https://docs.microsoft.com/en-us/windows/win32/properties/props
        public static PropertyKey s_pkDuration = new PropertyKey("64440490-4C8B-11D1-8B70-080036B11A03", 3);
        public static PropertyKey s_pkComment = new PropertyKey("F29F85E0-4FF9-1068-AB91-08002B27B3D9", 6); // 
        public static PropertyKey s_pkTitle = new PropertyKey("F29F85E0-4FF9-1068-AB91-08002B27B3D9", 2); // System.Title
        public static PropertyKey s_pkTrackNum = new PropertyKey("56A3372E-CE9C-11D2-9F0E-006097C686F6", 7); // System.Music.TrackNumber

    }
}
