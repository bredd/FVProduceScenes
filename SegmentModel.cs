using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FVSelectScenes
{
    class SegmentModel
    {
        public TimeSpan Position { get; set; }
        public SegmentDisposition Disposition { get; set; }
        public DateTime Date { get; set; }
        public string Subject { get; set; }
        public string Title { get; set; }

        public SegmentModel()
        {
            Position = new TimeSpan();
            Disposition = SegmentDisposition.Keep;
            Date = DateTime.MinValue;
            Subject = string.Empty;
            Title = string.Empty;
        }

        public SegmentModel(TimeSpan position, SegmentDisposition disposition, DateTime date, string subject, string title)
        {
            Position = position;
            Disposition = disposition;
            Date = date;
            Subject = subject;
            Title = title;
        }

        const string c_positionLabel = "seconds";
        const string c_dispositionLabel = "disposition";
        const string c_dateLabel = "date";
        const string c_subjectLabel = "subject";
        const string c_titleLabel = "title";

        static readonly TimeSpan s_minimumSegment = TimeSpan.FromSeconds(0.75);

        public static List<SegmentModel> LoadFromCsv(string filename)
        {
            using (var reader = new FileMeta.CsvReader(filename))
            {
                int positionIndex = int.MaxValue;
                int dispositionIndex = int.MaxValue;
                int dateIndex = int.MaxValue;
                int subjectIndex = int.MaxValue;
                int titleIndex = int.MaxValue;

                var list = new List<SegmentModel>();
                {
                    var header = reader.Read();
                    if (header == null)
                    {
                        return list;
                    }
                    for (int i = 0; i < header.Length; ++i)
                    {
                        switch (header[i].ToLowerInvariant())
                        {
                            case c_positionLabel:
                                positionIndex = i;
                                break;

                            case c_dispositionLabel:
                                dispositionIndex = i;
                                break;

                            case c_dateLabel:
                                dateIndex = i;
                                break;

                            case c_subjectLabel:
                                subjectIndex = i;
                                break;

                            case c_titleLabel:
                                titleIndex = i;
                                break;
                        }
                    }
                }

                for (; ; )
                {
                    var line = reader.Read();
                    if (line == null) break;

                    TimeSpan position;
                    if (positionIndex >= line.Length || !TryParseSeconds(line[positionIndex], out position))
                    {
                        position = new TimeSpan();
                    }

                    SegmentDisposition disposition;
                    if (dispositionIndex >= line.Length || !Enum.TryParse(line[dispositionIndex], out disposition))
                    {
                        disposition = SegmentDisposition.Keep;
                    }

                    DateTime date;
                    if (dateIndex >= line.Length || !TryParseDate(line[dateIndex], out date))
                    {
                        date = DateTime.MinValue;
                    }

                    string subject = (subjectIndex < line.Length) ? line[subjectIndex] : string.Empty;
                    string title = (titleIndex < line.Length) ? line[titleIndex] : string.Empty;

                    // If first and position isn't zero, add a zero entry
                    if (list.Count == 0 && position.Ticks != 0L)
                    {
                        list.Add(new SegmentModel());
                    }

                    // If length is less than 1/2 second, drop this segment
                    if (list.Count > 0 && position.Subtract(list.Last().Position) < s_minimumSegment)
                    {
                        continue;
                    }

                    list.Add(new SegmentModel(position, disposition, date, subject, title));
                }

                return list;
            }
        }

        public static void SaveToCsv(IEnumerable<SegmentModel> list, string filename)
        {
            using (var writer = new StreamWriter(filename, false, new UTF8Encoding(false)))
            {
                writer.WriteLine($"{c_positionLabel},{c_dispositionLabel},{c_dateLabel},{c_subjectLabel},{c_titleLabel}");
                foreach(var scene in list)
                {
                    writer.WriteLine(String.Join(",",
                        FormatSeconds(scene.Position),
                        scene.Disposition.ToString(),
                        FormatDate(scene.Date),
                        CsvEncodeEx(scene.Subject),
                        CsvEncodeEx(scene.Title)));
                }
            }
        }

        static bool TryParseSeconds(string value, out TimeSpan ts)
        {
            Decimal dv;
            if (!Decimal.TryParse(value, out dv))
            {
                ts = new TimeSpan();
                return false;
            }
            ts = new TimeSpan((long)(dv * 10000000L));
            return true;
        }

        static string FormatSeconds(TimeSpan ts)
        {
            decimal ticks = ts.Ticks;
            ticks /= 10000000L;
            return ticks.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        static bool TryParseDate(string value, out DateTime dt)
        {
            return DateTime.TryParse(value,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind
                | System.Globalization.DateTimeStyles.AllowWhiteSpaces
                | System.Globalization.DateTimeStyles.NoCurrentDateDefault,
                out dt);
        }

        static string FormatDate(DateTime dt)
        {
            return dt.ToString("o", System.Globalization.CultureInfo.InvariantCulture);
        }

        // Encode as CSV, minimizing whitespace that is outside of quotes.
        static string CsvEncodeEx(string value)
        {
            var writer = new StringWriter();
            writer.Write('"');  // Opening quote
            using (var reader = new StringReader(value))
            {
                for (; ; )
                {
                    int ci = reader.Read();
                    if (ci < 0) break;
                    char c = (char)ci;
                    if (c == '"')
                    {
                        // Double up quotes.
                        writer.Write("\"\"");

                        // Write quoted text literally
                        for (; ; )
                        {
                            ci = reader.Read();
                            if (ci < 0) break;
                            c = (char)ci;
                            if (c == '"') break;
                            writer.Write(c);
                        }

                        // Double up close quote
                        writer.Write("\"\"");
                    }
                    else if (char.IsWhiteSpace(c))
                    {
                        // Condense whitespace outside of quotes
                        writer.Write(' ');
                        while (char.IsWhiteSpace((char)reader.Peek()))
                            reader.Read();
                    }
                    else
                    {
                        writer.Write(c);
                    }
                }
            }

            writer.Write('"');  // Closing quote
            return writer.ToString();
        }

    }

    enum SegmentDisposition : int
    {
        Delete = 0,
        Keep = 1,
        AddToPrevious = 2
    }
}
