using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CodeBit
{
    static class PathEx
    {
        public static string[] GetFilesByPattern(string pattern, bool recursive = false)
        {
            string folder = Path.GetDirectoryName(pattern);
            folder = (string.IsNullOrEmpty(folder)) ? Environment.CurrentDirectory : Path.GetFullPath(folder);
            return Directory.GetFiles(folder, Path.GetFileName(pattern));
        }
    }
}
