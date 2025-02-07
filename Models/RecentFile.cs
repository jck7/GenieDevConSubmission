using System;
using System.IO;
using IOPath = System.IO.Path;

namespace ExcelGenie.Models
{
    public class RecentFile
    {
        public string? Name { get; set; }
        public string? FullPath { get; set; }
        public DateTime LastAccessed { get; set; }

        public RecentFile(string fullPath)
        {
            FullPath = fullPath;
            Name = IOPath.GetFileName(fullPath);
            LastAccessed = File.GetLastAccessTime(fullPath);
        }
    }
} 