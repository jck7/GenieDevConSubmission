using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using ExcelGenie.Models;
using IOPath = System.IO.Path;

namespace ExcelGenie.Services
{
    public class RecentFilesService
    {
        private readonly string genieForExcelFolderPath;
        private readonly string recentFilesListPath;
        private readonly ObservableCollection<RecentFile> recentFiles;
        private const int MaxRecentFiles = 5;

        public RecentFilesService(string genieForExcelFolderPath)
        {
            this.genieForExcelFolderPath = genieForExcelFolderPath;
            this.recentFilesListPath = IOPath.Combine(genieForExcelFolderPath, "recentFiles.txt");
            this.recentFiles = new ObservableCollection<RecentFile>();
            LoadRecentFiles();
        }

        public ObservableCollection<RecentFile> GetRecentFiles()
        {
            return recentFiles;
        }

        public void LoadRecentFiles()
        {
            try
            {
                recentFiles.Clear();
                
                if (File.Exists(recentFilesListPath))
                {
                    var lines = File.ReadAllLines(recentFilesListPath);
                    foreach (var line in lines)
                    {
                        if (File.Exists(line))
                        {
                            recentFiles.Add(new RecentFile(line));
                        }
                    }

                    // Sort by last accessed time and take only the last MaxRecentFiles
                    var sortedFiles = recentFiles.OrderByDescending(f => f.LastAccessed).Take(MaxRecentFiles).ToList();
                    recentFiles.Clear();
                    foreach (var file in sortedFiles)
                    {
                        recentFiles.Add(file);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading recent files: {ex.Message}");
            }
        }

        public void SaveRecentFiles()
        {
            try
            {
                var paths = recentFiles.Select(f => f.FullPath)
                    .Where(p => !string.IsNullOrEmpty(p))
                    .ToList();
                File.WriteAllLines(recentFilesListPath, paths!);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving recent files: {ex.Message}");
            }
        }

        public void AddToRecentFiles(string filePath)
        {
            var existingFile = recentFiles.FirstOrDefault(f => f.FullPath == filePath);
            if (existingFile != null)
            {
                recentFiles.Remove(existingFile);
            }

            recentFiles.Insert(0, new RecentFile(filePath));

            // Keep only the last MaxRecentFiles files
            while (recentFiles.Count > MaxRecentFiles)
            {
                recentFiles.RemoveAt(recentFiles.Count - 1);
            }

            SaveRecentFiles();
        }

        public void RemoveRecentFile(string filePath)
        {
            var recentFile = recentFiles.FirstOrDefault(f => f.FullPath == filePath);
            if (recentFile != null)
            {
                recentFiles.Remove(recentFile);
                SaveRecentFiles();
            }
        }
    }
} 