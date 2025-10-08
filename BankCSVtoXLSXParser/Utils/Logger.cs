// File: Utils/Logger.cs
using System;
using System.IO;

namespace BankCSVtoXLSXParser.Utils
{
    /// <summary>
    /// Provides lightweight, file-based logging for the application.
    /// </summary>
    /// <remarks>
    /// - Log entries are appended to a single text file located in the application's base directory.
    /// - All write failures are swallowed to avoid impacting the UI flow.
    /// - Intended for basic diagnostics rather than high-throughput scenarios.
    /// </remarks>
    public static class Logger
    {
        /// <summary>
        /// Gets the absolute path to the log file (error_log.txt) in the application's base directory.
        /// </summary>
        private static string LogPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error_log.txt");

        /// <summary>
        /// Appends an error entry to the log, including the location and exception details.
        /// </summary>
        /// <param name="location">Context where the error occurred (e.g., class/method name).</param>
        /// <param name="ex">The exception to log. Its message and stack trace are recorded.</param>
        /// <remarks>
        /// Timestamp format: yyyy-MM-dd HH:mm:ss.
        /// Any exceptions thrown while attempting to write the log are silently ignored.
        /// </remarks>
        public static void LogError(string location, Exception ex)
        {
            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {location}: {ex.Message}\n{ex.StackTrace}\n\n";
                File.AppendAllText(LogPath, logEntry);
            }
            catch
            {
                // Silently fail if can't write log
            }
        }

        /// <summary>
        /// Appends an informational entry to the log.
        /// </summary>
        /// <param name="message">The message to write.</param>
        /// <remarks>
        /// Timestamp format: yyyy-MM-dd HH:mm:ss.
        /// Any exceptions thrown while attempting to write the log are silently ignored.
        /// </remarks>
        public static void LogInfo(string message)
        {
            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] INFO: {message}\n";
                File.AppendAllText(LogPath, logEntry);
            }
            catch
            {
                // Silently fail if can't write log
            }
        }
    }
}