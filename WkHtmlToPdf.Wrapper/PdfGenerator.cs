using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace WkHtmlToPdf.Wrapper
{
    /// <summary>
    /// PDF generator - wrapper for the WkHtmlToPdf executable.
    /// </summary>
    /// <example>var pdfBytes = new PdfGenerator().GeneratePdf("https://www.google.com/", new PdfGenerator.Settings { PageOrientation = PdfGenerator.PageOrientation.Landscape });</example>
    public class PdfGenerator
    {
        /// <summary>
        /// The file path (directory+filename) field.
        /// </summary>
        private string filePath = null;

        /// <summary>
        /// Gets or sets the WkHtmlToPdf executable file path (directory+filename).
        /// </summary>
        /// <value>
        /// The WkHtmlToPdf executable file path (directory+filename).
        /// </value>
        public string WkHtmlToPdfExeFilePath
        {
            get
            {
                if (string.IsNullOrEmpty(this.filePath))
                    this.filePath = this.GetExecutableFullPath();

                return this.filePath;
            }
            set
            {
                this.filePath = value;
            }
        }

        /// <summary>
        /// Generates the PDF to the byte array.
        /// </summary>
        /// <param name="url">The URL of the page to generate.</param>
        /// <param name="settings">The generator settings.</param>
        /// <returns>Byte array containing generated PDF.</returns>
        /// <exception cref="System.Exception">Exception containing error message from WkHtmlToPdf.</exception>
        public byte[] GeneratePdf(string url, Settings settings = null)
        {
            var wkhtmlPath = Path.GetDirectoryName(this.WkHtmlToPdfExeFilePath);
            var wkhtmlFileName = Path.GetFileName(this.WkHtmlToPdfExeFilePath);

            var switches = settings?.GetSwitches();
            var arguments = $"-q {switches} {url} -";

            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = Path.Combine(wkhtmlPath, wkhtmlFileName),
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    RedirectStandardInput = true,
                    WorkingDirectory = wkhtmlPath,
                    CreateNoWindow = true
                }
            };
            proc.Start();

            using (var ms = new MemoryStream())
            {
                using (var wkhtmlOutput = proc.StandardOutput.BaseStream)
                {
                    wkhtmlOutput.CopyTo(ms);
                }

                var errorText = proc.StandardError.ReadToEnd();
                if (ms.Length == 0)
                    throw new Exception(errorText);

                proc.WaitForExit();
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Generates the PDF to the file.
        /// </summary>
        /// <param name="url">The URL of the page to generate.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="settings">The generator settings.</param>
        /// <exception cref="System.Exception">Exception containing error message from WkHtmlToPdf.</exception>
        public void GeneratePdfToFile(string url, string fileName, Settings settings = null)
        {
            var wkhtmlPath = Path.GetDirectoryName(this.WkHtmlToPdfExeFilePath);
            var wkhtmlFileName = Path.GetFileName(this.WkHtmlToPdfExeFilePath);

            var switches = settings?.GetSwitches();
            var arguments = $"-q {switches} {url} \"{fileName}\"";

            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = Path.Combine(wkhtmlPath, wkhtmlFileName),
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    RedirectStandardInput = true,
                    WorkingDirectory = wkhtmlPath,
                    CreateNoWindow = true
                }
            };
            proc.Start();

            using (var ms = new MemoryStream())
            {
                using (var wkhtmlOutput = proc.StandardOutput.BaseStream)
                {
                    wkhtmlOutput.CopyTo(ms);
                }

                var errorText = proc.StandardError.ReadToEnd();
                if (ms.Length == 0)
                    throw new Exception(errorText);

                proc.WaitForExit();

            }
        }

        //public byte[] GeneratePdfFromHtml(string html, Settings settings = null)
        //{
        //    throw new NotImplementedException();
        //}

        //public void GeneratePdfToFileFromHtml(string html, string fileName, Settings settings = null)
        //{
        //    throw new NotImplementedException();
        //}

        /// <summary>
        /// Gets the name of the executable full path - tries to search known locations where wkhtmltopdf.exe would usually reside.
        /// </summary>
        /// <returns>Full path (directory and filename) to the executable.</returns>
        /// <exception cref="Exception">Unable to locate WkHtmlToPdf executable. Please specify full filename for the WkHtmlToPdf executable</exception>
        private string GetExecutableFullPath()
        {
            string wkhtmlPath = null;
            string wkhtmlFullFileName = null;
            const string defaultWkhtmlExeName = "wkhtmltopdf.exe";

            // search current executable path
            wkhtmlPath = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
            if (wkhtmlPath is null)
                wkhtmlPath = AppDomain.CurrentDomain.BaseDirectory;

            wkhtmlFullFileName = Path.Combine(wkhtmlPath, defaultWkhtmlExeName);
            if (File.Exists(wkhtmlFullFileName))
                return wkhtmlFullFileName;

            // search current executable path +/wkhtmltopdf
            wkhtmlFullFileName = Path.Combine(wkhtmlPath, "wkhtmltopdf", defaultWkhtmlExeName);
            if (File.Exists(wkhtmlFullFileName))
                return wkhtmlFullFileName;

            // search programFiles
            wkhtmlPath = Environment.ExpandEnvironmentVariables("%ProgramW6432%");
            if (!string.IsNullOrEmpty(wkhtmlPath))
            {
                wkhtmlFullFileName = Path.Combine(wkhtmlPath, "wkhtmltopdf", "bin", defaultWkhtmlExeName);
                if (File.Exists(wkhtmlFullFileName))
                    return wkhtmlFullFileName;
            }

            wkhtmlPath = Environment.ExpandEnvironmentVariables("%ProgramFiles(x86)%");
            if (!string.IsNullOrEmpty(wkhtmlPath))
            {
                wkhtmlFullFileName = Path.Combine(wkhtmlPath, "wkhtmltopdf", "bin", defaultWkhtmlExeName);
                if (File.Exists(wkhtmlFullFileName))
                    return wkhtmlFullFileName;
            }

            wkhtmlPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            wkhtmlFullFileName = Path.Combine(wkhtmlPath, "wkhtmltopdf", "bin", defaultWkhtmlExeName);
            if (File.Exists(wkhtmlFullFileName))
                return wkhtmlFullFileName;

            wkhtmlPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            wkhtmlFullFileName = Path.Combine(wkhtmlPath, "wkhtmltopdf", "bin", defaultWkhtmlExeName);
            if (File.Exists(wkhtmlFullFileName))
                return wkhtmlFullFileName;

            throw new Exception("Unable to locate WkHtmlToPdf executable. Please specify full filename for the WkHtmlToPdf executable");
        }

        /// <summary>
        /// Settings for the PDF generator. 
        /// These settings are mapped to the WkHtmlToPdf executables' switches.
        /// </summary>
        public class Settings
        {
            /// <summary>
            /// Gets or sets the username to be used for authentication.
            /// Whis works for both Forms and Windows authentications.
            /// </summary>
            /// <value>
            /// The username.
            /// </value>
            /// <remarks>--username</remarks>
            public string Username { get; set; } = null;

            /// <summary>
            /// Gets or sets the password to be used for authentication.
            /// Whis works for both Forms and Windows authentications.
            /// </summary>
            /// <value>
            /// The password.
            /// </value>
            /// <remarks>--password</remarks>
            public string Password { get; set; } = null;

            /// <summary>
            /// Gets or sets the bottom margin (in mm).
            /// Default is 10mm.
            /// </summary>
            /// <value>
            /// The bottom margin.
            /// </value>
            /// <remarks>-B {value}</remarks>
            public int BottomMargin { get; set; } = 10;

            /// <summary>
            /// Gets or sets the left margin (in mm).
            /// Default is 10mm.
            /// </summary>
            /// <value>
            /// The left margin.
            /// </value>
            /// <remarks>-L {value}</remarks>
            public int LeftMargin { get; set; } = 10;

            /// <summary>
            /// Gets or sets the right margin (in mm).
            /// Default is 10mm.
            /// </summary>
            /// <value>
            /// The right margin.
            /// </value>
            /// <remarks>-R {value}</remarks>
            public int RightMargin { get; set; } = 10;

            /// <summary>
            /// Gets or sets the top margin (in mm).
            /// Default is 10mm.
            /// </summary>
            /// <value>
            /// The top margin.
            /// </value>
            /// <remarks>-T {value}</remarks>
            public int TopMargin { get; set; } = 10;

            /// <summary>
            /// Gets or sets the page orientation.
            /// Default is Portrait.
            /// </summary>
            /// <value>
            /// The page orientation.
            /// </value>
            /// <remarks>--orientation {value}</remarks>
            public PageOrientation PageOrientation { get; set; } = PageOrientation.Portrait;

            /// <summary>
            /// Gets or sets a value indicating whether to use the 'print media type'.
            /// Default is <c>false</c>.
            /// </summary>
            /// <value>
            ///   <c>True</c> if use print media type; otherwise, <c>false</c>.
            /// </value>
            /// <remarks>--print-media-type or --no-print-media-type</remarks>
            public bool UsePrintMediaType { get; set; } = false;

            /// <summary>
            /// Gets or sets a value indicating whether to disable smart shrinking.
            /// Default is <c>false</c>.
            /// </summary>
            /// <value>
            ///   <c>True</c> if disable smart shrinking; otherwise, <c>false</c>.
            /// </value>
            /// <remarks>--disable-smart-shrinking or --enable-smart-shrinking</remarks>
            public bool DisableSmartShrinking { get; set; } = false;

            /// <summary>
            /// Gets or sets the javascript delay (in miliseconds).
            /// Default is 200ms.
            /// </summary>
            /// <value>
            /// The javascript delay (in miliseconds).
            /// </value>
            /// <remarks>--javascript-delay {value}</remarks>
            public int JavascriptDelay { get; set; } = 200;

            /// <summary>
            /// Gets the switches string.
            /// </summary>
            /// <returns>Combined switches string (without url).</returns>
            internal string GetSwitches()
            {
                var lstSwitches = new List<string>();
                var defaultSettings = new Settings();

                if (defaultSettings.Username != this.Username)
                    lstSwitches.Add($"--username \"{this.Username}\"");

                if (defaultSettings.Password != this.Password)
                    lstSwitches.Add($"--password \"{this.Password}\"");

                if (defaultSettings.BottomMargin != this.BottomMargin)
                    lstSwitches.Add($"-B {this.BottomMargin.ToString("F0", CultureInfo.InvariantCulture)}");

                if (defaultSettings.LeftMargin != this.LeftMargin)
                    lstSwitches.Add($"-L {this.LeftMargin.ToString("F0", CultureInfo.InvariantCulture)}");

                if (defaultSettings.RightMargin != this.RightMargin)
                    lstSwitches.Add($"-R {this.RightMargin.ToString("F0", CultureInfo.InvariantCulture)}");

                if (defaultSettings.TopMargin != this.TopMargin)
                    lstSwitches.Add($"-T {this.TopMargin.ToString("F0", CultureInfo.InvariantCulture)}");

                if (defaultSettings.PageOrientation != this.PageOrientation)
                    lstSwitches.Add($"--orientation {this.PageOrientation.ToString()}");

                if (defaultSettings.UsePrintMediaType != this.UsePrintMediaType)
                    lstSwitches.Add(this.UsePrintMediaType ? "--print-media-type" : "--no-print-media-type");

                if (defaultSettings.DisableSmartShrinking != this.DisableSmartShrinking)
                    lstSwitches.Add(this.DisableSmartShrinking ? "--disable-smart-shrinking" : "--enable-smart-shrinking");

                if (defaultSettings.JavascriptDelay != this.JavascriptDelay)
                    lstSwitches.Add($"--javascript-delay {this.JavascriptDelay.ToString("F0", CultureInfo.InvariantCulture)}");

                return string.Join(" ", lstSwitches);
            }
        }

        /// <summary>
        /// Page orientation enumeration.
        /// </summary>
        public enum PageOrientation
        {
            /// <summary>
            /// The portrait orientation.
            /// </summary>
            Portrait,

            /// <summary>
            /// The landscape orientation.
            /// </summary>
            Landscape
        }
    }
}
