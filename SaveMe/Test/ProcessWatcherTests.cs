using src.Services;
using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;
using src.Events;
using src.Models;
using System.Reflection;
using System.IO;

namespace Test
{
    public class ProcessWatcherTests
    {
        private readonly IProcessWatcher _sut;

        public ProcessWatcherTests()
        {
            _sut = new ProcessWatcher();
        }

        [Fact]
        public void _processStartEvent_EventArrived_FiresWhenExcelOpened()
        {
            // Arrange
            var currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var testFile = File.Create(@$"{currentPath}\testApplication.xls");
            testFile.Dispose();
            var app = new Excel.Application();

            _sut.NewProcessEvent += (s, e) =>
            {
                Assert.Equal(OfficeAppType.Excel, e.ProcessType);
            };

            app.Workbooks.Open(testFile.Name);
        }

        [Fact]
        public void _processStartEvent_EventArrived_FiresWhenWordOpened()
        {
            // Arrange
            var currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var testFile = File.Create(@$"{currentPath}\testApplication.docx");
            testFile.Dispose();
            var app = new Word.Application();

            _sut.NewProcessEvent += (s, e) =>
            {
                Assert.Equal(OfficeAppType.Word, e.ProcessType);
            };

            app.Documents.Open(testFile.Name);
        }

        [Fact]
        public void _processStartEvent_EventArrived_FiresWhenPowerPointOpened()
        {
            // Arrange
            var currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var testFile = File.Create(@$"{currentPath}\testApplication.pptx");
            testFile.Dispose();
            var app = new PowerPoint.Application();

            _sut.NewProcessEvent += (s, e) =>
            {
                Assert.Equal(OfficeAppType.PowerPoint, e.ProcessType);
            };

            app.Presentations.Open(testFile.Name);
        }
    }
}