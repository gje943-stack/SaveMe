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
            var testFile = File.Create(@$"{currentPath}\testApplication.xlsx");
            var app = new Excel.Application();
            app.Visible = true;
            var expected = OfficeAppType.Word;
            _sut.NewProcessEvent += (s, e) => expected = e.ProcessType;

            // Act
            var wb = app.Workbooks.Open(testFile.Name);

            // Assert
            Assert.Equal(OfficeAppType.Excel, expected);
            
        }
    }
}
