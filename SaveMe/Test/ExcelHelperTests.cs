using Microsoft.Office.Interop.Excel;
using src.Services.Helpers;
using System;
using Xunit;

namespace Test
{
    public class ExcelHelperTests
    {
        [Fact]
        public void GetOpenApplicationNames_ReturnsCorrectNames()
        {
            // Arrange
            var sut = new ExcelHelper();
            var app = new Application();
            app.Workbooks.Open("test1");
            sut.OpenExcelInstances = app.Workbooks;
            var expected = ("Excel", "test1");

            // Act + Assert
            foreach (var (appName, instanceName) in sut.GetOpenApplicationNames())
            {
                Assert.Equal(expected.Item1, appName);
                Assert.Equal(expected.Item2, instanceName);
            };
        }
    }
}