using System;
using System.Collections.Generic;
using System.Text;

namespace src.Services
{
    public static class OfficeProcesses
    {
        public static List<string> NamesWithoutExe { get; private set; } = new List<string> { "EXCEL", "POWERPOINT", "WORD" };
        public static List<string> NamesWithExe { get; private set; } = new List<string> { "EXCEL.EXE", "POWERPOINT.EXE", "WORD.EXE" };
    }
}