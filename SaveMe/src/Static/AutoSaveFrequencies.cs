using System;
using System.Collections.Generic;
using System.Text;

namespace src.Static
{
    public static class AutoSaveFrequencies
    {

        /// <summary>
        ///     A dictonary with
        ///     Key: autosave frequencies in string format.
        ///     Value: autosave frequencies in milliseconds.
        /// </summary>
        public static Dictionary<string, int> Frequencies { get; private set; } = new()
        {
            { "10 Minutes", 10000 },
            { "30 Minutes", 30000 },
            { "1 Hour", 60000 }
        };
    }
}
