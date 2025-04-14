using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EepromReader.Scripts
{
    public class EepromMapping
    {
        public int Address { get; set; }
        public string Name { get; set; }
        public int Length { get; set; }
        public object Value { get; set; } // Stores the value, which can be cast based on its type.
        public Type EepromDataType { get; set; } // Stores the data type (e.g., byte, int, etc.)

        // Optional properties for multi-dimensional arrays
        public int? RowCount { get; set; } // Only used for 2D or 3D arrays
        public int? ColumnCount { get; set; } // Only used for 2D or 3D arrays
        public int? Depth { get; set; } // Only used for 3D arrays
    }
}
