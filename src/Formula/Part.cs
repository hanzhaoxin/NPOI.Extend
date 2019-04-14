using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Extend.Formula
{
    public class Part
    {
        private StringBuilder builder = new StringBuilder();

        public PartType Type { get; private set; }

        public Part(PartType type)
        {
            this.Type = type;
        }

        public void Append(char c)
        {
            builder.Append(c);
        }

        
        public override string ToString()
        {
            return builder.ToString();
        }
    }
}
