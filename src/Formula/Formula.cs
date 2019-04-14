using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Extend.Formula
{
    public class Formula:IEnumerable<Part>
    {
        private LinkedList<Part> parts = new LinkedList<Part>();

        private Part currentPart;

        public void Append(PartType type, char c)
        {
            if (currentPart == null || (!currentPart.Type.Equals(type)))
            {
                currentPart = new Part(type);
                parts.AddLast(currentPart);
            }

            currentPart.Append(c);
        }

        

        public override string ToString()
        {
            return this.ToString(part => part.ToString());
        }

        public string ToString(Func<Part, string> func)
        {
            StringBuilder builder = new StringBuilder();
            foreach (var part in parts)
            {
                builder.Append(func(part));
            }
            return builder.ToString();
        }

        public IEnumerator<Part> GetEnumerator()
        {
            return this.parts.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.parts.GetEnumerator();
        }
    }
}
