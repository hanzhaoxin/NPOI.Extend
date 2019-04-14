using NPOI.Extend.Formula.State;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.Extend.Formula
{
    public class FormulaContext
    {
        public string FormulaText { get; private set; }

        public TextReader Reader { get; private set; }

        public Formula Formula { get; private set; }

        public IState State { get; set; } = StateHolder.FormulaState;
        

        public FormulaContext(string formula)
        {
            this.FormulaText = formula;
            this.Reader = new StringReader(formula);
            this.Formula = new Formula();
        }

        public void Parse()
        {
            while (!(this.State is EndState))
            {
                int i = this.Reader.Read();
                this.State.Switch(this,i);
                this.State.Handle(this,(char)i);
            }
        }
    }
}
