using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Extend.Formula.State
{
    public interface IState
    {
        void Switch(FormulaContext context,int i);

        void Handle(FormulaContext context, char c);
    }
}
