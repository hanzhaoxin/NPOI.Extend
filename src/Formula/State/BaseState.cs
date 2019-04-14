using NPOI.Extend.Formula.State;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Extend.Formula.State
{
    abstract class BaseState : IState
    {
        public abstract void Handle(FormulaContext context, char c);
        public void Switch(FormulaContext context, int i)
        {
            if (-1 == i)
            {
                SwitchWhenEndCharacter(context);
            }
            else if ('"' == (char)i)
            {
                SwitchWhenEscapeCharacter(context);
            }
            else
            {
                SwitchWhenOtherCharacter(context);
            }
        }

        public abstract void SwitchWhenEndCharacter(FormulaContext context);

        public abstract void SwitchWhenEscapeCharacter(FormulaContext context);

        public abstract void SwitchWhenOtherCharacter(FormulaContext context);
    }
}
