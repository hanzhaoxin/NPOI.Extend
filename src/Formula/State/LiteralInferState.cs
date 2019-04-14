using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Extend.Formula.State
{
    class LiteralInferState : BaseState
    {
        public override void Handle(FormulaContext context, char c)
        {
            context.Formula.Append(PartType.Text, c);
        }

        public override void SwitchWhenEndCharacter(FormulaContext context)
        {
            context.State = StateHolder.EndState;
        }

        public override void SwitchWhenEscapeCharacter(FormulaContext context)
        {
            context.State = StateHolder.TextState;
        }

        public override void SwitchWhenOtherCharacter(FormulaContext context)
        {
            context.State = StateHolder.FormulaState;
        }
    }
}
