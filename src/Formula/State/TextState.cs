using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Extend.Formula.State
{
    class TextState : BaseState
    {
        public override void Handle(FormulaContext context, char c)
        {
            context.Formula.Append(PartType.Text, c);
        }

        public override void SwitchWhenEndCharacter(FormulaContext context)
        {
            throw new FormulaException($"formula exception:{context.FormulaText}");
        }

        public override void SwitchWhenEscapeCharacter(FormulaContext context)
        {
            context.State = StateHolder.LiteralInferState;
        }

        public override void SwitchWhenOtherCharacter(FormulaContext context)
        {
            
        }
    }
}
