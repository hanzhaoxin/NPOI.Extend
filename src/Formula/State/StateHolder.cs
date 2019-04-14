using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Extend.Formula.State
{
    static class StateHolder
    {
        public static IState FormulaState { get; } = new FormulaState();

        public static IState TextState { get; } = new TextState();

        public static IState LiteralInferState { get; } = new LiteralInferState();

        public static IState EndState { get; } = new EndState();
    }
}
