
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word
{
    //
    // 摘要:
    //     Defines the ShadingPatternValues enumeration.
    public enum ShadingPatternValues
    {
        //
        // 摘要:
        //     No Pattern.
        //     When the item is serialized out as xml, its value is "nil".
        Nil = 0,
        //
        // 摘要:
        //     No Pattern.
        //     When the item is serialized out as xml, its value is "clear".
        Clear = 1,
        //
        // 摘要:
        //     100% Fill Pattern.
        //     When the item is serialized out as xml, its value is "solid".
        Solid = 2,
        //
        // 摘要:
        //     Horizontal Stripe Pattern.
        //     When the item is serialized out as xml, its value is "horzStripe".
        HorizontalStripe = 3,
        //
        // 摘要:
        //     Vertical Stripe Pattern.
        //     When the item is serialized out as xml, its value is "vertStripe".
        VerticalStripe = 4,
        //
        // 摘要:
        //     Reverse Diagonal Stripe Pattern.
        //     When the item is serialized out as xml, its value is "reverseDiagStripe".
        ReverseDiagonalStripe = 5,
        //
        // 摘要:
        //     Diagonal Stripe Pattern.
        //     When the item is serialized out as xml, its value is "diagStripe".
        DiagonalStripe = 6,
        //
        // 摘要:
        //     Horizontal Cross Pattern.
        //     When the item is serialized out as xml, its value is "horzCross".
        HorizontalCross = 7,
        //
        // 摘要:
        //     Diagonal Cross Pattern.
        //     When the item is serialized out as xml, its value is "diagCross".
        DiagonalCross = 8,
        //
        // 摘要:
        //     Thin Horizontal Stripe Pattern.
        //     When the item is serialized out as xml, its value is "thinHorzStripe".
        ThinHorizontalStripe = 9,
        //
        // 摘要:
        //     Thin Vertical Stripe Pattern.
        //     When the item is serialized out as xml, its value is "thinVertStripe".
        ThinVerticalStripe = 10,
        //
        // 摘要:
        //     Thin Reverse Diagonal Stripe Pattern.
        //     When the item is serialized out as xml, its value is "thinReverseDiagStripe".
        ThinReverseDiagonalStripe = 11,
        //
        // 摘要:
        //     Thin Diagonal Stripe Pattern.
        //     When the item is serialized out as xml, its value is "thinDiagStripe".
        ThinDiagonalStripe = 12,
        //
        // 摘要:
        //     Thin Horizontal Cross Pattern.
        //     When the item is serialized out as xml, its value is "thinHorzCross".
        ThinHorizontalCross = 13,
        //
        // 摘要:
        //     Thin Diagonal Cross Pattern.
        //     When the item is serialized out as xml, its value is "thinDiagCross".
        ThinDiagonalCross = 14,
        //
        // 摘要:
        //     5% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct5".
        Percent5 = 15,
        //
        // 摘要:
        //     10% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct10".
        Percent10 = 16,
        //
        // 摘要:
        //     12.5% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct12".
        Percent12 = 17,
        //
        // 摘要:
        //     15% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct15".
        Percent15 = 18,
        //
        // 摘要:
        //     20% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct20".
        Percent20 = 19,
        //
        // 摘要:
        //     25% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct25".
        Percent25 = 20,
        //
        // 摘要:
        //     30% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct30".
        Percent30 = 21,
        //
        // 摘要:
        //     35% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct35".
        Percent35 = 22,
        //
        // 摘要:
        //     37.5% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct37".
        Percent37 = 23,
        //
        // 摘要:
        //     40% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct40".
        Percent40 = 24,
        //
        // 摘要:
        //     45% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct45".
        Percent45 = 25,
        //
        // 摘要:
        //     50% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct50".
        Percent50 = 26,
        //
        // 摘要:
        //     55% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct55".
        Percent55 = 27,
        //
        // 摘要:
        //     60% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct60".
        Percent60 = 28,
        //
        // 摘要:
        //     62.5% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct62".
        Percent62 = 29,
        //
        // 摘要:
        //     65% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct65".
        Percent65 = 30,
        //
        // 摘要:
        //     70% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct70".
        Percent70 = 31,
        //
        // 摘要:
        //     75% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct75".
        Percent75 = 32,
        //
        // 摘要:
        //     80% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct80".
        Percent80 = 33,
        //
        // 摘要:
        //     85% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct85".
        Percent85 = 34,
        //
        // 摘要:
        //     87.5% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct87".
        Percent87 = 35,
        //
        // 摘要:
        //     90% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct90".
        Percent90 = 36,
        //
        // 摘要:
        //     95% Fill Pattern.
        //     When the item is serialized out as xml, its value is "pct95".
        Percent95 = 37
    }
}
