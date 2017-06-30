using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word
{

    //
    // 摘要:
    //     Defines the UnderlineValues enumeration.
    public enum UnderlineValues
    {
        //
        // 摘要:
        //     Single Underline.
        //     When the item is serialized out as xml, its value is "single".
        Single = 0,
        //
        // 摘要:
        //     Underline Non-Space Characters Only.
        //     When the item is serialized out as xml, its value is "words".
        Words = 1,
        //
        // 摘要:
        //     Double Underline.
        //     When the item is serialized out as xml, its value is "double".
        Double = 2,
        //
        // 摘要:
        //     Thick Underline.
        //     When the item is serialized out as xml, its value is "thick".
        Thick = 3,
        //
        // 摘要:
        //     Dotted Underline.
        //     When the item is serialized out as xml, its value is "dotted".
        Dotted = 4,
        //
        // 摘要:
        //     Thick Dotted Underline.
        //     When the item is serialized out as xml, its value is "dottedHeavy".
        DottedHeavy = 5,
        //
        // 摘要:
        //     Dashed Underline.
        //     When the item is serialized out as xml, its value is "dash".
        Dash = 6,
        //
        // 摘要:
        //     Thick Dashed Underline.
        //     When the item is serialized out as xml, its value is "dashedHeavy".
        DashedHeavy = 7,
        //
        // 摘要:
        //     Long Dashed Underline.
        //     When the item is serialized out as xml, its value is "dashLong".
        DashLong = 8,
        //
        // 摘要:
        //     Thick Long Dashed Underline.
        //     When the item is serialized out as xml, its value is "dashLongHeavy".
        DashLongHeavy = 9,
        //
        // 摘要:
        //     Dash-Dot Underline.
        //     When the item is serialized out as xml, its value is "dotDash".
        DotDash = 10,
        //
        // 摘要:
        //     Thick Dash-Dot Underline.
        //     When the item is serialized out as xml, its value is "dashDotHeavy".
        DashDotHeavy = 11,
        //
        // 摘要:
        //     Dash-Dot-Dot Underline.
        //     When the item is serialized out as xml, its value is "dotDotDash".
        DotDotDash = 12,
        //
        // 摘要:
        //     Thick Dash-Dot-Dot Underline.
        //     When the item is serialized out as xml, its value is "dashDotDotHeavy".
        DashDotDotHeavy = 13,
        //
        // 摘要:
        //     Wave Underline.
        //     When the item is serialized out as xml, its value is "wave".
        Wave = 14,
        //
        // 摘要:
        //     Heavy Wave Underline.
        //     When the item is serialized out as xml, its value is "wavyHeavy".
        WavyHeavy = 15,
        //
        // 摘要:
        //     Double Wave Underline.
        //     When the item is serialized out as xml, its value is "wavyDouble".
        WavyDouble = 16,
        //
        // 摘要:
        //     No Underline.
        //     When the item is serialized out as xml, its value is "none".
        None = 17
    }
}
