using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word
{
    public enum JustificationValues
    {
        //
        // 摘要:
        //     Align Left.
        //     When the item is serialized out as xml, its value is "left".
        Left = 0,
        //
        // 摘要:
        //     start.
        //     When the item is serialized out as xml, its value is "start".
        //     This item is only available in Office2010.
        Start = 1,
        //
        // 摘要:
        //     Align Center.
        //     When the item is serialized out as xml, its value is "center".
        Center = 2,
        //
        // 摘要:
        //     Align Right.
        //     When the item is serialized out as xml, its value is "right".
        Right = 3,
        //
        // 摘要:
        //     end.
        //     When the item is serialized out as xml, its value is "end".
        //     This item is only available in Office2010.
        End = 4,
        //
        // 摘要:
        //     Justified.
        //     When the item is serialized out as xml, its value is "both".
        Both = 5,
        //
        // 摘要:
        //     Medium Kashida Length.
        //     When the item is serialized out as xml, its value is "mediumKashida".
        MediumKashida = 6,
        //
        // 摘要:
        //     Distribute All Characters Equally.
        //     When the item is serialized out as xml, its value is "distribute".
        Distribute = 7,
        //
        // 摘要:
        //     Align to List Tab.
        //     When the item is serialized out as xml, its value is "numTab".
        NumTab = 8,
        //
        // 摘要:
        //     Widest Kashida Length.
        //     When the item is serialized out as xml, its value is "highKashida".
        HighKashida = 9,
        //
        // 摘要:
        //     Low Kashida Length.
        //     When the item is serialized out as xml, its value is "lowKashida".
        LowKashida = 10,
        //
        // 摘要:
        //     Thai Language Justification.
        //     When the item is serialized out as xml, its value is "thaiDistribute".
        ThaiDistribute = 11
    }
}
