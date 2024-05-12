// table.cpp - table of values
#include "xll_mc.h"

using namespace xll;

AddIn xai_mc_range(
    Function(XLL_LPXLOPER, "xll_mc_range", CATEGORY ".RANGE")
    .Arguments({
        Arg(XLL_LPXLOPER, "range", "is a a 2-dimensional range of cells."),
        Arg(XLL_BOOL, "_cycle", "is an optional boolean indicating cycling over rows.")
        })
    .Uncalced()
    .Category(CATEGORY)
    .FunctionHelp("Return rows of a range during a simulation.")
    /*
    .Documentation(R"(Return the <i>n</i>-th row of a range on iteration <i>n</i>.
<p>
<code>mc.RANGE</code> allows for data driven simulations. By default a simulation
stops when the last row is returned. If <code>_cyclic</code> is <code>TRUE</code>
then simulation continues starting from the first row. 
<code>mc.RANGE</code>
returns the same number of rows as the output range.
</p>
)")
*/
);
LPXLOPER12 WINAPI xll_mc_range(LPXLOPER12 po, BOOL cyclic)
{
#pragma XLLEXPORT
    static XLOPER12 o = { .xltype = xltypeMulti };

    try {
        int r = rows(*po);
        int c = columns(*po);
        int r0 = rows(Excel(xlfCaller));

        o.val.array.rows = r0;
        o.val.array.columns = c;

        int count = mc.count;
        if (mc.state == INIT) {
            count = 0;
        }
        if (count == 0) {
            count = 1;
        }

        int n = (count - 1) % r;
        if (po->xltype == xltypeMulti) {
            o.val.array.lparray = po->val.array.lparray + n * c;
        }
        else {
            o = Excel(xlfOffset, *po, n, 0, r0, c);
        }

        if (!cyclic && (r + 1 == count + r0)) {
            mc.state_ = STOP;
        }
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        o = ErrNA;
    }

    return &o;
}