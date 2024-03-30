// mc_buffer.cpp - buffer results of a Monte Carlo simulation
#include "xll_mc.h"

using namespace xll;

static AddIn xai_mc_buffer(
	Function(XLL_LPOPER, "xll_mc_buffer", "MC.BUFFER")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is a cell or range to buffer."),
	})
	.Uncalced()
	.FunctionHelp(L"Buffer results of a Monte Carlo simulation to an output range.")
	.Category(L"MC")
);
LPOPER WINAPI xll_mc_buffer(LPOPER pr)
{
#pragma XLLEXPORT
	static OPER r;

	try {
		r = Excel(xlCoerce, Excel(xlfCaller));
		if (mc.state == IDLE) {
			ensure_message(rows(*pr) == 1,
				__FUNCTION__ ": range must have one row");
			ensure_message(columns(r) == columns(*pr),
				__FUNCTION__ ": range and ouput range must have the same number of columns");
		}

		if (mc.state == INIT) {
			for (int i = 0; i < size(r); ++i) {
				r[i] = L"";
			}
		}
		else if (mc.state == NEXT) {
			if (type(*pr) != xltypeMulti) {
				r(rows(r) - 1, 0) = *pr;
			}
			else {
				for (int j = 0; j < columns(r); ++j) {
					r(rows(r) - 1, j) = pr->val.array.lparray[j];
				}
			}
			auto a = r.val.array.lparray;
			std::rotate(a, a + size(r) - columns(r), a + size(r));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &r;
}