// mc_buffer.cpp - buffer results of a Monte Carlo simulation
#include "xll_mc.h"

using namespace xll;

static AddIn xai_mc_buffer(
	Function(XLL_LPOPER, "xll_mc_buffer", "MC.BUFFER")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is a cell or range to buffer."),
		//Arg(XLL_LONG, "n", "is the number rows to buffer."),
	})
	.Uncalced()
	.FunctionHelp(L"Buffer results of a Monte Carlo simulation.")
	.Category(L"MC")
);
LPOPER WINAPI xll_mc_buffer(LPOPER pr/*, LONG n*/)
{
#pragma XLLEXPORT
	static OPER r(10, 1, nullptr);

	try {
		if (mc.state == INIT) {
			for (int i = 0; i < rows(r); ++i) {
				r[i] = L"";
			}
		}
		else if (mc.state == NEXT) {
			r[rows(r) - 1] = *pr;
			std::rotate(r.val.array.lparray, r.val.array.lparray + rows(r) - 1, r.val.array.lparray + rows(r));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &r;
}