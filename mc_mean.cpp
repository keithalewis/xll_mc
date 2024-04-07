// mc_mean.cpp - mean value of a cell over a Monte Carlo simulation
#include "xll_mc.h"

using namespace xll;

void mean_init(LPXLOPER12& o, double x)
{

}

static AddIn xai_mc_mean(
	Function(XLL_LPOPER, L"xll_mc_mean", L"MC.MEAN")
	.Arguments({
		Arg(XLL_DOUBLE, "x", "is a number.")
		})
	.Uncalced()
	.FunctionHelp(L"Return the mean value of a cell over a Monte Carlo simulation.")
	.Category(CATEGORY)
);
LPOPER WINAPI xll_mc_mean(double x)
{
#pragma XLLEXPORT
	static OPER o;

	o = Excel(xlCoerce, Excel(xlfCaller));

	if (mc.state == INIT) {
		o[0] = Num(0);
		if (size(o) == 2) {
			o[1] = Num(0);
		}
	}
	else if (mc.state == NEXT) {
		double n = mc.count;
		if (size(o) == 2) {
			n = Num(o[1]);
		}
		Num(o[0]) += (x - Num(o[0])) / n;
	}

	return &o;
}