// mc_mean.cpp - mean value of a cell over a Monte Carlo simulation
#include "xll_mc.h"

struct mc_mean {
	double n, x; // count and running mean
	mc_mean(double x, double n)
		: n(n), x(x)
	{ }
	void init()
	{
		n = 0;
		x = 0;
	}
	void next(double x_)
	{
		n += 1;
		x += (x - x_)/n;
	}
	void done()
	{ }
};

using namespace xll;

static AddIn xai_mc_mean(
	Function(XLL_LPOPER, "xll_mc_mean", "MC.MEAN")
	.Arguments({
		Arg(XLL_FP, "x", "is an array of random variables."),
	})
	.FunctionHelp("Return the mean of the random variables.")
	.Category(CATEGORY)
);
LPOPER WINAPI xll_mc_mean(_FP12* px)
{
#pragma XLLEXPORT
	static OPER o;

	o = Excel(xlCoerce, Excel(xlfCaller));

	if (state == NEXT) {
		mc_mean m(px->array);
		m.next();
	}

	return &o;
}