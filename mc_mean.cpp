// mc_mean.cpp - mean value of a cell over a Monte Carlo simulation
#include "xll_mc.h"

struct mc_mean {
	double* state; // running mean, count
	mc_mean(double* state)
		: state(state)
	{ }
	void init()
	{
		state[0] = 0;
		state[1] = 0;
	}
	void next(double x)
	{
		state[1] += 1;
		state[0] += (x - state[1])/state[1];
	}
	void done()
	{ }
};

using namespace xll;

static AddIn xai_mean(
	Function(XLL_FP, "xll_mc_mean", "MC.MEAN")
	.Arguments({
		Arg(XLL_FP, "x", "is an array of numbers."),
		})
		.FunctionHelp("Return the mean of the numbers.")
	.Category(CATEGORY)
);
_FP12* WINAPI xll_mc_mean(_FP12* px)
{
#pragma XLLEXPORT
	static FPX m(1, 2);

	mc_mean m(m.array());

	m.next(px->array[0]);

	switch (state) {
	case INIT:
		m.init();
		break;
	case NEXT:
		m.next(px->array[0]);
		break;
	default:
		;
	}

	return m.get();
}