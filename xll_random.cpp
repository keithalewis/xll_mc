// xll_random.cpp - random number generator
#include <random>
#include "xll.h"

static std::random_device rd;
static std::default_random_engine dre(rd());
static std::uniform_real_distribution<double> urd;

using namespace xll;

static AddIn xai_std_rand(
	Function(XLL_DOUBLE, "xll_std_rand", "STD.RAND")
	.Arguments({
		Arg(XLL_HANDLEX, L"rng", L"is a handle to a PCG random number generator."),
		})
	.Volatile()
	.Category("STD")
	.FunctionHelp("Uniform [0, 1) random variate.")
);
double WINAPI xll_std_rand(HANDLEX h)
{
#pragma XLLEXPORT 
	if (h == 0) {
		return urd(dre);
	}
	else {
		return 0;
	}
}