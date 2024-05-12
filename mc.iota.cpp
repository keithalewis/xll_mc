// mc_mean.cpp - mean value of a cell over a Monte Carlo simulation
#include "xll_mc.h"

using namespace xll;

static AddIn xai_mc_iota(
    Function(XLL_DOUBLE, L"xll_mc_iota", L"MC.IOTA")
        .Arguments({
            Arg(XLL_DOUBLE, "start", "is a number."),
            Arg(XLL_DOUBLE, "step", "is an optional step size. Default is 1."),
            })
        .Volatile()
        .FunctionHelp(L"Return start + count * step.")
        .Category(CATEGORY));
double WINAPI xll_mc_iota(double x, double dx)
{
#pragma XLLEXPORT
     return x + mc.count * (dx ? dx : 1);
}