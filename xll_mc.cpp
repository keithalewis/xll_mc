// xll_mc.cpp - Monte Carlo simulation
#include <cmath>
#include "xll_mc.h"

using namespace xll;

monte mc;

static AddIn xai_count(
	Function(XLL_LONG, "xll_count", "MC.COUNT")
	.Arguments({
		Arg(XLL_LONG, "count", "is the optional number of iterations to run.")
	})
	.Volatile()
	.FunctionHelp("Return current number of iterations run. Set number of iterations if non-zero.")
	.Category(CATEGORY)
);
long WINAPI xll_count(long n)
{
#pragma XLLEXPORT
	if (n) {
		mc.limit = n;
	}

	return mc.count;
}

static AddIn xai_elapsed(
	Function(XLL_DOUBLE, "xll_elapsed", "MC.ELAPSED")
	.Arguments({})
	.Volatile()
	.FunctionHelp("Return elapsed time in seconds.")
	.Category(CATEGORY)
);
double WINAPI xll_elapsed()
{
#pragma XLLEXPORT
	return mc.elapsed_seconds();
}

static AddIn xai_stop(
	Function(XLL_BOOL, "xll_stop", "MC.STOP")
	.Arguments({
		Arg(XLL_BOOL, "stop", "stop the simulation.")
		})
	.Category(CATEGORY)
);
BOOL WINAPI xll_stop(BOOL b)
{
#pragma XLLEXPORT
	if (b && mc.state == NEXT) {
		mc.state_ = STOP;
	}

	return b;
}

static AddIn xai_pause(
	Function(XLL_BOOL, "xll_pause", "MC.PAUSE")
	.Arguments({
		Arg(XLL_BOOL, "pause", "pause the simulation.")
		})
	.Category(CATEGORY)
);
BOOL WINAPI xll_pause(BOOL b)
{
#pragma XLLEXPORT
	if (b && mc.state == NEXT) {
		mc.state_ = HALT;
	}

	return b;
}

// Reset a Monte Carlo simulation.
static AddIn xai_reset(Macro("xll_reset", "MC.RESET"));
int WINAPI xll_reset()
{
#pragma XLLEXPORT
	mc.reset();

	return 1;
}
/*
// On<Key> xlo_monte_reset(ON_CTRL "T", "MONTE.RESET");
Auto <OpenAfter> xaoa_monte_reset([]() {
	Excel(xlcOnKey, OPER(ON_CTRL L"T"), OPER(L"MC.RESET"));
	return TRUE;
	});
Auto <CloseBefore> xacb_monte_reset([]() {
	Excel(xlcOnKey, OPER(ON_CTRL "T"));
	return TRUE;
	});
*/

// Single step a Monte Carlo simulation.
static AddIn xai_step(Macro("xll_step", "MC.STEP"));
int WINAPI xll_step()
{
#pragma XLLEXPORT
	mc.step();

	return 1;
}

// Run a Monte Carlo simulation.
static AddIn xai_run(Macro("xll_run", "MC.RUN"));
int WINAPI xll_run()
{
#pragma XLLEXPORT
	mc.run();

	return 1;
}
