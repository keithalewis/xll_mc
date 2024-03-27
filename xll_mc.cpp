// xll_mc.cpp - Monte Carlo simulation
#include <cmath>
#include "xll_mc.h"

using namespace xll;

long mc_count = 0; // global count of iterations
long mc_limit = 1000; // maximum number of iterations
mc_state state = IDLE; // current state of simulation
mc_state state_ = IDLE; // next state of simulation
int calc = xlcCalculateDocument; // xlcCalculateNow - entire workbook

static AddIn xai_count(
	Function(XLL_LONG, "xll_count", "MC.COUNT")
	.Arguments({
		Arg(XLL_LONG, "count", "is the optional number of iterations to run.")
	})
	.Volatile()
	.FunctionHelp("Return current number of iterations run. Set number of iterations if non-zero.")
	.Category("MC")
);
long WINAPI xll_count(long n)
{
#pragma XLLEXPORT
	if (n) {
		mc_limit = n;
	}

	return mc_count;
}

static AddIn xai_stop(
	Function(XLL_BOOL, "xll_stop", "MC.STOP")
	.Arguments({
		Arg(XLL_BOOL, "stop", "stop the simulation.")
		})
	.Category("MC")
);
BOOL WINAPI xll_stop(BOOL b)
{
#pragma XLLEXPORT
	if (b) {
		state_ = STOP;
	}

	return b;
}


static AddIn xai_stop(
	Function(XLL_BOOL, "xll_stop", "MC.STOP")
	.Arguments({
		Arg(XLL_BOOL, "stop", "stop the simulation.")
	})
	.Category("MC")
);
BOOL WINAPI xll_stop(BOOL b)
{
#pragma XLLEXPORT
	if (b) {
		state_ = STOP;
	}

	return b;
}

// Reset a Monte Carlo simulation.
static AddIn xai_reset(Macro("xll_reset", "MC.RESET"));
int WINAPI xll_reset()
{
#pragma XLLEXPORT
	mc_count = 0;
	state = INIT;
	state_ = NEXT;
	Excel(calc);
	state = state_;

	return 1;
}

// Single step a Monte Carlo simulation.
static AddIn xai_step(Macro("xll_step", "MC.STEP"));
int WINAPI xll_step()
{
#pragma XLLEXPORT
	++mc_count;
	Excel(calc);
	state = state_;

	return 1;
}

// Run a Monte Carlo simulation.
static AddIn xai_run(Macro("xll_run", "MC.RUN"));
int WINAPI xll_run()
{
#pragma XLLEXPORT
	xll_reset();
	// fold moniods
	while (state == NEXT) {
		xll_step();
	}
	Excel(calc);
	state = IDLE;

	return 1;
}
