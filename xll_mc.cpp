// xll_mc.cpp - Monte Carlo simulation
#include <cmath>
#include "xll_mc.h"

using namespace xll;

monte mc;

// seconds after now
inline OPER timeout(double seconds)
{
	OPER o = Excel(xlfNow);
	
	Num(o) += seconds / 86400;

	return o;
}
static AddIn xai_update(Macro("xll_update", "MC.UPDATE"));
int WINAPI xll_update()
{
#pragma XLLEXPORT
	//!!! update message bar
	if (Excel(xlAbort) == true) {
		mc.state_ = HALT;
	}
	else if (mc.state == NEXT) {
		Excel(xlcOnTime, timeout(1), OPER("MC.UPDATE"));
	}
	Excel(xlfEcho, True);
	Excel(xlfEcho, False);

	return 1;
}

#ifdef _DEBUG
static AddIn xai_state(
	Function(XLL_LPOPER, "xll_state", "MC.STATE")
	.Arguments({})
	.Volatile()
	.FunctionHelp("Return current and next state of Monte Carlo simulation.")
);
LPOPER WINAPI xll_state()
{
#pragma XLLEXPORT
	static OPER o(1, 2, nullptr);

	o[0] = mc_state_name[mc.state];
	o[1] = mc_state_name[mc.state_];

	return &o;
}
#endif // _DEBUG

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

static AddIn xai_limit(
	Function(XLL_LONG, "xll_limit", "MC.LIMIT")
	.Arguments({
		Arg(XLL_LONG, "limit", "is the maximum number of iterations to run.")
		})
	.Volatile()
	.FunctionHelp("Set the maximum number of Monte Carlo simulations to run.")
	.Category(CATEGORY)
);
long WINAPI xll_limit(long n)
{
#pragma XLLEXPORT
	if (n) {
		mc.limit = n;
	}

	return mc.limit;
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
	.FunctionHelp("Cause the simulation to end.")
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
	.FunctionHelp("Cause the simulation to pause.")
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
OnKey xlo_monte_reset(ON_CTRL ON_SHIFT "T", "MC.RESET");

// Single step a Monte Carlo simulation.
static AddIn xai_step(Macro("xll_step", "MC.STEP"));
int WINAPI xll_step()
{
#pragma XLLEXPORT
	mc.step();

	return 1;
}
OnKey xlo_monte_step(ON_CTRL ON_SHIFT "S", "MC.STEP");

// Run a Monte Carlo simulation.
static AddIn xai_run(Macro("xll_run", "MC.RUN"));
int WINAPI xll_run()
{
#pragma XLLEXPORT
	Excel(xlcOnTime, timeout(1), OPER("MC.UPDATE"));
	mc.run();

	return 1;
}
OnKey xlo_monte_run(ON_CTRL ON_SHIFT "R", "MC.RUN");