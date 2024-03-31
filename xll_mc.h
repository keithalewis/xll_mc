// xll_mc.h - Monte Carlo simulation
#pragma once
#include "xll.h"

#ifndef CATEGORY
#define CATEGORY "MC"
#endif

// Simulation states
#define MC_STATE(X) \
	X(IDLE, "not running") \
	X(INIT, "starting") \
	X(NEXT, "running") \
	X(HALT, "pausing") \
	X(STOP, "ending") \

#define MC_STATE_ENUM(a, b) a,
enum mc_state {
	MC_STATE(MC_STATE_ENUM)
};
#undef MC_STATE_ENUM

#define MC_STATE_NAME(a, b) #a,
constexpr const char* mc_state_name[] = {
	MC_STATE(MC_STATE_NAME)
};
#undef MC_STATE_NAME

struct monte {
	long count, limit; // current and maximum number of simulations
	mc_state state, state_; // current and next state
	int calc; // calculation mode
	LARGE_INTEGER freq, start, stop, elapsed; // timing
	long update; // update interval in milliseconds
	void refresh()
	{
		// force screen to update
		xll::OPER echo = xll::Excel(xlfGetWorkspace, 40); // screen updating
		if (echo == false) {
			::Excel12(xlcEcho, 0, 1, &xll::True);
		}
		//xll::Excel(xlcWorkbookScroll, 1);
		//xll::Excel(xlcWorkbookScroll, -1);
		if (echo == false) {
			::Excel12(xlcEcho, 0, 1, &xll::False);
		}
	}
	void message()
	{
		xll::OPER o = xll::Excel(xlfText, xll::OPER(1.*count), xll::OPER("0"));
		//update = count/static_cast<long>(elapsed_seconds()*1000);
		xll::Excel(xlcMessage, o);
	}

	monte()
		: count(0), limit(LONG_MAX), 
		  state(IDLE), state_(IDLE), 
		  calc(xlcCalculateDocument), 
		  start({ 0 }), stop({ 0 }), elapsed({ 0 }), 
		  update(1)
	{
		QueryPerformanceFrequency(&freq);
	}

	// https://xlladdins.github.io/Excel4Macros/calculate.document.html
	// Calculates only the active worksheet.
	void calculateDocument()
	{
		calc = xlcCalculateDocument;
	}
	// https://xlladdins.github.io/Excel4Macros/calculate.now.html
	// Calculate all open workbooks when calculation is set to manual.
	void calculateNow()
	{
		calc = xlcCalculateNow;
	}

	void reset()
	{
		calculate(INIT);
		count = 0;
		state = IDLE;
		state_ = IDLE;
		elapsed = { 0 };
	}

	void step()
	{
		if (state == IDLE) {
			reset();
			state = NEXT;
		}
		else if (state == HALT) {
			state = NEXT;
		}

		if (state == NEXT && count < limit) {
			++count;
			calculate(state);
		}
		state = count == limit ? STOP : HALT;
	}
	void run()
	{
		if (state != HALT) {
			reset();
		}
		state = state_ = NEXT;

		::Excel12(xlcEcho, 0, 1, &xll::False);
		// fold moniods
		while (state == NEXT && count < limit) {
			++count;
			calculate(state);
			state = state_;
			refresh();
//			if (count % update == 0) {
//				::Excel12(xlcEcho, 0, 1, &xll::True);
//				message();
//				::Excel12(xlcEcho, 0, 1, &xll::False);
//			}
		}
		::Excel12(xlcEcho, 0, 1, &xll::True);
		//message();
		if (count == limit) {
			state = STOP;
		}
	}

	void calculate(mc_state curr)
	{
		QueryPerformanceCounter(&start);
		state = curr;
		::Excel12v(calc, 0, 0, nullptr);
		QueryPerformanceCounter(&stop);
		elapsed.QuadPart += stop.QuadPart - start.QuadPart;
	}

	// Elapsed time in seconds.
	double elapsed_seconds()
	{
		return elapsed.QuadPart / static_cast<double>(freq.QuadPart);
	}
};
extern monte mc;
