// xll_mc.h - Monte Carlo simulation
#pragma once
#include "xll.h"

#ifndef CATEGORY
#define CATEGORY "MC"
#endif

enum mc_state {
	IDLE,
	INIT,
	NEXT,
	HALT, // pause simulation
	STOP,
};

struct monte {
	long count, limit; // current and maximum number of simulations
	mc_state state, state_; // current and next state
	int calc; // calculation mode
	LARGE_INTEGER freq, start, stop, elapsed; // timing

	monte()
		: count(0), limit(LONG_MAX), state(IDLE), state_(IDLE), calc(xlcCalculateDocument), elapsed({ 0 })
	{
		QueryPerformanceFrequency(&freq);
	}

	void reset()
	{
		count = 0;
		state = INIT;
		state_ = NEXT;
		elapsed = { 0 };
		calculate();
		state = state_;
	}

	void step()
	{
		if (state == IDLE) {
			reset();
		}
		if (state == HALT) {
			state = NEXT;
			state_ = HALT;
		}
		if (state == NEXT && count < limit) {
			++count;
			calculate();
			state = state_;
		}
	}
	void run()
	{
		static XLOPER12 True = xll::Bool(true), False = xll::Bool(false);

		if (state != HALT) {
			reset();
		}
		else {
			state = state_ = NEXT;
		}

		::Excel12(xlcEcho, 0, 1, &False);
		// fold moniods
		while (state == NEXT && count < limit) {
			++count;
			calculate();
			state = state_;
		}
		::Excel12(xlcEcho, 0, 1, &True);
		
		calculate();
		state = state_ = IDLE;
	}

	void calculate()
	{
		QueryPerformanceCounter(&start);
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
