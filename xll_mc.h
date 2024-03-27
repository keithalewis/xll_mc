// xll_mc.h - Monte Carlo simulation
#pragma once
#include "xll.h"

#ifndef CATEGORY
#define CATEGORY "MC"
#endif

extern long mc_count;

enum mc_state {
	IDLE,
	INIT,
	NEXT,
	STOP,
	HALT, // pause simulation
};
extern mc_state state, state_; // current and next state
