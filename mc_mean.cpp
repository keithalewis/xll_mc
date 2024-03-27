// mc_mean.cpp - mean value of a cell over a Monte Carlo simulation
#include "xll_mc.h"

struct mc_mean {
	double* state; // count and running mean
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
		state[0] += 1;
		state[1] += (x - state[1])/state[0];
	}
	void done()
	{ }
};

using namespace xll;