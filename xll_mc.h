// xll_mc.h - Monte Carlo simulation
#pragma once
#include "xll.h"

#ifndef CATEGORY
#define CATEGORY "MC"
#endif

// https://devblogs.microsoft.com/oldnewthing/2005/02/22
inline void DoEvents()
{
    MSG msg;
    while (PeekMessage(&msg, nullptr, 0, 0, PM_NOREMOVE)) {
        BOOL b = GetMessage(&msg, nullptr, 0, 0);
        if (b == -1) {
            MessageBoxA(0, "GetMessage returned -1", "Error", MB_OK);
            break;
        } else if (b > 0) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        } else if (msg.message == WM_QUIT) {
#ifdef _DEBUG
            MessageBoxA(nullptr, "Calling PostQuitMessage", "Info", MB_OK);
#endif
            PostQuitMessage(static_cast<int>(msg.wParam)); // exits when macro is done?
        }
    }
}

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
    long update; // number of iterations between updates
		
	void message()
	{
        static wchar_t buf[1024];
		auto es = elapsed_seconds();
        auto sps = es ? static_cast<long>(count/es) : 0;
        buf[0] = static_cast<wchar_t>(swprintf(buf + 1, 1023, L"%ld [%ld/s] %.2fs", count, sps, es));
		xll::Excel(xlcMessage, true, xll::Str(buf));
        //DoEvents();
        if (xll::Excel(xlAbort) == true) {
            state_ = HALT;
        }
		// update 5 times per second
        update = 1 + sps / 5;
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
        message();
	}

	void next()
	{
		++count;
		calculate(state);
		state = state_;
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
			next();
			message();
		}
		state = (count == limit) ? STOP : HALT;
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
			next();
			if (count % update == 0) {
                ::Excel12(xlcEcho, 0, 1, &xll::True);
                message();
                ::Excel12(xlcEcho, 0, 1, &xll::False);
            }
		}
		::Excel12(xlcEcho, 0, 1, &xll::True);
		message();
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
