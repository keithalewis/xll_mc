# xll_mc

Monte Carlos simulation for Excel.

Compute fold of monoids.

Cell contains monoid state.

```cpp
    // (0, 0) -> (1, x1) -> (2, (x1 + x2)/2) -> ...
    // (n, x) -> (n + 1, (n*x + y)/(n + 1))
    struct mean {
	    double n, x;
		mean(n, x) : n(n), x(x) { }
		init() { n = x = 0; }
		next(double y) {
		    x += (y - x)/++n;
		}
		halt() { return n; }
		stop() { return x; }
    };
```

| current state | next state | action |
| ------------- | ---------- | ------ |
| INIT          | NEXT       | F9     |
| RUN           | STEP       | F9     |

RUN:
	state = INIT
	F9
	while state == STEP:
		F9
		state = state_;
		if STATE == STOP:
			F9
			state = IDLE
			continue
		else if state == HALT
