// xllsequence.cpp - a sequence
#include "range.h"

using namespace xll;

static AddInX xai_range_sequence(
	FunctionX(XLL_LPOPERX, _T("?xll_range_sequence"), RANGE_PREFIX _T("SEQUENCE"))
	.Arg(XLL_DOUBLEX, _T("Start"), _T("is the initial value of the sequence."), 1)
	.Arg(XLL_DOUBLEX, _T("Stop"), _T("is the final value of the sequence."), 3)
	.Arg(XLL_DOUBLEX, _T("_Step"), _T("is the step size of the sequence. Default value is 1. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a range from Start to Stop with increment Step. "))
	.Documentation(
		_T("")
	)
);
LPOPERX WINAPI
xll_range_sequence(double start, double stop, double step)
{
#pragma XLLEXPORT
	static OPERX o;

	if (step == 0)
		step = 1;

	try {
		o.resize(static_cast<xword>((stop - start)/step + 1), 1);
		for (xword i = 0; i < o.size(); ++i)
			o[i] = start + i*step;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &o;
}