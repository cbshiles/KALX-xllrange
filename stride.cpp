// stride.cpp - extract arithmetic progression
#include "range.h"

using namespace xll;
using namespace range;

static AddInX xai_range_stride(
	FunctionX(XLL_LPOPERX, _T("?xll_range_stride"), RANGE_PREFIX _T("STRIDE"))
	.Arg(XLL_LONGX, _T("Step"), _T("is the step size for successive elements."), 2)
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE, _T("=RANGE.SET({1,2,3,4,5,6})"))
	.Arg(XLL_LONGX, _T("_Offset"), _T("is the optional offset into the range. "), 1)
	.Category(CATEGORY)
	.FunctionHelp(_T("Extracts elements from Start to the end of the array skipping by Step. "))
	.Documentation(
		_T("")
	)
);
LPOPERX WINAPI
xll_range_stride(int step, LPOPERX pr, int offset)
{
#pragma XLLEXPORT
	static OPERX r;

	if (offset == 0 && step == 0)
		return pr;

	try {
		OPERX* px = range::handle(pr);

		if (px) {
			stride_(*px, step, offset);
			r = *pr; // handle
		}
		else {
			r = *pr;
			stride_(r, step, offset);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &r;
}
