// take.cpp - take rows/columns from range
#include "range.h"

using namespace xll;
using namespace range;

static AddInX xai_range_take(
	FunctionX(XLL_LPOPERX, _T("?xll_range_take"), RANGE_PREFIX _T("TAKE"))
	.Arg(XLL_LONGX, _T("Rows"), _T("is the number of rows to take."), -1)
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE _T(" "), _T("=RANGE.SET({1,2,3})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Take Rows rows from front, if positive, or back, if negative, from Range. "))
	.Documentation(
		_T("If Range has one row then columns are taken. ")
	)
);
LPOPERX WINAPI
xll_range_take(int n, const LPOPERX pr)
{
#pragma XLLEXPORT
	static OPERX r;

	try {
		OPERX* px = range::handle(pr);

		if (px) {
			take_(n, *px);
			r = *pr; // handle
		}
		else {
			r = *pr;
			take_(n, r);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &r;
}
