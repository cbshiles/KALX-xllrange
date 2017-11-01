// stack.cpp - stack ranges with equal columns
#include "range.h"

using namespace xll;

static AddInX xai_range_stack(
	FunctionX(XLL_LPOPERX, _T("?xll_range_stack"), _T("RANGE.STACK"))
	.Arg(XLL_LPOPERX, _T("Range1"), IS_RANGE,
		_T("=RANGE.SET({1,2,3})"))
	.Arg(XLL_LPOPERX, _T("Range2"), IS_RANGE _T(" "),
		_T("=RANGE.SET({4,5,6})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Stack ranges having the same number of columns."))
	.Documentation(
		_T("")
	)
);
LPOPERX WINAPI
xll_range_stack(LPOPERX pr1, const LPOPERX pr2)
{
#pragma XLLEXPORT
	static OPERX r;

	try {
		OPERX* px1 = range::handle(pr1);
		OPERX* px2 = range::handle(pr2);

		if (px1) {
			px1->push_back(px2 ? *px2 : *pr2);
			r = *pr1; // handle
		}
		else {
			r = *pr1;
			r.push_back(px2 ? *px2 : *pr2);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		r = Err(xlerrNA);
	}

	return &r;
}
