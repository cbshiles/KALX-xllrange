// index.cpp - cyclic indices for INDEX
#include "range.h"

using namespace xll;

static AddInX xai_range_index(
	FunctionX(XLL_LPOPERX, _T("?xll_range_index"), RANGE_PREFIX _T("INDEX"))
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE,
		_T("=RANGE.SET({1,2,3})"))
	.Arg(XLL_LPOPERX, _T("Rows"), _T("is an array of rows to select."),
		1)
	.Arg(XLL_LPOPERX, _T("_Columns"), _T("is an optional array of columns to select. "),
		_T("=RANGE.SET({2,-1})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a sub-array of Range indicated by Rows and Columns. "))
	.Documentation(_T("If Rows is missing, select all rows. "))
);
LPOPERX WINAPI
xll_range_index(LPOPERX pa, LPOPERX pr, LPOPERX pc)
{
#pragma XLLEXPORT
	static OPERX a;

	try {
		pa = range::lookup(pa);
		pr = range::lookup(pr);
		pc = range::lookup(pc);

		if (pr->xltype == xltypeMissing && pc->xltype== xltypeMissing)
			return pa;

		a = range::index_(*pa, *pr, *pc);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &a;
}