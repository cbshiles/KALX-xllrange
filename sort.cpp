// sort.cpp - sort a range
#include "range.h"

using namespace xll;
using namespace range;

static AddInX xai_range_sort(
	FunctionX(XLL_LPOPERX, _T("?xll_range_sort"), RANGE_PREFIX _T("SORT"))
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE, _T("=RANGE.SET({1,2,3})"))
	.Arg(XLL_LONGX, _T("_Count"), _T("is the number of items in case of a partial sort. Negative means decreasing sort"))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Sort a range by rows based on the first Column. "))
	.Documentation(
		_T("A count of -1 sorts all elements in decreasing order. ")
	)
); 
LPOPERX WINAPI
xll_range_sort(LPOPERX pr, int n)
{
#pragma XLLEXPORT
	static OPERX r;

	try {
		OPERX* px = range::handle(pr);

		if (px) {
			sort_(*px, n);
			r = *pr;
		}
		else {
			r = *pr;	
			sort_(r, n);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &r;
}

static AddInX xai_range_sort_columns(
	FunctionX(XLL_LPOPERX, _T("?xll_range_sort_columns"), RANGE_PREFIX _T("SORT.COLUMNS"))
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE, 
		_T("=RANGE.SET({4,5,6;1,2,3})"))
	.Arg(XLL_LPOPERX, _T("Columns"), _T("is an optional array of columns on which to sort. Default is first. "),
		1)
	.Arg(XLL_LONGX, _T("_Count"), _T("is then number of items in case of a partial sort. Negative means decreasing sort"))
	.ThreadSafe()
	.Category(CATEGORY)
	.FunctionHelp(_T("Sort a range by rows based on Column. "))
	.Documentation(
		_T("A count of -1 sorts all elements in decreasing order. ")
		_T("Set columns to 0 to sort on all columns. ")
	)
); 
LPOPERX WINAPI
xll_range_sort_columns(const LPOPERX pr, LPOPERX pc, int n)
{
#pragma XLLEXPORT
	static OPERX r;

	try {
		OPERX* px = range::handle(pr);

		if (px) {
			OPERX r = index_(*px, OPERX(xltype::Missing), *pc);
			OPERX s = grade_(r, n);
			derange(*px, s);
			r = *px;
		}
		else {
			OPERX r = index_(*pr, OPERX(xltype::Missing), *pc);
			OPERX s = grade_(r, n);
			r = *pr;
			derange(r, s);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &r;
}
