// drop.cpp - drop rows/columns from range
#include "range.h"

using namespace xll;
using namespace range;

static AddInX xai_range_drop(
	FunctionX(XLL_LPOPERX, _T("?xll_range_drop"), RANGE_PREFIX _T("DROP"))
	.Arg(XLL_LONGX, _T("Rows"), _T("is the number of rows to drop."),
		2)
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE _T(" "),
		_T("=RANGE.SET({1,2,3})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Drop Rows from front, if positive, or back, if negative, of Range."))
	.Documentation(
		_T("If Range has one row then columns are dropped. ")
	)
);
LPOPERX WINAPI
xll_range_drop(int n, LPOPERX pr)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		OPERX* px = range::handle(pr);

		if (px) {
			drop_(n, *px);
			o = *pr; // handle
		}
		else {
			o = *pr;
			drop_(n, o);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &o;
}

#ifdef _DEBUG

int test_drop(void)
{
	try {
		OPERX o;
		o = Excel<XLOPERX>(xlfEvaluate, OPERX(_T("RANGE.DROP(1, {1,2,3})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 1 && o.columns() == 2);
		ensure (o[0] == 2 && o[1] == 3);

		o = Excel<XLOPERX>(xlfEvaluate, OPERX(_T("RANGE.DROP(-1, {1,2,3})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 1 && o.columns() == 2);
		ensure (o[0] == 1 && o[1] == 2);

		o = Excel<XLOPERX>(xlfEvaluate, OPERX(_T("RANGE.DROP(1, {1;2;3})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 2 && o.columns() == 1);
		ensure (o[0] == 2 && o[1] == 3);

		o = Excel<XLOPERX>(xlfEvaluate, OPERX(_T("RANGE.DROP(-1, {1;2;3})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 2 && o.columns() == 1);
		ensure (o[0] == 1 && o[1] == 2);

		o = Excel<XLOPERX>(xlfEvaluate, OPERX(_T("RANGE.DROP(1, {1,2,3;4,5,6})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 1 && o.columns() == 3);
		ensure (o[0] == 4 && o[2] == 6);

		o = Excel<XLOPERX>(xlfEvaluate, OPERX(_T("RANGE.DROP(-1, {1,2,3;4,5,6})")));
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 1 && o.columns() == 3);
		ensure (o[0] == 1 && o[2] == 3);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

static Auto<OpenAfterX> xao_test_drop(test_drop);

#endif // _DEBUG