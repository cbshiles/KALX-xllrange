// reshape.cpp - reshape a range
#include "range.h"

using namespace xll;

static AddInX xai_range_reshape(
	FunctionX(XLL_LPOPERX, _T("?xll_range_reshape"), _T("RANGE.RESHAPE"))
	.Arg(XLL_USHORTX, _T("Rows"), _T("is the number of rows in the new range."),2)
	.Arg(XLL_USHORTX, _T("Columns"), _T("is the number of columns in the new range."),3)
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE _T(" "), _T("=RANGE.SET({1,2;3,4;5,6})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Reshape a range."))
	.Documentation(
		_T("If Rows or Columns are 0 or omitted approriate values are calculated for you. ")
	)
);
LPXLOPERX WINAPI
xll_range_reshape(xword r, xword c, LPOPERX po)
{
#pragma XLLEXPORT
	LPOPERX pr = range::handle(po);

	try {
		if (!pr) {
			if (po->xltype != xltypeMulti)
				return 0;
			pr = po;
		}

		if (!r) 
			r = 1;
		if (!c)
			c = 1;

		ensure (pr->size() == r*c);

		pr->val.array.rows = r;
		pr->val.array.columns = c;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return pr;
}

#ifdef _DEBUG

int xll_test_reshape(void)
{
	try {
		OPERX o(2,3);

		o = *xll_range_reshape(3,2, &o);
		ensure (o.rows() == 3 && o.columns() == 2);

		o = *xll_range_reshape(0,6, &o);
		ensure (o.rows() == 1 && o.columns() == 6);

		o = *xll_range_reshape(6, 0, &o);
		ensure (o.rows() == 6 && o.columns() == 1);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfterX> xao_test_reshape(xll_test_reshape);

#endif // _DEBUG