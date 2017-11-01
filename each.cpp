// each.cpp - apply function to each row of a range
#include "range.h"

using namespace range;
using namespace xll;

static AddInX xai_range_each(
	FunctionX(XLL_LPOPERX, _T("?xll_range_each"), RANGE_PREFIX _T("EACH"))
	.Arg(XLL_HANDLEX, _T("Function"), _T("is a handle to a function."))
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE)
	.Category(CATEGORY)
	.FunctionHelp(_T("Apply Function to each row of range using RANGE.CALL. "))
	.Documentation(
		_T("For example, RANGE.EACH(RANGE.BIND(RANGE.DROP,2,), Range) will ")
		_T("remove the first two columns of Range. ")
		_T("The result has the same number of rows as Range with each ")
		_T("row being the result of calling the function using the row ")
		_T("as arguments. ")
	)
);
LPOPERX WINAPI
xll_range_each(HANDLEX hf, LPXLOPERX pr)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		if (!hf)
			o = ErrX(xlerrNA);
		else {
			o = OPERX();

			xword c = pr->val.array.columns;
			for (xword i = 0; i < pr->val.array.rows; ++i) {
				XLOPERX x(wrap_(1, c, pr->val.array.lparray + i*c));
				o.push_back(range::call(hf, &x));
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}