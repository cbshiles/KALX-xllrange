// override.cpp - override values in a range
#include "range.h"

using namespace xll;

static AddInX xai_range_override(
	FunctionX(XLL_LPOPERX, _T("?xll_range_override"), _T("RANGE.OVERRIDE"))
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE,
		_T("=RANGE.SET({1, 2, 3})"))
	.Arg(XLL_LPOPERX, _T("Override"), _T("is a range the same size as Range containing overriding data. "),
		_T("=RANGE.SET({4, \"\", \"\"})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Non-empty cells in Override replace the contents of Range. "))
	.Documentation()
);
LPOPERX WINAPI
xll_range_override(LPOPERX pr1, LPOPERX pr2)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		pr1 = range::lookup(pr1);
		pr2 = range::lookup(pr2);

		ensure (pr1->columns() == pr2->columns());

		o = *pr1;
		for (xword i = 0; i < pr1->size(); ++i)
			if ((*pr2)[i].xltype != xltypeNil)
				o[i] = (*pr2)[i];
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}