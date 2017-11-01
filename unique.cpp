// unique.cpp
#include "range.h"

using namespace range;
using namespace xll;

static AddInX xai_range_unique(
	FunctionX(XLL_LPOPERX, _T("?xll_range_unique"), RANGE_PREFIX _T("UNIQUE"))
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE, _T("=RANGE.SET({1,1,2,2,2,3})"))
	.Arg(XLL_LPOPERX, _T("_Columns"), _T("is an optional list of columns to use. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Remove consecutive duplicates from Range."))
	.Documentation(
		_T("For example, <codeInline>RANGE.UNIQUE({\"one\", \"two\", \"two\", \"three\", \"two\", \"two\"})</codeInline> is ")
		_T("<codeInline>{1, \"two\", \"three\", \"two\"}</codeInline>. ")
		_T("This function calls the C++ Standard Library routine <codeInline>std::unique()</codeInline> ")
		_T("declared in the &lt;algorithm&gt; header. ")
		_T("You should sort the routine before calling <codeInline>RANGE.UNIQUE</codeInline> to remove all duplicates. ")
	)
);
LPOPERX WINAPI
xll_range_unique(LPOPERX pa, LPOPERX pc)
{
#pragma XLLEXPORT
	static OPERX r;

	try {
		OPERX *px = range::handle(pa);

		if (px) {
			unique_(*px);
			r = *pa;
		}
		else {
			r = *pa;
			unique_(r);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &r;
}

#ifdef _DEBUG

#endif // _DEBUG