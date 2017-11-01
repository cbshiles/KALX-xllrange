// find.cpp - get indices of elements in array matching criteria
#include "range.h"

using namespace xll;

static AddInX xai_range_find(
	FunctionX(XLL_LPOPERX, _T("?xll_range_find"), RANGE_PREFIX _T("FIND"))
	.Arg(XLL_LPOPERX, _T("What"), _T("is What you want to find."),
		_T("two"))
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE,
		_T("=RANGE.SET({\"one\", \"two\", \"three\"})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return index of What in Range."))
	.Documentation(
		_T("For example, INDEX(Column1, RANGE.FIND(Key, Column2)) is similar to the ")
		_T("Excel VLOOKUP function. ")
		_T("An index of 0 means not found. The output has the same shape as What. ")
	)
);
LPOPERX WINAPI
xll_range_find(LPOPERX pw, LPOPERX pr)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		pr = range::lookup(pr);

		o.resize(pw->rows(), pw->columns());
		for (xword i = 0; i < o.size(); ++i) {
			xword j = std::find(pr->begin(), pr->end(), (*pw)[i]) - pr->begin();
			o[i] = j == pr->size() ? 0 : j + 1;
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}

#ifdef _DEBUG

#define E(x) Excel<XLOPERX>(xlfEvaluate, OPERX(_T(x)))

int xll_test_find(void)
{
	try {
		ensure (E("2 = RANGE.FIND(\"bar\",{\"foo\", \"bar\", \"baz\"})"));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfterX> xao_test_find(xll_test_find);


#endif // _DEBUG