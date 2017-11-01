// call.cpp - call a functions given arguments as a range
#include "range.h"

using namespace xll;

static AddInX xai_range_call(
	FunctionX(XLL_LPOPERX XLL_UNCALCEDX, _T("?xll_range_call"), RANGE_PREFIX _T("CALL"))
	.Arg(XLL_HANDLEX, _T("Function"), _T("is the function to be called."),
		_T("=RANGE.MAKE"))
	.Arg(XLL_LPOPERX, _T("Range"), _T("is a range containing the function arguments. "),
		_T("={2,3,\"2x3\"}"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Calls Function and supplies elements of Range as arguments."))
	.Documentation(
		_T("For example, <codeInline>RANGE.CALL(RANGE.MAKE, {2, 3})</codeInline> will return an array having ")
		_T("two rows and three columns, i.e., <codeInline>RANGE.MAKE(2,3)</codeInline>. ")
		_T("If an argument is a range, use the result of <codeInline>RANGE.SET</codeInline>. ")
		_T("It will be looked up using <codeInline>RANGE.GET</codeInline> prior to the function being called. ")
	)
);
LPOPERX WINAPI
xll_range_call(HANDLEX hf, LPOPERX po)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = range::call(hf, po);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}

#ifdef _DEBUG

#define E(s) Excel<XLOPERX>(xlfEvaluate, OPERX(_T(s)));

int xll_test_call(void)
{
	try {
		OPERX o;

		o = E("RANGE.CALL(RANGE.MAKE,{2, 3})");
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 2);
		ensure (o.columns() == 3);

	} 
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfterX> xao_test_call(xll_test_call);

#endif // _DEBUG