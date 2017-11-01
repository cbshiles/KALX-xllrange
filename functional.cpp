// functional.cpp - functions for comparing ranges
#include "range.h"

using namespace xll;

static AddInX xai_range_less(
	FunctionX(XLL_BOOLX, _T("?xll_range_less"), RANGE_PREFIX _T("LESS"))
	.Arg(XLL_LPOPERX, _T("Range1"), IS_RANGE)
	.Arg(XLL_LPOPERX, _T("Range2"), IS_RANGE _T(" "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns TRUE if Range1 is less than Range2."))
	.Documentation()
);
BOOL WINAPI
xll_range_less(const LPOPERX pr1, const LPOPERX pr2)
{
#pragma XLLEXPORT

	return *pr1 < *pr2;
}
static AddInX xai_range_greater(
	FunctionX(XLL_BOOLX, _T("?xll_range_greater"), RANGE_PREFIX _T("GREATER"))
	.Arg(XLL_LPOPERX, _T("Range1"), IS_RANGE)
	.Arg(XLL_LPOPERX, _T("Range2"), IS_RANGE _T(" "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns TRUE if Range1 is greater than Range2."))
	.Documentation()
);
BOOL WINAPI
xll_range_greater(const LPOPERX pr1, const LPOPERX pr2)
{
#pragma XLLEXPORT

	return *pr1 > *pr2;
}
static AddInX xai_range_equal_to(
	FunctionX(XLL_BOOLX, _T("?xll_range_equal_to"), RANGE_PREFIX _T("EQUAL.TO"))
	.Arg(XLL_LPOPERX, _T("Range1"), IS_RANGE)
	.Arg(XLL_LPOPERX, _T("Range2"), IS_RANGE _T(" "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns TRUE if Range1 is equal to than Range2."))
	.Documentation()
);
BOOL WINAPI
xll_range_equal_to(const LPOPERX pr1, const LPOPERX pr2)
{
#pragma XLLEXPORT

	return *pr1 == *pr2;
}
static AddInX xai_range_not_equal_to(
	FunctionX(XLL_BOOLX, _T("?xll_range_not_equal_to"), RANGE_PREFIX _T("NOT.EQUAL.TO"))
	.Arg(XLL_LPOPERX, _T("Range1"), IS_RANGE)
	.Arg(XLL_LPOPERX, _T("Range2"), IS_RANGE _T(" "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns TRUE if Range1 is not equal to than Range2."))
	.Documentation()
);
BOOL WINAPI
xll_range_not_equal_to(const LPOPERX pr1, const LPOPERX pr2)
{
#pragma XLLEXPORT

	return *pr1 != *pr2;
}
static AddInX xai_range_greater_equal(
	FunctionX(XLL_BOOLX, _T("?xll_range_greater_equal"), RANGE_PREFIX _T("GREATER.EQUAL"))
	.Arg(XLL_LPOPERX, _T("Range1"), IS_RANGE)
	.Arg(XLL_LPOPERX, _T("Range2"), IS_RANGE _T(" "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns TRUE if Range1 is greater than or equal than Range2."))
	.Documentation()
);
BOOL WINAPI
xll_range_greater_equal(const LPOPERX pr1, const LPOPERX pr2)
{
#pragma XLLEXPORT

	return *pr1 >= *pr2;
}
static AddInX xai_range_less_equal(
	FunctionX(XLL_BOOLX, _T("?xll_range_less_equal"), RANGE_PREFIX _T("LESS.EQUAL"))
	.Arg(XLL_LPOPERX, _T("Range1"), IS_RANGE)
	.Arg(XLL_LPOPERX, _T("Range2"), IS_RANGE _T(" "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns TRUE if Range1 is less than or equal than Range2."))
	.Documentation()
);
BOOL WINAPI
xll_range_less_equal(const LPOPERX pr1, const LPOPERX pr2)
{
#pragma XLLEXPORT

	return *pr1 <= *pr2;
}

#ifdef _DEBUG

#define E(x) Excel<XLOPERX>(xlfEvaluate, OPERX(_T(x)))

int xll_test_functional(void)
{
	try {
		ensure (E("RANGE.LESS(1,2)"));
		ensure (E("RANGE.LESS({1,1},{1,2})"));
		ensure (!E("RANGE.LESS(3,2)"));
		ensure (!E("RANGE.LESS({1,3},{1,2})"));

		ensure (!E("RANGE.GREATER(1,2)"));
		ensure (!E("RANGE.GREATER({1,1},{1,2})"));
		ensure (E("RANGE.GREATER(3,2)"));
		ensure (E("RANGE.GREATER({1,3},{1,2})"));

		ensure (!E("RANGE.EQUAL.TO(1,2)"));
		ensure (!E("RANGE.EQUAL.TO({1,1},{1,2})"));
		ensure (!E("RANGE.EQUAL.TO(3,2)"));
		ensure (!E("RANGE.EQUAL.TO({1,3},{1,2})"));

		ensure (!E("RANGE.GREATER.EQUAL(1,2)"));
		ensure (!E("RANGE.GREATER.EQUAL({1,1},{1,2})"));
		ensure (E("RANGE.GREATER.EQUAL(3,2)"));
		ensure (E("RANGE.GREATER.EQUAL({1,3},{1,2})"));

		ensure (E("RANGE.LESS.EQUAL(1,2)"));
		ensure (E("RANGE.LESS.EQUAL({1,1},{1,2})"));
		ensure (!E("RANGE.LESS.EQUAL(3,2)"));
		ensure (!E("RANGE.LESS.EQUAL({1,3},{1,2})"));

		ensure (E("RANGE.NOT.EQUAL.TO(1,2)"));
		ensure (E("RANGE.NOT.EQUAL.TO({1,1},{1,2})"));
		ensure (E("RANGE.NOT.EQUAL.TO(3,2)"));
		ensure (E("RANGE.NOT.EQUAL.TO({1,3},{1,2})"));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfterX> xao_test_functional(xll_test_functional);

#endif // _DEBUG