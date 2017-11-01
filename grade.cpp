// grade.cpp - indices that sort an array
// Copyright (c) 2006-2013 KALX, LLC. All rights reserved. No warranty is made.
/*
Use RANGE.GRADE(RANGE.COLUMNS(a, c), n) to grade on columns
*/
#include <algorithm>
#include <functional>
#include "range.h"

using namespace xll;
using namespace range;

static AddInX xai_range_grade(
	FunctionX(XLL_LPOPERX, _T("?xll_range_grade"), RANGE_PREFIX _T("GRADE"))
	.Arg(XLL_LPOPERX, _T("Range"), _T("is a range to grade."),
		_T("=RANGE.SET({3, 2, 1})"))
	.Arg(XLL_WORDX, _T("_Count"), _T("is the optional number of items in case of a partial grade. If missing a full grade is done."),
		0)
	.Arg(XLL_LPOPERX, _T("_Columns"), _T("is an optional array of columns to use for the grade. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the indices that would sort the rows of Range."))
	.Documentation(
		_T("For example, RANGE.INDEX(Range, RANGE.GRADE(Range)) returns RANGE.SORT(Range). ")
	)
);
LPOPERX WINAPI
xll_range_grade(LPOPERX pr, xword n, LPOPERX pc)
{
#pragma XLLEXPORT
	static OPERX s;

	try {
		pr = range::lookup(pr);

		s = grade_(*pr, n, *pc);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &s;
}


#ifdef _DEBUG

#define E(x) Excel<XLOPERX>(xlfEvaluate, OPERX(_T(x)))

int xll_test_grade(void)
{
	try {
		OPERX o;

		o = E("RANGE.GRADE({3,1,2})");
		ensure (o.rows() == 1 && o.columns() == 3);
		ensure (o[0] == 2 && o[1] == 3 && o[2] == 1);

		o = E("RANGE.GRADE({3;1;2})");
		ensure (o.rows() == 3 && o.columns() == 1);
		ensure (o[0] == 2 && o[1] == 3 && o[2] == 1);

		o = E("RANGE.GRADE({5,6;1,2;3,4})");
		ensure (o.rows() == 3 && o.columns() == 1);
		ensure (o[0] == 2 && o[1] == 3 && o[2] == 1);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfterX> xao_test_grade(xll_test_grade);
#endif // _DEBUG
