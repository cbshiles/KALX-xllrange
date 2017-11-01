// member.cpp - index of elements in a set
#include <algorithm>
#include "range.h"

using namespace xll;

static AddInX xai_range_member(
	FunctionX(XLL_LPOPERX, _T("?xll_range_member"), RANGE_PREFIX _T("MEMBER"))
	.Arg(XLL_LPOPERX, _T("Element"), _T("is a range of elements. "))
	.Arg(XLL_LPOPERX, _T("Set"), _T("is a range specifying a set. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns an Element shaped range of the index of Element in Set. "))
	.Documentation(
		_T("Note this can be used as a mask. ")
	)
);
LPOPERX WINAPI
xll_range_member(LPOPERX pe, LPOPERX ps)
{
#pragma XLLEXPORT
	static OPERX m(pe->rows(), pe->columns());

	for (xword i = 0; i < m.size(); ++i) {
		// STL style
//		m[i] = (std::find(ps->begin(), ps->end(), (*pe)[i]) - ps->begin() + 1)%(ps->size() + 1);
		// Excel macro style
		m[i] = ExcelX(xlfIferror, ExcelX(xlfMatch, (*pe)[i], *ps, OPERX(0)), OPERX(0));
	}

	return &m;
}