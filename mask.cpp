// mask.cpp - select elements from range based on a mask
#include "range.h"

using namespace xll;
using namespace range;

static AddInX xai_range_mask(
	FunctionX(XLL_LPOPERX, _T("?xll_range_mask"), _T("RANGE.MASK"))
	.Arg(XLL_LPOPERX, _T("Mask"), _T("is an array of the same size as Range indicating which elements to select. "),
		_T("=RANGE.SET({0, 1, 0})"))
	.Arg(XLL_LPOPERX, _T("Range"), IS_RANGE,
		_T("=RANGE.SET({\"a\", \"b\", \"c\"})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Choose elements from Range correponding to non-zero elements of Mask. "))
	.Documentation(
		_T("For example, RANGE.MASK(Range, Range=\"foo\" removes all cells containing the value \"foo\"." )
	)
);
LPOPERX WINAPI
xll_range_mask(LPOPERX pm, LPOPERX pr)
{
#pragma XLLEXPORT
	static OPERX r;

	try {
		OPERX* px = range::handle(pr);
		pm = range::lookup(pm);

		if (px) {
			mask_(*px, *pm);
			r = *pr; // handle
		}
		else {
			r = *pr;
			mask_(r, *pm);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &r;
}
