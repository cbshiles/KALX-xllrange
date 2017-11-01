// range.cpp - Excel range functions
#include "range.h"

using namespace xll;

#ifdef _DEBUG

static AddInX xai_range(
	DocumentX(CATEGORY)
	.Documentation(
		_T("<para>")
		_T("Ranges are two dimensional ranges of cells in Excel. ")
		_T("This add-in provides functions for slicing, dicing, and splicing ranges. ")
		_T("The convention is to have the range be the last argument to each function ")
		_T("except for the optional arguments which always trail the range argument. ")
		_T("This makes formulas having multiple function calls more intuitive - think ")
		_T("of them as a pipeline where the array comes in from the back of the forumula ")
		_T("and has successive transformations applied. ")
		_T("</para><para>")
		_T("For example, a data vendor might provide a call schedule as a string in the following format: ")
		_T("\";2;16;2;5;01/18/2013;3;100.000000;5;04/18/2013;3;100.000000;...\" and you want ")
		_T("to extract a two column array of call dates and call prices. ")
		_T("The function <codeInline>RANGE.SPLIT(String)</codeInline> uses a semicolone, \";\", as the default record separator ")
		_T("to produce a one column arry of data. ")
		_T("The leading \";2;16;2;\" can be dropped using <codeInline>RANGE.DROP(4, Range)</codeInline> to drop ")
		_T("the first 4 rows and 0 columns. ")
		_T("Next, <codeInline>RANGE.STRIDE(2, Range, 1)</codeInline> will extract the dates and prices. ")
		_T("Finally, <codeInline>RANGE.RESHAPE(0, 2, Range)</codeInline> will create ")
		_T("a 2 column array. ")
		_T("</para>")
	)
);
#endif 

static AddInX xai_range_set(
	FunctionX(XLL_HANDLEX XLL_UNCALCEDX, _T("?xll_range_set"), _T("RANGE.SET"))
	.Arg(XLL_LPOPERX, _T("Range"), _T("is a range. "),
		_T("=RANGE.SET({1,2,3;4,5,6})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to a range."))
	.Documentation(
		_T("Ranges are two dimensional arrays of cells in Excel.")
		,
		xml::element()
			.content(xml::xlink(_T("RANGE.GET")))
			.content(xml::xlink(_T("RANGE.MAKE")))
	)
);
HANDLEX WINAPI
xll_range_set(LPOPERX pr)
{
#pragma XLLEXPORT
	if (pr->xltype == xltypeErr)
		return handlex();

	handle<OPERX> ho(new OPERX(*pr));
	
	return ho.get();
}

static AddInX xai_range_get(
	FunctionX(XLL_LPOPERX, _T("?xll_range_get"), _T("RANGE.GET"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a range created using RANGE.SET. "),
		_T("=RANGE.SET({1,2,3})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the range corresponding to Handle."))
	.Documentation(
		_T("Ranges are two dimensional ranges of cells in Excel.")
		,
		xml::element()
			.content(xml::xlink(_T("RANGE.SET")))
			.content(xml::xlink(_T("RANGE.MAKE")))
	)
);
LPOPERX WINAPI
xll_range_get(HANDLEX hr)
{
#pragma XLLEXPORT

	handle<OPERX> h(hr);
	
	return h.ptr();
}

static AddInX xai_range_make(
	FunctionX(XLL_LPOPERX XLL_UNCALCEDX, _T("?xll_range_make"), _T("RANGE.MAKE"))
	.Arg(XLL_LONGX, _T("Rows"), _T("is the number of Rows in the range."), 2)
	.Arg(XLL_LONGX, _T("_Columns"), _T("is the number of Columns in the range."), 3)
	.Arg(XLL_LPOPERX, _T("_Value"), _T("is the default value for all cells in the range. "), _T("2x3"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a range."))
	.Documentation(
		_T("Ranges are two dimensional ranges of cells in Excel.")
		,
		xml::element()
			.content(xml::xlink(_T("RANGE.GET")))
			.content(xml::xlink(_T("RANGE.SET")))
	)
);
LPOPERX WINAPI
xll_range_make(ULONG rows, ULONG cols, LPOPERX po)
{
#pragma XLLEXPORT
	static OPERX o;
	
	o = OPERX(rows, cols ? cols : 1);

	if (po->xltype != xltypeMissing) {
		for (xword i = 0; i < o.size(); ++i)
			o[i] = (*po)[range::index0_(i, po->size())]; // cyclic fill		
	}
	
	return &o;
}

static AddInX xai_range_key(
	FunctionX(XLL_LPOPERX XLL_UNCALCEDX, _T("?xll_range_key"), _T("RANGE.KEY"))
	.Arg(XLL_LPOPERX, _T("Key"), _T("is the string key to look up in the first column of Range."),
		_T("two"))
	.Arg(XLL_LPOPERX, _T("Range"), _T("is a range or handle to a two column range of key-value pairs. "),
		_T("=RANGE.SET({\"one\", 1; \"two\", 2})"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the value in the second column of Range corresponding the the Key in the first column of Range."))
	.Documentation(
		_T("Uses VLOOKUP with lookup flag FALSE so the keys do not have to be sorted. ")
		_T("Returns <codeInline>xltypeMissing</codeInline> if the key cannont be found or value is <codeInline>xltypeNil</codeInline> ")
		_T("so functions taking <codeInline>OPER</codeInline>s work correctly. ")
		_T("</para><para>")
		_T("If <codeInline>Key</codeInline> is a one column array then a one column array of values is returned. ")
		_T("If <codeInline>Key</codeInline> is a one row array, then nested values are looked up. ")
	)
//	.Alias(_T("KEY")) // presumptuous?
);
LPOPERX WINAPI
xll_range_key(LPOPERX pk, LPOPERX pr)
{
#pragma XLLEXPORT
	static OPERX v;

	try {
		pr = range::lookup(pr);
		pk = range::lookup(pk);

		v.resize(pk->rows(), 1);

		ensure (pr->xltype == xltypeMulti && pr->val.array.columns == 2);

		for (xword i = 0; i < v.rows(); ++i) {
			v[i] = Excel<XLOPERX>(xlfVlookup, (*pk)(i,0), *pr, OPERX(2), OPERX(false));

			if (v.xltype == xltypeErr || v.xltype == xltypeNil) {
				v.xltype = xltypeMissing; // so functions taking OPER args work correctly
			}
			else if (v[i].xltype == xltypeNum) {
				handle<OPERX> h(v[i].val.num);
				if (h) {
					v[i] = *h;
				}
			}
			else if (v[i].xltype == xltypeStr) {
				OPERX o = range::eval(range::bang(v[i]));
				if (o)
					v[i] = o;
			}

			if (v[i].xltype == xltypeMulti && v[i].columns() == 2 && pk->columns() > 1) {
				OPERX k = *pk;
				range::drop_(1, k);
				v[i] = *xll_range_key(&k, &v[i]);
			}
		}
		
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &v;
}

#ifdef _DEBUG

int test_row(void)
{
	try {
		OPERX o = Excel<XLOPERX>(xlfEvaluate, OPERX(_T("RANGE.SEQUENCE(1,6)")));
		o.resize(2,3);
		OPERX o0 = Excel<XLOPERX>(xlfEvaluate, OPERX(_T("RANGE.SEQUENCE(1,3)")));
		o0.resize(1, 3);
		OPERX o1 = Excel<XLOPERX>(xlfEvaluate, OPERX(_T("RANGE.SEQUENCE(4,6)")));
		o1.resize(1, 3);

	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfterX> xao(test_row);

#endif // _DEBUG