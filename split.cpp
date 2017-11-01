// split.cpp - split a string
#include "range.h"

using namespace xll;

inline xcstr tcschr(xcstr s, xchar delim)
{
	xcstr t = _tcschr(s, delim);

	return t ? t : s + _tcslen(s);
}
inline xcstr tcspbrk(xcstr s, xcstr delim)
{
	xcstr t = _tcspbrk(s, delim);

	return t ? t : s + _tcslen(s);
}


// number of fields between s and t
inline xword count(xcstr s, xcstr t, xcstr delim)
{
	xword n = 1;

	for (s = tcspbrk(s, delim); s < t; s = tcspbrk(s + 1, delim))
		++n;

	return n;
}

static AddInX xai_range_split(
	FunctionX(XLL_LPOPERX, _T("?xll_range_split"), _T("RANGE.SPLIT"))
	.Arg(XLL_CSTRINGX, _T("String"), _T("is a string to split."),
		_T("1,2,3;4,5,6"))
	.Arg(XLL_CSTRINGX, _T("FS"), _T("are the field delimiters. Default is \",\"."))
	.Arg(XLL_CSTRINGX, _T("RS"), _T("are the record delimiters. Default is \";\". "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Split a String into a range based on field and record delimiters."))
	.Documentation(_T(""))
);
LPOPERX WINAPI
xll_range_split(xcstr s, xcstr fs, xcstr rs)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		size_t ns = _tcslen(s);
		fs = *fs ? fs : _T(",");
		rs = *rs ? rs : _T(";");
		xstring frs_(fs); frs_.append(rs);
		xcstr frs = frs_.c_str();

		xword r = count(s, s + ns, rs);
		xword c = 0;
		
		xcstr t = s, u = tcspbrk(s, rs);
		for (xword i = 0; i < r; ++i) {
			xword n = count(t, u, fs);
			if (n > c)
				c = n;
			t = u + 1;
			u = tcspbrk(t, rs);
		}

		o.resize(0, 0); // clear out old data
		o.resize(r, c);

		t = s; u = tcspbrk(t, frs);
		for (xword i = 0; i < r; ++i) {
			for (xword j = 0; t != s + ns; ++j, t = u + 1, u = tcspbrk(t, frs)) {
				OPERX o_ = OPERX(t, static_cast<char>(u - t));
				OPERX v_ = Excel<XLOPERX>(xlfValue, o_); // handles dates
				if (v_)
					o(i, j) = v_;
				else if (o_.xltype != xltypeStr) {
					v_ = Excel<XLOPERX>(xlfEvaluate, o_); // numbers and booleans
					o(i, j) = v_ ? v_ : o_;
				}
				else {
					o(i, j) = o_;
				}

				if (*tcschr(fs, *u) == 0 || !*u)
					break; // short rows
			}
			t = u + 1;
			u = tcspbrk(t, frs);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return &o;
}