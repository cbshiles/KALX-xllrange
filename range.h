// range.h - Excel range functions.
// Uncomment the following line to use features for Excel2007 and above.
/*
UNIQUE(a,{3,2})
GRADE(a,n), GRADE(a,{1,2}), GRADE(a,{1,2},n), GRADE(a,n,{1,2})
FILTER(a,Pred)
*/
#pragma once
#pragma warning(disable: 4244)
//#define EXCEL12
#include <algorithm>
#include <functional>
//#include "../xll8/xll/xll.h"
#include "xll/xll.h"

#ifndef CATEGORY
#define CATEGORY _T("Range")
#endif
#define RANGE_PREFIX _T("RANGE.")

typedef xll::traits<XLOPERX>::xchar xchar;
typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xstring xstring;
typedef xll::traits<XLOPERX>::xword xword;
typedef xll::traits<XLOPERX>::xfp xfp;

#define IS_RANGE _T("is a range or a handle to a range.")

namespace range {

	// cyclic 0-based index in the range [0, n)
	template<class I>
	inline I index0_(int i, I n)
	{
		return (i >= 0) ? (i < n ? i : i%n) : n - 1 - (-i - 1)%n;
	}

	// Convert cyclic 1-based index to the range [0, n)
	inline xword index1_(int i, xword n)
	{
		ensure (i != 0);

		return i > 0 ? index0_(i - 1, n) : index0_(i, n);
	}

	inline OPERX index_(const OPERX& a, const OPERX& r, const OPERX& c)
	{
		OPERX o;

		if (r.xltype == xltypeMissing) {
			if (c.xltype == xltypeMissing)
				return OPERX(xltype::Missing);

			o.resize(a.rows(), c.size());
			for (xword i = 0; i < a.rows(); ++i)
				for (xword j = 0; j < c.size(); ++j)
					o(i, j) = a(i, range::index1_(c[j], a.columns()));
		}
		else if (c.xltype == xltypeMissing) {
			o.resize(r.size(), a.columns());
			for (xword i = 0; i < r.size(); ++i)
				for (xword j = 0; j < a.columns(); ++j)
					o(i, j) = a(range::index1_(r[i], a.rows()), j);
		}
		else {
			o.resize(r.size(), c.size());
			for (xword i = 0; i < o.rows(); ++i)
				for (xword j = 0; j < o.columns(); ++j)
					o(i, j) = a(range::index1_(r[i], a.rows()), range::index1_(c[j], a.columns()));			
		}

		return o;
	}

	// if px not a multi, look up using xll::handle
	inline LPOPERX handle(LPOPERX px)
	{
		if (px->xltype != xltypeNum)
			return 0;
	
		xll::handle<OPERX> h(px->val.num, true);	

		return h;
	}
	// range or pointer to in memory range
	inline LPOPERX lookup(LPOPERX pr)
	{
		LPOPERX px = handle(pr);

		return px ? px : pr;
	}

	inline XLOPERX wrap_(xword r, xword c, XLOPERX* a)
	{
		XLOPERX x;

		x.xltype = xltypeMulti;
		x.val.array.rows = r;
		x.val.array.columns = c;
		x.val.array.lparray = a;

		return x;
	}
	inline const XLOPERX wrap_(xword r, xword c, const XLOPERX* a)
	{
		XLOPERX x;

		x.xltype = xltypeMulti;
		x.val.array.rows = r;
		x.val.array.columns = c;
		x.val.array.lparray = const_cast<xloper*>(a);

		return x;
	}

	inline bool is_vector(const XLOPERX& a)
	{
		return a.val.array.rows == 1 || a.val.array.columns == 1;
	}
	// vectors 1-d, ranges stride by column
	inline xword stride(const XLOPERX& a)
	{
		return is_vector(a) ? 1 : a.val.array.columns;
	}

	class oper_iter : public std::iterator<std::random_access_iterator_tag, XLOPERX>
	{
		XLOPERX r_;
	public:
		oper_iter()
			: r_()
		{ }
		oper_iter(const oper_iter& oi)
			: r_(oi.r_)
		{ }
		oper_iter(xword n, XLOPERX* a)
		{
			r_.xltype = xltypeMulti;
			r_.val.array.rows = 1;
			r_.val.array.columns = n;
			r_.val.array.lparray = a;
		}
		oper_iter(const XLOPERX& x)
		{
			ensure (x.val.array.rows == 1);

			r_.xltype = xltypeMulti;
			r_.val.array.rows = 1;
			r_.val.array.columns = x.val.array.columns;
			r_.val.array.lparray = x.val.array.lparray;
		}
/*		oper_iter(oper_iter&& oi)
			: r_(oi.r_)
		{ }
*/		oper_iter& operator=(const oper_iter& oi)
		{
			if (this != &oi)
				r_ = oi.r_;
		
			return *this;
		}
/*		oper_iter& operator=(oper_iter&& oi)
		{
			if (this != &oi)
				r_ = oi.r_;
		
			return *this;
		}
*/		~oper_iter()
		{ }

		// pointer, not value, equality
		bool operator==(const oper_iter& ia) const
		{
			return r_.val.array.lparray == ia.r_.val.array.lparray;
		}
		bool operator!=(const oper_iter& ia) const
		{
			return !operator==(ia);
		}
		bool operator<(const oper_iter& ia) const
		{
			return r_.val.array.lparray < ia.r_.val.array.lparray;
		}

		// trivial iterator returning one row
		XLOPERX operator*(void)
		{
			return r_;
		}

		// forward iterator
		oper_iter& operator++(void)
		{
			r_.val.array.lparray += r_.val.array.columns;

			return *this;
		}
		oper_iter operator++(int)
		{
			oper_iter i(*this);

			++i;

			return i;
		}

		// bidirectional iterator
		oper_iter& operator--(void)
		{
			r_.val.array.lparray -= r_.val.array.columns;

			return *this;
		}
		oper_iter operator--(int)
		{
			oper_iter i(*this);

			--i;

			return *this;
		}
		ptrdiff_t operator-(const oper_iter& r) const
		{
			return (r_.val.array.lparray - r.r_.val.array.lparray)/r_.val.array.columns;
		}

		// random access
		oper_iter& operator+=(size_t n)
		{
			r_.val.array.lparray += n*r_.val.array.columns;

			return *this;
		}
		oper_iter& operator-=(size_t n)
		{
			r_.val.array.lparray -= n*r_.val.array.columns;

			return *this;
		}
		oper_iter operator+(size_t n) const
		{
			oper_iter i(r_);

			i += n;

			return i;
		}
		oper_iter operator-(size_t n) const
		{
			oper_iter i(r_);

			i -= n;

			return i;
		}
		void iter_swap(oper_iter& oi)
		{
			std::swap_ranges(r_.val.array.lparray, r_.val.array.lparray + r_.val.array.columns, oi.r_.val.array.lparray); 
		}
	};

	inline oper_iter begin_iter(XLOPERX& x)
	{
		oper_iter i(x.val.array.columns, x.val.array.lparray);

		return i;
	}
	inline oper_iter end_iter(XLOPERX& x)
	{
		oper_iter i(x.val.array.columns, x.val.array.lparray + ::size(x));

		return i;
	}

	inline void drop_(int n, OPERX& o)
	{
		if (o.rows() > 1) {
			if (n > 0) {
				xword r = (std::max)(o.rows() - n, 0);
				std::copy(o.end() - r*o.columns(), o.end(), o.begin());
				o.resize(r, o.columns());
			}
			else if (n < 0) {
				xword r = (std::max)(o.rows() + n, 0);
				o.resize(r, o.columns());
			}
		}
		else {
			if (n > 0) {
				xword c = (std::max)(o.size() - n, 0);
				std::copy(o.end() - c, o.end(), o.begin());
				o.resize(1, c);
			}
			else if (n < 0) {
				xword c = (std::max)(o.size() + n, 0);
				o.resize(1, c);
			}
		}
	}


	inline void take_(int n, OPERX& o)
	{
		if (o.rows() > 1) {
			if (n > 0) {
				xword r = (std::min<xword>)(o.rows(), n);
				o.resize(r, o.columns());
			}
			else if (n < 0) {
				xword r = (std::min<xword>)(o.rows(), -n);
				std::copy(o.end() - r*o.columns(), o.end(), o.begin());
				o.resize(r, o.columns());
			}
		}
		else {
			if (n > 0) {
				xword c = (std::min<xword>)(o.size(), n);
				std::copy(o.end() - c, o.end(), o.begin());
				o.resize(1, c);
			}
			else if (n < 0) {
				xword c = (std::min<xword>)(o.size(), -n);
				o.resize(1, c);
			}
		}
	}

	inline void sequence(OPERX& o, int step = 1, int offset = 0)
	{
		for (xword i = 0; i < o.size(); ++i)
			o[i] = offset + i*step;
	}

	// compare all columns in order
	inline std::function<bool(const OPERX&, const OPERX&)> compare(const OPERX& o,
		const std::function<bool(const OPERX&, const OPERX&)>& lg)
	{
		return [&o,&lg](const OPERX& i, const OPERX& j) -> bool {
			xword oi = static_cast<xword>(i.val.num) - 1;
			xword oj = static_cast<xword>(j.val.num) - 1;

			for (xword k = 0; k < o.columns(); ++k) {
				if (lg(o(oi, k), o(oj, k)))
					return true;
				else if (lg(o(oj, k), o(oi, k)))
					return false;
			}

			return false;
		};
	}
	// compare o based on columns c
	inline std::function<bool(const OPERX&, const OPERX&)> compare(const OPERX& o, const OPERX& c, 
		const std::function<bool(const OPERX&, const OPERX&)>& lg)
	{
		return [&o,&c,&lg](const OPERX& i, const OPERX& j) -> bool {
			xword oi = static_cast<xword>(i.val.num) - 1;
			xword oj = static_cast<xword>(j.val.num) - 1;

			for (xword k = 0; k < c.size(); ++k) {
				ensure (c[k].xltype == xltypeNum);
				xword ck = static_cast<xword>(c[k].val.num) - 1;
				if (lg(o(oi, ck), o(oj, ck)))
					return true;
				else if (lg(o(oj, ck), o(oi, ck)))
					return false;
			}

			return false;
		};
	}

	// partial grade all columsn of o
	inline OPERX partial_grade(const OPERX& o, int n, const OPERX& c)
	{
		OPERX s;

		if (is_vector(o)) {
			s.resize(o.size(), 1);
			sequence(s, 1, 1);
			XLOPERX o_ = wrap_(o.size(), 1, o.val.array.lparray);

			if (n > 0) {
				n = (std::min<int>)(n, o.size());
				if (c.xltype == xltypeMissing)
					std::partial_sort(s.begin(), s.begin() + n, s.end(), compare(o_, std::less<OPERX>()));
				else
					std::partial_sort(s.begin(), s.begin() + n, s.end(), compare(o_, c, std::less<OPERX>()));
			}
			else if (n < 0) {
				n = (std::max<int>)(n, -o.size());
				if (c.xltype == xltypeMissing)
					std::partial_sort(s.begin(), s.begin() - n, s.end(), compare(o_, std::greater<OPERX>()));
				else
					std::partial_sort(s.begin(), s.begin() - n, s.end(), compare(o_, c, std::greater<OPERX>()));
			}

			s.resize(o.rows(), o.columns());
		}
		else {
			s.resize(o.rows(), 1);
			sequence(s, 1, 1);
			if (n > 0) {
				n = (std::min<int>)(n, o.size());
				if (c.xltype == xltypeMissing)
					std::partial_sort(s.begin(), s.begin() + n, s.end(), compare(o, std::less<OPERX>()));
				else
					std::partial_sort(s.begin(), s.begin() + n, s.end(), compare(o, c, std::less<OPERX>()));
			}
			else if (n < 0) {
				n = (std::max<int>)(n, -o.size());
				if (c.xltype == xltypeMissing)
					std::partial_sort(s.begin(), s.begin() - n, s.end(), compare(o, std::greater<OPERX>()));
				else
					std::partial_sort(s.begin(), s.begin() - n, s.end(), compare(o, c, std::greater<OPERX>()));
			}
		}

		return s;
	}

	inline OPERX grade_(const OPERX& o, int n = 0, const OPERX& c = OPERX(xltype::Missing))
	{
		OPERX s;

		if (n > 0 || n < -1)
			return partial_grade(0, n, c);

		if (is_vector(o)) {
			s.resize(o.size(), 1);
			sequence(s, 1, 1);
			XLOPERX o_ = wrap_(o.size(), 1, o.val.array.lparray);

			if (n >= 0) {
				if (c.xltype == xltypeMissing)
					std::sort(s.begin(), s.end(), compare(o_, std::less<XLOPER>()));
				else
					std::sort(s.begin(), s.end(), compare(o_, c, std::less<XLOPER>()));
			}
			else {
				if (c.xltype == xltypeMissing)
					std::sort(s.begin(), s.end(), compare(o_, std::greater<XLOPER>()));
				else
					std::sort(s.begin(), s.end(), compare(o_, c, std::greater<XLOPER>()));
			}

			s.resize(o.rows(), o.columns());
		}
		else {
			s.resize(o.rows(), 1);
			sequence(s, 1, 1);

			if (n == 0) {
				if (c.xltype == xltypeMissing)
					std::sort(s.begin(), s.end(), compare(o, std::less<XLOPER>()));
				else
					std::sort(s.begin(), s.end(), compare(o, c, std::less<XLOPER>()));
			}
			else {
				if (c.xltype == xltypeMissing)
					std::sort(s.begin(), s.end(), compare(o, std::greater<XLOPER>()));
				else
					std::sort(s.begin(), s.end(), compare(o, c, std::greater<XLOPER>()));
			}
		}

		return s;
	}

	inline void derange(OPERX& o, const OPERX& g)
	{
		for (xword i = 0; i < g.size(); i++) {
			xword j = g[i] - 1;

			while (j < i)
				j = g[j] - 1;
			
			if (o.rows() > 1 && o.columns() > 1)
				std::swap_ranges(o.begin() + i*o.columns(), o.begin() + (i + 1)*o.columns(), o.begin() + j*o.columns());
			else
				std::swap(o[i], o[j]);
		}
	}

	inline void sort_(OPERX& o, xword n = 0)
	{
		derange(o, grade_(o, n));
	}

	inline void mask_(OPERX& o, const OPERX& m)
	{
		ensure (o.rows() > 1 ? o.rows() == m.size() : o.size() == m.size());

		if (o.rows() > 1 && o.columns() > 1) {
			xword c = o.columns();
			xword j = 0;
			for (xword i = 0; i < o.rows(); ++i) {
				if (m[i]) {
					std::copy(o.begin() + i*c, o.begin() + (i + 1)*c, o.begin() + j*c);
					++j;
				}
			}
			o.resize(j, c);
		}
		else {
			xword j = 0;
			for (xword i = 0; i < o.size(); ++i) {
				if (m[i]) {
					o[j++] = o[i];
				}
			}
			if (o.rows() == 1)
				o.resize(1, j);
			else
				o.resize(j, 1);
		}
	}

	inline void stride_(OPERX& o, int step = 1, int off = 0)
	{
		xword i = 0;
		for (; off + i*step < o.size(); ++i)
			o[i] = o[off + i * step];

		if (o.rows() == 1)
			o.resize(1, i);
		else if (o.columns() == 1)
			o.resize(i, 1);
		else { // assume we are pulling out columns
			o.resize(o.rows(), i/o.rows());
		}
	}

	inline void unique_(OPERX& o, const OPERX& c = OPERX(xltype::Missing))
	{
		if (is_vector(o)) {
			LPOPERX pe = std::unique(o.begin(), o.end());
			if (o.rows() == 1)
				o.resize(1, pe - o.begin());
			else
				o.resize(pe - o.begin(), 1);
		}
		else {
			oper_iter b(begin_iter(o)), e(end_iter(o));
			oper_iter ei = std::unique(b, e);
			o.resize(ei - b, o.columns());
		}
	}
	inline OPERX prepend(const OPERX& o, xll::traits<XLOPERX>::xcstr s)
	{
		ensure (o.xltype == xltypeStr);

		return xll::Excel<XLOPERX>(xlfConcatenate, OPERX(s), o);
	}
	inline OPERX formula(const OPERX& o)
	{
		return prepend(o, _T("="));
	}
	inline OPERX bang(const OPERX& o)
	{
		return prepend(o, _T("!"));
	}
	inline OPERX eval(const OPERX& o)
	{
		return xll::Excel<XLOPERX>(xlfEvaluate, o);
	}

	// call UDF using elements of range as arguments
	inline OPERX call(HANDLEX f, LPXLOPERX po, bool command = false)
	{
		OPERX x(1,1), _x;
		x[0] = f;

		if (po->xltype == xltypeStr) {
			_x = eval(bang(*po));
			po = &_x;
		}
		else {
			x.push_back(*po);
			x.resize(1, x.size());
		}

		if (!command) {
			// convert handles to values
			for (xword i = 0; i < x.size(); ++i) {
				if (x[i].xltype == xltypeNum) {
					xll::handle<OPERX> h(x[i].val.num);
					if (h)
						x[i] = *h;
				}
			}
		}

		return xll::traits<XLOPERX>::Excelv(xlUDF, x.size(), x.begin());
	}

} // namespace range

