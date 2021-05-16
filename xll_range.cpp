// xll_range.cpp - two-dimensional range of cells
// ??? Can range.scan(monoid, range.get(h)) act on underlying handle ???
#include "xll/xll/xll.h"

#ifndef CATEGORY
#define CATEGORY "Range"
#endif

using namespace xll;

// pointer to underlying range
inline LPOPER ptr(LPOPER po)
{
	if (po->size() == 1 and (*po)[0].is_num()) {
		handle<OPER> h(po->as_num());
		if (h) {
			po = h.ptr();
		}
	}

	return po;
}

// size of row or column, or number of rows of 2-d range
inline unsigned items(const XLOPERX& x)
{
	return rows(x) > 1 ? rows(x) : columns(x);
}
inline unsigned item_index(const XLOPERX& x, unsigned i)
{
	return xmod(i * (rows(x) > 1 ? columns(x) : 1), size(x));
}
// i-th element or i-th row
inline XLOPERX item(XLOPERX x, unsigned i)
{
	if (x.xltype & xltypeMulti) {
		x.val.array.lparray = &index(x, item_index(x, i));
		if (rows(x) > 1) {
			x.val.array.rows = 1;
		}
		else {
			x.val.array.columns = 1;
		}
	}

	return x;
}

#ifdef _DEBUG
int test_item()
{
	OPER o({ OPER(1), OPER(2), OPER(3), OPER(4) });
	XLOPERX x;
	try {
		o.resize(o.size(), 1);
		x = item(o, 0);
		ensure(x.xltype == xltypeMulti);
		ensure(x.val.array.rows == 1);
		ensure(x.val.array.columns == 1);
		ensure(x.val.array.lparray[0].xltype == xltypeNum);
		ensure(x.val.array.lparray[0].val.num == 1);

		x = item(o, 3);
		ensure(x.xltype == xltypeMulti);
		ensure(x.val.array.rows == 1);
		ensure(x.val.array.columns == 1);
		ensure(x.val.array.lparray[0].xltype == xltypeNum);
		ensure(x.val.array.lparray[0].val.num == 4);

		o.resize(1, o.size());
		x = item(o, 0);
		ensure(x.xltype == xltypeMulti);
		ensure(x.val.array.rows == 1);
		ensure(x.val.array.columns == 1);
		ensure(x.val.array.lparray[0].xltype == xltypeNum);
		ensure(x.val.array.lparray[0].val.num == 1);

		x = item(o, 3);
		ensure(x.xltype == xltypeMulti);
		ensure(x.val.array.rows == 1);
		ensure(x.val.array.columns == 1);
		ensure(x.val.array.lparray[0].xltype == xltypeNum);
		ensure(x.val.array.lparray[0].val.num == 4);

		o.resize(2, o.size()/2);
		x = item(o, 0);
		ensure(x.xltype == xltypeMulti);
		ensure(x.val.array.rows == 1);
		ensure(x.val.array.columns == 2);
		ensure(x.val.array.lparray[0].xltype == xltypeNum);
		ensure(x.val.array.lparray[0].val.num == 1);

		x = item(o, 1);
		ensure(x.xltype == xltypeMulti);
		ensure(x.val.array.rows == 1);
		ensure(x.val.array.columns == 2);
		ensure(x.val.array.lparray[1].xltype == xltypeNum);
		ensure(x.val.array.lparray[1].val.num == 4);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_test_item(test_item);
#endif // _DEBUG

AddIn xai_range_set(
	Function(XLL_HANDLEX, "xll_range_set_", "\\RANGE")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is the range to set.", "{1,2;3,4}")
		})
	.Uncalced()
	.FunctionHelp("Return a handle to a range.")
	.Category(CATEGORY)
	.Documentation(R"(
Create a handle to a two dimensional range of cells.
)")
);
HANDLEX WINAPI xll_range_set_(LPOPER px)
{
#pragma XLLEXPORT
	handle<OPER> h(new OPER(*px));

	return h.get();
}

AddIn xai_range_get(
	Function(XLL_LPOPER, "xll_range_get", "RANGE.GET")
	.Arguments({
		Arg(XLL_HANDLEX, "handle", "is a handle returned by RANGE.SET.", "\\RANGE.SET({0,1;2,3})")
		})
	.FunctionHelp("Return the range held by a handle.")
	.Category(CATEGORY)
	.Documentation(R"(
Return a two dimensional range of cells.
)")
);
LPOPER WINAPI xll_range_get(HANDLEX h)
{
#pragma XLLEXPORT
	handle<OPER> h_(h);

	if (!h_) {
		XLL_ERROR("RANGE.GET: unknown handle");

		return nullptr;
	}

	return h_.ptr();
}

AddIn xai_range_fold(
	Function(XLL_LPOPER, "xll_range_fold", "RANGE.FOLD")
	.Arguments({
		Arg(XLL_DOUBLE, "monoid", "is the monoid used to fold."),
		Arg(XLL_LPOPER, "range", "is a range to be folded."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return the right fold of range using monoid.")
	.Documentation(R"(
Accumulate elements of <code>range</code> using <code>monoid</code>.
If <code>range</code> has more than one row and more than one column
then return one row range of fold applied to each column.
)")
);
LPOPER WINAPI xll_range_fold(double m, LPOPER pr)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		OPER M(m);
		o = item(*pr, 0);
		for (unsigned i = 1; i < items(*pr); ++i) {
			o = Excel(xlUDF, M, o, item(*pr, i));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}

AddIn xai_range_scan(
	Function(XLL_LPOPER, "xll_range_scan", "RANGE.SCAN")
	.Arguments({
		Arg(XLL_DOUBLE, "monoid", "is the monoid used to scan."),
		Arg(XLL_LPOPER, "range", "is a range or a handle to a range to be scanned."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return the right scan of range using monoid.")
	.Documentation(R"(
Scan elements of <code>range</code> using <code>monoid</code>.
If <code>range</code> has more than one row and more than one column
then return scan applied to each column.
)")
);
LPOPER WINAPI xll_range_scan(double m, LPOPER pr)
{
#pragma XLLEXPORT
	static OPER o;
	try {
		o.resize(rows(*pr), columns(*pr));
		OPER M(m);
		OPER o0 = item(*pr, 0);
		std::copy(o0.begin(), o0.end(), begin(o));
		for (unsigned i = 1; i < items(*pr); ++i) {
			o0 = Excel(xlUDF, M, o0, item(*pr, i));
			std::copy(o0.begin(), o0.end(), o.begin() + item_index(o, i));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrNA;
	}

	return &o;
}

AddIn xai_range_max(
	Function(XLL_LPOPER, "xll_range_max", "RANGE.MAX")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is a range."),
		Arg(XLL_LPOPER, "_range", "is an optional second range."),
		})
	.FunctionHelp("Maximum of ranges.")
	.Category(CATEGORY)
	.Documentation(R"(
Maximum monoid for ranges.
)")
);
LPOPER WINAPI xll_range_max(const LPOPER pr, const LPOPER p_r)
{
#pragma XLLEXPORT
	static OPER r;

	if (p_r->is_missing()) {
		r.resize(rows(*pr), columns(*pr));
		for (unsigned i = 0; i < size(*pr); ++i) {
			switch (type((*pr)[i])) {
			case xltypeNum:
				r[i] = -std::numeric_limits<double>::infinity();
				break;
			case xltypeStr:
				r[i] = OPER{};
				break;
			default:
				r[i] = OPER{};
			}
		}
	}
	else {
		r = *pr;
		for (unsigned i = 0; i < size(*pr); ++i) {
			r[i] = std::max(r[i], (*p_r)[i]);
		}
	}

	return &r;
}

AddIn xai_range_sum(
	Function(XLL_LPOPER, "xll_range_sum", "RANGE.SUM")
	.Arguments({
		Arg(XLL_LPOPER, "range", "is a range."),
		Arg(XLL_LPOPER, "_range", "is an optional second range."),
		})
	.FunctionHelp("Sum of ranges.")
	.Category(CATEGORY)
	.Documentation(R"(
Sum or concatenate monoid for ranges.
)")
);
LPOPER WINAPI xll_range_sum(const LPOPER pr, const LPOPER p_r)
{
#pragma XLLEXPORT
	static OPER r; // !!! return pr ???

	try {
		if (p_r->is_missing()) {
			r.resize(rows(*pr), columns(*pr));
			for (unsigned i = 0; i < size(*pr); ++i) {
				r[i] = 0;
			}
		}
		else {
			r = *pr;
			for (unsigned i = 0; i < size(*pr); ++i) {
				r[i] = r[i].as_num() + (*p_r)[i].as_num();
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		r = ErrNA;
	}

	return &r;
}

inline XLOPERX range_take(XLOPERX x, long i)
{
	if (i > 0) {
		unsigned n = i;
		if (rows(x) > 1 and n <= rows(x)) {
			x.val.array.rows = n;
		}
		else if (n <= columns(x)) {
			x.val.array.columns = n;
		}
	}
	else if (i < 0) {
		unsigned n = -i;
		if (rows(x) > 1 and n <= rows(x)) {
			x.val.array.lparray = end(x) - n * columns(x);
			x.val.array.rows = n;
		}
		else if (n <= columns(x)) {
			x.val.array.lparray = end(x) - n;
			x.val.array.columns = n;
		}
	}
	else {
		x.val.array.rows = 0;
		x.val.array.columns = 0;
	}

	return x;
}
AddIn xai_range_take(
	Function(XLL_LPOPER, "xll_range_take", "RANGE.TAKE")
	.Arguments({
		Arg(XLL_LONG, "count", "is the number of items to take."),
		Arg(XLL_LPOPER, "range", "is the range to take items from.")
		})
	.FunctionHelp("Take count items for beginning (count > 0) or end (count < 0) of range.")
	.Category(CATEGORY)
	.Documentation(R"xyzyx(
Take <code>count</code> items from <code>range</code>.
)xyzyx")
);
LPOPER WINAPI xll_range_take(LONG n, const LPOPER pr)
{
#pragma XLLEXPORT
	static OPER o;

	o = range_take(*pr, n);

	return &o;
}

#if 0
#ifdef _DEBUG

int xll_test_range()
{
	try {
		OPER o({ OPER(1), OPER("a"), OPER(true), ErrNA });
		HANDLEX h = xll_range_set(&o);
		LPOPER po = xll_range_get(h);
		const OPER& o_ = *po;
		ensure(o_ == o);
		ensure(o == *po);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return FALSE;
	}

	return TRUE;
}
Auto<OpenAfter> xaoa_test_range([]() { return xll_test_range(); });

#endif // _DEBUG
#endif // 0