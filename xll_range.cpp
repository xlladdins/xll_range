// xll_range.cpp - two-dimensional range of cells
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

// i-th item or row
inline const XLOPERX item(const OPER& o, unsigned i)
{
	XLOPERX x;

	if (o.type() != xltypeMulti) {
		ensure(i == 0);
		x.xltype = o.xltype;
		x.val = o.val;
	}
	else {
		x.xltype = xltypeMulti;
		x.val.array.rows = 1;

		if (o.rows() > 1 and o.columns() > 1) {
			x.val.array.columns = o.columns();
			x.val.array.lparray = o.val.array.lparray + i * o.columns();
		}
		else {
			x.val.array.columns = 1;
			x.val.array.lparray = o.val.array.lparray + i;
		}
	}

	return x;
}

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
		Arg(XLL_LPOPER, "monoid", "is the monoid used to fold."),
		Arg(XLL_LPOPER, "range", "is a range or a handle to a range to be folded."),
		})
		.Category(CATEGORY)
	.FunctionHelp("Return the right fold of range using monoid.")
	.Documentation(R"(
Accumulate elements of <code>range</code> using <code>monoid</code>.
If <code>range</code> has more than one row and more than one column
then return one row range of fold applied to each column.
)")
);
LPOPER WINAPI xll_range_fold(const LPOPER pm, const LPOPER pr)
{
#pragma XLLEXPORT
	static OPER o;

	try {
		LPOPER pr_ = ptr(pr);

		o = Excel(xlUDF, *pm, item(*pr_, 0));
		unsigned n = pr_->size() / o.size();
		for (unsigned i = 0; i < n; ++i) {
			o = Excel(xlUDF, *pm, o, item(*pr_, i));
		}

		if (pr_ != pr) {
			*pr_ = o; // assign to handle value
			o = *pr; // return handle
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
			r[i] = -std::numeric_limits<double>::infinity();
		}
	}
	else {
		r = *pr;
		for (unsigned i = 0; i < size(*pr); ++i) {
			r[i] = std::max(r[i].as_num(), (*p_r)[i].as_num());
		}
	}

	return &r;
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