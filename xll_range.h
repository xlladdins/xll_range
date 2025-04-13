// xll_range.h - common header file
#pragma once
#include "xll24/include/xll.h"

#ifndef CATEGORY
#define CATEGORY "Range"
#endif

namespace xll {

	// pointer to underlying range or nullptr if handle is unknown
	inline LPOPER ptr(LPXLOPER12 po)
	{
		HANDLEX h = 0;
		
		if (isNum(*po)) {
			h = Num(*po);
		} 
		else if (isMulti(*po) and size(*po) == 1 and isNum(po->val.array.lparray[0])) {
			h = Num(po->val.array.lparray[0]);
		}

		handle<OPER> h_(h); // TODO: make sure h = 0 returns nullptr

		return h_.ptr();
	}
	// Const reference to underlying range or ErrNA if handle is unknown
	inline const OPER& ref(LPXLOPER12 px)
	{
		LPOPER po = ptr(px);

		return po ? *po : ErrNA;
	}

}