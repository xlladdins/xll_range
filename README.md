# xll_range

This library provides functions for operating on 2-dimesional _range_ of `OPER`s having type `xltypeMulti`.
A _scalar range_ is a 1x1 multi.
A _vector range_ is a range having only one row or only one column.

Algorithms on ranges act on each _row_ unless it is a vector range, in which case they act on each _element_ of the vector range.  
Ranges in Excel are stored in row major order so acting on columns is not so easy.

Excel is inherently 2-dimensional (or maybe 2.5-dimensional using tabs) but `OPER`s in the library can be
nested ranges. Any element of the range can be a range itself. This allows JSON to be represent as
two column ranges with strings (keys) in the first column. The second column can be a scalar, vector, or a two column
array of another JSON object.

