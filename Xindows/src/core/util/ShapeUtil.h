
#ifndef __XINDOWS_CORE_UTIL_SHAPEUTIL_H__
#define __XINDOWS_CORE_UTIL_SHAPEUTIL_H__

#define SHAPE_TYPE_RECT     0
#define SHAPE_TYPE_CIRCLE   1
#define SHAPE_TYPE_POLY     2

#define DELIMS              _T(", ;")

DECLARE_CDataAry(CPointAry, POINT)

typedef struct
{
    LONG lx, ly, lradius;
} SCircleCoords;

typedef struct
{
    HRGN hPoly;
} SPolyCoords;

union CoordinateUnion
{
    RECT Rect;
    SCircleCoords Circle;
    SPolyCoords Polygon;
};

HRESULT NextNum(LONG* plNum, TCHAR** ppch);

BOOL PointInCircle(POINT pt, LONG lx, LONG ly, LONG lradius);
BOOL Contains(POINT pt, union CoordinateUnion coords, UINT nShapeType);

#endif //__XINDOWS_CORE_UTIL_SHAPEUTIL_H__