
#ifndef __XINDOWS_CORE_UTIL_UINTVALUE_H__
#define __XINDOWS_CORE_UTIL_UINTVALUE_H__

struct PROPERTYDESC;

#define MAKEUNITVALUE(val, unit)    ((((long)val)<<CUnitValue::NUM_TYPEBITS)|CUnitValue::unit)

#define MAKE_PIXEL_UNITVALUE(val)   \
    ((((long)val * CUnitValue::TypeNames[CUnitValue::UNIT_PIXELS].wScaleMult)<<CUnitValue::NUM_TYPEBITS) | CUnitValue::UNIT_PIXELS)

#define MAKE_EM_UNITVALUE(val)      \
    ((((long)(val * CUnitValue::TypeNames[CUnitValue::UNIT_EM].wScaleMult))<<CUnitValue::NUM_TYPEBITS) | CUnitValue::UNIT_EM)

#define NULL_UNITVALUE              MAKEUNITVALUE(0, UNIT_NULLVALUE)

class CUnitValue : public CVoid
{
public:
    // Unit-Value types, stored in bits 0..3 of the value long
    enum UNITVALUETYPE
    {
        UNIT_NULLVALUE,     // Value was not specified on input and should not be persisted
        UNIT_POINT,         // Points
        UNIT_PICA,          // Picas
        UNIT_INCH,          // Inches
        UNIT_CM,            // Centimeters
        UNIT_MM,            // Millimeters
        UNIT_EM,            // em's ( relative to width of an "m" character in current-font )
        UNIT_EN,            // en's ( relative to width of an "n" character in current-font )
        UNIT_EX,            // ex's ( relative to height of an "x" character in current-font )
        UNIT_PIXELS,        // Good old-fashioned pixels
        UNIT_PERCENT,       // Percentage
        UNIT_TIMESRELATIVE, // Times something - *
        UNIT_FLOAT,         // A float value (like integer, but with 4 digits of precision)
        UNIT_INTEGER,       // An integer
        UNIT_RELATIVE,      // A relative integer - the sign of the value was specified
        UNIT_ENUM,          // Font size keywords - "large", "small", etc.
        NUM_TYPES,
        MAX_TYPES = 0xFFFFFFFF  // To make this a long.
    };
    // When we get a value that's stored in pels, need to convert using log pixels
    // thus we need to know the direction
    enum DIRECTION
    {
        DIRECTION_CX,
        DIRECTION_CY
    };
    struct TypeDesc
    {
        TCHAR* pszName;
        UNITVALUETYPE uvt;
        WORD wShiftOffset;  // Number of places right to shift dec point when scaling
        WORD wScaleMult;    // Scale mult - equals 10^wShiftOffset
    };
    static CUnitValue::TypeDesc TypeNames[NUM_TYPES];
    struct ConvertTable
    {
        long lMul;
        long lDiv;
    };
    static ConvertTable BasicConversions[6][6];

    CUnitValue() { _lValue = 0; }
    CUnitValue(ULONG lValue) { _lValue = lValue; }

private:
    // The _lValue is the actual long value that encodes the unit & value
    long _lValue;

    enum { TYPE_MASK = 0x0F };
    HRESULT NumberFromString(LPCTSTR pStr, const PROPERTYDESC* pPropDesc);

    static long ConvertTo(long lValue, UNITVALUETYPE uvtFromUnits,
        UNITVALUETYPE uvtTo, DIRECTION direction, long lFontHeight=1);

    // Set the value in current working units
    BOOL SetHimetricMeasureValue(long lNewValue, DIRECTION direction, long lFontHeight);
    BOOL SetPercentValue(long lNewValue, DIRECTION direction, long lPercentOfValue);
    float GetFloatValueInUnits(UNITVALUETYPE uvt, DIRECTION dir, long lFontHeight=1);
    HRESULT SetFloatValueKeepUnits(float fValue, UNITVALUETYPE uvt,
        long lCurrentPixelValue, DIRECTION dir, long lFontHeight);
public:
    static BOOL IsScalerUnit(UNITVALUETYPE uvt);
    HRESULT FormatBuffer(LPTSTR szBuffer, UINT uMaxLen,
        const PROPERTYDESC* pPropDesc, BOOL fAlwaysAppendUnit=FALSE) const;
    enum { NUM_TYPEBITS = 4 };
    long SetValue(long lVal, UNITVALUETYPE uvt);
    long SetRawValue(long lVal) { return (_lValue = lVal); }
    long SetDefaultValue(void) { return SetValue(0, UNIT_NULLVALUE); }
    long GetRawValue(void) const { return _lValue; }

    // Get the value in current working units
    long GetMeasureValue(DIRECTION direction,
        UNITVALUETYPE uvtConvertToUnits=UNIT_PIXELS, long lFontHiehgt=1) const;
    long GetPercentValue(DIRECTION direction, long lPercentOfValue) const;

    HRESULT ConvertToUnitType(UNITVALUETYPE uvt, long lCurrentPixelSize, DIRECTION dir, long lFontHeight=1);
    HRESULT SetFloatUnitValue(float fValue);
    float XGetFloatValueInUnits(UNITVALUETYPE uvt, long lFontHeight=1)
    {
        return GetFloatValueInUnits(uvt, DIRECTION_CX, lFontHeight);
    }
    float YGetFloatValueInUnits(UNITVALUETYPE uvt, long lFontHeight=1)
    {
        return GetFloatValueInUnits(uvt, DIRECTION_CY, lFontHeight);
    }
    HRESULT XConvertToUnitType(UNITVALUETYPE uvt, long lCurrentPixelSize, long lFontHeight)
    {
        return ConvertToUnitType(uvt, lCurrentPixelSize, DIRECTION_CX, lFontHeight);
    }
    HRESULT YConvertToUnitType(UNITVALUETYPE uvt, long lCurrentPixelSize, long lFontHeight)
    {
        return ConvertToUnitType(uvt, lCurrentPixelSize, DIRECTION_CY, lFontHeight);
    }
    HRESULT XSetFloatValueKeepUnits(float fValue, UNITVALUETYPE uvt,
        long lCurrentPixelValue, long lFontHeight=1)
    {
        return SetFloatValueKeepUnits(fValue, uvt, lCurrentPixelValue, DIRECTION_CX, lFontHeight);
    }
    HRESULT YSetFloatValueKeepUnits(float fValue, UNITVALUETYPE uvt,
        long lCurrentPixelValue, long lFontHeight=1)
    {
        return SetFloatValueKeepUnits(fValue, uvt, lCurrentPixelValue, DIRECTION_CY, lFontHeight);
    }

    // Get the PIXEL equiv.
    long XGetPixelValue(CTransform* pTransform, long lPercentOfValue, long yHeight) const
    {
        return GetPixelValue(pTransform, DIRECTION_CX, lPercentOfValue, yHeight);
    }
    long YGetPixelValue(CTransform* pTransform, long lPercentOfValue, long yHeight) const
    {
        return GetPixelValue(pTransform, DIRECTION_CY, lPercentOfValue, yHeight);
    }
    long ScaleTo(UNITVALUETYPE uvt)
    {
        return SetValue(MulDivQuick(GetUnitValue(), TypeNames[uvt].wScaleMult , TypeNames[GetUnitType()].wScaleMult), uvt);
    }
    // GetUnitValue - we need to know whether it was Null
    long GetUnitValue(void) const { return _lValue>>NUM_TYPEBITS; }
    HRESULT FromString(LPCTSTR pStr, const PROPERTYDESC* pPropDesc);
    UNITVALUETYPE GetUnitType(void) const { return (UNITVALUETYPE)(_lValue&TYPE_MASK); }

    HRESULT Persist(IStream* pStream, const PROPERTYDESC* pPropDesc) const;
    HRESULT IsValid(const PROPERTYDESC* pPropDesc) const;

    // These are the 3 primary methods for interfacing
    // Check for Null if you allowed Null values in your PROPDESC
    BOOL IsNull(void) const { return GetUnitType()==UNIT_NULLVALUE; }
    void SetNull(void) { SetValue(0, UNIT_NULLVALUE); }
    BOOL IsNullOrEnum(void) const
    {
        return (GetUnitType()==UNIT_NULLVALUE)||(GetUnitType()==UNIT_ENUM);
    }

    // Store the equiv of a HIMETRIC value
    BOOL SetFromHimetricValue(long lNewValue, DIRECTION direction=DIRECTION_CX,
        long lPercentOfValue=0, long lFontHeight=1);

    BOOL IsPercent() const
    {
        return GetUnitType()==UNIT_PERCENT&&GetUnitValue();
    }

    long GetPercent() const
    {
        Assert(GetUnitType() == UNIT_PERCENT);
        return GetUnitValue()/TypeNames[UNIT_PERCENT].wScaleMult;
    }

    void SetPercent(long l)
    {
        SetValue(l*TypeNames[UNIT_PERCENT].wScaleMult, UNIT_PERCENT);
    }
    // Set from a whole number of points
    void SetPoints(long l)
    {
        SetValue(l*TypeNames[UNIT_POINT].wScaleMult, UNIT_POINT);
    }
    // Get the whole number of points
    long GetPoints() const
    {
        Assert(GetUnitType() == UNIT_POINT);
        return GetUnitValue()/TypeNames[UNIT_POINT].wScaleMult;
    }

    // Get the whole number of EM
    long GetEM() const
    {
        Assert(GetUnitType() == UNIT_EM);
        return GetUnitValue()/TypeNames[UNIT_EM].wScaleMult;
    }

    // Get the whole number of EX
    long GetEX() const
    {
        Assert(GetUnitType() == UNIT_EX);
        return GetUnitValue()/TypeNames[UNIT_EX].wScaleMult;
    }

    // Get the whole number of PICA
    long GetPica() const
    {
        Assert(GetUnitType() == UNIT_PICA);
        return GetUnitValue()/TypeNames[UNIT_PICA].wScaleMult;
    }

    // Get the whole number of Inch
    long GetInch() const
    {
        Assert(GetUnitType() == UNIT_INCH);
        return GetUnitValue()/TypeNames[UNIT_INCH].wScaleMult;
    }

    // Get the whole number of CM
    long GetCM() const
    {
        Assert(GetUnitType() == UNIT_CM);
        return GetUnitValue()/TypeNames[UNIT_CM].wScaleMult;
    }

    // Get the whole number of MM
    long GetMM() const
    {
        Assert(GetUnitType() == UNIT_MM);
        return GetUnitValue()/TypeNames[UNIT_MM].wScaleMult;
    }

    long GetTimesRelative() const
    {
        Assert(GetUnitType() == UNIT_TIMESRELATIVE);
        return GetUnitValue()/TypeNames[UNIT_TIMESRELATIVE].wScaleMult;
    }

    void SetTimesRelative(long l)
    {
        SetValue(l*TypeNames[UNIT_TIMESRELATIVE].wScaleMult, UNIT_TIMESRELATIVE);
    }

    long GetPixelValue(CTransform* pTransform=NULL, DIRECTION direction=DIRECTION_CX,
        long lPercentOfValue=0, long lBaseFontHeight=1) const
    {
        return (_lValue&~TYPE_MASK)==0?0:
            GetPixelValueCore(pTransform, direction, lPercentOfValue, lBaseFontHeight);
    };

    long GetPixelValueCore(CTransform* pTransform=NULL, DIRECTION direction=DIRECTION_CX,
        long lPercentOfValue=0, long lBaseFontHeight=1) const;
};

#endif //__XINDOWS_CORE_UTIL_UINTVALUE_H__