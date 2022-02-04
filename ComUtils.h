#pragma once

struct SourceLocation
{
    SourceLocation(int line, const TCHAR* file, const TCHAR* function, const TCHAR* funcsig)
        : line(line)
        , file(file)
        , function(function)
        , funcsig(funcsig)
    {
    }

    int line;
    const TCHAR* file;
    const TCHAR* function;
    const TCHAR* funcsig;
};

#define SRC_LOC SourceLocation(__LINE__, _T(__FILE__), _T(__FUNCTION__), _T(__FUNCSIG__))

inline bool Validate(HRESULT hr, const WCHAR* msg, const SourceLocation& src_loc)
{
    if (FAILED(hr))
    {
        _ftprintf(stderr, L"FAILED [%s(%d)] %#10.8x: %s\n", src_loc.file, src_loc.line, hr, msg);
        return false;
    }
    else
        return true;
}

inline bool ValidateIgnore(HRESULT hr, HRESULT ignore, const WCHAR* msg, const SourceLocation& src_loc)
{
    return hr == ignore ? true : Validate(hr, msg, src_loc);
}

#define CHECK(x, m, r) if (!Validate(x, m, SRC_LOC)) return r;
#define CHECK_IGNORE(x, i, m, r) if (!ValidateIgnore(x, i, m, SRC_LOC)) return r;

template<class T>
struct ObjectArrayIndex
{
    UINT i;
    IObjectArray* pObjectArray;

    ObjectArrayIndex& operator++()
    {
        ++i;
        return *this;
    }

    CComPtr<T> operator*()
    {
        CComPtr<T> p;
        if (FAILED(pObjectArray->GetAt(i, IID_PPV_ARGS(&p))))
            return nullptr;
        return p;
    }
};

template<class T>
bool operator!=(const ObjectArrayIndex<T>& l, const ObjectArrayIndex<T>& r)
{
    assert(l.pObjectArray == r.pObjectArray);
    return l.i != r.i;
}

template<class T>
class ObjectArrayRange
{
private:
    IObjectArray* pObjectArray;

public:
    ObjectArrayRange(IObjectArray* pObjectArray)
        : pObjectArray(pObjectArray)
    {
    }

    ObjectArrayIndex<T> begin() const
    {
        return { 0, pObjectArray };
    }

    ObjectArrayIndex<T> end() const
    {
        UINT count;
        if (FAILED(pObjectArray->GetCount(&count)))
            count = 0;
        return { count, pObjectArray };
    }
};

template<class T>
struct ObjectArrayIndexRev
{
    UINT i;
    IObjectArray* pObjectArray;

    ObjectArrayIndexRev& operator++()
    {
        --i;
        return *this;
    }

    CComPtr<T> operator*()
    {
        CComPtr<T> p;
        if (FAILED(pObjectArray->GetAt(i - 1, IID_PPV_ARGS(&p))))
            return nullptr;
        return p;
    }
};

template<class T>
bool operator!=(const ObjectArrayIndexRev<T>& l, const ObjectArrayIndexRev<T>& r)
{
    assert(l.pObjectArray == r.pObjectArray);
    return l.i != r.i;
}

template<class T>
class ObjectArrayRangeRev
{
private:
    IObjectArray* pObjectArray;

public:
    ObjectArrayRangeRev(IObjectArray* pObjectArray)
        : pObjectArray(pObjectArray)
    {
    }

    ObjectArrayIndexRev<T> begin() const
    {
        UINT count;
        if (FAILED(pObjectArray->GetCount(&count)))
            count = 0;
        return { count, pObjectArray };
    }

    ObjectArrayIndexRev<T> end() const
    {
        return { 0, pObjectArray };
    }
};
