
#include <Windows.h>
#include <shlwapi.h>

#pragma comment (lib, "Shlwapi.lib")

BOOL WINAPI DllMain( HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved) {
    return TRUE;
}

int __cdecl SumInt(int a, int b) {
	return a + b;
}

__int64 __cdecl SumCur(__int64 a, __int64 b) {
	return a + b;
}

double __cdecl SumDbl(double a, double b) {
	return a + b;
}

float __cdecl SumFlt(float a, float b) {
	return a + b;
}

short __cdecl SumShort(short a, short b) {
	return a + b;
}

void __cdecl Void(short a, short b) {
	return;
}

BYTE __cdecl SumByte(BYTE a, BYTE b) {
	return a + b;
}

HRESULT __cdecl SumHresult(int a, int b) {
	return a + b;
}

IUnknown* __cdecl CreateStream(PVOID a, int b) {
	return SHCreateMemStream((BYTE*)a, b);
}

BSTR __cdecl SumStr(int a, int b) {
	BSTR s1, s2, s3;
	VarBstrFromI4(a, GetUserDefaultLCID(), 0, &s1);
	VarBstrFromI4(b, GetUserDefaultLCID(), 0, &s2);
	s3 = SysAllocStringLen(NULL, SysStringLen(s1) + SysStringLen(s2));
	lstrcpy(s3, s1);
	lstrcat(s3, s2);
	SysFreeString(s1);
	SysFreeString(s2);
	return s3;
}

SAFEARRAY* __cdecl GetSA(int a, int b) {
	SAFEARRAY *saRet;

	SafeArrayAllocDescriptor(1, &saRet);

	saRet->cbElements = 4;
	saRet->rgsabound->lLbound = 0;
	saRet->rgsabound->cElements = 2;

	SafeArrayAllocData(saRet);

	((DWORD*)saRet->pvData)[0] = a;
	((DWORD*)saRet->pvData)[1] = b;

	return saRet;

}

VARIANT __cdecl SumVnt(VARIANT a, VARIANT b) {
	VARIANT r;
	VariantInit(&r);
	VarAdd(&a, &b, &r);
	return r;
}