	[dllname("MSVBVM60.DLL")]
	module msvbvm60 {

		[entry("GetMem1")]
		void GetMem1(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("GetMem2")]
		void GetMem2(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("GetMem4")]
		void GetMem4(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("GetMem8")]
		void GetMem8(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("GetMem4")]
		void GetMemPtr(
			[in] void *pSrc,
			[in, out] void *pDst);

		[entry("PutMem1")]
		void PutMem1(
			[in, out] void *pDst,
			[in] unsigned char v);

		[entry("PutMem2")]
		void PutMem2(
			[in, out] void *pDst,
			[in] short v);

		[entry("PutMem4")]
		void PutMem4(
			[in, out] void *pDst,
			[in] long v);

		[entry("PutMem4")]
		void PutMemPtr(
			[in, out] void *pDst,
			[in] PTR v);

		[entry("PutMem8")]
		void PutMem8(
			[in, out] void *pDst,
			[in] long vL,
			[in] long vH);

		[entry("__vbaAryMove")]
		void vbaAryMove(
			[in] void* pDst,
			[in] void* Source);

		[entry("__vbaObjSetAddref")]
		void vbaObjSetAddref(
			[in, out] void *pDstObj,
			[in] void *pSrcObj);

		[entry("__vbaObjSet")]
		void vbaObjSet(
			[in, out] void *pDstObj,
			[in] void *pSrcObj);

		[entry("__vbaObjAddref")]
		void vbaObjAddref(
			[in] void *pDstObj);

		[entry("__vbaCopyBytesZero")]
		void vbaCopyBytesZero(
			[in] long lSize,
			[in] void *pDst,
			[in] void *pSrc);

		[entry("rtcCallByName")]
		void rtcCallByName(
			[in, out] VARIANT *vRet,
			[in] IDispatch *cObj,
			[in] BSTR sMethod,
			[in] long lCallType,
			[in] SAFEARRAY(VARIANT) *pArgs,
			[in] long lcid);

		[entry("__vbaVarSetObjAddref")]
		void vbaVarSetObjAddref(
			[in, out] void *pDstVarObj,
			[in] void *pSrcObj);

		[entry("__vbaVarSetObj")]
		void vbaVarSetObj(
			[in, out] void *pDstVarObj,
			[in] void *pSrcObj);

		[entry("VarPtr")]
		PTR ArrPtr(
			[in] SAFEARRAY(void) *pDst);

	}