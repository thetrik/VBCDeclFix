	[dllname("KERNEL32.DLL")] 
	module kernel32 {

		const int MEM_RELEASE = 0x8000;
		const int MEM_COMMIT = 0x1000;
		const int MEM_RESERVE = 0x2000;

		const int PAGE_EXECUTE_READWRITE = 0x40;
		const int PAGE_READWRITE = 4;
		const int PAGE_READONLY = 0x02;
		const int PAGE_WRITECOPY=0x08;
		const int PAGE_EXECUTE_WRITECOPY=0x80;
		const int PAGE_EXECUTE_READ=0x20;

		const int CP_ACP = 0;

		[entry("GetModuleHandleW")]
		HANDLE GetModuleHandle(LPWSTR pszModule);

		[entry("LoadLibraryW")]
		HANDLE LoadLibrary(LPWSTR pszModule);

		[entry("FreeLibrary")]
		LONG FreeLibrary(HANDLE hLib);

		[entry("GetProcAddress")]
		PTR GetProcAddress (
			HANDLE hLib,
			LPSTR pszFunction);

		[entry("RtlMoveMemory")]
		long memcpy (
			void *pDest,
			void *pSrc,
			long ByteLen);

		[entry("RtlFillMemory")]
		void memset(
			[in] void* Destination,
			[in] long Length,
			[in] long Fill
		);

		[entry("VirtualProtect")]
		long VirtualProtect (
			long lpAddress,
			long dwSize,
			long flNewProtect,
			long* lpflOldProtect);

		[entry("VirtualAlloc")]
		long VirtualAlloc (
			long lpAddress,
			long dwSize,
			long flAllocationType,
			long flProtect);

		[entry("VirtualFree")]
		long VirtualFree (
			long lpAddress,
			long dwSize,
			long dwFreeType);

		[entry("WideCharToMultiByte")]
		long WideCharToMultiByte(
			[in] long CodePage,
			[in] long dwFlags,
			[in] void* lpWideCharStr,
			[in] long cchWideChar,
			[in] void* lpMultiByteStr,
			[in] long cchMultiByte,
			[in] void* lpDefaultChar,
			[in] long* lpUsedDefaultChar
		);

		[entry("MultiByteToWideChar")]
		long MultiByteToWideChar(
			[in] long CodePage,
			[in] long dwFlags,
			[in] void* lpMultiByteStr,
			[in] long cchMultiByte,
			[in] void* lpWideCharStr,
			[in] long cchWideChar
		);

	}
