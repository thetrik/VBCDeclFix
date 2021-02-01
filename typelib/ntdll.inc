
[dllname("NTDLL.DLL")] 
module ntdll {

		[entry("RtlCompareMemory")]
		long memcmp (
			void *pSource1,
			void *pSource2,
			long uLength);

};