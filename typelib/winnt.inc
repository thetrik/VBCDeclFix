typedef struct {
	short e_magic;
	short e_cblp;
	short e_cp;
	short e_crlc;
	short e_cparhdr;
	short e_minalloc;
	short e_maxalloc;
	short e_ss;
	short e_sp;
	short e_csum;
	short e_ip;
	short e_cs;
	short e_lfarlc;
	short e_ovno;
	short e_res[4];
	short e_oemid;
	short e_oeminfo;
	short e_res2[10];
	long e_lfanew;
} IMAGE_DOS_HEADER;

typedef struct {
	long VirtualAddress;
	long Size;
} IMAGE_DATA_DIRECTORY;

typedef struct { 
	short Magic;
	unsigned char MajorLinkerVersion;
	unsigned char MinorLinkerVersion;
	long SizeOfCode;
	long SizeOfInitializedData;
	long SizeOfUnitializedData;
	long AddressOfEntryPoint;
	long BaseOfCode;
	long BaseOfData;
	PTR ImageBase;
	long SectionAlignment;
	long FileAlignment;
	short MajorOperatingSystemVersion;
	short MinorOperatingSystemVersion;
	short MajorImageVersion;
	short MinorImageVersion;
	short MajorSubsystemVersion;
	short MinorSubsystemVersion;
	long W32VersionValue;
	long SizeOfImage;
	long SizeOfHeaders;
	long CheckSum;
	short SubSystem;
	short DllCharacteristics;
	SIZE_T SizeOfStackReserve;
	SIZE_T SizeOfStackCommit;
	SIZE_T SizeOfHeapReserve;
	SIZE_T SizeOfHeapCommit;
	long LoaderFlags;
	long NumberOfRvaAndSizes;
	IMAGE_DATA_DIRECTORY DataDirectory[16];
} IMAGE_OPTIONAL_HEADER;

typedef struct { 
	short Machine;
	short NumberOfSections;
	long TimeDateStamp;
	long PointerToSymbolTable;
	long NumberOfSymbols;
	short SizeOfOptionalHeader;
	short Characteristics;
} IMAGE_FILE_HEADER;


typedef struct {
    long Signature;
    IMAGE_FILE_HEADER FileHeader;
    IMAGE_OPTIONAL_HEADER OptionalHeader;
} IMAGE_NT_HEADERS;

typedef struct {
	unsigned char SectionName[8];
	long VirtualSize;
	long VirtualAddress;
	long SizeOfRawData;
	long PointerToRawData;
	long PointerToRelocations;
	long PointerToLinenumbers;
	short NumberOfRelocations;
	short NumberOfLinenumbers;
	long Characteristics;
} IMAGE_SECTION_HEADER;

typedef struct {
	long Characteristics;
	long TimeDateStamp;
	short MajorVersion;
	short MinorVersion;
	short NumberOfNamedEntries;
	short NumberOfIdEntries;
} IMAGE_RESOURCE_DIRECTORY;

typedef struct {
	long NameId;
	long OffsetToData;
} IMAGE_RESOURCE_DIRECTORY_ENTRY;

typedef struct {
	long OffsetToData;
	long Size;
	long CodePage;
	long Reserved;
} IMAGE_RESOURCE_DATA_ENTRY;

typedef struct {
	long VirtualAddress;
	long SizeOfBlock;
} IMAGE_BASE_RELOCATION;

typedef struct {
	long cb;
	short MajorRuntimeVersion;
	short MinorRuntimeVersion;
	IMAGE_DATA_DIRECTORY MetaData;
	long flags;
	long EntryPointRVA;
	IMAGE_DATA_DIRECTORY Resources;
	IMAGE_DATA_DIRECTORY StrongNameSignature;
	IMAGE_DATA_DIRECTORY CodeManagerTable;
	IMAGE_DATA_DIRECTORY VTableFixups;
	IMAGE_DATA_DIRECTORY ExportAddressTableJumps;
	IMAGE_DATA_DIRECTORY ManagedNativeHeader;
} IMAGE_COR20_HEADER;

typedef struct {
	long lSignature;
	short iMajorVer;
	short iMinorVer;
	long iExtraData;
	long iVersionString;
} STORAGESIGNATURE;

typedef struct {
	long VirtualAddress;
	long SymbolTableIndex;
	short Type;
} IMAGE_RELOCATION;

typedef struct {
	unsigned char SymbolName[8];
	long Value;
	short SectionNumber;
	short Type;
	unsigned char StorageClass;
	unsigned char NumberOfAuxSymbols;
} IMAGE_SYMBOL;

module winnt {

	const int IMAGE_DOS_SIGNATURE = 0x5A4D;
	const int IMAGE_NT_SIGNATURE = 0x4550;

	const int IMAGE_NT_OPTIONAL_HDR64_MAGIC = 0x20B;
	const int IMAGE_NT_OPTIONAL_HDR32_MAGIC = 0x10B;

	const int IMAGE_DIRECTORY_ENTRY_SECURITY = 4;
	const int IMAGE_DIRECTORY_ENTRY_BASERELOC = 5;
	const int IMAGE_DIRECTORY_ENTRY_TLS = 9;
	const int IMAGE_DIRECTORY_ENTRY_DEBUG = 6;
	const int IMAGE_DIRECTORY_ENTRY_EXCEPTION = 3;
    const int IMAGE_DIRECTORY_ENTRY_ARCHITECTURE = 7;
    const int IMAGE_DIRECTORY_ENTRY_GLOBALPTR = 8;
	const int IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR = 14;
	const int IMAGE_DIRECTORY_ENTRY_RESOURCE = 2;

	const int IMAGE_REL_BASED_ABSOLUTE = 0;
	const int IMAGE_REL_BASED_HIGHLOW = 3;

	const int IMAGE_FILE_RELOCS_STRIPPED = 0x0001;  /*  Relocation info stripped from file. */
	const int IMAGE_FILE_EXECUTABLE_IMAGE = 0x0002;  /*  File is executable  (i.e. no unresolved externel references). */
	const int IMAGE_FILE_LINE_NUMS_STRIPPED = 0x0004;  /*  Line nunbers stripped from file. */
	const int IMAGE_FILE_LOCAL_SYMS_STRIPPED = 0x0008;  /*  Local symbols stripped from file. */
	const int IMAGE_FILE_BYTES_REVERSED_LO = 0x0080;  /*  Bytes of machine word are reversed. */
	const int IMAGE_FILE_32BIT_MACHINE = 0x0100;  /*  32 bit word machine. */
	const int IMAGE_FILE_DEBUG_STRIPPED = 0x0200;  /*  Debugging info stripped from file in .DBG file */
	const int IMAGE_FILE_SYSTEM = 0x1000;  /*  System File. */
	const int IMAGE_FILE_DLL = 0x2000;  /*  File is a DLL. */
	const int IMAGE_FILE_BYTES_REVERSED_HI = 0x8000;  /*  Bytes of machine word are reversed. */

	const int IMAGE_FILE_MACHINE_UNKNOWN = 0;
	const int IMAGE_FILE_MACHINE_I386 = 0x14c;   /*  Intel 386. */
	const int IMAGE_FILE_MACHINE_R3000 = 0x162;   /*  MIPS little-endian, 0540 big-endian */
	const int IMAGE_FILE_MACHINE_R4000 = 0x166;   /*  MIPS little-endian */
	const int IMAGE_FILE_MACHINE_ALPHA = 0x184;   /*  Alpha_AXP */
	const int IMAGE_FILE_MACHINE_POWERPC = 0x1F0;   /*  IBM PowerPC Little-Endian */
	const int IMAGE_FILE_MACHINE_TAHOE = 0x7cc;   /*  Intel EM machine */
	const int IMAGE_FILE_MACHINE_AMD64 = 0x8664;

	const int IMAGE_DIRECTORY_ENTRY_IMPORT  = 1;
	const int IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT  = 11;

	const int IMAGE_SCN_CNT_CODE = 0x00000020;
	const int IMAGE_SCN_CNT_INITIALIZED_DATA = 0x00000040;
	const int IMAGE_SCN_CNT_UNINITIALIZED_DATA = 0x00000080;
	const int IMAGE_SCN_MEM_DISCARDABLE = 0x02000000;
	const int IMAGE_SCN_MEM_NOT_CACHED = 0x04000000;
	const int IMAGE_SCN_MEM_NOT_PAGED = 0x08000000;
	const int IMAGE_SCN_MEM_SHARED = 0x10000000;
	const int IMAGE_SCN_MEM_EXECUTE = 0x20000000;
	const int IMAGE_SCN_MEM_READ = 0x40000000;
	const int IMAGE_SCN_MEM_WRITE = 0x80000000;

	const int IMAGE_REL_I386_REL32 = 0x14;
	const int IMAGE_SYM_CLASS_EXTERNAL = 2;

	const int COR20MetadataSignature = 1112167234;

	const int IMAGE_SIZEOF_FILE_HEADER =            20;
	const int IMAGE_SIZEOF_ROM_OPTIONAL_HEADER =      56;
	const int IMAGE_SIZEOF_STD_OPTIONAL_HEADER =     28;
	const int IMAGE_SIZEOF_NT_OPTIONAL_HEADER =     224;
	const int IMAGE_SIZEOF_NT_OPTIONAL64_HEADER =   244;
	const int IMAGE_SIZEOF_SHORT_NAME              = 8;
	const int IMAGE_SIZEOF_SECTION_HEADER           = 40;
	const int IMAGE_SIZEOF_SYMBOL                  = 18;
	const int IMAGE_SIZEOF_RELOCATION              = 10;
	const int IMAGE_SIZEOF_BASE_RELOCATION         = 8;
	const int IMAGE_SIZEOF_LINENUMBER              = 6;
	const int IMAGE_SIZEOF_ARCHIVE_MEMBER_HDR      = 60;


}