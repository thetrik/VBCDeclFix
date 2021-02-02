# VBCDeclFix

 The Add-in allows you to use Cdecl functions in VB6 declared in type libraries.
 
 If you have ever tried to use CDECL-functions declared in a TLB then you know that debugging (in IDE) is impossible. The project just has crashed and doesn't even start although compilation to native code works without any issues. This Add-in fixes this behavior and you can debug your code in IDE.
 
 ## What's the crash?
 
 When VBA6 compiles the code it calls [ITypeComp::Bind](https://docs.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-itypecomp-bind) for each external function:
 
 ![ITypeComp::Bind](/images/type_bind.png)
 
 This method returns [the FUNCDESC](https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-funcdesc) structure where the callconv member constains the calling convention of the external function. VBA6 then uses this information inside its own structiures which describes the functions in the project. During P-code generation the runtime extracts this information to map the specified function with the corresponding p-codes. There is the table which manages the process of the codegenetraion:
 
![ITypeComp::Bind](/images/bug_table.png)

Each entry in the table contains 30 bytes. This table contains 0x19 value at 0x1D offset in the first entry which causes the jump to procedure which causes the crash. If you try to compile a stdcall function you'll see it uses the second entry in the table and it contains 0x09 value at 0x1D offset. So if you replace 0x19 value to 0x09 the bug dissapeares, but the code won't work yet. Why? It's because VBA6 doesn't have the P-code handlers for CDECL functions. It contains the stubs which generates 0x33 error (internal error):

![ITypeComp::Bind](/images/internal_error.png)

 ## How does the Add-in work?
 
 The Add-in search the table and replaces 0x19 value to 0x09 then it replaces the P-code handlers (which raises 0x33 error) with new ones. It uses the signature search based on machine code sequences to extract all the unexported entities.

All the information obtained during my own superficial reverse engineering experience therefore, everything described above may not work exactly like that, or even not at all.

Thank you for attention,
The trick.
