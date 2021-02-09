# VBCDeclFix

 The Add-in allows you to use Cdecl functions in VB6.
 
 If you have ever tried to use CDECL-functions declared in a TLB then you know that debugging (in IDE) is impossible. The project just has crashed and doesn't even start although compilation to native code works without any issues. A similar problem occurs when using the **CDecl** keyword - VB6 always generates the code with the **0x31** error (*Bad Dll Calling Convention*) so you can use such functions neither IDE nor compiled executable. This Add-in fixes this behavior and you can debug your code in IDE and compile the code to an executable file. Moreover this Add-in adds the ability to use the CDecl keyword for user functions.
 
 ## What's the crash?
 
 When VBA6 compiles the code it calls [ITypeComp::Bind](https://docs.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-itypecomp-bind) for each external function:
 
 ![ITypeComp::Bind](/images/type_bind.png)
 
 This method returns [the FUNCDESC](https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-funcdesc) structure where the callconv member contains the calling convention of the external function. VBA6 then uses this information inside its own structures which describe the functions in the project. During P-code generation the runtime extracts this information to map the specified function with the corresponding p-codes. There is the table which manages the process of the codegeneration:
 
![BugTable](/images/bug_table.png)

Each entry in the table contains 30 bytes. This table contains 0x19 value at 0x1D offset in the first entry which causes the jump to procedure which causes the crash. If you try to compile a stdcall function you'll see it uses the second entry in the table and it contains 0x09 value at 0x1D offset. So if you replace 0x19 value to 0x09 the bug dissapeares, but the code won't work yet. Why? It's because VBA6 doesn't have the P-code handlers for CDECL functions. It contains the stubs which generates 0x33 error (internal error):

![InternalError](/images/internal_error.png)

 ## Why CDecl keyword doesn't work for "Declare" functions?
 
 I guess it's because there is no P-code handlers in the runtime. There is the point where the runtime check the calling convention of a declared (declared with **Declare** keyword) function and it unconditionally generates the P-code which raises the 0x31 error:
 
 ![DeclareCCCheck](/images/cdecl_declare_code.png)

 If you change this logic you can force the runtime to compile a CDECL function like a STDCALL one. In fact, if you patch the above JMP to the CompileSTDCALL label the runtime will successfully generate the code! Moreover the native code will be compiled successfully along with the **add esp, xx** instruction after each **CDecl** call:
 
  ![DeclareCDeclCompiled](/images/cdecl_declare_compiled.png)
 
 The first problem is with the "Bug table" which will crash the IDE but we already know how to bypass it. The second problem is related to the unimplemented p-code handlers which generate 0x33 error. 
 
 ## How to enable CDecl keyword for user functions?
 
 The native VB6 code-parser rejects the CDecl keyword after function name so to change this behavior the Add-in patches some internal code which parses the source. There are 2 points where the VB6 parser process procedures. The first point is when it process the source text file. It extracts all the tokens (lexems) from the source line step-by-step and validates their type. If there is some wrong sequence / error it rejects the parsing process and mark the line in red. Regarding to the user procedure declaration we can just hijack the process when it has parsed the procedure name to pass the CDecl keyword. The good thing is Declared procedures and user procedures fall to the same function that makes a PCR-node which describes them. Since we can specify the CDecl keyword for a Declared function and it falls to its PCR-node we also can modify the process to set the same info for a user function.
 
 The first issue is the PCR-node is translated to a BSCR-node which already directly describes the procedure in the VB6-project. The problem related to user functions is the parser doesn't check the calling convention field for them and sets it to STDCALL inside the BSCR-node. So we should modify this process to specify the CDECL calling convention inside a BSCR-node if such convention is specified in a PCR-node.
 
 At this point we already have user CDECL functions they are compiled to the proper native code (with ADD ESP, XX / RET instructions) but they isn't visible in the Code Pane. So the third patch should affect the displaying procedure. This isn't a big issue if we remember about Declared functions which are displayed correctly.
 
 The last issue is related to P-code building. Because of it can't compile CDECL-procedures (it always does STDCALL ones) we can't debug the code in IDE. The good news is that a P-code procedure descriptor has the value which shows how many bytes it should pop from the stack when procedure ends. So we only need to fix this value for our CDECL functions when the project are compiled to P-code in IDE.
 
 ## How does the Add-in work?
 
 The Add-in search the table and replaces 0x19 value to 0x09 then it replaces the P-code handlers (which raises 0x33 error) with new ones. It also patches the JMP in order to enable the **CDecl** keyword normal behavior. To enable the CDecl keyword in user functions it patches the VB6 code parser and the P-code builder. It uses the signature search based on machine code sequences to extract all the unexported entities.

All the information obtained during my own superficial reverse engineering experience therefore, everything described above may not work exactly like that, or even not at all.

Thank you for attention,

The trick.
