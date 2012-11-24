BASeParser XP Expression parsing/evaluation library

WORK SAMPLE

BASeParser XP is one of my flagship products. What it is is an extendable library that can parse expressions based on
a standard operator/function type of syntax. Function and Operator definitions can be contained in external "Plugins" installed by adding their progID's to the appropriate location in the registry. By default, each CParser Object uses a internal plugin implementor that handles the core operators and functions. These are a big varied bunch, from simple +,-,*, etc, to Pascal's "In" operator, and Perl's "spaceship" operator, <=> (for sorting comparisions).

This version is a work sample I literally slapped together. I cannot even be sure it works, since I haven't yet inspected the dependency information. I included the less-common libraries, such as those from VBAccelerator, that I used. Others, such as Microsoft's VBScript Regular Expression library and the ScriptControl, are more common and quite likely to be present on the target machine.

The sample program was written in FreeBasic, which I call B--. This is a freeware compiler for QuickBasic compatible code in Win32 that can use all the newer features of todays systems.
This code has been thoroughly tested on windows XP, with blazen runs on my Windows 98SE laptop.

*****NOTICE*******

This version of BASeParser is a initial "ALPHA" version. It is the first time the code has in any way left my computer. Currently, it writes a vast amount of Debugging and profiling information into a "Debug" directory under the folder containing the dll. THis "DebugLog" Directory can be safely deleted, or inspected as necessary. For the curious, the Debug output and profiling information is created almost effortlessly (well, a global find and replace for Debug.Print, and a procedure call at the start and end of each profiled procedure). These are implemented using conditional compilation. I keep them intact because this is an alpha version. Since the Debuglog is typically 500KB for a good-sized session, the constant disk-writes slow down the engine somewhat, so the straight-up speed of this library should not be compared to other, "Commercial" implementations, at least not without removing the debug logging behaviour.

*****SOURCE CODE********

I do not exclude the source code because I am snobby. Rather, it is because I have not yet discovered the dependency information for the source. The source code requires several type-libraries not required on machines running the compiled versions, and these type-libraries are not extremely common on most machines I have seen. 

If you contact me, I will of course provide the source code for inspection. I personally think it is quite well written, adhering pretty strictly to a interface-based design (where applicable), and reaping the benefits. Over 40% of the source lines consist of Comments, as well. Although I will admit, there are way to many "way after" comments that are actually useless. I'd glance over the code, and then comment, "'Not sure what this is doing."


****IF IT DOESN"T WORK******

If the code doesn't run, or gives some dhCreateObject error, I have just left a bad impression. I assure you it works on my computer.  I swear. Oh, And don't forget to Regsvr32 the DLL file!