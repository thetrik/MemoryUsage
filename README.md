# Get information about memory usage

This example is intended for retrieving the information about memory usage. This project is compatible with x64 environment.

The information is based on the **GetProcessMemoryInfo** function for 32 bit processes. This function has the disadvantage such you can't get correct information for 64 bit processes. The 64-bit approach is based on the internal structure of the **GetProcessMemoryInfo** function which just calls the **NtQueryInformationProcess** one with the **ProcessVmCounters** type information.

Because of this function goes thru WoW64 layer we also can't extract the correct information. To avoid this i used switching to x64 (Long-Mode) and calling the appropriate function in the 64-bit ntdll. Therefore we have no WoW64 layer and can extract the correct information.

I've added the useful module to working with 64-bit ntdll (and others dlls). You can call the most of functions in 64-bit dlls using **CallX64** function. Notice, all the pointers in x64 are 64-bit one so you need to pass them as a Currency variable if you want to work with other 64-bit applications. If you work with 32-bit applications you can pass usual pointers (like VarPtr casted to Long) because they are expanded to 64-bit ones with the zeroed high part.

**GetModuleHandle64** allows to get a handle of a 64-bit library in the current process by its name.

**GetProcAddress64** allows to get the address of function from a 64-bit dll in the current process (it has not-full functional because it doesn't support redirects but i don't know imagine a situation when it can be applicable in the current mode). I didin't checked but i think you can load another 64-bit dll to the current process but it should use only the **Native APIs**.

Thanks for your attention!

The trick,
2014 - 2020.



