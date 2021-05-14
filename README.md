# Add-in Template

This project gets you up and running with the [xll library](https://github.com/xlladdins/xll.git)
for creating high performance Excel add-ins written in C++, or C, or any language
that can be called from C. This includes Fortran, JavaScript, Python, Node and Deno, Lua, Zig, any .Net langurage, WASM,
but maybe not COBOL [1].

Clone this project in Visual Studio 2019,
open the `xllproject1.sln` solution, 
and press `F5` to build and start Excel in the debugger. You can set breakpoints in the debugger
but the great thing about the xll library is that it lets you use Excel as a debugger on steriods.
Hook up your code and call it from Excel to see the data instead of looking the algorithm.
Copy, paste, and create graphs to get a better understanding of the code you write or incorporate
from other libraries.

If you use 32-bit Excel you have 
let Visual Studio know by selecting `Build` from the top level menu and choosing `Configuration Manager...`.
Set the `Platform` to `Win32`.

If you want to create your own project, fork this repository and rename it before 
cloning to your computer.

[1] Don't do this except to call a pure function. That is what Excel was built for. 
