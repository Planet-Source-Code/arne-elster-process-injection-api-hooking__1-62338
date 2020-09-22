1) Download Compile Controller from http://johnchamberlain.com/ccupdates.html and install it.
2) compile "testapp"
3) open "inject"
4) start the compile controller addin.
5) file->hook compilation
6) compile the project
   when "link.exe" is called, look for /BASE and replace 0x400000 with 0x13140000.
7) Now run both testapp and inject, and hook testapp`s MessageBoxA. :)