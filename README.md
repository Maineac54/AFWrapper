# AFWrapper
Wrapper Classes for OSISoft's AFSDK

You should be able to copy the three files, AFWrapper.dll, AFWrapper.pdb, and AFWrapper.tlb to a folder on your 
system and register the type library (*.tlb) with regtlib12.exe located in C:\windows\Microsoft.NET\Framework\v4.0.30319.

The application was compiled with .Net Framework 4.5 on VS2013

I found it helpful to fully qualify the object names such as "dim wafVal as AFWrapper.WAFValue".

The wrapper is built with the ClassInterfaceType AutoDual which means that you should be able to early bind with VBA and
see the methods and their required parameters.  

Some Caveats:
I focused mostly on being able to edit PI Archive values from Excel.  So there are a lot of AF Classes and Functions that
are not in this wrapper.  Please feel free to add additional capability to the wrapper.  Please branch the project to do this.

