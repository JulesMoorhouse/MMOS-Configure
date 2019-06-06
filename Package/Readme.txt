You need to remove Ex from DLLSelfRegisterEx

e.g.

From ;_

File11=@MSCAL.OCX,$(WinSysPath),$(DLLSelfRegisterEx),$(Shared),4/5/19 9:57:44 AM,125528,11.0.5510.0

To :-

File11=@MSCAL.OCX,$(WinSysPath),$(DLLSelfRegister),$(Shared),4/5/19 9:57:44 AM,125528,11.0.5510.0

