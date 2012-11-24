VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BPDllPlug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



'This class is the interface between BASeParser itself and the non-ActiveX DLL's that implement
'a plugin. There are limitations to this method:

'I apologize. I simply do NOT understand how to use the VARIANT structure in C++. sure, I can use VariantInit() to clear it's memory,
'but after that I am confused. I think I have to assign the vartype myself, and then populate the respective member of the union. No thanks,
'I'll stick to what I know for now.

'In fact, I'll bet you this class get's removed and you never read this.





'Anyway-  I won't enable this until I actually get my test DLL to compile.
'grrr. VARIANT apparently isn't a identifier.



