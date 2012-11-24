VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CList
'represents a list.
'NOTE: this is designed to replace my current implementation of lists, which simply
'packs a variant array into the variable. Rather, I should create a Ioperable implementor that
'implements the appropriate methods and interfaces. the beauty of this is that I shouldn't need to change very much code at all
'that relates to my previous implementation, other then changing the assignment from the creation of an array to
' the creation of an instance of this class.

'Also, whenever a "pure" array is identified, it should be converted to an instance of this class.

