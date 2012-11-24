VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDynamicDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'IDynamicDispatch interface.
'This interface is optionally implementable by Objects that can be returned
'to the BASeparser engine (for example, the CComplex class). Wht does it do? Well,
'it allows BASeParser to implement a fall-back scheme for when your object doesn't contain a particular method.
'in essence, it is my lame-o implementation of IDispatchEx. To be honest, it would be best to
'create a type-library definition for IDispatchEx and implement it in your object, I'd imagine that TLI
'would handle the details...


'Also, I apologize if I got your hopes up for a easy IDispatchEx implementation. Sorry. I tried for
'hours fiddling with a IDL source, and even though I could get the interface to be implemented
'in a VB Class, I couldn't get the RemoteInvokeObject() to implement. Besides, I had no clue
'what to do in the other ones anyway :)
'if anybody knows how to create an implementation of IDispatchEx that isn't <TOO> messy, I'd
'be most pleased. I'm sure it'll be kludgy, similiar to the Implementing IEnumVariant.

'For now, the IdynamicDispatch Interface is essentially a kludge. It won't exactly be used by
'development environments, but IOperable objects aren't exactly targeting use in a development environment anyway.



'Anyway, if BASeParser fails to call the method, it will check to see if the object in question can
'be cast to a IDynamicDispatch Interface. If so, it will call the InvokeMember() Function on this interface.
'(I didn't even try to call it Invoke(), since VB automatically implements that as part of
'it's behind the scenes IDispatch stuff).

'passes WithParser so the routine can use "PerformOperation".
'Another advantage over the default IDispatch/IDispatchEx via TLI.

'You gotta love this- the Parser tries all sorts of convolutions before giving up.
'beautiful. :).

'"hmm, that doesn't work. What about this... no? well, maybe... damn- what about... there we go!"

Public Sub InvokeMember(withparser As CParser, ByVal Strmember As String, Params As Variant, ByVal PreferProperty As Boolean, ByRef returnValue As Variant)
    'Invoke a method on this object.
    'Note: this will only be invoked when the
    'given objects doesn't have the given member pre-defined, as well as
    'if an error occured in the call to that member (IE, not enough parameters).
    'so this can also serve as a way to kludge out method overloading, since different parameter
    'lists then the standard method will invoke this member, given the object that implements it a chance to perform it.
    'if the method is unsupported, the appropriate error should be raised within this implementation.
    'I suggest 438: Object Doesn't support this property or method.
    'also, since BASeParser will consider any error at  the point it calls this as a error
    'that should be propagated back to the client, errors in this routine should
    'only be those that prevent normal operation. For example a routine that multiplies two numbers should not
    'automatically fail if one or both numbers are strings- It should, AT LEAST, attempt to
    'convert those strings to numbers. Also, too many arguments should NEVER be an error- just
    'use the arguments you need, ignoring the extras.



End Sub
'Public Function GetMembers() As String()
'
'
'
'End Function

