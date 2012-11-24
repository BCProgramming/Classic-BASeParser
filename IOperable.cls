VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOperable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IOperable Interface.
'BASeParser XP- Feb 10th, 2007

'Purpose:
'The IOperable interface is an Interface that should be supported by any and all objects that wish to be used in BASeParser expressions.
'as a example, the CComplex and Cmatrix Classes (do/will) implement this interface, allowing the Core Plugin to delegate the actual
'work of dealing with these objects to something that knows the most about them- themselves. Generally, this interface
'is not actually part of BASeParser XP, but is a component of the Core plugins, which are included in the same library simply because
'the library is useless without them. Oh, and because you cannot have circular references between projects.
Public Enum PerformFunctionReturnConstants
    Perform_Success = 0 'successful execution.
    Perform_Unknown = 1 'can't be used against the given function.
    

End Enum

    

Public Sub PerformOp(withparser As CParser, ByVal Operation As String, OtherOperand As Variant, ByVal DoUnary As IEvalUnaryOpFlags, ByRef retval As Variant)
    'performOp: interface method that should perform the given operation on the given operand.
    'Withparser:
    'Standard parameter, included with most interface methods from this library.
    'contains a reference to the current executing Parser Context. If the implementor requires
    'the use of operations, it should use this objects "PerformOperation" method to do so. Keep in mind
    'that this procedure may be called again at that point if the provided operands to the
    'PerformOperation method contain a object that is of the implementors type. Also, it is not a good idea
    'to try to cheat with this method and call PerformOperation() with the very same arguments passed- This will cause stack-overflow as
    'the two procedures will mutually call each other. Besides, what the hell would make you think that would actually work anyway?
    'Huh? What is wrong with you! BASeparser does all this work to parse and build an expression stack, and kindly asks you
    'if you can perform this operation, and you turn around and ask it to! Well, I have news for you, You can do your OWN god damn math!
    'That's right. You put yourself in BASeParser's Shoes, and do all this work, and ask a hoity-toity IOperable implementor for
    'a measly operation, and then the bastard refuses. I think maybe BASeParser should show a MessageBox
    'Telling the user how retarded you are to even try.
    'The implementor of this method should try it's darndest to interpret the given operands. If it either doesn't support the operation or one of the
    'operands are unrecognized, however, it is safe, and recommended, to raise an error. Keep in mind that any descriptive text that
    'is given will be lost, and it is entirely possible, if the client using the BASeParser library doesn't
    'implement error handling, to Crash the client app. (but thats their fault...)
    
End Sub
Public Function ToString(withparser As CParser) As String
    'converts the objects internal data, whatever it may be, to a string. Used by "ResultToString"
    'in CParser. IOperable has this instead of a separate "IFormattable class because if an expression can operate
    'on it, then the result of a evaluation very well could be it.
End Function
Public Function PerformFunc(withparser As CParser, ByVal FuncName As String, AllArguments As Variant, ByRef retval As Variant) As PerformFunctionReturnConstants
Attribute PerformFunc.VB_Description = "Called when a Function is called with an object implementing IOperable as the first (or only) argument."
    'performs a function. Works only for those functions that have a single argument.
    'not yet implemented.- Thus the private.
    'PerformFunc <NOW IMPLEMENTED>
    'purpose:
    'I realized recently that although my CComplex class contains code to perform the Sin() function, the actual Sin() function
    'won't end up calling the method. Rather the manually add this code to all the functions that can call
    'IOperable, however, the code to do so will be contained (rather appropriately) in CParser's
    '"PerformFunction" procedure.
    'that routine will determine the number of passed arguments. if a single argument was passed,
    'it will determine if that single argument was a object, and if so, wether it supports this interface.
    'if so, it will cast the object to this interface and call this interface routine.
    'which will, appropriately, raise an error if it doesn't understand how to perform the function. At this point,
    'the PerformFunction procedure will continue with default processing.

    'return value determines how to proceed.
    
    'This method will be called if the first argument is determined to implement IOperable.
    'If the method gets confused, it should ignore the request.




End Function

