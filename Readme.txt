

16/Jan/2007 17:11

ReadMe file for BASeParser Expression Parsing Library.
(c) 2006-2007 BASeCamp Corporation, All rights reserved.


	BASeParser is a full-featured expression evaluator designed for Microsoft Visual Basic 6. I have purposely created
the library with the sole goal of mixing together flexibility and speed.



The following files should have been included in the distribution of BASeParser, which may be an installer or archive:

FileName				Purpose


BASeParser.dll				The BASeParser Library.
BPXPGUI.ocx					BASeParser XP GUI components (configuration dialog,as well as settings pages for intrinsic plugins, for example)
BPCoreFunc.dll				a few useful plugins I whipped up, such as calling functions within Script files, and environment variables.
Readme.txt				This file.
devlog.txt				My personal Development log, for the curious.
vbalColumnTreeview6.ocx VBAccelerator Multi-Column treeview control 
vbaldTab6.ocx           VBAccelerator Tab control




Getting Started:
	If your using Visual Basic, this is easy. All you need to do is open the References dialog for the project you wish to add parser support for. Then look for an item in the list named "BASeParser XP Expression Parsing Library" or something similiar. (It WILL have BASeParser XP in the name). make sure it is selected. If the item is not there, use the Browse... button to locate the aforementioned "BASeParser.dll" library. The item will be registered and selected. The MSI installation, By default, Places this DLL in your %Program Files%\BASeCamp\BASeParser directory.

There are two ways you could use this library- for quick & dirty parsing, or you can instantiate a CParser object and set the options and the expression. The Quick & dirty method is to use the global Function EvaluateExpression (or the EvaluateExpressionByRef procedure). This allows you to simply pass in a string and retrieve the value. Of course, there are trade-offs to using this method for expression evaluation. First off, if you are using two different expressions in a loop and execute them over and over again out of sequence (Expr1,Expr2,Expr1,Expr2) etc.), the CParser object intrinsically created will have it's stack torn down and then rebuilt over and over. In addition, you gain no benefit from the optimizations the Parser attempts to make. It is possible to store variables within this implicit parser object, however, this requires the use of an expression utilizing the "STORE()" function.

Expressions

It is the authors humble opinion that this expression parser is one of the best freeware, if not the best componentized expression parsers, available for use in VB6. The parser supports a multitude of functions and operators, as well as intrinsic support for Complex numbers,Matrices, and arrays. The functions include pretty much every Visual Basic function, as well as a few C-run-time functions. In addition, due to it's support for Complex numbers and Arrays (informally known as "lists" due to the {} syntax that can be used), it has several "utility" functions, such as the Complex() function that creates a complex number(pretty useless, since you can simply use i anywhere, assuming evaluation won't require a invalid operation on the complex number), as well as an Array() function, which transforms it's arguments into an array. It also supports the creation of objects as well as using a "object access" operator, @, with which you can access members of that object. This extends into the use of Complex numbers, because I use a CComplex object to represent them, as well as into Arrays, where I specifically added support for a few "methods" on the array.


ERRORS:

	When BASeParser encounters an error either parsing or executing the expression, it will raise the Error() event, and then break out of its current operation. However, this is one zone where you can't expect the error description to match exactly what happened. The most common one I get irregardless of the actual error is "Subscript out of range". The parser tries it's best to resolve what REALLY happened, so as not to always leave you with a generic VB error and an empty feeling.


BASeParser touts the following features:


-Access to a vast number of pre-defined number, intrinsic complex number support for most operators, and intrinsic support for lists and arrays. As well, it supports, in addition to numbers, Strings, objects, and almost anything you can throw into a Variant.

-Extreme extensiblity
	Not only is it relatively easy to create your own functions, but it is just as easy to create your own operators, including unary prefix and unary suffix operators.

-Optimizations

	BASeParser parses an expression into a ParseStack, which can be subsequently accessed with a minimum, if any, of slower string manipulation. The result? Evaluating a expression after parsing is much quicker, regardless of wether any of the variables therein have changed. In fact, this makes it suitable for situations where a single variable is changed throughout a number of iterations. A prime example is a function graphing application.

Unlike my previous version of BASeParser, which hopefully nobody has downloaded <g>, The parser attempts to do even further optimization, such as when it contains a constant expression. The old version of BASeParser possessed this architecture, but it also construed a fatal flaw- every time a parenthesized expression was encountered, the contents would be placed immediately into a new Parser object, evaluated, and the parser object deleted, effectively discarding the optimizations that were performed on that sub-expression, only to do it over again when the expression was executed later. This version goes so far as to optimize every little cranny. For example, the expression:

(Sqr(4-3*A/(Sqr(2)-Sqr(3^3))))

won't optimize the entire expression. due to A, a variable, being present. However, it notices that the "Sqr(2)-Sqr(2^3)"
parenthesized expression contains no variables or dynamic functions/operators. As such, it goes right ahead and replaced the IT_SUBEXPRESSION within the parsestack with a IT_VALUE that equal that value- or, the value -1.4142135623731. It is obvious why this is beneficial to multiple executions, due to the Sqr Function requiring so much more then a single value assignment. The cool thing, from my view, is that I didn't need to add extra code- the fact that a CParser object can optimize itself, paired with the use of "child" Cparser objects to handle parenthesized expressions and function arguments, is alone enough to enable it. In fact, I was quite pleased to find this unintended bonus in the Debug output- the CParser was optimizing a expression even though another part had a variable.- On the topic of debug output, if you downloaded the debug version, don't expect it to go so fast...



In addition, the Variable Access is extremely easy- Simply retrieve a CVariable Object from the variables collection. Don't worry! It doesn't have to exist! If it doesn't exist, the collection creates a new variable with that name and returns it. Also, I have had concerns mentioned as to the possiblity that a parenthesized STORE() function call will not work with the base-level parser, since it will be modifying the variables of the IT_SUBEXPRESSION parser. I have it covered. The sub-parsers are created with the CParsers Clone() method (better named, GiveBirth), which sets the parsers reference variables to it's own, so they reference the same variables and functions. Here are a few tips that I learned during the creation of the BASeParser Core and auxilary projects, such as GraphCtl:


	-If you have an expression, say "Sin(X)", make sure that a variable named "X" is already added BEFORE you assign the "Expression" property (or, more precisely, before it is parsed). If you don't, you may get odd results- usually zero. The problem? Well, it is way to technical to talk about here. There should be a devlog.txt file included, refer to that to get inside my head (my first log was only 9 hours after selecting "New Project" from the VB File menu!






EASY FUNCTION ADDITION:

On top of all of this, I have added to ability to add simpler derived functions without the need for interfaces. This is constituted in the CParser objects "Functions" Collection, which holds a collection of CFunction objects. Also, I have implemented a "DataFunc" plugin that enables the use of a Access database to define even more functions.

MORE TO COME!

Just when I think I have finally reached a plateau of awesomeness with BASeParser, one of two things happen:

A: I find a bug. One of those really shameful bugs that would make you wish you never started the app in the first place.
B: I think of another feature that BASeParser just cannot do without.



SPECIAL THANKS

Although this particular project doesn't use any of his controls, I'd like to applaud Steve McMahon of VBaccelerator.com for being cool. Several times I almost added the VBAccelerator controls to the project. The only reason I didn't was because I wanted to make BASeParser as light as possible. I think I might use SGrid2 though-

Eduardo Morcillo. (EdanmosVBPage.com) Sure, he's moved into VB.NET and won't look back, but the VB Classic stuff he did do is quite stellar. All sorts of COM stuff that should probably only be done from C++, such as Subclassing the VB Control Extender, and not to mention the OLE type libraries, which I believe I used (hey, TLB's are compiled in, why not?)


A REQUEST
------------------------

	Although I believe I am a pretty good VB programmer, I often hit roadblocks when I try to do the advanced stuff. Sure, the VBAccelerator components and COM libraries tear down about a billion barriers, in fact, where VB6 could do 95 percent of what you wanted, VB6 with VBaccelerator and the COM typelibs can do 99.999999999% of what you want.

There is ONE thing that I think would REALLY make BASeParser cool. The ability to declare DLL functions.  It is obviously possible, since VB does it at run-time when  you use Declare, and MSVBVM60.dll  exports a "DLLFunctionCall", whose arguments, If I knew them, would most definitely solve my problem. This is where I must put out an APB. I REALLY want  this feature. In fact, I want it so bad, I have even embedded a transcript of my attempts to add  this ability in devlog.txt, check it out! If you have any information on how to do this, I am sure I would be interested. Do, however, read the devlog to see if I didn't try it already.
