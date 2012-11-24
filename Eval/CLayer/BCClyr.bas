option explicit
'No it isn't just you- this code looks like some kind
'of horrible creature created with a mix of C++ and Basic.
'just perfect for me :)

'This is a small program I package with my Expression Evaluator,
'and gives it a cheap front end.
'the tricky bit was IDispatch.
#define UNICODE
#include "disphelper/disphelper.bi"
Common Shared ParserPtr as IDispatch Ptr
Common Shared DoBreak as Integer
Common Shared NoShowBanner as Integer 
Common Shared ConsoleMode as Integer 

Function ParseExpression(Byval StrExpression as String) as String
    
    
	dim as zstring ptr szResponse
    On Error goto ObjectError
    
    dhPutValue(ParserPtr,".Expression = %s ",StrExpression)
	
    dhCallMethod(ParserPtr,".Execute()")
	dhGetValue( "%s", @szResponse, ParserPtr, ".ResultAsString" )
    
	'print "Result: "; *szResponse
    ParseExpression = *szResponse
    Exit Function
    ObjectError:
    Print "Error:" + Str$(ERR)

end function

sub showhelp()
    print "BASeCamp BASeParser Command-Line evaluator front-end."
    print "Copyright 2005-2006 BASeCamp corporation, all rights reserved."
    print 
    print "Possible arguments:"
    print 
    print "/H,/?,-H,-?","This Help screen.
    print 
    print "/NB","Don't display anything but the result.(IE, no ver info)"
    print 
    print "/CL","Break into Command-line mode. BCEval will show a prompt and "
    print    ,   "You can freely define and mess around with variables."
    print "This Program is a Win32 Console Application written in FreeBasic."
    print "it requires that BASeParser.dll be properly registered."
    
end sub
Sub ParseArguments(Byval StrParse as String)
    Dim Argc as Integer
    argc = 1
    Do
            select case ucase$(Command$(argc))
            case "/H","-H","-?","/?"
                showhelp
                System
            Case "/NB","/nb"
                    noshowBanner=true 
                Case "/CL","-cl"
                consolemode = true 
            case ""
                Exit Do
            End Select 
                
            
       argc+=1 
    loop
    
End Sub 
Sub ShowConsoleHelp()
    Print "Commands:"
    print "Enter Expressions to have them evaluated"
    print "Type q to quit"
    print "type help for help."
End Sub
Sub PerformConsoleMode(ParseObj As IDispatch Ptr)
    Dim CurrCommand as String
    Dim QuitEntry as String 
    Dim strshow as String 
    Print "BCEval Expression Evaluator"
    Print "CommandLine mode Active."
    Print "Type Q to quit."
    
    Do
        Line Input "->";CurrCommand
        CurrCommand=Ucase$(CurrCommand)
        if Currcommand="HELP" Then
            
                ShowConsoleHelp
            
        end if 
        if currcommand = "Q" then
            quitentry=""
            Do Until Instr("YN",left$(QuitEntry,1))
                LINE INPUT "Are you sure you want to quit[Y/N]?";QuitEntry
            
            
            Loop
            if quitentry = "Y" then
                print "BCEvaluator Command-Line terminated."
                Exit Do
            end if 
        end if 
            'otherwise- the tricky stuff. parse it!
            strshow = ParseExpression(CurrCommand)
            print CurrCommand
            print String$(len(currcommand),"-")
            print strshow
            print
        
    Loop



    
end sub
Function ExprParse CDECL Alias "ExprParse" (StrExpression as String) as String EXPORT
	DISPATCH_OBJ(objParser)
	dim as zstring ptr szResponse
  
    Call ParseArguments(Command$)
    
    dhInitialize( TRUE )
    dhToggleExceptions( TRUE )

	dhCreateObject( "BASeParserXP.CParser", NULL, @objParser )
    ParserPtr = objParser
	dhCallMethod( objParser, ".Create()")
    if (ConsoleMode) then
            PerformConsoleMode(objParser)
        System
    end if 
    
    dhPutValue(objParser,".Expression = %s ",StrExpression)
	
    dhCallMethod(objParser,".Execute()")
	dhGetValue( "%s", @szResponse, objParser, ".ResultAsString" )
    SAFE_RELEASE( objParser )
	return(*szResponse)
	

	
	
End Function