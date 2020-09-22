Sub Main()
	Scripting.Include "Scripts\BuffyBots.c"
	Scripting.Include "Scripts\Startup.c"
	Scripting.Include "Scripts\ChanJoin.c"
	Scripting.Include "Scripts\Bans.c"
	Scripting.Include "Scripts\Quote.c"
End Sub

Sub Test(ProcVar)
msgbox "testing"
End Sub

Sub PrivMSG(WhoSaid, What, Where, Ident, Host, InChan)
	if left(What, 1) = "!" then
		' Channel command ok, get out parameters and continue...
		Dim sCommand
		Dim sParam
		if instr(1, cstr(What), " ", vbTextCompare) <> 0 then
		sCommand = botcore.leftstring(cstr(What), Instr(1, cstr(What), " ", vbTextCompare)-1)
		sParam = botcore.RightString(cstr(What), Len(What)-Len(sCommand)-1)
		else
		sCommand = What
		sParam = ""
		end if
		Select Case lcase(sCommand)
		Case "!codestats"
		ircfunctions.postmsg cstr(where), "BuffyBot Runtime Statistics:"
		ircfunctions.postmsg cstr(where), "Current Procedures: " & scripting.GetProcedureCount
		ircfunctions.postmsg cstr(where), "Current Modules: " & scripting.GetModuleCount
		ircfunctions.postmsg cstr(where), "End of Runtime Statistics."
		end select
	end if

End Sub