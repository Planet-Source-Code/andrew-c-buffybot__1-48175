' This module is still in development
' Dynamic Ban Script

Sub OnJoin(Nick, Ident, Host, Where)
Dim UserIdent
UserIdent = Ident & "@" & Host
if botcore.leftstring(cstr(userident), 1) = "~" then
	userident = botcore.rightstring(cstr(userident), len(userident)-1)
	if botcore.checkban(cstr(where), cstr(userident)) = 1 then
		ircfunctions.sendirc "KICK " & where & " " & nick
		ircfunctions.setmode cstr(where), cstr(nick), "+b"
	end if
end if
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
		Case "!ban"
			if instr(1, sParam, " ", vbTextCompare) = 0 then
				' Just a ban, with no reason code given
				botcore.addban cstr(where), cstr(ident) & "@" & host
				userident = botcore.leftstring(cstr(whoSaid), instr(1, cstr(WhoSaid), "!", vbTextCompare)-1)
				ircfunctions.sendirc "KICK " & where & " :" & userident
				ircfunctions.sendirc "MODE " & where & " :" & userident & " +b"
			else
				' extract the reason
			end if			
		end select
	end if
End Sub