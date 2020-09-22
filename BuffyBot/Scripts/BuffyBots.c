'// #BuffyBots channel script file

' The following subfunction handles a privatemessage either to a channel or in PM.
' If in a channel the InChan variable will be 'True', if 'False' the Where var will
' be 'privatemessage'.

Sub PrivMSG(WhoSaid, What, Where, Ident, Host, InChan)
'if InChan = "True" and lcase(Where) = "#buffybots" then
if lcase(where) = "#buffybots" then 'Just a dummy if statement
	' Ok something was said inside the channel
	' -
	' First check for any special commands:
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
		Case "!seen"
			' The !seen command
			if sParam = "" or sParam = " " then
				' Just a single seen with no parameters, so notice the user
				userident = botcore.leftstring(cstr(whoSaid), instr(1, cstr(WhoSaid), "!", vbTextCompare)-1)
				IRCFunctions.UserNotice cstr(userident), "Incorrect !seen command. Usage: !Seen [Username]"
				exit sub
			end if
			if inChan = True then
				seenchan = Where
			else
				seenchan = "*"
			end if
			'if instr(1, cstr(sParam), " ", vbTextCompare) <> 0 then
		'		sParam = botcore.leftstring(cstr(sParam), instr(1, cstr(sParam), " ", vbTextCompare)-1)
		'	end if
			userident = botcore.leftstring(cstr(whoSaid), instr(1, cstr(WhoSaid), "!", vbTextCompare)-1)
			searchresults = getuservalue(cstr(sParam), "Lastonline")
			whosaid = userident
			if lcase(whosaid) = lcase(cstr(sparam)) then
				if inchan = "True" then
				ircfunctions.postmsg cstr(where), Whosaid & ", Why the hell are you looking for yourself? That's pretty stupid!"
				exit sub
				else
				ircfunctions.usernotice cstr(whosaid), Whosaid & ", Why the hell are you looking for yourself? That's pretty stupid!"
				exit sub
				end if
			end if
			if datalibrary.scandata(cstr(sparam), "BadWords.dat") <> 0 and botcore.CheckUser(cstr(sparam)) = False then
				if inchan = "True" then
				ircfunctions.postmsg cstr(where), Whosaid & ", " & datalibrary.getrandomsentence("putdowns.dat")
				exit sub
				else
				ircfunctions.usernotice cstr(whosaid), Whosaid & ", " & datalibrary.getrandomsentence("putdowns.dat")
				exit sub
				end if
			end if
			if searchresults = "-1" then
				IRCFunctions.UserNotice cstr(WhoSaid), WhoSaid & ", I have not seen " & sparam & " in the channel."
				exit sub
			else
				if datediff("s", searchresults, GetNowTime) < 60 then
				IRCFunctions.UserNotice cstr(WhoSaid), WhoSaid & ", " & sparam & " was just here! You just missed them. They left " & MathFunctions.ConvertTime(cstr(searchresults)) & " ago."
				else				
				IRCFunctions.UserNotice cstr(WhoSaid), WhoSaid & ", I last saw " & sparam & " " & MathFunctions.ConvertTime(cstr(searchresults)) & " ago."
				end if
			end if
		end select
	end if
end if	
End Sub

Sub OnPart(WhoSaid, Nick, Ident, Host, Where)
userident = botcore.leftstring(cstr(whoSaid), instr(1, cstr(WhoSaid), "!", vbTextCompare)-1)
writeuservalue cstr(userident), "Lastonline", cstr(botcore.getnowtime)
End Sub

BotCore.WriteConsole "Channel script file loaded!"