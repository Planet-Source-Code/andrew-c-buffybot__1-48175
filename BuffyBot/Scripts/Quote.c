' Quote script
' This script is designed to only operate on the #flirts chatroom

Sub PrivMSG(WhoSaid, What, Where, Ident, Host, InChan)
if lcase(where) = "#flirts" and inchan = "True" then
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
		Case "!quote"
		dim quotenick
		dim quotestring
		if left(sparam, 1) = "<" then
			' This means they've prolly put "<SexyBabe> Hi" or similar
			quotenick = left(sparam, instr(1, sparam, ">", vbTextCompare)-1)
			quotenick = right(quotenick, len(quotenick)-1)
			quotestring = right(sparam, len(sparam)-len(quotenick)-3)
		else
		quotenick = left(sparam, instr(1, sparam, " ", vbTextCompare)-1)
		quotestring = right(sparam, len(sparam)-len(quotenick)-1)

		end if
		DataLibrary.WriteDataFile cstr(quotenick) & ".dat", "Quotes", "[" & cstr(quotenick) & "] " & cstr(quotestring)
		
		Case "!randquote"
		dim rndquote
		rndquote = datalibrary.GetRankedMsg("Quotes", sparam & ".dat")
		ircfunctions.postmsg cstr(where), cstr(rndquote)

		Case "!listquotes"
		if datalibrary.fileexists(cstr(sparam) & ".dat") = 0 then
			ircfunctions.postmsg cstr(where), "No quotes found for " & cstr(sparam)
			exit sub
		end if
		dim quotecount	
		quotecount = datalibrary.getlinecount(cstr(sparam) & ".dat")
		for x = 0 to quotecount
		quotestring = getline(cstr(sparam) & ".dat", x)
		ircfunctions.postmsg cstr(where), cstr(quotestring)
		next
		end select
	end if
end if




End Sub