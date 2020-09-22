Sub OnJoin(Nick, Ident, Host, Where)
	' Ok see if they have a channel flag and act accordingly
	if lcase(Nick) <> lcase(Botnick) then

		' Is the channel on autovoice?
		if botcore.getchanflag(cstr(Where), "AutoVoice") = "Yes" then
			ircfunctions.setmode cstr(where), cstr(nick), "+v"
		end if
		if botcore.getchanflag(cstr(Where), "AutoOp") = "Yes" then
			ircfunctions.setmode cstr(where), cstr(nick), "+o"
		end if
		' Are they registered for a voice?
		if botcore.getchanflag(cstr(Where), "RegVoice") = "Yes" then
		if botcore.VerifyFlag(cstr(Nick), cstr(Where), "v") = 1 then
			ircfunctions.setmode cstr(where), cstr(nick), "+v"
		end if
		end if

		' Are they registered for a operator flag??
		if botcore.getchanflag(cstr(Where), "RegOp") = "Yes" then
		if botcore.VerifyFlag(cstr(Nick), cstr(Where), "o") = 1 then
			ircfunctions.setmode cstr(where), cstr(nick), "+o"
		end if
		end if
		
		if botcore.getchanflag(cstr(Where), "ShowInfo") = "Yes" then
			' Show their info line
			if botcore.Getuservalue(cstr(Nick), "Info") <> "" then
				'' Ok they have an info line
				ircfunctions.PostMsg cstr(Where), "[" & cstr(Nick) & "] """ & botcore.getuservalue(cstr(Nick), "Info") & """"
			end if
		end if
	end if
End Sub