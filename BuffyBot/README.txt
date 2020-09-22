============================
BUFFYBOT OPEN-SOURCE	   |
Readme File		   |
Revision 1.1		   |
============================


Hi, thanks for downloading the source-code for BuffyBot

A lil' history about this bot.
Buffy was originally developed as a script for eggdrop, but i soon got bored again.
The result of too much spare time was this automation 'bot for EFNET irc networks.

A NOTE ON COMPILING:
You must compile this project to run. The executable must also be located in the main
directory containing the /scripts and /data directory as well as the config.ini file.
If these are not present, some wacky-stuff could happen...



She currently has !seen and !quote functions, however all of these features happen via
the scripting engine. The scripting engine works like this:

	-> The engine reads main.c
		-> From main.c, other script modules are loaded and 'included'

Each script module has a procedure which is called at various times.

For example, a module may have a PrivMSG() procedure, which will be called when a message
is received by the client from either a user or a channel.
Similar procedures are also available such as OnJoin() and OnPart(). Others probably exist,
but i can't remember them at this stage ;)

Channels.cfg is a configuration routine for setting up channel parameters.
Config.ini is the bots configuration file. Holds info such as nickname, servers,
ident-daemon, fullname and partyline options (not complete)

Also included in this release is the Export Utility. This utility will scan the main
project file *.vbp and generate a .HTML report on the procedures and functions that have
been exported to the scripting engine via available class modules. Pretty useful... Took me
about 15 minutes to code, so check it out...

Some suggestions and stuff to do:
Maybe create a module to link the bot to the ALICE artificial intelligence systems.
Maybe create a sexbot. u that bored? ;)
Fix the dynamic ban system for me?
Anything decent, but if you make a really good script, ya gotta send it to me too :)

--------------------------------------------------------------------------------------
Please note that this source-code is released for your educational leisure, and i take
no responsibility for what you do with it.
--------------------------------------------------------------------------------------

-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
				|
Well, thats all for now...	|
HAVE FUN!			|
				|
	Cheers, AndyR007	|
	BuffyBot Developer	|
	ALMI Studios		|
				|
-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=