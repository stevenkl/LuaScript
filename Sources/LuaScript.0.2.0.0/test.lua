


function fTest()
	return "TEST!!!"
end



WScript:echo("yes!!!")



o = luacom.CreateObject("scripting.dictionary")

o:add("abc", "xyz")

WScript:echo( o:item("abc") )



WScript:echo("Hello A")

local s = "Hello B"
WScript:echo(s)



p = {}

p.xxx = function(n) return(n+1) end

WScript:echo(p.xxx(0))



WScript:echo(fTest())



