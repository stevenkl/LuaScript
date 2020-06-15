<%@ language="LuaScript" %>
<html>

<%
	n=0
	function fTest2()
		o = luacom.CreateObject("scripting.dictionary")

		o:add("abc", "xyz")

		return( o("abc") .. ", " .. n )
	end
%>

<body>
	LuaScript: <% Response:write("yes!!!") %><br/>
	<%= "Hello" %>, <%= fTest2() %>
</body>

</html>
