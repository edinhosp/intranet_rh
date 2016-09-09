<html>
<head>
<title>Example 1: Hello World</title>
</head>
<body bgcolor=white>
<%
dim strUsersBrowser as string
strUsersBrowser+=request.browser.browser
strUsersBrowser+=cstr(request.browser.majorversion)
strUsersBrowser+="."
strUsersBrowser+=cstr(request.browser.minorversion)
response.write("<h1>Your web browser is " & strUsersBrowser & "</h1>")
%>
</body>
</html>