<div align="center">

## Count Words In a Texbox


</div>

### Description

This code will count the number of words in a textbox. A friend asked how to do this so I thought I'd post my reply.
 
### More Info
 
Any string.

Assumes a word is any set of characters separated by a space. The code can be easily modified to use any delineators.

The number of words in the string.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Oliver French](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/oliver-french.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/oliver-french-count-words-in-a-texbox__4-6662/archive/master.zip)





### Source Code

```
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<SCRIPT LANGUAGE = VBSCRIPT>
<!--
sub AddWords()
	dim strWhole
	dim x
	dim i
	dim currentLetter
	dim prevLetter
	dim total
	strWhole = Trim(form1.txtString.value) ' Remove leading and trailing spaces of string if any.
	if strWhole = "" then		' The text box is empty.
		msgbox "There are no words in the text box!"
		exit sub
	end if
	x = Len(strWhole)		' Store the total number of letters in the string.
	total = 1		' Start the count at one.
	' Loop through the string two letters at a time. If the current letter is a space, and the previous
	' letter was not a space, then another word is counted.
	For i = 2 to x
		currentLetter = mid(strWhole, i, 1)
		prevLetter = mid(strWhole, i - 1, 1)
		if currentLetter = Chr(32) and prevLetter <> Chr(32) then
			total = total + 1
		end if
	next
	msgbox "There are " & total & " words in the text box."
end sub
' -->
</Script>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p>Enter the string you wish counted:</p>
<form name="form1" method="post" action="">
 <p>
  <textarea name=txtString rows=5 cols=51></textarea>
 </p>
 <p>&nbsp;</p>
 <p>
  <input type="button" name="cmdExecute" value="Add Words" onclick=AddWords()>
 </p>
 <p>&nbsp;</p>
</form>
</body>
</html>
```

