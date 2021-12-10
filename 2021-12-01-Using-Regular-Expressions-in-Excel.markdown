---
layout: post
title: "Using Regular Expressions in Excel (Draft)"
date: 2021-10-01
categories: Functions, VBA
---

Excel has a lot of useful text functions, but sometimes a regular expression can be more effective for matching or replacing text strings. The following user defined function makes it possible to use regular expressions in a worksheet.

<!--more-->

```
Function RegexMatch(MyRange As Range, _
            strMatch As String, _
            booGlobal As Boolean) As Boolean
     
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")

        With regex
            .global = booGlobal
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strMatch
        End With

If regex.Test(MyRange) Then
            RegexMatch = True
        Else: RegexMatch = False
End If

End Function
```

Declaring a regular expression object means that you don't need to add a reference to Microsoft VBScript Regular Expressiosn 5.5 on every PC that uses this function:

```
Dim regex As Object
Set regex = CreateObject("VBScript.RegExp")
```



A function in the form RegexMatch(CELL_REF,Expression_to_match) can now be used in worksheets.

```
=RegexMatch(A3,"^T")
```
The above expression will return TRUE if the first character of cell A3 is a "T".


Using regular expressions to replace text is just as straightforwards. 

```
Function RegexReplace(MyRange As Range, _
            strMatch As String, _
            strReplace As String, _
            booGlobal As Boolean) As String
     
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
            .global = booGlobal
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strMatch
    End With

    RegexReplace = regex.Replace(MyRange, strReplace)

End Function
```

A function in the form regexreplace(CELL_REF,expression_to_match,replacement_text,Global_Replace) can now be used. Global replace can either be true or false.

```
=RegexReplace(A2,"^h","C",FALSE)
```

So here the text "hat" would be changed to "Cat" in cell A2.

Regular expressions can also be used to count matches.

```
Function RegexMatchCount(MyRange As Range, _
            strMatch As String) As Integer
     
Dim regex As Object
Dim RegexMatches As Object
Set regex = CreateObject("VBScript.RegExp")

        With regex
            .global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strMatch
        End With

Set RegexMatches = regex.Execute(MyRange)

RegexMatchCount = RegexMatches.Count()
           
End Function
```

This gives a function in the form:
RegexMatchCount(CELL_REFERENCE,Expression)

For example: 

RegexMatchCount(A4,"t")
