VBScript code minifier written in VBScript

# USAGE

```vbs
Dim text
text = "x = 1" & vbCrLF & "y = 2"
text = New VBSMin.minify(text)
' Results to: "x=1:y=2"
```

# DESCRIPTION

This class implements the minification feature of VBScript codes
written in VBScript.

I found the only two implementations: one is written on Python [1]
and has partial support of minification, second is Ruby-based library
[2] having wide support.

I thought, it's good idea to implement this feature on VBScript
itself. This class does all the best to minify the code. Initially
it was translated from Ruby and implemented in a procedural style
[3]. Now it is OOP-styled and uses RegExp actively and produces a bit
shorter output than the Ruby library.

# LIMITATIONS

In most cases minification works and a minified code is executed
fine. However two cases exist when a minified code fails completely.

Both are having to do with the `If ... Then .. Else ...` statement [4].

No more investigations were done.

## Execution failure with the single-line `If ... Then ... Else ...`

VBScript supports a single-line syntax for `If ... Then ... Else ...`. A
minified code having a single-line construction doesn't start at all
and fails with the compilation error `Expected 'If'`.

There is short example demonstrating a failure:

```vbs
Function max(a, b)
	If a > b Then max = a Else max = b
End Function
```

The minified version of the code above doesn't work. To make it working,
it was minified, further the minified version was modified manually
until it stopped throwing compilation errors.

```vbs
Function max(a,b):If a>b Then max=a Else max=b:
End Function
```

## Execution failure with the `ElseIf` and `Else` keywords

Also the minified code doesn't work, if it contains `ElseIf` and
`Else` keywords within a line among other commands. The VBScript
engine requires them to be placed in the beginning of lines (to be the
first keyword in the line). The minified code stops execution with the
compilation errors `Must be first statement on the line` and `Expected
'End'`.

There is another example:

```vbs
Function sgn(a)
	If a > 0 Then
		sgn = +1
	ElseIf a<0 Then
		sgn = -1
	Else
		sgn = 0
	End If
End Function
```

The minified version of the code above doesn't work. To make it working,
it was minified, further the minified version was modified manually
until it stopped throwing compilation errors.

```vbs
Function sign(a):If a>0 Then
sign=+1:
ElseIf a<0 Then:sign=-1:
Else:sign=0:End If:End Function
```

To resolve these issues a deeper analysis is needed to distinguish
these specialties and put the line breaks immediately before `ElseIf`
and `Else` and after the single-line construction. But I am not sure
if it's really possible with regexps only.

*None of existsing minifiers resolve these issues.*

# SEE ALSO

* [1] Python implementation:
https://github.com/freginold/thinIt
* [2] Ruby implementation:
https://github.com/noraj/vbsmin
* [3] VBScript implementation (my first public version):
https://forum.script-coding.com/viewtopic.php?pid=143579#p143579
* [4] `If .. Then ... Else ...` statement:
https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/if-then-else-statement

# AUTHORS

Copyright 2021, Ildar Shaimordanov

MIT

<!-- sed -n "/^'/ s/..//p" <lib/vbsmin.vbs >README.md -->
