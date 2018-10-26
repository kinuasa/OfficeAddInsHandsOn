Option Explicit

Dim ret

Do
  ret = InputBox("Please copy a generated GUID." & vbNewLine & _
                 "You can then press Enter to regenerate GUID." & vbNewLine & vbNewLine & _
                 "Click the Cancel button or press ESC to finish.", _
                 "GUID Generator", _
                 LCase(Mid(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)))
Loop Until Len(Trim(ret)) < 1