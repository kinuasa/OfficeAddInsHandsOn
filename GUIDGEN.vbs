Option Explicit

Dim ret

Do
  ret = InputBox("終了する場合は「キャンセル」ボタンをクリックしてください。", _
                 "GUID作成", _
                 LCase(Mid(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)))
Loop Until Len(Trim(ret)) < 1