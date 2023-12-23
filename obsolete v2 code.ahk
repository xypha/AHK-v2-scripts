;  = (obsolete) Min Max Array Function
;  = (obsolete) Min Max Ternary Function

#Requires AutoHotkey v2.0

;-------------------------------------------------------------------------------
;  = (obsolete) Min Max Ternary Function
; Use built-in functions - https://www.autohotkey.com/docs/v2/lib/Math.htm#Max

; Below code was modified from https://www.autohotkey.com/boards/viewtopic.php?p=22455#p22455

Min(x, y) {
  Return x < y ? x : y
}
; MsgBox Min(27, 64)    ; ==> 27

Max(x, y) {
  Return x < y ? y : x
}
; MsgBox Max(27, 64)    ; ==> 64

;-------------------------------------------------------------------------------
;  = (obsolete) Min Max Array Function
; Use built-in functions - https://www.autohotkey.com/docs/v2/lib/Math.htm#Max

; Below code was modified from https://www.autohotkey.com/boards/viewtopic.php?p=22481#p22481
; using decimals (and/or floating point numbers?) causes error that I haven't figured out yet

MinInArray(array) {
  Min := array.get(1)
  for index, value in array
    if (value < Min)
      Min := value
  return Min
}

MaxInArray(array) {
  Max := array.get(1)
  for index, value in array
    if (value > Max)
      Max := value
  return Max
}
; temp := [73, 12, 29, 83, 99, -1, -5, 3]
; Msgbox MinInArray(temp)    ; ==> -5
; Msgbox MaxInArray(temp)    ; ==> 99
