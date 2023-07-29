# AutoHotkeyv2 FuzzyMatch Library

## Description

This library provides fuzzy matching functionalities using C# and .NET Framework Interop (CLR) in AutoHotkeyv2. It allows you to calculate the similarity between two strings based on various fuzzy matching algorithms.

### Credit

CLR.ahk provided by Lexikos; Version: 2.0

## Usage

```autohotkey
#include "FuzzyMatch.ahk"
Fuz := Fuzzy()

;Fuzzy Match Score
val1 := Fuz.Match("M4tch", "Match") ; 0.80
val2 := Fuz.Match("Match", string_to_search) ; 0.0014710208884965992


;Levenshtein Distance
val3 := Fuz.LevenshteinDistance("hel1lo", "hello") ; 1

;Jaro Distance
val4 := Fuz.JaroDistance("Jaroooo", string_to_search) ; 0

;Jaro-Winkler Distance
val5 := Fuz.JaroWinklerDistance("W1nkler", string_to_search) ; 1
```
## Example: Finding the Best Match Line

```ahk
BestMatchLines := {
    Line: 0,
    Score: 0
}

; Looking for the Fuzzy.Match function by line and rank in the C# code
; Exited with code=0 in 0.256 seconds
Loop Parse, string_to_search, "`n" "`r"
{
    test := Fuz.Match("J4roW1nklerD1stance", A_LoopField)
    if BestMatchLines.Score < test {
        BestMatchLines.Line := A_Index
        BestMatchLines.Score := test
    }
}

MsgBox("The best match for 'J4roW1nklerD1stance' was found on line: " BestMatchLines.Line "`nwith a score of: " BestMatchLines.Score)
```


