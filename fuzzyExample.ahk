#Include fuzzyMatch.ahk

Fuz := Fuzzy()
string_to_search := csharp_code() ;entire csharp fuzzymatch script

val4 := Fuz.Match("M4tch", "Match") ; 0.80
val5 := Fuz.Match("Match", string_to_search) ; 0.0014710208884965992

val1 := Fuz.LevenshteinDistance("hel1lo", "hello") ;1
val2 := Fuz.JaroDistance("Jaroooo", string_to_search) ;0
val3 := Fuz.JaroWinklerDistance("W1nkler", string_to_search) ;1


BestMatchLines := {
    Line: 0,
    Score: 0
}

;looking for the Fuzzy.Match function by line and rank in the csharp code
;exited with code=0 in 0.256 seconds
Loop Parse, string_to_search, "`n" "`r"
{
    test := Fuz.Match("J4roW1nklerD1stance", A_LoopField)
    if BestMatchLines.Score < test {
        BestMatchLines.Line := A_Index
        BestMatchLines.Score := test
    }
}


Loop 100
{
    ;exited with code=0 in 1.728 seconds
    test := Fuz.Match("hello", "he1lo")
    F := FileOpen("test.txt", "a", "utf-8")
    F.write(test "`n")
    F.Close()
}
MsgBox("The best match for 'J4roW1nklerD1stance' was found on line: " BestMatchLines.Line "`nwith a score of: " BestMatchLines.Score)
