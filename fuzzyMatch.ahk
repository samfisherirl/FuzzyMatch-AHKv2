; ; ==========================================================
; ;                  .NET Framework Interop
; ;   https://www.autohotkey.com/boards/viewtopic.php?t=4633
; ; ==========================================================
; ;
; ;   Author:     Lexikos
; ;   Version:    2.0
; ;   Requires:	AutoHotkey v2.0-beta.1

/*
#include "FuzzyMatch.ahk"
Fuz := Fuzzy()

val1 := Fuz.Match("M4tch", "Match") ; 0.80
val2 := Fuz.Match("Match", string_to_search) ; 0.0014710208884965992

;Levenshtein Distance
val3 := Fuz.LevenshteinDistance("hel1lo", "hello") ; 1

;Jaro Distance
val4 := Fuz.JaroDistance("Jaroooo", string_to_search) ; 0

;Jaro-Winkler Distance
val5 := Fuz.JaroWinklerDistance("W1nkler", string_to_search) ; 1
*/
class Fuzzy
{
    static Call() {
        asm := CLR_CompileCS(csharp_code(), "System.dll")
        return CLR_CreateObject(asm, "FuzzyMatcher")
    }
}


csharp_code(){
    return c := "
(
using System;

class FuzzyMatcher
{
    public int LevenshteinDistance(string string1, string string2)
    {
        if (string.IsNullOrEmpty(string1))
        {
            if (string.IsNullOrEmpty(string2))
                return 0;
            return string2.Length;
        }

        if (string.IsNullOrEmpty(string2))
        {
            if (string.IsNullOrEmpty(string1))
                return 0;
            return string1.Length;
        }

        var distance = new int[string1.Length + 1, string2.Length + 1];

        for (var i = 0; i <= string1.Length; i++)
            distance[i, 0] = i;

        for (var j = 0; j <= string2.Length; j++)
            distance[0, j] = j;

        for (var i = 1; i <= string1.Length; i++)
        {
            for (var j = 1; j <= string2.Length; j++)
            {
                var cost = (string1[i - 1] == string2[j - 1]) ? 0 : 1;

                distance[i, j] = Math.Min(
                    Math.Min(distance[i - 1, j] + 1, distance[i, j - 1] + 1),
                    distance[i - 1, j - 1] + cost);
            }
        }

        return distance[string1.Length, string2.Length];
    }
    public int JaroWinklerDistance(string str1, string str2)
    {
        int prefixLength = 0;
        int maxPrefixLength = Math.Min(4, Math.Min(str1.Length, str2.Length));

        for (int i = 0; i < maxPrefixLength; i++)
        {
            if (str1[i] == str2[i])
                prefixLength++;
            else
                break;
        }

        int jaroDistance = JaroDistance(str1, str2);
        int winklerDistance = jaroDistance + (int)Math.Round(prefixLength * 0.1 * (1 - jaroDistance));

        return 1 - winklerDistance;
    }

    public int JaroDistance(string str1, string str2)
    {
        int matchingCharacters = 0;
        int transpositions = 0;
        int str1Length = str1.Length;
        int str2Length = str2.Length;
        int searchRange = Math.Max(0, Math.Max(str1Length, str2Length) / 2 - 1);

        bool[] str1Matched = new bool[str1Length];
        bool[] str2Matched = new bool[str2Length];

        for (int i = 0; i < str1Length; i++)
        {
            int start = Math.Max(0, i - searchRange);
            int end = Math.Min(i + searchRange + 1, str2Length);

            for (int j = start; j < end; j++)
            {
                if (!str2Matched[j] && str1[i] == str2[j])
                {
                    str1Matched[i] = true;
                    str2Matched[j] = true;
                    matchingCharacters++;
                    break;
                }
            }
        }

        if (matchingCharacters == 0)
            return 0;

        int k = 0;
        for (int i = 0; i < str1Length; i++)
        {
            if (str1Matched[i])
            {
                while (!str2Matched[k]) k++;
                if (str1[i] != str2[k])
                    transpositions++;
                k++;
            }
        }

        return (matchingCharacters / str1Length + matchingCharacters / str2Length + (matchingCharacters - transpositions / 2) / matchingCharacters) / 3;
    }

    public double Match(string string1, string string2)
    {
        var maxLen = Math.Max(string1.Length, string2.Length);
        if (maxLen == 0)
            return 1;

        var dist = LevenshteinDistance(string1, string2);
        return (1 - (double)dist / maxLen);
    }
}
)"
}


; ==========================================================
;                  .NET Framework Interop
;   https://www.autohotkey.com/boards/viewtopic.php?t=4633
; ==========================================================
;
;   Author:     Lexikos
;   Version:    2.0
;   Requires:	AutoHotkey v2.0-beta.1
;

CLR_LoadLibrary(AssemblyName, AppDomain:=0)
{
	if !AppDomain
		AppDomain := CLR_GetDefaultDomain()
	try
		return AppDomain.Load_2(AssemblyName)
	static null := ComValue(13,0)
	args := ComObjArray(0xC, 1),  args[0] := AssemblyName
	typeofAssembly := AppDomain.GetType().Assembly.GetType()
	try
		return typeofAssembly.InvokeMember_3("LoadWithPartialName", 0x158, null, null, args)
	catch
		return typeofAssembly.InvokeMember_3("LoadFrom", 0x158, null, null, args)
}

CLR_CreateObject(Assembly, TypeName, Args*)
{
	if !(argCount := Args.Length)
		return Assembly.CreateInstance_2(TypeName, true)
	
	vargs := ComObjArray(0xC, argCount)
	Loop argCount
		vargs[A_Index-1] := Args[A_Index]
	
	static Array_Empty := ComObjArray(0xC,0), null := ComValue(13,0)
	
	return Assembly.CreateInstance_3(TypeName, true, 0, null, vargs, null, Array_Empty)
}

CLR_CompileCS(Code, References:="", AppDomain:=0, FileName:="", CompilerOptions:="")
{
	return CLR_CompileAssembly(Code, References, "System", "Microsoft.CSharp.CSharpCodeProvider", AppDomain, FileName, CompilerOptions)
}

CLR_CompileVB(Code, References:="", AppDomain:=0, FileName:="", CompilerOptions:="")
{
	return CLR_CompileAssembly(Code, References, "System", "Microsoft.VisualBasic.VBCodeProvider", AppDomain, FileName, CompilerOptions)
}

CLR_StartDomain(&AppDomain, BaseDirectory:="")
{
	static null := ComValue(13,0)
	args := ComObjArray(0xC, 5), args[0] := "", args[2] := BaseDirectory, args[4] := ComValue(0xB,false)
	AppDomain := CLR_GetDefaultDomain().GetType().InvokeMember_3("CreateDomain", 0x158, null, null, args)
}

; ICorRuntimeHost::UnloadDomain
CLR_StopDomain(AppDomain) => ComCall(20, CLR_Start(), "ptr", ComObjValue(AppDomain))

; NOTE: IT IS NOT NECESSARY TO CALL THIS FUNCTION unless you need to load a specific version.
CLR_Start(Version:="") ; returns ICorRuntimeHost*
{
	static RtHst := 0
	; The simple method gives no control over versioning, and seems to load .NET v2 even when v4 is present:
	; return RtHst ? RtHst : (RtHst:=COM_CreateObject("CLRMetaData.CorRuntimeHost","{CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}"), DllCall(NumGet(NumGet(RtHst+0)+40),"uint",RtHst))
	if RtHst
		return RtHst
	if Version = ""
		Loop Files EnvGet("SystemRoot") "\Microsoft.NET\Framework" (A_PtrSize=8?"64":"") "\*","D"
			if (FileExist(A_LoopFilePath "\mscorlib.dll") && StrCompare(A_LoopFileName, Version) > 0)
				Version := A_LoopFileName
	static CLSID_CorRuntimeHost := CLR_GUID("{CB2F6723-AB3A-11D2-9C40-00C04FA30A3E}")
	static IID_ICorRuntimeHost  := CLR_GUID("{CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}")
	DllCall("mscoree\CorBindToRuntimeEx", "wstr", Version, "ptr", 0, "uint", 0
		, "ptr", CLSID_CorRuntimeHost, "ptr", IID_ICorRuntimeHost
		, "ptr*", &RtHst:=0, "hresult")
	ComCall(10, RtHst) ; Start
	return RtHst
}

;
; INTERNAL FUNCTIONS
;

CLR_GetDefaultDomain()
{
	; ICorRuntimeHost::GetDefaultDomain
	static defaultDomain := (
		ComCall(13, CLR_Start(), "ptr*", &p:=0),
		ComObjFromPtr(p)
	)
	return defaultDomain
}

CLR_CompileAssembly(Code, References, ProviderAssembly, ProviderType, AppDomain:=0, FileName:="", CompilerOptions:="")
{
	if !AppDomain
		AppDomain := CLR_GetDefaultDomain()
	
	asmProvider := CLR_LoadLibrary(ProviderAssembly, AppDomain)
	codeProvider := asmProvider.CreateInstance(ProviderType)
	codeCompiler := codeProvider.CreateCompiler()

	asmSystem := (ProviderAssembly="System") ? asmProvider : CLR_LoadLibrary("System", AppDomain)

	; Convert | delimited list of references into an array.
	Refs := References is String ? StrSplit(References, "|", " `t") : References
	aRefs := ComObjArray(8, Refs.Length)
	Loop Refs.Length
		aRefs[A_Index-1] := Refs[A_Index]
	
	; Set parameters for compiler.
	prms := CLR_CreateObject(asmSystem, "System.CodeDom.Compiler.CompilerParameters", aRefs)
	, prms.OutputAssembly          := FileName
	, prms.GenerateInMemory        := FileName=""
	, prms.GenerateExecutable      := SubStr(FileName,-4)=".exe"
	, prms.CompilerOptions         := CompilerOptions
	, prms.IncludeDebugInformation := true
	
	; Compile!
	compilerRes := codeCompiler.CompileAssemblyFromSource(prms, Code)
	
	if error_count := (errors := compilerRes.Errors).Count
	{
		error_text := ""
		Loop error_count
			error_text .= ((e := errors.Item[A_Index-1]).IsWarning ? "Warning " : "Error ") . e.ErrorNumber " on line " e.Line ": " e.ErrorText "`n`n"
		throw Error("Compilation failed",, "`n" error_text)
	}
	; Success. Return Assembly object or path.
	return FileName="" ? compilerRes.CompiledAssembly : compilerRes.PathToAssembly
}

; Usage 1: pGUID := CLR_GUID(&GUID, "{...}")
; Usage 2: GUID := CLR_GUID("{...}"), pGUID := GUID.Ptr
CLR_GUID(a, b:=unset)
{
	DllCall("ole32\IIDFromString"
		, "wstr", sGUID := IsSet(b) ? b : a
		, "ptr", GUID := Buffer(16,0), "hresult")
	return IsSet(b) ? GUID.Ptr : GUID
}
