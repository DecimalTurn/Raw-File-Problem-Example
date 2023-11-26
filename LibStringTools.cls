VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibStringTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId

'===============================================================================
' VBA StringTools
' ------------------------------------------
' https://github.com/guwidoe/VBA-StringTools
' ------------------------------------------
' MIT License
'
' Copyright (c) 2023 Guido Witt-Döring
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'===============================================================================

Option Explicit
Option Base 0
Option Compare Binary

#If Mac Then
    #If VBA7 Then 'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv.3.html
        Private Declare PtrSafe Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As LongPtr, ByVal fromCode As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr) As Long
        Private Declare PtrSafe Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr, ByRef InBuf As LongPtr, ByRef InBytesLeft As LongPtr, ByRef OutBuf As LongPtr, ByRef OutBytesLeft As LongPtr) As LongPtr

        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
        Private Declare PtrSafe Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As LongPtr
    #Else
        Private Declare Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long, ByRef inBuf As Long, ByRef inBytesLeft As Long, ByRef outBuf As Long, ByRef outBytesLeft As Long) As Long
        Private Declare Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As Long, ByVal fromCode As Long) As Long
        Private Declare Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long) As Long

        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
        Private Declare Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal DwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
        Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal DwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal LpUsedDefaultChar As LongPtr) As Long

        Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
        Private Declare PtrSafe Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    #Else
        Private Declare Function MultiByteToWideChar Lib "kernel32" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
        Private Declare Function WideCharToMultiByte Lib "kernel32" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

        Private Declare Function GetLastError Lib "kernel32" () As Long
        Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    #End If
#End If

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

'Flag used to simulate ByRef Variants
Private Const VT_BYREF As Long = &H4000

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    #If Win64 Then
        dummyPadding As Long
        pvData As LongLong
    #Else
        pvData As Long
    #End If
    rgsabound0 As SAFEARRAYBOUND
End Type
Private Const FADF_HAVEVARTYPE As Long = &H80

Private Const BYTE_SIZE As Long = 1
Private Const INT_SIZE As Long = 2

Private Type EscapeSequence
    ueFormat As UnicodeEscapeFormat
    ueSignature As String
    letSngSurrogate As Boolean
    buffPosition As Long
    currPosition As Long
    sigSize As Long
    escSize As Long
    Codepoint As Long
    unEscSize As Long
End Type

Private Type TwoCharTemplate
    s As String * 2
End Type
Private Type LongTemplate
    l As Long
End Type

#If Win64 Then
    #If Mac Then
        Private Const vbLongLong As Long = 20 'Apparently missing for x64 on Mac
    #End If
    Private Const vbLongPtr As Long = vbLongLong
#Else
    Private Const vbLongLong As Long = 20 'Useful in Select Case logic
    Private Const vbLongPtr As Long = vbLong
#End If

Private Type StringificationSettings
    maxChars As Long
    escapeNonPrintable As Boolean
    Delimiter As String
    maxCharsPerElement As Long
    maxCharsPerLine As Long
    maxLines As Long
    inklColIndices As Boolean
    inklRowIndices As Boolean
End Type

Dim PrintfSettings As StringificationSettings
Dim PrintfSettingsAreInitialized As Boolean

Private Const WC_ERR_INVALID_CHARS As Long = &H80&
Private Const MB_ERR_INVALID_CHARS As Long = &H8&

Private Const ERROR_INVALID_PARAMETER      As Long = 87
Private Const ERROR_INSUFFICIENT_BUFFER    As Long = 122
Private Const ERROR_INVALID_FLAGS          As Long = 1004
Private Const ERROR_NO_UNICODE_TRANSLATION As Long = 1113

Private Const MAC_API_ERR_EILSEQ As Long = 92 'Illegal byte sequence
Private Const MAC_API_ERR_EINVAL As Long = 22 'Invalid argument
Private Const MAC_API_ERR_E2BIG  As Long = 7  'Argument list too long

Private Const vbErrInternalError As Long = 51

Public Enum UnicodeEscapeFormat
    [_efNone] = 0
    efPython = 1 '\uXXXX \u00XXXXXX (4 or 8 hex digits, 8 for chars outside BMP)
    efRust = 2   '\u{X} \U{XXXXXX}  (1 to 6 hex digits)
    efUPlus = 4  'u+XXXX u+XXXXXX   (4 or 6 hex digits)
    efMarkup = 8 '&#ddddddd;        (1 to 7 decimal digits)
    efAll = 15
    [_efMin] = efPython
    [_efMax] = efAll
End Enum

'https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
Public Enum CodePageIdentifier
    [_first] = -1 '(Is initialized)
  'Enum_Name   Identifier             '.NET Name               Additional information
    cpIBM037 = 37                     'IBM037                  IBM EBCDIC US-Canada
    cpIBM437 = 437                    'IBM437                  OEM United States
    cpIBM500 = 500                    'IBM500                  IBM EBCDIC International
    cpASMO_708 = 708                  'ASMO-708                Arabic (ASMO 708)
    cpASMO_449 = 709                  '                        Arabic (ASMO-449+, BCON V4)
    cpTransparent_Arabic = 710        '                        Arabic - Transparent Arabic
    cpDOS_720 = 720                   'DOS-720                 Arabic (Transparent ASMO); Arabic (DOS)
    cpIbm737 = 737                    'ibm737                  OEM Greek (formerly 437G); Greek (DOS)
    cpIbm775 = 775                    'ibm775                  OEM Baltic; Baltic (DOS)
    cpIbm850 = 850                    'ibm850                  OEM Multilingual Latin 1; Western European (DOS)
    cpIbm852 = 852                    'ibm852                  OEM Latin 2; Central European (DOS)
    cpIBM855 = 855                    'IBM855                  OEM Cyrillic (primarily Russian)
    cpIbm857 = 857                    'ibm857                  OEM Turkish; Turkish (DOS)
    cpIBM00858 = 858                  'IBM00858                OEM Multilingual Latin 1 + Euro symbol
    cpIBM860 = 860                    'IBM860                  OEM Portuguese; Portuguese (DOS)
    cpIbm861 = 861                    'ibm861                  OEM Icelandic; Icelandic (DOS)
    cpDOS_862 = 862                   'DOS-862                 OEM Hebrew; Hebrew (DOS)
    cpIBM863 = 863                    'IBM863                  OEM French Canadian; French Canadian (DOS)
    cpIBM864 = 864                    'IBM864                  OEM Arabic; Arabic (864)
    cpIBM865 = 865                    'IBM865                  OEM Nordic; Nordic (DOS)
    cpCp866 = 866                     'cp866                   OEM Russian; Cyrillic (DOS)
    cpIbm869 = 869                    'ibm869                  OEM Modern Greek; Greek, Modern (DOS)
    cpIBM870 = 870                    'IBM870                  IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2
    cpWindows_874 = 874               'windows-874             Thai (Windows)
    cpCp875 = 875                     'cp875                   IBM EBCDIC Greek Modern
    cpShift_jis = 932                 'shift_jis               ANSI/OEM Japanese; Japanese (Shift-JIS)
    cpGb2312 = 936                    'gb2312                  ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)
    cpKs_c_5601_1987 = 949            'ks_c_5601-1987          ANSI/OEM Korean (Unified Hangul Code)
    cpBig5 = 950                      'big5                    ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
    cpIBM1026 = 1026                  'IBM1026                 IBM EBCDIC Turkish (Latin 5)
    cpIBM01047 = 1047                 'IBM01047                IBM EBCDIC Latin 1/Open System
    cpIBM01140 = 1140                 'IBM01140                IBM EBCDIC US-Canada (037 + Euro symbol); IBM EBCDIC (US-Canada-Euro)
    cpIBM01141 = 1141                 'IBM01141                IBM EBCDIC Germany (20273 + Euro symbol); IBM EBCDIC (Germany-Euro)
    cpIBM01142 = 1142                 'IBM01142                IBM EBCDIC Denmark-Norway (20277 + Euro symbol); IBM EBCDIC (Denmark-Norway-Euro)
    cpIBM01143 = 1143                 'IBM01143                IBM EBCDIC Finland-Sweden (20278 + Euro symbol); IBM EBCDIC (Finland-Sweden-Euro)
    cpIBM01144 = 1144                 'IBM01144                IBM EBCDIC Italy (20280 + Euro symbol); IBM EBCDIC (Italy-Euro)
    cpIBM01145 = 1145                 'IBM01145                IBM EBCDIC Latin America-Spain (20284 + Euro symbol); IBM EBCDIC (Spain-Euro)
    cpIBM01146 = 1146                 'IBM01146                IBM EBCDIC United Kingdom (20285 + Euro symbol); IBM EBCDIC (UK-Euro)
    cpIBM01147 = 1147                 'IBM01147                IBM EBCDIC France (20297 + Euro symbol); IBM EBCDIC (France-Euro)
    cpIBM01148 = 1148                 'IBM01148                IBM EBCDIC International (500 + Euro symbol); IBM EBCDIC (International-Euro)
    cpIBM01149 = 1149                 'IBM01149                IBM EBCDIC Icelandic (20871 + Euro symbol); IBM EBCDIC (Icelandic-Euro)
    cpUTF_16 = 1200                   'utf-16                  Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications
    cpUnicodeFFFE = 1201              'unicodeFFFE             Unicode UTF-16, big endian byte order; available only to managed applications
    cpWindows_1250 = 1250             'windows-1250            ANSI Central European; Central European (Windows)
    cpWindows_1251 = 1251             'windows-1251            ANSI Cyrillic; Cyrillic (Windows)
    cpWindows_1252 = 1252             'windows-1252            ANSI Latin 1; Western European (Windows)
    cpWindows_1253 = 1253             'windows-1253            ANSI Greek; Greek (Windows)
    cpWindows_1254 = 1254             'windows-1254            ANSI Turkish; Turkish (Windows)
    cpWindows_1255 = 1255             'windows-1255            ANSI Hebrew; Hebrew (Windows)
    cpWindows_1256 = 1256             'windows-1256            ANSI Arabic; Arabic (Windows)
    cpWindows_1257 = 1257             'windows-1257            ANSI Baltic; Baltic (Windows)
    cpWindows_1258 = 1258             'windows-1258            ANSI/OEM Vietnamese; Vietnamese (Windows)
    cpJohab = 1361                    'Johab                   Korean (Johab)
    cpMacintosh = 10000               'macintosh               MAC Roman; Western European (Mac)
    cpX_mac_japanese = 10001          'x-mac-japanese          Japanese (Mac)
    cpX_mac_chinesetrad = 10002       'x-mac-chinesetrad       MAC Traditional Chinese (Big5); Chinese Traditional (Mac)
    cpX_mac_korean = 10003            'x-mac-korean            Korean (Mac)
    cpX_mac_arabic = 10004            'x-mac-arabic            Arabic (Mac)
    cpX_mac_hebrew = 10005            'x-mac-hebrew            Hebrew (Mac)
    cpX_mac_greek = 10006             'x-mac-greek             Greek (Mac)
    cpX_mac_cyrillic = 10007          'x-mac-cyrillic          Cyrillic (Mac)
    cpX_mac_chinesesimp = 10008       'x-mac-chinesesimp       MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)
    cpX_mac_romanian = 10010          'x-mac-romanian          Romanian (Mac)
    cpX_mac_ukrainian = 10017         'x-mac-ukrainian         Ukrainian (Mac)
    cpX_mac_thai = 10021              'x-mac-thai              Thai (Mac)
    cpX_mac_ce = 10029                'x-mac-ce                MAC Latin 2; Central European (Mac)
    cpX_mac_icelandic = 10079         'x-mac-icelandic         Icelandic (Mac)
    cpX_mac_turkish = 10081           'x-mac-turkish           Turkish (Mac)
    cpX_mac_croatian = 10082          'x-mac-croatian          Croatian (Mac)
    cpUTF_32 = 12000                  'utf-32                  Unicode UTF-32, little endian byte order; available only to managed applications
    cpUTF_32BE = 12001                'utf-32BE                Unicode UTF-32, big endian byte order; available only to managed applications
    cpX_Chinese_CNS = 20000           'x-Chinese_CNS           CNS Taiwan; Chinese Traditional (CNS)
    cpX_cp20001 = 20001               'x-cp20001               TCA Taiwan
    cpX_Chinese_Eten = 20002          'x_Chinese-Eten          Eten Taiwan; Chinese Traditional (Eten)
    cpX_cp20003 = 20003               'x-cp20003               IBM5550 Taiwan
    cpX_cp20004 = 20004               'x-cp20004               TeleText Taiwan
    cpX_cp20005 = 20005               'x-cp20005               Wang Taiwan
    cpX_IA5 = 20105                   'x-IA5                   IA5 (IRV International Alphabet No. 5, 7-bit); Western European (IA5)
    cpX_IA5_German = 20106            'x-IA5-German            IA5 German (7-bit)
    cpX_IA5_Swedish = 20107           'x-IA5-Swedish           IA5 Swedish (7-bit)
    cpX_IA5_Norwegian = 20108         'x-IA5-Norwegian         IA5 Norwegian (7-bit)
    cpUs_ascii = 20127                'us-ascii                US-ASCII (7-bit)
    cpX_cp20261 = 20261               'x-cp20261               T.61
    cpX_cp20269 = 20269               'x-cp20269               ISO 6937 Non-Spacing Accent
    cpIBM273 = 20273                  'IBM273                  IBM EBCDIC Germany
    cpIBM277 = 20277                  'IBM277                  IBM EBCDIC Denmark-Norway
    cpIBM278 = 20278                  'IBM278                  IBM EBCDIC Finland-Sweden
    cpIBM280 = 20280                  'IBM280                  IBM EBCDIC Italy
    cpIBM284 = 20284                  'IBM284                  IBM EBCDIC Latin America-Spain
    cpIBM285 = 20285                  'IBM285                  IBM EBCDIC United Kingdom
    cpIBM290 = 20290                  'IBM290                  IBM EBCDIC Japanese Katakana Extended
    cpIBM297 = 20297                  'IBM297                  IBM EBCDIC France
    cpIBM420 = 20420                  'IBM420                  IBM EBCDIC Arabic
    cpIBM423 = 20423                  'IBM423                  IBM EBCDIC Greek
    cpIBM424 = 20424                  'IBM424                  IBM EBCDIC Hebrew
    cpX_EBCDIC_KoreanExtended = 20833 'x-EBCDIC-KoreanExtended IBM EBCDIC Korean Extended
    cpIBM_Thai = 20838                'IBM-Thai                IBM EBCDIC Thai
    cpKoi8_r = 20866                  'koi8-r                  Russian (KOI8-R); Cyrillic (KOI8-R)
    cpIBM871 = 20871                  'IBM871                  IBM EBCDIC Icelandic
    cpIBM880 = 20880                  'IBM880                  IBM EBCDIC Cyrillic Russian
    cpIBM905 = 20905                  'IBM905                  IBM EBCDIC Turkish
    cpIBM00924 = 20924                'IBM00924                IBM EBCDIC Latin 1/Open System (1047 + Euro symbol)
    cpEuc_jp = 20932                  'EUC-JP                  Japanese (JIS 0208-1990 and 0212-1990)
    cpX_cp20936 = 20936               'x-cp20936               Simplified Chinese (GB2312); Chinese Simplified (GB2312-80)
    cpX_cp20949 = 20949               'x-cp20949               Korean Wansung
    cpCp1025 = 21025                  'cp1025                  IBM EBCDIC Cyrillic Serbian-Bulgarian
    cpDeprecated = 21027                       '                        (deprecated)
    cpKoi8_u = 21866                  'koi8-u                  Ukrainian (KOI8-U); Cyrillic (KOI8-U)
    cpIso_8859_1 = 28591              'iso-8859-1              ISO 8859-1 Latin 1; Western European (ISO)
    cpIso_8859_2 = 28592              'iso-8859-2              ISO 8859-2 Central European; Central European (ISO)
    cpIso_8859_3 = 28593              'iso-8859-3              ISO 8859-3 Latin 3
    cpIso_8859_4 = 28594              'iso-8859-4              ISO 8859-4 Baltic
    cpIso_8859_5 = 28595              'iso-8859-5              ISO 8859-5 Cyrillic
    cpIso_8859_6 = 28596              'iso-8859-6              ISO 8859-6 Arabic
    cpIso_8859_7 = 28597              'iso-8859-7              ISO 8859-7 Greek
    cpIso_8859_8 = 28598              'iso-8859-8              ISO 8859-8 Hebrew; Hebrew (ISO-Visual)
    cpIso_8859_9 = 28599              'iso-8859-9              ISO 8859-9 Turkish
    cpIso_8859_13 = 28603             'iso-8859-13             ISO 8859-13 Estonian
    cpIso_8859_15 = 28605             'iso-8859-15             ISO 8859-15 Latin 9
    cpX_Europa = 29001                'x-Europa                Europa 3
    cpIso_8859_8_i = 38598            'iso-8859-8-i            ISO 8859-8 Hebrew; Hebrew (ISO-Logical)
    cpIso_2022_jp = 50220             'iso-2022-jp             ISO 2022 Japanese with no halfwidth Katakana; Japanese (JIS)
    cpCsISO2022JP = 50221             'csISO2022JP             ISO 2022 Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)
    cpIso_2022_jp_w_1b_Kana = 50222   'iso-2022-jp             ISO 2022 Japanese JIS X 0201-1989; Japanese (JIS-Allow 1 byte Kana - SO/SI)
    cpIso_2022_kr = 50225             'iso-2022-kr             ISO 2022 Korean
    cpX_cp50227 = 50227               'x-cp50227               ISO 2022 Simplified Chinese; Chinese Simplified (ISO 2022)
    cpISO_2022_Trad_Chinese = 50229   '                        ISO 2022 Traditional Chinese
    cpEBCDIC_Jap_Katakana_Ext = 50930 '                        EBCDIC Japanese (Katakana) Extended
    cpEBCDIC_US_Can_and_Jap = 50931   '                        EBCDIC US-Canada and Japanese
    cpEBCDIC_Kor_Ext_and_Kor = 50933  '                        EBCDIC Korean Extended and Korean
    cpEBCDIC_Simp_Chin_Ext = 50935    '                        EBCDIC Simplified Chinese Extended and Simplified Chinese
    cpEBCDIC_Simp_Chin = 50936        '                        EBCDIC Simplified Chinese
    cpEBCDIC_US_Can_Trad_Chin = 50937 '                        EBCDIC US-Canada and Traditional Chinese
    cpEBCDIC_Jap_Latin_Ext = 50939    '                        EBCDIC Japanese (Latin) Extended and Japanese
    euc_jp = 51932                    'euc-jp                  EUC Japanese
    cpEUC_CN = 51936                  'EUC-CN                  EUC Simplified Chinese; Chinese Simplified (EUC)
    cpEuc_kr = 51949                  'euc-kr                  EUC Korean
    cpEUC_Traditional_Chinese = 51950 '                        EUC Traditional Chinese
    cpHz_gb_2312 = 52936              'hz-gb-2312              HZ-GB2312 Simplified Chinese; Chinese Simplified (HZ)
    cpGB18030 = 54936                 'GB18030                 Windows XP and later: GB18030 Simplified Chinese (4 byte); Chinese Simplified (GB18030)
    cpX_iscii_de = 57002              'x-iscii-de              ISCII Devanagari
    cpX_iscii_be = 57003              'x-iscii-be              ISCII Bangla
    cpX_iscii_ta = 57004              'x-iscii-ta              ISCII Tamil
    cpX_iscii_te = 57005              'x-iscii-te              ISCII Telugu
    cpX_iscii_as = 57006              'x-iscii-as              ISCII Assamese
    cpX_iscii_or = 57007              'x-iscii-or              ISCII Odia
    cpX_iscii_ka = 57008              'x-iscii-ka              ISCII Kannada
    cpX_iscii_ma = 57009              'x-iscii-ma              ISCII Malayalam
    cpX_iscii_gu = 57010              'x-iscii-gu              ISCII Gujarati
    cpX_iscii_pa = 57011              'x-iscii-pa              ISCII Punjabi
    cpUTF_7 = 65000                   'utf-7                   Unicode (UTF-7)
    cpUTF_8 = 65001                   'utf-8                   Unicode (UTF-8)
    [_last]
End Enum

'According to documentation:
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-multibytetowidechar
'Note: The documentation doesn't seem to list all codepages for which certain
'      flags are disallowed. This can lead to 'Library implementation erroneous'
'      errors when calling Encode, Decode or Transcode with 'raiseErrors = True'
Private Static Function CodePageAllowsFlags(ByVal cpID As Long) As Boolean
    Dim Arr(CodePageIdentifier.[_first] To CodePageIdentifier.[_last]) As Boolean

    If Arr(CodePageIdentifier.[_first]) Then
        CodePageAllowsFlags = Arr(cpID)
        Exit Function
    End If

    Dim i As Long
    For i = CodePageIdentifier.[_first] To CodePageIdentifier.[_last]
        Arr(i) = True
    Next i

    'According to docs:
    Arr(cpIso_2022_jp) = False
    Arr(cpCsISO2022JP) = False
    Arr(cpIso_2022_jp_w_1b_Kana) = False
    Arr(cpIso_2022_kr) = False
    Arr(cpX_cp50227) = False
    Arr(cpISO_2022_Trad_Chinese) = False
    For i = cpX_iscii_de To cpX_iscii_pa
        Arr(i) = False
    Next i
    Arr(cpUTF_7) = False

    'According to trial and error, it is easier to whitelist:
    For i = CodePageIdentifier.[_first] + 1 To CodePageIdentifier.[_last]
        Arr(i) = False
    Next i
    Arr(cpUTF_32) = True   'Not sure about this one
    Arr(cpUTF_32BE) = True 'Not sure about this one
    Arr(cpGB18030) = True  'This one is definitely allowed
    Arr(cpUTF_8) = True    'This one is definitely allowed

    CodePageAllowsFlags = Arr(cpID)
End Function

'According to documentation:
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-multibytetowidechar
Private Static Function CodePageAllowsQueryReversible(ByVal cpID As Long) As Boolean
    Dim Arr(CodePageIdentifier.[_first] To CodePageIdentifier.[_last]) As Boolean

    If Arr(CodePageIdentifier.[_first]) Then
        CodePageAllowsQueryReversible = Arr(cpID)
        Exit Function
    End If

    Dim i As Long
    For i = CodePageIdentifier.[_first] To CodePageIdentifier.[_last]
        Arr(i) = True
    Next i

    'According to docs:
    Arr(cpUTF_7) = False
    Arr(cpUTF_8) = False

    'According to trial and error there are a bunch more:
    Arr(cpIso_2022_jp) = False
    Arr(cpCsISO2022JP) = False
    Arr(cpIso_2022_jp_w_1b_Kana) = False
    Arr(cpIso_2022_kr) = False
    Arr(cpX_cp50227) = False
    Arr(cpISO_2022_Trad_Chinese) = False
    Arr(cpHz_gb_2312) = False
    Arr(cpGB18030) = False
    Arr(cpX_iscii_de) = False
    Arr(cpX_iscii_be) = False
    Arr(cpX_iscii_ta) = False
    Arr(cpX_iscii_te) = False
    Arr(cpX_iscii_as) = False
    Arr(cpX_iscii_or) = False
    Arr(cpX_iscii_ka) = False
    Arr(cpX_iscii_ma) = False
    Arr(cpX_iscii_gu) = False
    Arr(cpX_iscii_pa) = False

    CodePageAllowsQueryReversible = Arr(cpID)
End Function

'Returns an array for converting CodePageIDs to ConversionDescriptorNames
Private Static Function ConvDescriptorName(ByVal cpID As Long) As String
    Dim Arr(CodePageIdentifier.[_first] To CodePageIdentifier.[_last]) As String

    If Arr(CodePageIdentifier.[_first]) Then
        ConvDescriptorName = StrConv(Arr(cpID), vbFromUnicode)
        Exit Function
    End If

    Dim i As Long
    For i = CodePageIdentifier.[_first] To CodePageIdentifier.[_last]
        Arr(i) = -1
    Next i

    'Source:
    'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv_open.3.html#//apple_ref/doc/man/3/iconv_open
    'European languages
    Arr(cpIso_8859_1) = "ISO-8859-1"
    Arr(cpIso_8859_2) = "ISO-8859-2"
    Arr(cpIso_8859_3) = "ISO-8859-3"
    Arr(cpIso_8859_4) = "ISO-8859-4"
    Arr(cpIso_8859_5) = "ISO-8859-5"
    Arr(cpIso_8859_7) = "ISO-8859-7"
    Arr(cpIso_8859_9) = "ISO-8859-9"
    Arr(28600) = "ISO-8859-10"
    Arr(cpIso_8859_13) = "ISO-8859-13"
    Arr(28604) = "ISO-8859-14"
    Arr(cpIso_8859_15) = "ISO-8859-15"
    Arr(28606) = "ISO-8859-16"
    Arr(20866) = "KOI8-R"
    Arr(cpKoi8_u) = "KOI8-U"
    'arr( ) =  "KOI8-RU" 'No equivalent ID, variation of KOI8-R
    Arr(cpWindows_1250) = "CP1250"
    Arr(cpWindows_1251) = "CP1251"
    Arr(cpWindows_1252) = "CP1252"
    Arr(cpWindows_1253) = "CP1253"
    Arr(cpWindows_1254) = "CP1254"
    Arr(cpWindows_1257) = "CP1257"
    Arr(cpIbm850) = "CP850"
    Arr(cpCp866) = "CP866"
    Arr(cpMacintosh) = "MacRoman"   'duplicate
    Arr(cpX_mac_ce) = "MacCentralEurope"
    Arr(cpX_mac_icelandic) = "MacIceland"
    Arr(cpX_mac_croatian) = "MacCroatian"
    Arr(cpX_mac_romanian) = "MacRomania"
    Arr(cpX_mac_cyrillic) = "MacCyrillic"
    Arr(cpX_mac_ukrainian) = "MacUkraine"
    Arr(cpX_mac_greek) = "MacGreek"
    Arr(cpX_mac_turkish) = "MacTurkish"
    Arr(cpMacintosh) = "Macintosh"

    'Semitic languages
    Arr(cpIso_8859_6) = "ISO-8859-6"
    Arr(cpIso_8859_8) = "ISO-8859-8"
    Arr(cpWindows_1255) = "CP1255"
    Arr(cpWindows_1256) = "CP1256"
    Arr(cpDOS_862) = "CP862"
    Arr(cpX_mac_hebrew) = "MacHebrew"
    Arr(cpX_mac_arabic) = "MacArabic"

    'Japanese
    Arr(euc_jp) = "EUC-JP"
    Arr(cpShift_jis) = "SHIFT_JIS"
    Arr(cpShift_jis) = "CP932" '(duplicate)
    Arr(cpIso_2022_jp) = "ISO-2022-JP"
    Arr(cpCsISO2022JP) = "ISO-2022-JP-2"
    Arr(cpIso_2022_jp_w_1b_Kana) = "ISO-2022-JP-1"

    'Chinese
    Arr(cpEUC_CN) = "EUC-CN"
    'arr( ) =  "HZ" 'No equivalent ID, 7-bit encoding method for GB2312
    Arr(cpGb2312) = "GBK" 'duplicate
    Arr(cpGb2312) = "CP936"
    Arr(cpGB18030) = "GB18030"
    'arr( ) =  "EUC-TW" 'No equivalent ID, extended UNIX Code for Traditional Chinese
    Arr(cpBig5) = "BIG5"
    Arr(cpBig5) = "CP950" '(duplicate)
    Arr(951) = "BIG5-HKSCS"
    Arr(951) = "BIG5-HKSCS:2001"
    Arr(951) = "BIG5-HKSCS:1999"
    Arr(cpX_cp50227) = "ISO-2022-CN"
    'arr( ) =  "ISO-2022-CN-EXT" 'No equivalent ID, extended version of ISO-2022-CN

    'Korean
    Arr(cpEuc_kr) = "EUC-KR"
    Arr(cpKs_c_5601_1987) = "CP949"
    Arr(cpIso_2022_kr) = "ISO-2022-KR"
    Arr(cpJohab) = "JOHAB"

    'Armenian
    'arr( ) =  "ARMSCII-8" '8-bit Armenian character encoding

    'Georgian
    'arr( ) =  "Georgian-Academy" 'No equivalent ID
    'arr( ) =  "Georgian-PS" 'No equivalent ID

    'Tajik
    'arr( ) =  "KOI8-T" 'No equivalent ID

    'Kazakh
    'arr( ) =  "PT154" 'No equivalent ID, Paratype KZ

    'Thai
    Arr(cpWindows_874) = "TIS-620" 'duplicate
    Arr(cpWindows_874) = "CP874"
    Arr(cpX_mac_thai) = "MacThai"

    'Laotian
    'arr( ) =  "MuleLao-1" 'No equivalent ID, MULE (MULtilingual Enhancement to GNU Emacs) internal encoding for the Lao script
    Arr(1133) = "CP1133"

    'Vietnamese
    'arr( ) =  "VISCII" 'No equivalent ID, 8-bit encoding for the Vietnamese alphabet
    'arr( ) =  "TCVN" 'No equivalent ID, Vietnamese national standard for character encoding
    Arr(cpWindows_1258) = "CP1258"

    'Platform specifics
    'arr( ) =  "HP-ROMAN8" 'No equivalent ID, 8-bit character encoding used by Hewlett-Packard for their workstations and printers.
    'arr( ) =  "NEXTSTEP" 'No equivalent ID, encoding is associated with the NeXTSTEP operating system developed by NeXT, the company founded by Steve Jobs after leaving Apple in the 1980s.

    'Full Unicode
    'arr( ) =  "UCS-2"
    Arr(cpUnicodeFFFE) = "UCS-2BE" '(duplicate)
    Arr(cpUTF_16) = "UCS-2LE" '(duplicate)
    'arr( ) =  "UCS-4"
    Arr(cpUTF_32BE) = "UCS-4BE" '(duplicate)
    Arr(cpUTF_32) = "UCS-4LE" '(duplicate)
    'arr( ) =  "UTF-16"
    Arr(cpUnicodeFFFE) = "UTF-16BE"
    Arr(cpUTF_16) = "UTF-16LE"
    'arr( ) =  "UTF-32"
    Arr(cpUTF_32BE) = "UTF-32BE"
    Arr(cpUTF_32) = "UTF-32LE"
    Arr(cpUTF_7) = "UTF-7"
    Arr(cpUTF_8) = "UTF-8"
    'arr( ) =  "C99"
    'arr( ) =  "JAVA"

    'Full Unicode in terms of uint16_t or uint32_t
    '(with machine dependent endianness and alignment)
    'arr( ) =  "UCS-2-INTERNAL"
    'arr( ) =  "UCS-4-INTERNAL"

    'Locale dependent in terms of char or wchar_t
    '(with  machine  dependent  endianness  and  alignment and with
    'semantics depending on the OS and the  current  LC_CTYPE  locale facet)
    'arr( ) =  "char"
    'arr( ) =  "wchar_t"

    'When  configured with the option --enable-extra-encodings
    'it also pro-vides provides vides support for a few extra encodings:

    'European languages
    Arr(cpIBM437) = "CP437"
    Arr(cpIbm737) = "CP737"
    Arr(cpIbm775) = "CP775"
    Arr(cpIbm852) = "CP852"
    Arr(853) = "CP853"
    Arr(cpIBM855) = "CP855"
    Arr(cpIbm857) = "CP857"
    Arr(cpIBM00858) = "CP858"
    Arr(cpIBM860) = "CP860"
    Arr(cpIbm861) = "CP861"
    Arr(cpIBM863) = "CP863"
    Arr(cpIBM865) = "CP865"
    Arr(cpIbm869) = "CP869"
    Arr(1125) = "CP1125"

    'Semitic languages
    Arr(cpIBM864) = "CP864"

    'Japanese
    'arr( ) =  "EUC-JISX0213" 'No equivalent ID
    'arr( ) =  "Shift_JISX0213" 'No equivalent ID
    'arr( ) =  "ISO-2022-JP-3" 'No equivalent ID

    'Chinese
    'arr( ) = "BIG5-2003" '(experimental 'No equivalent ID

    'Turkmen
    'arr( ) =  "TDS565" 'No equivalent ID

    'Platform specifics
    'arr( ) =  "ATARIST" 'No equivalent ID, 8-bit character encoding used on Atari ST computers, which were a series of personal computers released in the 1980s.
    'arr( ) =  "RISCOS-LATIN1" 'No equivalent ID, 8-bit character encoding used on the RISC OS operating system, which was developed by Acorn Computers in the late 1980s.

    'The empty encoding name is equivalent to "char":
    'it denotes the locale dependent character encoding.
    ConvDescriptorName = ConvDescriptorName(cpID)
End Function

''Returns a Collection for converting ConversionDescriptorNames to CodePageIDs
'Private Function ConvDescriptorNameToCodePage() As Collection
'    Static c As Collection
'
'    If Not c Is Nothing Then
'        Set ConvDescriptorNameToCodePage = c
'        Exit Function
'    End If
'
'    Dim cpID As Long
'    Dim conversionDescriptor As Variant
'    Set c = New Collection
'
'    On Error Resume Next
'    For cpID = CodePageIdentifier.[_first_] To CodePageIdentifier.[_last_]
'        conversionDescriptor = CodePageToConvDescriptorName(CStr(cpID))
'        If conversionDescriptor <> "" Then
'            c.Add Item:=cpID, Key:=conversionDescriptor
'        End If
'        conversionDescriptor = ""
'    Next cpID
'
'    cpID = -1
'    For Each conversionDescriptor In CodePageToConvDescriptorName
'        cpID = c(conversionDescriptor)
'        If cpID = -1 Then c.Add Item:=-1, Key:=conversionDescriptor
'        cpID = -1
'    Next conversionDescriptor
'    On Error GoTo 0
'
'    Set ConvDescriptorNameToCodePage = c
'End Function

Private Function GetApiErrorNumber() As Long
    #If Mac Then
        CopyMemory GetApiErrorNumber, ByVal errno_location(), 4
    #Else
        GetApiErrorNumber = err.LastDllError 'GetLastError
    #End If
End Function

Private Function SetApiErrorNumber(ByVal errNumber As Long) As Long
    #If Mac Then
        CopyMemory ByVal errno_location(), errNumber, 4
    #Else
        SetLastError errNumber
    #End If
End Function

#If Mac = 0 Then
Public Function GetBstrFromWideStringPtr(ByVal lpwString As LongPtr) As String
    Dim Length As Long
    If lpwString Then Length = lstrlenW(lpwString)
    If Length Then
        GetBstrFromWideStringPtr = Space$(Length)
        CopyMemory ByVal StrPtr(GetBstrFromWideStringPtr), ByVal lpwString, Length * 2
    End If
End Function
#End If

'This function attempts to transcode 'str' from codepage 'fromCodePage' to
'codepage 'toCodePage' using the appropriate API functions on the platform.
'Calling it with 'raiseErrors = True' will raise an error if either:
'   - the string 'str' contains byte sequences that do not represent a valid
'     string of codepage 'fromCodePage', or
'   - the string contains codepoints that can not be represented in 'toCodePage'
'     and will lead to the insertion of a "default character".
'E.g.: Transcode("°", cpUTF_16, cpUs_ascii, True) will raise an error, because
'      "°" is not an ASCII character.
'Note that even calling the function with 'raiseErrors = True' doesn't guarantee
'that the conversion is reversible, because sometimes codepoints are replaced
'with more generic characters that aren't the default character (raise no error)
'E.g.:Decode(Transcode("³", cpUTF_16, cpUs_ascii, True), cpUs_ascii) returns "3"
Public Function Transcode(ByRef Str As String, _
                          ByVal fromCodePage As CodePageIdentifier, _
                          ByVal toCodePage As CodePageIdentifier, _
                 Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "Transcode"
    'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv.3.html
    #If Mac Then
        Dim InBytesLeft As LongPtr:  InBytesLeft = LenB(Str)
        Dim OutBytesLeft As LongPtr: OutBytesLeft = InBytesLeft * 4
        Dim Buffer As String:        Buffer = Space$(CLng(InBytesLeft) * 2)
        Dim InBuf As LongPtr:        InBuf = StrPtr(Str)
        Dim OutBuf As LongPtr:       OutBuf = StrPtr(Buffer)
        Dim cd As LongPtr: cd = GetConversionDescriptor(fromCodePage, toCodePage)
        Dim IrrevConvCount As Long
        Dim ReplacementChar As String

        Do While InBytesLeft > 0
            SetApiErrorNumber = 0
            IrrevConvCount = iconv(cd, InBuf, InBytesLeft, OutBuf, OutBytesLeft)

            If IrrevConvCount = -1 Then 'Error occurred
                If StrPtr(ReplacementChar) = 0 Then _
                    ReplacementChar = GetReplacementCharForCodePage(toCodePage)

                Select Case GetApiErrorNumber
                    Case MAC_API_ERR_EILSEQ
                        If raiseErrors Then err.Raise 5, methodName, _
                            "Input is invalid byte sequence of " & _
                            "CodePage " & fromCodePage

                        CopyMemory ByVal OutBuf, ReplacementChar(0), _
                                   LenB(ReplacementChar)
                        OutBuf = OutBuf + LenB(ReplacementChar)
                        OutBytesLeft = OutBytesLeft - LenB(ReplacementChar)
                        InBuf = InBuf + 1
                        InBytesLeft = InBytesLeft - 1
                    Case MAC_API_ERR_EINVAL
                        If raiseErrors Then err.Raise 5, methodName, _
                            "Input is incomplete byte sequence of" & _
                            "CodePage " & fromCodePage

                        CopyMemory ByVal OutBuf, ReplacementChar(0), _
                                   LenB(ReplacementChar)
                        OutBuf = OutBuf + OutBytesLeft
                        InBuf = InBuf + InBytesLeft
                        OutBytesLeft = 0
                        InBytesLeft = 0
                End Select
            End If
        Loop

        If IrrevConvCount > 0 And raiseErrors Then err.Raise 5, _
            methodName, "Default char would be used, encoding would be irreversible"

        Transcode = LeftB$(Buffer, LenB(Buffer) - CLng(OutBytesLeft))

        'These errors are bugs and should be raised even if raiseErrors = False:
        Select Case GetApiErrorNumber
            Case MAC_API_ERR_E2BIG
                err.Raise vbErrInternalError, methodName, _
                    "Output buffer overrun while transcoding from CodePage " _
                    & fromCodePage & " to CodePage " & toCodePage
            Case Is <> 0
                err.Raise vbErrInternalError, methodName, "Unknown error " & _
                    "occurred during transcoding with 'iconv'. API Error" & _
                    "Code: " & GetApiErrorNumber
        End Select
        If iconv_close(cd) <> 0 Then
            err.Raise vbErrInternalError, methodName, "Unknown error occurred" _
                & " when calling 'iconv_close'. API ErrorCode: " & _
                GetApiErrorNumber
        End If
    #Else
        If toCodePage = cpUTF_16 Then
            Transcode = Decode(Str, fromCodePage, raiseErrors)
        ElseIf fromCodePage = cpUTF_16 Then
            Transcode = Encode(Str, toCodePage, raiseErrors)
        Else
            Transcode = Encode(Decode(Str, fromCodePage, raiseErrors), _
                               toCodePage, raiseErrors)
        End If
    #End If
End Function

#If Mac Then
Private Function GetConversionDescriptor( _
                            ByVal fromCodePage As CodePageIdentifier, _
                            ByVal toCodePage As CodePageIdentifier) As LongPtr
    Dim ToCpCdName As String:   ToCpCdName = ConvDescriptorName(toCodePage)
    Dim FromCpCdName As String: FromCpCdName = ConvDescriptorName(fromCodePage)
    'Todo: potentially implement custom error numbers
    If LenB(ToCpCdName) = 0 Then err.Raise 5, methodName, _
        "No conversion descriptor name assigned to CodePage " & toCodePage
    If LenB(FromCpCdName) = 0 Then err.Raise 5, methodName, _
        "No conversion descriptor name assigned to CodePage " & fromCodePage

    SetApiErrorNumber = 0 'Clear previous errors
    GetConversionDescriptor = iconv_open(StrPtr(ToCpCdName), StrPtr(fromCodePage))

    If Not GetConversionDescriptor Then
        Select Case GetApiErrorNumber
            Case MAC_API_ERR_EINVAL
                err.Raise 5, methodName, "The conversion from CodePage " & _
                    fromCodePage & " to CodePage " & toCodePage & " is not " & _
                    "supported by the implementation of 'iconv' on this platform"
            Case Is <> 0
                err.Raise vbErrInternalError, methodName, "Unknown error " & _
                    "trying to create a conversion descriptor. API Error" & _
                    "Code: " & GetApiErrorNumber
        End Select
    End If
End Function
#End If

#If Mac Then
'On Mac, replacement character must be manually inserted when using iconv
Private Function GetReplacementCharForCodePage( _
                                    ByVal cpID As CodePageIdentifier) As String
    Static replacementChars As Collection
    If replacementChars Is Nothing Then replacementChars = New Collection
    On Error GoTo 0
    On Error Resume Next
    GetReplacementCharForCodePage = replacementChars(CStr(cpID))
    If err.Number = 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' U+FFFD (0xEF 0xBF 0xBD) in UTF-8
    Dim ReplacementCharUtf8() As Byte: ReDim ReplacementCharUtf8(0 To 2)
    ReplacementCharUtf8(0) = &HEF
    ReplacementCharUtf8(1) = &HBF
    ReplacementCharUtf8(2) = &HBD

    replacementChars.Add Transcode(CStr(ReplacementCharUtf8), cpUTF_8, cpID), _
                         CStr(cpID)
    GetReplacementCharForCodePage = replacementChars(CStr(cpID))
End Function
#End If

'This function tries to encode utf16leStr from vba-internal codepage UTF-16LE to
'codepage 'toCodePage' using the appropriate API functions on the platform.
'Calling it with 'raiseErrors = True' will raise an error if either:
'   - the string 'utf16leStr' contains byte sequences that do not represent a
'     valid UTF-16LE string, or
'   - the string contains codepoints that can not be represented in 'toCodePage'
'     and will lead to the insertion of a "default character".
'E.g.: Encode("°", cpUs_ascii, True) will raise an error, because
'      "°" is not an ASCII character.
'Note that even calling the function with 'raiseErrors = True' doesn't guarantee
'that the conversion is reversible, because sometimes codepoints are replaced
'with more generic characters that aren't the default character (raise no error)
'E.g.: Decode(Encode("³", cpUTF_16, cpUs_ascii, True), cpUs_ascii) returns "3"
Public Function Encode(ByRef utf16leStr As String, _
                       ByVal toCodePage As CodePageIdentifier, _
              Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "Encode"

    If toCodePage = cpUTF_16 Then err.Raise 5, methodName, _
        "Input string should already be UTF-16. Can't encode UTF-16 to UTF-16."

    If utf16leStr = vbNullString Then Exit Function
    #If Mac Then
        Encode = Transcode(utf16leStr, cpUTF_16, toCodePage, raiseErrors)
    #Else
        Dim ByteCount As Long
        Dim DwFlags As Long
        Dim UsedDefaultChar As Boolean
        Dim LpUsedDefaultChar As LongPtr
        If raiseErrors And CodePageAllowsQueryReversible(toCodePage) Then _
            LpUsedDefaultChar = VarPtr(UsedDefaultChar)

        If raiseErrors And CodePageAllowsFlags(toCodePage) Then _
            DwFlags = WC_ERR_INVALID_CHARS

        SetApiErrorNumber 0
        ByteCount = WideCharToMultiByte(toCodePage, DwFlags, StrPtr(utf16leStr), _
                                    Len(utf16leStr), 0, 0, 0, LpUsedDefaultChar)
        If ByteCount = 0 Then
            Select Case GetApiErrorNumber
                Case ERROR_NO_UNICODE_TRANSLATION
                    err.Raise 5, methodName, _
                        "Input is invalid byte sequence of CodePage " & cpUTF_16
                Case ERROR_INVALID_PARAMETER
                    err.Raise 5, methodName, _
                        "Conversion to CodePage " & toCodePage & " is not " & _
                        "supported by the API on this platform."
                Case ERROR_INSUFFICIENT_BUFFER, ERROR_INVALID_FLAGS
                    err.Raise vbErrInternalError, methodName, _
                        "Library implementation erroneous. API Error: " & _
                        GetApiErrorNumber
                Case Else
                    err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
            End Select
        End If

        If raiseErrors And UsedDefaultChar Then _
            err.Raise 5, methodName, "Default char would be used, encoding " & _
                "would be irreversible."

        Dim b() As Byte: ReDim b(0 To ByteCount - 1)
        Encode = b
        WideCharToMultiByte toCodePage, DwFlags, StrPtr(utf16leStr), _
                Len(utf16leStr), StrPtr(Encode), ByteCount, 0, LpUsedDefaultChar

        Select Case GetApiErrorNumber
            Case Is <> 0
                err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
        End Select
    #End If
End Function

'This function tries to decode 'str' from codepage 'fromCodePage' to the vba-
'internal codepage UTF-16LE using the appropriate API functions on the platform.
'Calling it with 'raiseErrors = True' will raise an error if the string 'str'
'contains byte sequences that does not represent a valid encoding in codepage
'fromCodePage.
'E.g.: If 'str' is an UTF-8 encoded string that was read from an external file
'      using 'Open' and 'Get', you can convert it to the VBA-internal UTF-16LE
'      like this:
'      Decode(str, cpUTF_8)
'      If you are afraid 'str' might contain invalid UTF-8 data, use it like so:
'      Decode(str, cpUTF_8, True)
'      The function will now raise an error if invalid UTF-8 data is encountered
Public Function Decode(ByRef Str As String, _
                       ByVal fromCodePage As CodePageIdentifier, _
              Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "Decode"

    If fromCodePage = cpUTF_16 Then err.Raise 5, methodName, _
        "VBA strings are UTF-16 by default. No need to decode string from UTF-16."

    If Str = vbNullString Then Exit Function
    #If Mac Then
        Decode = Transcode(Str, fromCodePage, cpUTF_16, raiseErrors)
    #Else
        Dim CharCount As Long
        Dim DwFlags As Long

        SetApiErrorNumber 0
        If raiseErrors And CodePageAllowsFlags(fromCodePage) Then _
            DwFlags = MB_ERR_INVALID_CHARS

        CharCount = MultiByteToWideChar(fromCodePage, DwFlags, StrPtr(Str), _
                                        LenB(Str), 0, 0)
        If CharCount = 0 Then
            Select Case GetApiErrorNumber
                Case ERROR_NO_UNICODE_TRANSLATION
                    err.Raise 5, methodName, _
                        "Input is invalid byte sequence of CodePage " & cpUTF_16
                Case ERROR_INVALID_PARAMETER
                    err.Raise 5, methodName, _
                        "Conversion from CodePage " & fromCodePage & " is not" _
                        & " supported by the API on this platform."
                Case ERROR_INSUFFICIENT_BUFFER, ERROR_INVALID_FLAGS
                    err.Raise vbErrInternalError, methodName, _
                        "Library implementation erroneous. API Error: " & _
                        GetApiErrorNumber
                Case Else
                    err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
            End Select
        End If

        Decode = Space$(CharCount)
        MultiByteToWideChar fromCodePage, DwFlags, StrPtr(Str), LenB(Str), _
                            StrPtr(Decode), CharCount

        Select Case GetApiErrorNumber
            Case Is <> 0
                err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
        End Select
    #End If
End Function

'Returns strings defined as hex literal as string
'Accepts the following formattings:
'   0xXXXXXX...
'   &HXXXXXX...
'   XXXXXX...
'Where:
'   - prefixes 0x and &H are case sensitive
'   - there's an even number of Xes, X = 0-9 or a-f or A-F (case insensitive)
'Raises error 5 if:
'   - Length is not even / partial bytes
'   - Invalid characters are found (outside prefix and 0-9 / a-f / A-F ranges)
'Examples:
'   - HexToString("0x610062006300") returns "abc"
'   - StrConv(HexToString("0x616263"), vbUnicode) returns "abc"
'   - HexToString("0x61626t") or HexToString("0x61626") both raise error 5
Public Function HexToString(ByRef hexStr As String) As String
    Const methodName As String = "HexToString"
    Const errPrefix As String = "Invalid Hex string literal. "
    Dim Size As Long: Size = Len(hexStr)

    If Size = 0 Then Exit Function
    If Size Mod 2 = 1 Then err.Raise 5, methodName, errPrefix & "Uneven length"

    Static nibbleMap(0 To 255) As Long 'Nibble: 0 to F. Byte: 00 to FF
    Static charMap(0 To 255) As String
    Dim i As Long

    If nibbleMap(0) = 0 Then
        For i = 0 To 255
            nibbleMap(i) = -256 'To force invalid character code
            charMap(i) = ChrB$(i)
        Next i
        For i = 0 To 9
            nibbleMap(Asc(CStr(i))) = i
        Next i
        For i = 10 To 15
            nibbleMap(i + 55) = i 'Asc("A") to Asc("F")
            nibbleMap(i + 87) = i 'Asc("a") to Asc("f")
        Next i
    End If

    Dim Prefix As String: Prefix = Left$(hexStr, 2)
    Dim StartPos As Long: StartPos = -4 * CLng(Prefix = "0x" Or Prefix = "&H")
    Dim b() As Byte:      b = hexStr
    Dim j As Long
    Dim CharCode As Long

    HexToString = MidB$(hexStr, 1, Size / 2 - Sgn(StartPos))
    For i = StartPos To UBound(b) Step 4
        j = j + 1
        CharCode = nibbleMap(b(i)) * &H10& + nibbleMap(b(i + 2))
        If CharCode < 0 Or b(i + 1) > 0 Or b(i + 3) > 0 Then
            err.Raise 5, methodName, errPrefix & "Expected a-f/A-F or 0-9"
        End If
        MidB$(HexToString, j, 1) = charMap(CharCode)
    Next i
End Function

'Converts the input string into a string of hex literals.
'e.g.: "abc" will be turned into "0x610062006300" (UTF-16LE)
'e.g.: StrConv("ABC", vbFromUnicode) will be turned into "0x414243"
Public Function StringToHex(ByRef s As String) As String
    Static map(0 To 255) As String
    Dim b() As Byte: b = s
    Dim i As Long

    If LenB(map(0)) = 0 Then
        For i = 0 To 255
            map(i) = Right$("0" & Hex$(i), 2)
        Next i
    End If

    StringToHex = Space$(LenB(s) * 2 + 2)
    Mid$(StringToHex, 1, 2) = "0x"

    For i = LBound(b) To UBound(b)
        Mid$(StringToHex, (i + 1) * 2 + 1, 2) = map(b(i))
    Next i
End Function

'Replaces all occurences of unicode characters outside the codePoint range
'defined by maxNonEscapedCharCode with literals of the following formats
'specified by `escapeFormat`:
' efPython = 1 ... \uXXXX \u00XXXXXX   (4 or 8 hex digits, 8 for chars outside BMP)
' efRust   = 2 ... \u{XXXX} \U{XXXXXX} (1 to 6 hex digits)
' efUPlus  = 4 ... u+XXXX u+XXXXXX     (4 or 6 hex digits)
' efMarkup = 8 ... &#ddddddd;          (1 to 7 decimal digits)
'Where:
'   - prefixes \u is case insensitive
'   - Xes are the digits of the codepoint in hexadecimal. (X = 0-9 or A-F/a-f)
'Note:
'   - Avoid u+XXXX syntax if string contains literals without delimiters as it
'     can be misinterpreted if adjacent to text starting with 0-9 or a-f.
'   - This function accepts all combinations of UnicodeEscapeFormats:
'     If called with, e.g. `escapeFormat = efRust Or efPython`, every character
'     in the scope will be escaped with in either format, efRust or efPython,
'     chosen at random for each replacement.
'   - If `escapeFormat` is set to efAll, it will replace every character in the
'     scope with a randomly chosen format of all available fotrmats.
'   - To escape every character, set `maxNonEscapedCharCode = -1`
Public Function EscapeUnicode(ByRef Str As String, _
                     Optional ByVal maxNonEscapedCharCode As Long = &HFF, _
                     Optional ByVal escapeFormat As UnicodeEscapeFormat _
                                                = efPython) As String
    Const methodName As String = "EscapeUnicode"
    If maxNonEscapedCharCode < -1 Then err.Raise 5, methodName, _
        "`maxNonEscapedCharCode` must be greater or equal -1."
    If escapeFormat < [_efMin] Or escapeFormat > [_efMax] Then _
        err.Raise 5, methodName, "Invalid escape type."
    If Len(Str) = 0 Then Exit Function
    Dim i As Long
    Dim j As Long:                j = 1
    Dim Result() As String:       ReDim Result(1 To Len(Str))
    Dim CopyChunkSize As Long
    Dim RndEscapeFormat As Boolean
    RndEscapeFormat = ((escapeFormat And (escapeFormat - 1)) <> 0) 'eFmt <> 2^n
    Dim NumescapeFormats As Long
    If RndEscapeFormat Then
        Dim EscapeFormats() As Long
        For i = 0 To (Log(efAll + 1) / Log(2)) - 1
            If 2 ^ i And escapeFormat Then
                ReDim Preserve EscapeFormats(0 To NumescapeFormats)
                EscapeFormats(NumescapeFormats) = 2 ^ i
                NumescapeFormats = NumescapeFormats + 1
            End If
        Next i
    End If
    For i = 1 To Len(Str)
        Dim Codepoint As Long: Codepoint = AscU(Mid$(Str, i, 2))
        If Codepoint > maxNonEscapedCharCode Then
            If CopyChunkSize > 0 Then
                Result(j) = Mid$(Str, i - CopyChunkSize, CopyChunkSize)
                CopyChunkSize = 0
                j = j + 1
            End If
            If RndEscapeFormat Then
                escapeFormat = EscapeFormats(Int(NumescapeFormats * Rnd))
            End If
            Select Case escapeFormat
                Case efPython
                    If Codepoint > &HFFFF& Then 'Outside BMP
                        Result(j) = "\u" & "00" & Right$("0" & Hex(Codepoint), 6)
                    Else 'BMP
                        Result(j) = "\u" & Right$("000" & Hex(Codepoint), 4)
                    End If
                Case efRust
                    Result(j) = "\u{" & Hex(Codepoint) & "}"
                Case efUPlus
                    If Codepoint < &H1000& Then
                        Result(j) = "u+" & Right$("000" & Hex(Codepoint), 4)
                    Else
                        Result(j) = "u+" & Hex(Codepoint)
                    End If
                Case efMarkup
                    Result(j) = "&#" & Codepoint & ";"
            End Select
            If RndEscapeFormat Then
                If Int(2 * Rnd) = 1 Then Result(j) = UCase(Result(j))
            End If
            j = j + 1
        Else
            If Codepoint < &H10000 Then
                CopyChunkSize = CopyChunkSize + 1
            Else
                CopyChunkSize = CopyChunkSize + 2
            End If
        End If
        If Codepoint > &HFFFF& Then i = i + 1
    Next i
    If CopyChunkSize > 0 Then _
        Result(j) = Mid$(Str, i - CopyChunkSize, CopyChunkSize)
    EscapeUnicode = Join(Result, "")
End Function

'Replaces all occurences of unicode literals
'Accepts the following formattings `escapeFormat`:
'   efPython = 1 ... \uXXXX \u000XXXXX    (4 or 8 hex digits, 8 for chars outside BMP)
'   efRust   = 2 ... \u{XXXX} \U{XXXXXXX} (1 to 6 hex digits)
'   efUPlus  = 4 ... u+XXXX u+XXXXXX      (4 or 6 hex digits)
'   efMarkup = 8 ... &#ddddddd;           (1 to 7 decimal digits)
'Where:
'   - prefixes \u is case insensitive
'   - Xes are the digits of the codepoint in hexadecimal. (X = 0-9 or A-F/a-f)
'Example:
'   - "abcd &#97;u+0062\U0063xy\u{64}", efAll returns "abcd abcxyd"
'Notes:
'   - Avoid u+XXXX syntax if string contains literals without delimiters as it
'     can be misinterpreted if adjacent to text starting with 0-9 or a-f.
'   - This function also accepts all combinations of UnicodeEscapeFormats:
'       E.g.:
'UnescapeUnicode("abcd &#97;u+0062\U0063xy\u{64}", efMarkup Or efRust)
'       will return:
'"abcd au+0062\U0063xyd"
'   - By default, this function will not invalidate UTF-16 strings if they are
'     currently valid, but this can happen if `allowSingleSurrogates = True`
'     E.g.: EscapeUnicode(ChrU(&HD801&, True)) returns "\uD801", but this string
'     can no longer be un-escaped with UnescapeUnicode because "\uD801"
'     represents a surrogate halve which is invalid unicode on its own.
'     So UnescapeUnicode("\uD801") returns "\uD801" again, unless called with
'     the optional parameter `allowSingleSurrogates = False` like this
'     `UnescapeUnicode("\uD801", , True)`. This will return invalid UTF-16.
Public Function UnescapeUnicode(ByRef Str As String, _
                       Optional ByVal escapeFormat As UnicodeEscapeFormat = efAll, _
                       Optional ByVal allowSingleSurrogates As Boolean = False) _
                                As String
    If escapeFormat < [_efMin] Or escapeFormat > [_efMax] Then
        err.Raise 5, "EscapeUnicode", "Invalid escape format"
    End If

    Dim escapes() As EscapeSequence: escapes = NewEscapes()
    Dim lb As Long: lb = LBound(escapes)
    Dim ub As Long: ub = UBound(escapes)
    Dim i As Long

    For i = lb To ub 'Find first signature for each wanted format
        With escapes(i)
            If escapeFormat And .ueFormat Then
                .buffPosition = InStr(1, Str, .ueSignature, vbBinaryCompare)
                .letSngSurrogate = allowSingleSurrogates
            End If
        End With
    Next i
    UnescapeUnicode = Str 'Allocate buffer
    
    Const posByte As Byte = &H80
    Const buffSize As Long = 1024
    Dim BuffSignaturePos(1 To buffSize) As Byte
    Dim BuffFormat(1 To buffSize) As UnicodeEscapeFormat
    Dim BuffEscIndex(1 To buffSize) As Long
    Dim PosOffset As Long
    Dim Diff As Long
    Dim HighSur As Long
    Dim LowSur As Long
    Dim RemainingLen As Long: RemainingLen = Len(Str)
    Dim PosChar As String:    PosChar = ChrB$(posByte)
    Dim OutPos As Long:       OutPos = 1
    Dim InPos As Long:        InPos = 1

    Do
        Dim UpperLimit As Long: UpperLimit = PosOffset + buffSize
        For i = lb To ub 'Find all signatures within buffer size
            With escapes(i)
                Do Until .buffPosition = 0 Or .buffPosition > UpperLimit
                    .buffPosition = .buffPosition - PosOffset
                    BuffSignaturePos(.buffPosition) = posByte
                    BuffFormat(.buffPosition) = .ueFormat
                    BuffEscIndex(.buffPosition) = i
                    .buffPosition = .buffPosition + .sigSize + PosOffset
                    .buffPosition = InStr(.buffPosition, Str, .ueSignature)
                Loop
            End With
        Next i

        Dim Temp As String:  Temp = BuffSignaturePos
        Dim NextPos As Long: NextPos = InStrB(1, Temp, PosChar)

        Do Until NextPos = 0 'Unescape all found signatures from buffer
            i = BuffEscIndex(NextPos)
            escapes(i).currPosition = NextPos + PosOffset
            Select Case BuffFormat(NextPos)
                Case efPython: TryPythonEscape escapes(i), Str
                Case efRust:   TryRustEscape escapes(i), Str
                Case efUPlus:  TryUPlusEscape escapes(i), Str
                Case efMarkup: TryMarkupEscape escapes(i), Str
            End Select
            With escapes(i)
                If .unEscSize > 0 Then
                    Diff = .currPosition - InPos
                    If OutPos > 1 Then
                        Mid$(UnescapeUnicode, OutPos) = Mid$(Str, InPos, Diff)
                    End If
                    OutPos = OutPos + Diff
                    If .unEscSize = 1 Then
                        Mid$(UnescapeUnicode, OutPos) = ChrW$(.Codepoint)
                    Else
                        .Codepoint = .Codepoint - &H10000
                        HighSur = &HD800& Or (.Codepoint \ &H400&)
                        LowSur = &HDC00& Or (.Codepoint And &H3FF&)
                        Mid$(UnescapeUnicode, OutPos) = ChrW$(HighSur)
                        Mid$(UnescapeUnicode, OutPos + 1) = ChrW$(LowSur)
                    End If
                    OutPos = OutPos + .unEscSize
                    InPos = .currPosition + .escSize
                    NextPos = NextPos + .escSize - .sigSize
                End If
                NextPos = InStrB(NextPos + .sigSize, Temp, PosChar)
            End With
        Loop
        RemainingLen = RemainingLen - buffSize
        PosOffset = PosOffset + buffSize
        Erase BuffSignaturePos
    Loop Until RemainingLen < 1
    
    If OutPos > 1 Then
        Diff = Len(Str) - InPos + 1
        If Diff > 0 Then
            Mid$(UnescapeUnicode, OutPos, Diff) = Mid$(Str, InPos, Diff)
        End If
        UnescapeUnicode = Left$(UnescapeUnicode, OutPos + Diff - 1)
    End If
End Function
Private Function NewEscapes() As EscapeSequence()
    Static escapes(0 To 6) As EscapeSequence
    If escapes(0).ueFormat = [_efNone] Then
        InitEscape escapes(0), efPython, "\U"
        InitEscape escapes(1), efPython, "\u"
        InitEscape escapes(2), efRust, "\U{"
        InitEscape escapes(3), efRust, "\u{"
        InitEscape escapes(4), efUPlus, "U+"
        InitEscape escapes(5), efUPlus, "u+"
        InitEscape escapes(6), efMarkup, "&#"
    End If
    NewEscapes = escapes
End Function
Private Sub InitEscape(ByRef escape As EscapeSequence, _
                       ByVal ueFormat As UnicodeEscapeFormat, _
                       ByRef ueSignature As String)
    With escape
        .ueFormat = ueFormat
        .ueSignature = ueSignature
        .sigSize = Len(ueSignature)
    End With
End Sub

Private Sub TryPythonEscape(ByRef escape As EscapeSequence, ByRef Str As String)
    Const H As String = "[0-9A-Fa-f]"
    Const PYTHON_ESCAPE_PATTERN_NOT_BMP = "00[01]" & H & H & H & H & H
    Const PYTHON_ESCAPE_PATTERN_BMP As String = H & H & H & H & "*"
    Dim PotentialEscape As String

    With escape
        .unEscSize = 0
        PotentialEscape = Mid$(Str, .currPosition + 2, 8) 'Exclude leading \[Uu]
        If PotentialEscape Like PYTHON_ESCAPE_PATTERN_NOT_BMP Then
            .escSize = 10 '\[Uu]00[01]HHHHH
            .Codepoint = CLng("&H" & PotentialEscape) 'No extra Mid$ needed
            If .Codepoint < &H10000 Then
                If IsValidBMP(.Codepoint, .letSngSurrogate) Then
                    .unEscSize = 1
                    Exit Sub
                End If
            ElseIf .Codepoint < &H110000 Then
                .unEscSize = 2
                Exit Sub
            End If
        End If
        If PotentialEscape Like PYTHON_ESCAPE_PATTERN_BMP Then
            .escSize = 6 '\[Uu]HHHH
            .Codepoint = CLng("&H" & Left$(PotentialEscape, 4))
            If IsValidBMP(.Codepoint, .letSngSurrogate) Then .unEscSize = 1
        End If
    End With
End Sub
Private Function IsValidBMP(ByVal Codepoint As Long, _
                            ByVal letSingleSurrogate As Boolean) As Boolean
    IsValidBMP = (Codepoint < &HD800& Or Codepoint >= &HE000& Or letSingleSurrogate)
End Function

Private Sub TryRustEscape(ByRef escape As EscapeSequence, ByRef Str As String)
    Static rustEscPattern(1 To 6) As String
    Static isPatternInit As Boolean
    Dim PotentialEscape As String
    Dim NextBrace As Long
    
    If Not isPatternInit Then
        Dim i As Long
        rustEscPattern(1) = "[0-9A-Fa-f]}*"
        For i = 2 To 6
            rustEscPattern(i) = "[0-9A-Fa-f]" & rustEscPattern(i - 1)
        Next i
        isPatternInit = True
    End If
    With escape
        .unEscSize = 0
        PotentialEscape = Mid$(Str, .currPosition + 3, 7) 'Exclude leading \[Uu]{
        NextBrace = InStr(2, PotentialEscape, "}", vbBinaryCompare)
        
        If NextBrace = 0 Then Exit Sub
        If Not PotentialEscape Like rustEscPattern(NextBrace - 1) Then Exit Sub
        
        .Codepoint = CLng("&H" & Left$(PotentialEscape, NextBrace - 1))
        .escSize = NextBrace + 3
        If .Codepoint < &H10000 Then
            If IsValidBMP(.Codepoint, .letSngSurrogate) Then .unEscSize = 1
        ElseIf .Codepoint < &H110000 Then
            .unEscSize = 2
        End If
    End With
End Sub

Private Sub TryUPlusEscape(ByRef escape As EscapeSequence, _
                           ByRef Str As String)
    Const H As String = "[0-9A-Fa-f]"
    Const UPLUS_ESCAPE_PATTERN_4_DIGITS = H & H & H & H & "*"
    Const UPLUS_ESCAPE_PATTERN_5_DIGITS = H & H & H & H & H & "*"
    Const UPLUS_ESCAPE_PATTERN_6_DIGITS = H & H & H & H & H & H
    Dim PotentialEscape As String
    
    With escape
        .unEscSize = 0
        PotentialEscape = Mid$(Str, .currPosition + 2, 6) 'Exclude leading [Uu]+
        If PotentialEscape Like UPLUS_ESCAPE_PATTERN_6_DIGITS Then
            .escSize = 8
            .Codepoint = CLng("&H" & PotentialEscape)
            If .Codepoint < &H10000 Then
                If IsValidBMP(.Codepoint, .letSngSurrogate) Then
                    .unEscSize = 1
                    Exit Sub
                End If
            ElseIf .Codepoint < &H110000 Then
                .unEscSize = 2
                Exit Sub
            End If
        End If
        If PotentialEscape Like UPLUS_ESCAPE_PATTERN_5_DIGITS Then
            .escSize = 7
            .Codepoint = CLng("&H" & Left$(PotentialEscape, 5))
            If .Codepoint < &H10000 Then
                If IsValidBMP(.Codepoint, .letSngSurrogate) Then
                    .unEscSize = 1
                    Exit Sub
                End If
            Else
                .unEscSize = 2
                Exit Sub
            End If
        End If
        If PotentialEscape Like UPLUS_ESCAPE_PATTERN_4_DIGITS Then
            .escSize = 6
            .Codepoint = CLng("&H" & Left$(PotentialEscape, 4))
            If IsValidBMP(.Codepoint, .letSngSurrogate) Then .unEscSize = 1
        End If
    End With
End Sub
Private Sub TryMarkupEscape(ByRef escape As EscapeSequence, _
                            ByRef Str As String)
    Static mEscPattern(1 To 7) As String
    Static isPatternInit As Boolean
    Dim PotentialEscape As String
    Dim NextSemicolon As Long
    
    If Not isPatternInit Then
        Dim i As Long
        For i = 1 To 6
            mEscPattern(i) = String$(i, "#") & ";*"
        Next i
        mEscPattern(7) = "1######;"
        isPatternInit = True
    End If
    With escape
        .unEscSize = 0
        PotentialEscape = Mid$(Str, .currPosition + 2, 8) 'Exclude leading &[#]
        NextSemicolon = InStr(2, PotentialEscape, ";", vbBinaryCompare)
        
        If NextSemicolon = 0 Then Exit Sub
        If Not PotentialEscape Like mEscPattern(NextSemicolon - 1) Then Exit Sub
        
        .Codepoint = CLng(Left$(PotentialEscape, NextSemicolon - 1))
        .escSize = NextSemicolon + 2
        If .Codepoint < &H10000 Then
            If IsValidBMP(.Codepoint, .letSngSurrogate) Then .unEscSize = 1
        ElseIf .Codepoint < &H110000 Then
            .unEscSize = 2
        End If
    End With
End Sub

'Returns the given unicode codepoint as standard VBA UTF-16LE string
Public Function ChrU(ByVal Codepoint As Long, _
             Optional ByVal allowSingleSurrogates As Boolean = False) As String
    Const methodName As String = "ChrU"
    Static st As TwoCharTemplate
    Static lt As LongTemplate

    If Codepoint < &H8000 Then err.Raise 5, methodName, "Codepoint < -32768"
    If Codepoint < 0 Then Codepoint = Codepoint And &HFFFF& 'Incase of uInt input

    If Codepoint < &HD800& Then
        ChrU = ChrW$(Codepoint)
    ElseIf Codepoint < &HE000& And Not allowSingleSurrogates Then
        err.Raise 5, methodName, "Range reserved for surrogate pairs"
    ElseIf Codepoint < &H10000 Then
        ChrU = ChrW$(Codepoint)
    ElseIf Codepoint < &H110000 Then
        lt.l = (&HD800& Or (Codepoint \ &H400& - &H40&)) _
            Or (&HDC00 Or (Codepoint And &H3FF&)) * &H10000 '&HDC00 with no &
        LSet st = lt
        ChrU = st.s
    Else
        err.Raise 5, methodName, "Codepoint outside of valid Unicode range."
    End If
End Function

'Returns a given characters unicode codepoint as long.
'Note: One unicode character can consist of two VBA "characters", a so-called
'      "surrogate pair" (input string of length 2, so Len(char) = 2!)
Public Function AscU(ByRef Char As String) As Long
    AscU = AscW(Char) And &HFFFF&
    If Len(Char) > 1 Then
        Dim Lo As Long: Lo = AscW(Mid$(Char, 2, 1)) And &HFFFF&
        If &HDC00& > Lo Or Lo > &HDFFF& Then Exit Function
        AscU = (AscU - &HD800&) * &H400& + (Lo - &HDC00&) + &H10000
    End If
End Function

'Function transcoding a VBA-native UTF-16LE encoded string to an ASCII string
'Note: Information will be lost for codepoints > 127!
Public Function EncodeASCII(ByRef utf16leStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim Utf16le() As Byte: Utf16le = utf16leStr
    Dim Ascii() As Byte

    ReDim Ascii(1 To Len(utf16leStr))
    For i = LBound(Ascii) To UBound(Ascii)
        If Utf16le(j) < 128 And Utf16le(j + 1) = 0 Then
            Ascii(i) = Utf16le(j)
            j = j + 2
        Else
            Ascii(i) = &H3F 'Chr(&H3F) = "?"
            j = j + 2
        End If
    Next i
    EncodeASCII = Ascii
End Function

'Function transcoding an ASCII encoded string to the VBA-native UTF-16LE
Public Function DecodeASCII(ByRef asciiStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim Ascii() As Byte:   Ascii = asciiStr
    Dim Utf16le() As Byte: ReDim Utf16le(0 To LenB(asciiStr) * 2 - 1)

    For i = LBound(Ascii) To UBound(Ascii)
        Utf16le(j) = Ascii(i)
        j = j + 2
    Next i
    DecodeASCII = Utf16le
End Function

'Function transcoding a VBA-native UTF-16LE encoded string to an ANSI string
'Note: Information will be lost for codepoints > 255!
Public Function EncodeANSI(ByRef utf16leStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim Utf16le() As Byte: Utf16le = utf16leStr
    Dim Ansi() As Byte

    ReDim Ansi(1 To Len(utf16leStr))
    For i = LBound(Ansi) To UBound(Ansi)
        If Utf16le(j + 1) = 0 Then
            Ansi(i) = Utf16le(j)
            j = j + 2
        Else
            Ansi(i) = &H3F 'Chr(&H3F) = "?"
            j = j + 2
        End If
    Next i
    EncodeANSI = Ansi
End Function

'Function transcoding an ANSI encoded string to the VBA-native UTF-16LE
Public Function DecodeANSI(ByRef ansiStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim Ansi() As Byte:    Ansi = ansiStr
    Dim Utf16le() As Byte: ReDim Utf16le(0 To LenB(ansiStr) * 2 - 1)

    For i = LBound(Ansi) To UBound(Ansi)
        Utf16le(j) = Ansi(i)
        j = j + 2
    Next i
    DecodeANSI = Utf16le
End Function

'Function transcoding an VBA-native UTF-16LE encoded string to UTF-8
Public Function EncodeUTF8(ByRef utf16leStr As String, _
                  Optional ByVal raiseErrors As Boolean = False) _
                                  As String
    Const methodName As String = "EncodeUTF8native"
    Dim Codepoint As Long
    Dim LowSurrogate As Long
    Dim i As Long:            i = 1
    Dim j As Long:            j = 0
    Dim Utf8() As Byte:       ReDim Utf8(Len(utf16leStr) * 4 - 1)

    Do While i <= Len(utf16leStr)
        Codepoint = AscW(Mid$(utf16leStr, i, 1)) And &HFFFF&

        If Codepoint >= &HD800& And Codepoint <= &HDBFF& Then 'high surrogate
            LowSurrogate = AscW(Mid$(utf16leStr, i + 1, 1)) And &HFFFF&

            If &HDC00& <= LowSurrogate And LowSurrogate <= &HDFFF& Then
                Codepoint = (Codepoint - &HD800&) * &H400& + _
                            (LowSurrogate - &HDC00&) + &H10000
                i = i + 1
            Else
                If raiseErrors Then _
                    err.Raise 5, methodName, _
                        "Invalid Unicode codepoint. (Lonely high surrogate)"
                Codepoint = &HFFFD&
            End If
        End If

        If Codepoint < &H80& Then
            Utf8(j) = Codepoint
            j = j + 1
        ElseIf Codepoint < &H800& Then
            Utf8(j) = &HC0& Or ((Codepoint And &H7C0&) \ &H40&)
            Utf8(j + 1) = &H80& Or (Codepoint And &H3F&)
            j = j + 2
        ElseIf Codepoint < &HDC00 Then
            Utf8(j) = &HE0& Or ((Codepoint And &HF000&) \ &H1000&)
            Utf8(j + 1) = &H80& Or ((Codepoint And &HFC0&) \ &H40&)
            Utf8(j + 2) = &H80& Or (Codepoint And &H3F&)
            j = j + 3
        ElseIf Codepoint < &HE000 Then
            If raiseErrors Then _
                err.Raise 5, methodName, _
                    "Invalid Unicode codepoint. (Lonely low surrogate)"
            Codepoint = &HFFFD&
        ElseIf Codepoint < &H10000 Then
            Utf8(j) = &HE0& Or ((Codepoint And &HF000&) \ &H1000&)
            Utf8(j + 1) = &H80& Or ((Codepoint And &HFC0&) \ &H40&)
            Utf8(j + 2) = &H80& Or (Codepoint And &H3F&)
            j = j + 3
        Else
            Utf8(j) = &HF0& Or ((Codepoint And &H1C0000) \ &H40000)
            Utf8(j + 1) = &H80& Or ((Codepoint And &H3F000) \ &H1000&)
            Utf8(j + 2) = &H80& Or ((Codepoint And &HFC0&) \ &H40&)
            Utf8(j + 3) = &H80& Or (Codepoint And &H3F&)
            j = j + 4
        End If

        i = i + 1
    Loop
    EncodeUTF8 = MidB$(Utf8, 1, j)
End Function

'Function transcoding an UTF-8 encoded string to the VBA-native UTF-16LE
'Function transcoding an VBA-native UTF-16LE encoded string to UTF-8
Public Function DecodeUTF8(ByRef utf8Str As String, _
                  Optional ByVal raiseErrors As Boolean = False) As String

    Const methodName As String = "DecodeUTF8native"
    Dim i As Long
    Dim NumBytesOfCodePoint As Byte

    Static numBytesOfCodePoints(0 To 255) As Byte
    Static mask(2 To 4) As Long
    Static minCp(2 To 4) As Long

    If numBytesOfCodePoints(0) = 0 Then
        For i = &H0& To &H7F&: numBytesOfCodePoints(i) = 1: Next i '0xxxxxxx
        '110xxxxx - C0 and C1 are invalid (overlong encoding)
        For i = &HC2& To &HDF&: numBytesOfCodePoints(i) = 2: Next i
        For i = &HE0& To &HEF&: numBytesOfCodePoints(i) = 3: Next i '1110xxxx
       '11110xxx - 11110100, 11110101+ (= &HF5+) outside of valid Unicode range
        For i = &HF0& To &HF4&: numBytesOfCodePoints(i) = 4: Next i
        For i = 2 To 4: mask(i) = (2 ^ (7 - i) - 1): Next i
        minCp(2) = &H80&: minCp(3) = &H800&: minCp(4) = &H10000
    End If

    Dim Codepoint As Long
    Dim CurrByte As Byte
    Dim Utf8() As Byte:  Utf8 = utf8Str
    Dim Utf16() As Byte: ReDim Utf16(0 To (UBound(Utf8) - LBound(Utf8) + 1) * 2)
    Dim j As Long:       j = 0
    Dim k As Long

    i = LBound(Utf8)
    Do While i <= UBound(Utf8)
        Codepoint = Utf8(i)
        NumBytesOfCodePoint = numBytesOfCodePoints(Codepoint)

        If NumBytesOfCodePoint = 0 Then
            If raiseErrors Then err.Raise 5, methodName, "Invalid byte"
            GoTo insertErrChar
        ElseIf NumBytesOfCodePoint = 1 Then
            Utf16(j) = Codepoint
            j = j + 2
        ElseIf i + NumBytesOfCodePoint - 1 > UBound(Utf8) Then
            If raiseErrors Then err.Raise 5, methodName, _
                    "Incomplete UTF-8 codepoint at end of string."
            GoTo insertErrChar
        Else
            Codepoint = Utf8(i) And mask(NumBytesOfCodePoint)

            For k = 1 To NumBytesOfCodePoint - 1
                CurrByte = Utf8(i + k)

                If (CurrByte And &HC0&) = &H80& Then
                    Codepoint = (Codepoint * &H40&) + (CurrByte And &H3F)
                Else
                    If raiseErrors Then _
                        err.Raise 5, methodName, "Invalid continuation byte"
                    GoTo insertErrChar
                End If
            Next k
            'Convert the Unicode codepoint to UTF-16LE bytes
            If Codepoint < minCp(NumBytesOfCodePoint) Then
                If raiseErrors Then err.Raise 5, methodName, "Overlong encoding"
                GoTo insertErrChar
            ElseIf Codepoint < &HD800& Then
                Utf16(j) = CByte(Codepoint And &HFF&)
                Utf16(j + 1) = CByte(Codepoint \ &H100&)
                j = j + 2
            ElseIf Codepoint < &HE000& Then
                If raiseErrors Then err.Raise 5, methodName, _
                "Invalid Unicode codepoint.(Range reserved for surrogate pairs)"
                GoTo insertErrChar
            ElseIf Codepoint < &H10000 Then
                If Codepoint = &HFEFF& Then GoTo nextCp '(BOM - will be ignored)
                Utf16(j) = Codepoint And &HFF&
                Utf16(j + 1) = Codepoint \ &H100&
                j = j + 2
            ElseIf Codepoint < &H110000 Then 'Calculate surrogate pair
                Dim m As Long:           m = Codepoint - &H10000
                Dim LoSurrogate As Long: LoSurrogate = &HDC00& Or (m And &H3FF)
                Dim HiSurrogate As Long: HiSurrogate = &HD800& Or (m \ &H400&)

                Utf16(j) = HiSurrogate And &HFF&
                Utf16(j + 1) = HiSurrogate \ &H100&
                Utf16(j + 2) = LoSurrogate And &HFF&
                Utf16(j + 3) = LoSurrogate \ &H100&
                j = j + 4
            Else
                If raiseErrors Then err.Raise 5, methodName, _
                        "Codepoint outside of valid Unicode range"
insertErrChar:  Utf16(j) = &HFD
                Utf16(j + 1) = &HFF
                j = j + 2

                If NumBytesOfCodePoint = 0 Then NumBytesOfCodePoint = 1
            End If
        End If
nextCp: i = i + NumBytesOfCodePoint 'Move to the next UTF-8 codepoint
    Loop
    DecodeUTF8 = MidB$(Utf16, 1, j)
End Function

#If Mac = 0 Then
'Transcoding a VBA-native UTF-16LE encoded string to UTF-8 using ADODB.Stream
'Much faster than EncodeUTF8native, but only available on Windows
Public Function EncodeUTF8usingAdodbStream(ByRef utf16leStr As String) _
                                            As String
    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText
        .CharSet = "utf-8"
        .Open
        .WriteText utf16leStr
        .Position = 0
        .Type = 1 ' adTypeBinary
        .Position = 3 ' Skip BOM (Byte Order Mark)
        EncodeUTF8usingAdodbStream = .Read
        .Close
    End With
End Function

'Transcoding an UTF-8 encoded string to VBA-native UTF-16LE using ADODB.Stream
'Faster than DeocdeUTF8native for some strings but only available on Windows
'Warning: This function performs extremely slow for strings bigger than ~5MB
Public Function DecodeUTF8usingAdodbStream(ByRef utf8Str As String) As String
    Dim b() As Byte: b = utf8Str
    With CreateObject("ADODB.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .Write b
        .Position = 0
        .Type = 2 ' adTypeText
        .CharSet = "utf-8"
        DecodeUTF8usingAdodbStream = .ReadText
        .Close
    End With
End Function
#End If

'Function transcoding an VBA-native UTF-16LE encoded string to UTF-32
Public Function EncodeUTF32LE(ByRef utf16leStr As String, _
                     Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "EncodeUTF32LE"

    If utf16leStr = "" Then Exit Function

    Dim Codepoint As Long
    Dim LowSurrogate As Long
    Dim Utf32() As Byte:      ReDim Utf32(Len(utf16leStr) * 4 - 1)
    Dim i As Long:            i = 1
    Dim j As Long:            j = 0

    Do While i <= Len(utf16leStr)
        Codepoint = AscW(Mid$(utf16leStr, i, 1)) And &HFFFF&

        If Codepoint >= &HD800& And Codepoint <= &HDBFF& Then 'high surrogate
            LowSurrogate = AscW(Mid$(utf16leStr, i + 1, 1)) And &HFFFF&

            If &HDC00& <= LowSurrogate And LowSurrogate <= &HDFFF& Then
                Codepoint = (Codepoint - &HD800&) * &H400& + _
                            (LowSurrogate - &HDC00&) + &H10000
                i = i + 1
            Else
                If raiseErrors Then err.Raise 5, methodName, _
                    "Invalid Unicode codepoint. (Lonely high surrogate)"
                Codepoint = &HFFFD&
            End If
        End If

        If Codepoint >= &HD800& And Codepoint < &HE000& Then
            If raiseErrors Then err.Raise 5, methodName, _
                "Invalid Unicode codepoint. (Lonely low surrogate)"
            Codepoint = &HFFFD&
        ElseIf Codepoint > &H10FFFF Then
            If raiseErrors Then err.Raise 5, methodName, _
                "Codepoint outside of valid Unicode range"
            Codepoint = &HFFFD&
        End If

        Utf32(j) = Codepoint And &HFF&
        Utf32(j + 1) = (Codepoint \ &H100&) And &HFF&
        Utf32(j + 2) = (Codepoint \ &H10000) And &HFF&
        i = i + 1: j = j + 4
    Loop
    EncodeUTF32LE = MidB$(Utf32, 1, j)
End Function

'Function transcoding an UTF-32 encoded string to the VBA-native UTF-16LE
Public Function DecodeUTF32LE(ByRef utf32str As String, _
                     Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "DecodeUTF32LE"

    If utf32str = "" Then Exit Function

    Dim Codepoint As Long
    Dim Utf32() As Byte:   Utf32 = utf32str
    Dim Utf16() As Byte:   ReDim Utf16(LBound(Utf32) To UBound(Utf32))
    Dim i As Long: i = LBound(Utf32)
    Dim j As Long: j = i

    Do While i < UBound(Utf32)
        If Utf32(i + 2) = 0 And Utf32(i + 3) = 0 Then
            Utf16(j) = Utf32(i): Utf16(j + 1) = Utf32(i + 1): j = j + 2
        Else
            If Utf32(i + 3) <> 0 Then
                If raiseErrors Then _
                    err.Raise 5, methodName, _
                    "Codepoint outside of valid Unicode range"
                Codepoint = &HFFFD&
            Else
                Codepoint = Utf32(i + 2) * &H10000 + _
                            Utf32(i + 1) * &H100& + Utf32(i)
                If Codepoint >= &HD800& And Codepoint < &HE000& Then
                    If raiseErrors Then _
                        err.Raise 5, methodName, _
                        "Invalid Unicode codepoint. " & _
                        "(Range reserved for surrogate pairs)"
                    Codepoint = &HFFFD&
                ElseIf Codepoint > &H10FFFF Then
                    If raiseErrors Then _
                        err.Raise 5, methodName, _
                        "Codepoint outside of valid Unicode range"
                    Codepoint = &HFFFD&
                End If
            End If

            Dim n As Long:             n = Codepoint - &H10000
            Dim HighSurrogate As Long: HighSurrogate = &HD800& Or (n \ &H400&)
            Dim LowSurrogate As Long:  LowSurrogate = &HDC00& Or (n And &H3FF)

            Utf16(j) = HighSurrogate And &HFF&
            Utf16(j + 1) = HighSurrogate \ &H100&
            Utf16(j + 2) = LowSurrogate And &HFF&
            Utf16(j + 3) = LowSurrogate \ &H100&
            j = j + 4
        End If
        i = i + 4
    Loop
    DecodeUTF32LE = MidB$(Utf16, 1, j)
End Function

'Returns a UTF-16 string containing all alphanumeric characters randomly equally
'distributed. (0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz)
Public Function RandomStringAlphanumeric(ByVal Length As Long) As String
    Const methodName As String = "RandomStringAlphanumeric"
    Const INKL_CHARS As String = _
        "01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Static Chars() As Byte
    Static numPossChars As Long
    Static isInitialized As Boolean
    If Not isInitialized Then
        Chars = StrConv(INKL_CHARS, vbFromUnicode)
        numPossChars = UBound(Chars) - LBound(Chars) + 1
        isInitialized = True
    End If
    
    If Length = 0 Then Exit Function
    If Length < 0 Then err.Raise 5, methodName, "Length must be >= 0"
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)
    Dim i As Long
    For i = 0 To Length * 2 - 1 Step 2
        b(i) = Chars(Int(Rnd * numPossChars))
    Next i
    RandomStringAlphanumeric = b
End Function

'Returns a UTF-16 string containing all ASCII characters equally,
'randomly distributed.
Public Function RandomStringASCII(ByVal Length As Long) As String
    Const methodName As String = "RandomStringASCII"
    Const MAX_ASC As Long = &H7F&
    If Length = 0 Then Exit Function
    If Length < 0 Then err.Raise 5, methodName, "Length must be >= 0"
    Dim i As Long
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)
    For i = 0 To Length * 2 - 1 Step 2
        b(i) = Int(MAX_ASC * Rnd) + 1
    Next i
    RandomStringASCII = b
End Function

'Function returning a UTF-16 string containing all characters from the BMP
'(Basic Multilingual Plane, so from all 2 byte UTF-16 chars) equally, randomly
'distributed. Excludes surrogate range and BOM.
Public Function RandomStringBMP(ByVal Length As Long) As String
    Const methodName As String = "RandomStringBMP"
    Const MAX_UINT As Long = &HFFFF&
    If Length = 0 Then Exit Function
    If Length < 0 Then err.Raise 5, methodName, "Length must be >= 0"

    Dim i As Long
    Dim Char As Long
    Dim b() As Byte:  ReDim b(0 To Length * 2 - 1)

    For i = 0 To Length * 2 - 1 Step 2
        Do
            Char = Int(MAX_UINT * Rnd) + 1
        Loop Until (Char < &HD800& Or Char > &HDFFF&) _
               And (Char <> &HFEFF&)
        b(i) = Char And &HFF
        b(i + 1) = Char \ &H100& And &HFF
    Next i
    RandomStringBMP = b
End Function

'Returns a UTF-16 string containing all valid unicode characters equally,
'randomly distributed. Excludes surrogate range and BOM.
'Length in UTF-16 codepoints, (Len(result) = length)
Public Function RandomStringUnicode(ByVal Length As Long) As String
    Const methodName As String = "RandomStringUnicode"
    Const MAX_UNICODE As Long = &H10FFFF
    If Length = 0 Then Exit Function
    If Length < 0 Then err.Raise 5, methodName, "Length must be >= 0"

    Dim i As Long
    Dim Char As Long
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)

    If Length > 1 Then
        For i = 0 To Length * 2 - 3 Step 2
            Do
                Char = Int(MAX_UNICODE * Rnd) + 1
            Loop Until (Char < &HD800& Or Char > &HDFFF&) _
                   And (Char <> &HFEFF&)
            If Char < &H10000 Then
                b(i) = Char And &HFF
                b(i + 1) = Char \ &H100& And &HFF
            Else
                Dim m As Long: m = Char - &H10000
                Dim HighSurrogate As Long: HighSurrogate = &HD800& + (m \ &H400&)
                Dim LowSurrogate As Long: LowSurrogate = &HDC00& + (m And &H3FF)
                b(i) = HighSurrogate And &HFF&
                b(i + 1) = HighSurrogate \ &H100&
                i = i + 2
                b(i) = LowSurrogate And &HFF&
                b(i + 1) = LowSurrogate \ &H100&
            End If
        Next i
    End If
    RandomStringUnicode = b
    
    Const MAX_UINT As Long = &HFFFF&
    If CInt(b(UBound(b) - 1)) + b(UBound(b)) = 0 Then
        Do
            Char = Int(MAX_UINT * Rnd) + 1
        Loop Until (Char < &HD800& Or Char > &HDFFF&) _
               And (Char <> &HFEFF&)
        Mid$(RandomStringUnicode, Len(RandomStringUnicode), 1) = ChrW(Char)
    End If
End Function

'Returns a string containing random byte data
Public Function RandomBytes(ByVal numBytes As Long) As String
    Const methodName As String = "RandomBytes"
    If numBytes = 0 Then Exit Function
    If numBytes < 0 Then err.Raise 5, methodName, "numBytes must be >= 0"

    Dim Bytes() As Byte: ReDim Bytes(0 To numBytes - 1)
    Dim i As Long
    For i = 0 To numBytes - 1
        Bytes(i) = Int(Rnd * &H100)
    Next i
    RandomBytes = Bytes
End Function

'Returns a UTF-16 string containing random characters from the codepoint range
'between 'minCodepoint' and 'maxCodepoint'.
'E.g.: RandomString(10, 48, 57) will return a string of length 100 containing
'      all the digit characters randomly, e.g. "3239107914"
Public Function RandomString(ByVal Length As Long, _
                    Optional ByVal minCodepoint As Long = 1, _
                    Optional ByVal maxCodepoint As Long = &H10FFFF) As String
    Const methodName As String = "RandomString"
    Const MAX_UNICODE As Long = &H10FFFF
    Const MAX_UINT As Long = &HFFFF&
    If Length = 0 Then Exit Function
    If Length < 0 Then err.Raise 5, methodName, "Length must be >= 0"
    If maxCodepoint > MAX_UNICODE Or maxCodepoint < 0 Then err.Raise 5, _
        methodName, "'maxCodepoint' outside of valid unicode range."
    If minCodepoint > MAX_UNICODE Or minCodepoint < 0 Then err.Raise 5, _
        methodName, "'minCodepoint' outside of valid unicode range."
    If minCodepoint > maxCodepoint Then err.Raise 5, methodName, _
        "'minCodepoint' can't be greater than 'maxCodepoint'."
    If minCodepoint > MAX_UINT And Length Mod 2 = 1 Then err.Raise 5, methodName, _
        "Can't build string of uneven length from only Surrogate Pairs."
        
    Dim CpRange As Long: CpRange = maxCodepoint - minCodepoint + 1

    Dim i As Long
    Dim Char As Long
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)

    If Length > 1 Then
        For i = 0 To Length * 2 - 3 Step 2
            Do
                Char = Int(CpRange * Rnd) + minCodepoint
            Loop Until (Char < &HD800& Or Char > &HDFFF&) _
                   And (Char <> &HFEFF&)

            If Char < &H10000 Then
                b(i) = Char And &HFF
                b(i + 1) = Char \ &H100& And &HFF
            Else
                Dim m As Long: m = Char - &H10000
                Dim HighSurrogate As Long: HighSurrogate = &HD800& + (m \ &H400&)
                Dim LowSurrogate As Long: LowSurrogate = &HDC00& + (m And &H3FF)
                b(i) = HighSurrogate And &HFF&
                b(i + 1) = HighSurrogate \ &H100&
                i = i + 2
                b(i) = LowSurrogate And &HFF&
                b(i + 1) = LowSurrogate \ &H100&
            End If
        Next i
    End If
    RandomString = b
    
    If CInt(b(UBound(b) - 1)) + b(UBound(b)) = 0 Then
        Do
            Char = Int(CpRange * Rnd) + minCodepoint
        Loop Until (Char < &HD800& Or Char > &HDFFF&) _
               And (Char <> &HFEFF&) _
               And (Char <= MAX_UINT)
        Mid$(RandomString, Len(RandomString), 1) = ChrW(Char)
    End If
End Function

'Returns a UTF-16 string containing all characters in `inklChars` randomly
'equally distributed.
'E.g. if 'inklChars = "aab"', the returned string will, on average, contain
'     about twice as many "a"s as "b"s
Public Function RandomStringFromChars(ByVal Length As Long, _
                             Optional ByRef inklChars As String = _
    "01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz") As String
    Const methodName As String = "RandomStringFromChars"
    If Length = 0 Then Exit Function
    If Len(inklChars) = 0 Then err.Raise 5, methodName, _
        "No characters to build a string from specified in 'inklChars'"
    If Length < 0 Then err.Raise 5, methodName, "Length must be >= 0"
    
    Dim Chars() As String:    Chars = StringToCodepointStrings(inklChars)
    Dim Codepoints() As Long: Codepoints = StringToCodepointNums(inklChars)
    Dim NumChars As Long:  NumChars = UBound(Chars) - LBound(Chars) + 1
    If NumChars * 2 = Len(inklChars) And Length Mod 2 = 1 Then err.Raise 5, _
    methodName, "Can't build string of uneven length from only Surrogate Pairs."
        
    RandomStringFromChars = Space$(Length)

    Dim i As Long
    For i = 1 To Length - 1
        Dim Idx As String: Idx = Int(Rnd * NumChars)
        Mid$(RandomStringFromChars, i) = Chars(Idx)
        If Codepoints(Idx) > &HFFFF& Then i = i + 1
    Next i
    If Mid$(RandomStringFromChars, Length) = Space$(1) Then
        Do
            Idx = Int(Rnd * NumChars)
        Loop Until Codepoints(Idx) < &H10000
        Mid$(RandomStringFromChars, Length) = Chars(Idx)
    End If
End Function

'Returns an array of strings containing the individual UTF-16 characters
'Surrogate pairs remain together.
Public Function StringToCodepointStrings(ByRef Str As String) As Variant
    If Len(Str) = 0 Then Exit Function
    Dim Arr() As String: ReDim Arr(0 To Len(Str) - 1)
    Dim i As Long, j As Long
    For i = 1 To Len(Str)
        If AscU(Mid$(Str, i, 2)) > &HFFFF& Then
            Arr(j) = Mid$(Str, i, 2)
            i = i + 1
        Else
            Arr(j) = Mid$(Str, i, 1)
        End If
        j = j + 1
    Next i
    ReDim Preserve Arr(0 To j - 1)
    StringToCodepointStrings = Arr
End Function

'Returns an array of numbers representing the individual UTF-16 codepoints from
'the string 'str'
Public Function StringToCodepointNums(ByRef Str As String) As Variant
    If Len(Str) = 0 Then Exit Function
    Dim Arr() As Long: ReDim Arr(0 To Len(Str) - 1)
    Dim i As Long, j As Long
    Dim Codepoint As Long
    For i = 1 To Len(Str)
        Codepoint = AscU(Mid$(Str, i, 2))
        Arr(j) = Codepoint
        If Codepoint > &HFFFF& Then i = i + 1
        j = j + 1
    Next i
    ReDim Preserve Arr(0 To j - 1)
    StringToCodepointNums = Arr
End Function

'Returns an array of 'numElements' random UTF-16 strings
Public Function RandomStringArray(ByVal NumElements As Long, _
                         Optional ByVal maxElementLength As Long = 10, _
                         Optional ByVal minElementLength As Long = 0, _
                         Optional ByVal minCodepoint As Long = 1, _
                         Optional ByVal maxCodepoint As Long = &H10FFFF) _
                                  As String()
    Const methodName As String = "RandomStringArray"
    Const MAX_UNICODE As Long = &H10FFFF
    Const MAX_UINT As Long = &HFFFF&
    If NumElements < 0 Then err.Raise 5, methodName, "numElements must be >= 0"
    If maxCodepoint > MAX_UNICODE Or maxCodepoint < 0 Then err.Raise 5, _
        methodName, "'maxCodepoint' outside of valid unicode range."
    If minCodepoint > MAX_UNICODE Or minCodepoint < 0 Then err.Raise 5, _
        methodName, "'minCodepoint' outside of valid unicode range."
    If minCodepoint > maxCodepoint Then err.Raise 5, methodName, _
        "'minCodepoint' can't be greater than 'maxCodepoint'."
    If minElementLength > maxElementLength Then err.Raise 5, methodName, _
        "'minElementLength' can't be greater than 'maxElementLength'."
    If minCodepoint > &HFFFF& And maxElementLength = minElementLength _
    And maxElementLength Mod 2 = 1 Then err.Raise 5, methodName, _
        "Can't build string of uneven length from only Surrogate Pairs."
    
    Dim StringArray() As String: ReDim StringArray(0 To NumElements - 1)
    Dim i As Long
    Dim StrLength As Long
    
    For i = 0 To NumElements - 1
        Do
            StrLength = Int((maxElementLength - minElementLength + 1) _
                            * Rnd + minElementLength)
            If minCodepoint > &HFFFF& Then
                If StrLength Mod 2 = 0 Then Exit Do
            Else
                Exit Do
            End If
        Loop
        StringArray(i) = RandomString(StrLength, minCodepoint, maxCodepoint)
    Next i
    
    RandomStringArray = StringArray
End Function

'Removes all characters from a string (str) that are not in the string inklChars
'Default inklChars are all alphanumeric characters including dot and space
Public Function CleanString(ByRef Str As String, _
                   Optional ByVal inklChars As String = _
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ01234567890. ") _
                            As String
    Dim SChr As String
    Dim i As Long
    Dim j As Long: j = 1

    For i = 1 To Len(Str)
        SChr = Mid$(Str, i, 1)

        If InStr(1, inklChars, SChr, vbBinaryCompare) Then
            Mid$(Str, j, 1) = SChr
            j = j + 1
        End If
    Next i
    CleanString = Left$(Str, j - 1)
End Function

#If Mac = 0 Then
'Removes all non-numeric characters from a string.
'Only keeps codepoints U+0030 - U+0039
Public Function RegExNumOnly(ByRef s As String) As String
    With CreateObject("VBScript.RegExp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "[^0-9]+"
         RegExNumOnly = .Replace(s, "")
    End With
End Function
#End If

'Removes all non-numeric characters from a string.
'Keeps only codepoints U+0030 - U+0039 AND ALSO
'keeps the Unicode "Fullwidth Digits" (U+FF10 - U+FF19)!
Public Function RemoveNonNumeric(ByVal Str As String) As String
    Dim SChr As String
    Dim i As Long
    Dim j As Long: j = 1
    For i = 1 To Len(Str)
        SChr = Mid$(Str, i, 1)
        If SChr Like "#" Then _
            Mid$(Str, j, 1) = SChr: j = j + 1
    Next i
    RemoveNonNumeric = Left$(Str, j - 1)
End Function

'Inserts a string into another string at a specified position
'Insert("abcd", "ff", 0) = "ffabcd"
'Insert("abcd", "ff", 1) = "affbcd"
'Insert("abcd", "ff", 3) = "abcffd"
'Insert("abcd", "ff", 4) = "abcdff"
'Insert("abcd", "ff", 9) = "abcdff"
Public Function Insert(ByRef Str As String, _
                       ByRef strToInsert As String, _
                       ByRef afterPos As Long) As String
    Const methodName As String = "Insert"
    If afterPos < 0 Then err.Raise 5, methodName, _
        "Argument 'afterPos' = " & afterPos & " < 0, invalid"

    Insert = Mid$(Str, 1, afterPos) & strToInsert & Mid$(Str, afterPos + 1)
End Function

'Works like Insert but interprets 'afterPos' as byte-index, not char-index
'Inserting at uneven byte positions likely invalidates an utf-16 string!
Public Function InsertB(ByRef Str As String, _
                        ByRef strToInsert As String, _
                        ByRef afterPos As Long) As String
    Const methodName As String = "InsertB"
    If afterPos < 0 Then err.Raise 5, methodName, _
        "Argument 'afterPos' = " & afterPos & " < 0, invalid"

    InsertB = MidB$(Str, 1, afterPos) & strToInsert & MidB$(Str, afterPos + 1)
End Function

'Counts the number of times a substring exists in a string. Does not count
'overlapping occurrences of substring.
''lLimit' can define a number after which the function should stop counting,
'enables premature exiting of the procedure if for example the calling code
'just needs the information if count >= lLimit
'E.g.: CountSubstring("abababab", "abab") -> 2
Public Function CountSubstring(ByRef Str As String, _
                               ByRef subStr As String, _
                      Optional ByVal lStart As Long = 1, _
                      Optional ByVal lLimit As Long = -1, _
                      Optional ByVal lCompare As VbCompareMethod _
                                                 = vbBinaryCompare) As Long
    Const methodName As String = "CountSubstring"
    If lStart < 1 Then err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
    If lLimit < -1 Then err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    If subStr = vbNullString Then Exit Function

    Dim LenSubStr As Long: LenSubStr = Len(subStr)
    Dim i As Long:         i = InStr(lStart, Str, subStr, lCompare)

    CountSubstring = 0
    Do Until i = 0 Or lLimit = CountSubstring
        CountSubstring = CountSubstring + 1
        i = InStr(i + LenSubStr, Str, subStr, lCompare)
    Loop
End Function

'Like CountSubstring but scans a string bytewise.
'Example illustrating the difference to CountSubstring:
'                       |c1||c2|
'bytes = HexToString("0x00610061")
'                         |c3|
'sFind =   HexToString("0x6100")
'CountSubstring(bytes, sFind) -> 0
'CountSubstringB(bytes, sFind) -> 1
Public Function CountSubstringB(ByRef Bytes As String, _
                                ByRef subStr As String, _
                       Optional ByVal lStart As Long = 1, _
                       Optional ByVal lLimit As Long = -1, _
                       Optional ByVal lCompare As VbCompareMethod _
                                               = vbBinaryCompare) As Long
    Const methodName As String = "CountSubstringB"
    If lStart < 1 Then err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
    If lLimit < -1 Then err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    If subStr = vbNullString Then Exit Function
    
    Dim LenBSubStr As Long: LenBSubStr = LenB(subStr)
    Dim i As Long:          i = InStrB(lStart, Bytes, subStr, lCompare)

    CountSubstringB = 0
    Do Until i = 0 Or lLimit = CountSubstringB
        CountSubstringB = CountSubstringB + 1
        i = InStrB(i + LenBSubStr, Bytes, subStr, lCompare)
    Loop
End Function

'Counts the number of times a substring exists in a string unless they are
'escaped' (appear twice in a row). Does not count overlapping occurrences of
'substring.
'E.g.: CountSubstringUnlessEscaped("abababababab", "abab") -> 1
Public Function CountSubstringUnlessEscaped(ByRef Str As String, _
                                            ByRef subStr As String, _
                                   Optional ByVal lStart As Long = 1, _
                                   Optional ByVal lLimit As Long = -1, _
                                   Optional ByVal lCompare As VbCompareMethod _
                                                            = vbBinaryCompare) _
                                            As Long
    Const methodName As String = "CountSubstringUnlessEscaped"
    If lStart < 1 Then err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
    If lLimit < -1 Then err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
        
    Dim LenSubStr As Long: LenSubStr = Len(subStr)
    Dim i As Long:         i = InStr(lStart, Str, subStr, lCompare)

    CountSubstringUnlessEscaped = 0
    Do Until i = 0 Or lLimit = CountSubstringUnlessEscaped
        If StrComp(subStr, Mid(Str, i + LenSubStr, LenSubStr), lCompare) = 0 Then
            i = i + LenSubStr
        Else
            CountSubstringUnlessEscaped = CountSubstringUnlessEscaped + 1
        End If
        i = InStr(i + LenSubStr, Str, subStr, lCompare)
    Loop
End Function

'Like CountSubstringUnlessEscaped but scans a string bytewise.
'Example illustrating the difference to CountSubstring:
'                       |c1||c2||c3||c4|
'bytes = HexToString("0x0061006100610061")
'                         |escape||ct|
'sFind =   HexToString("0x6100")
'CountSubstringUnlessEscaped(bytes, sFind) -> 0
'CountSubstringUnlessEscapedB(bytes, sFind) -> 1
Public Function CountSubstringUnlessEscapedB(ByRef Bytes As String, _
                                             ByRef subStr As String, _
                                    Optional ByVal lStart As Long = 1, _
                                    Optional ByVal lLimit As Long = -1, _
                                    Optional ByVal lCompare As VbCompareMethod _
                                                            = vbBinaryCompare) _
                                             As Long
    Const methodName As String = "CountSubstringUnlessEscaped"
    If lStart < 1 Then err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
    If lLimit < -1 Then err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"

    Dim LenBSubStr As Long: LenBSubStr = LenB(subStr)
    Dim i As Long:          i = InStrB(lStart, Bytes, subStr, lCompare)

    CountSubstringUnlessEscapedB = 0
    Do Until i = 0 Or lLimit = CountSubstringUnlessEscapedB
        If StrComp(subStr, MidB(Bytes, i + LenBSubStr, LenBSubStr), _
                   lCompare) = 0 Then
            i = i + LenBSubStr
        Else
            CountSubstringUnlessEscapedB = CountSubstringUnlessEscapedB + 1
        End If
        i = InStrB(i + LenBSubStr, Bytes, subStr, lCompare)
    Loop
End Function

'Works like the inbuilt 'Replace', but parses the string bytewise, not charwise.
'Example illustrating the difference:
'bytes = HexToString("0x00610061")
'sFind = HexToString("0x6100")
'? StringToHex(ReplaceB(bytes, sFind, "")) -> "0x0061"
'? StringToHex(Replace(bytes, sFind, "")) -> "0x00610061"
Public Function ReplaceB(ByRef Bytes As String, _
                         ByRef sFind As String, _
                         ByRef sReplace As String, _
                Optional ByVal lStart As Long = 1, _
                Optional ByVal lCount As Long = -1, _
                Optional ByVal lCompare As VbCompareMethod _
                                        = vbBinaryCompare) As String
    Const methodName As String = "ReplaceB"
    If lStart < 1 Then err.Raise 5, methodName, _
        "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then err.Raise 5, methodName, _
        "Argument 'lCount' = " & lCount & " < -1, invalid"
    lCount = lCount And &H7FFFFFFF

    If LenB(Bytes) = 0 Or LenB(sFind) = 0 Then
        ReplaceB = MidB$(Bytes, lStart)
        Exit Function
    End If

    Dim LenBFind As Long:    LenBFind = LenB(sFind)
    Dim LenBReplace As Long: LenBReplace = LenB(sReplace)

    Dim Buffer() As Byte
    ReDim Buffer(0 To LenB(Bytes) - lStart + CountSubstringB(Bytes, sFind, _
                 lStart, lCount, lCompare) * (LenBReplace - LenBFind))
    ReplaceB = Buffer

    Dim i As Long:              i = InStrB(lStart, Bytes, sFind, lCompare)
    Dim j As Long:              j = 1
    Dim LastOccurrence As Long: LastOccurrence = lStart
    Dim Count As Long:          Count = 1

    Do Until i = 0 Or Count > lCount
        Dim Diff As Long: Diff = i - LastOccurrence
        If Diff > 0 Then _
            MidB$(ReplaceB, j, Diff) = MidB$(Bytes, LastOccurrence, Diff)
        j = j + Diff
        If LenBReplace <> 0 Then
            MidB$(ReplaceB, j, LenBReplace) = sReplace
            j = j + LenBReplace
        End If
        Count = Count + 1
        LastOccurrence = i + LenBFind
        i = InStrB(LastOccurrence, Bytes, sFind, lCompare)
    Loop
    If j <= LenB(ReplaceB) Then MidB$(ReplaceB, j) = MidB$(Bytes, LastOccurrence)
End Function

'Replaces consecutive occurrences of 'substring' that repeat more than 'limit'
'times with exactly 'limit' consecutive occurrences
'E.g.: LimitConsecutiveSubstringRepetition("aaaabaaac", "a", 1)  -> "abac"
'      LimitConsecutiveSubstringRepetition("aaaabaaac", "aa", 1) -> "aabaaac"
'      LimitConsecutiveSubstringRepetition("aaaabaaac", "a", 2)  -> "aabaac"
'      LimitConsecutiveSubstringRepetition("aaaabaaac", "ab", 0) -> "aaaaaac"
Public Function LimitConsecutiveSubstringRepetition( _
                                           ByRef Str As String, _
                                  Optional ByRef subStr As String = vbNewLine, _
                                  Optional ByVal Limit As Long = 1, _
                                  Optional ByVal Compare As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                           As String
    Const methodName As String = "LimitConsecutiveSubstringRepetition"

    If Limit < 0 Then err.Raise 5, methodName, _
        "Argument 'limit' = " & Limit & " < 0, invalid"
    If Limit = 0 Then
        LimitConsecutiveSubstringRepetition = Replace(Str, subStr, _
                                                      vbNullString, , , Compare)
        Exit Function
    Else
        LimitConsecutiveSubstringRepetition = Str
    End If
    If Len(Str) = 0 Then Exit Function
    If Len(subStr) = 0 Then Exit Function

    Dim i As Long:                i = InStr(1, Str, subStr, Compare)
    Dim j As Long:                j = 1
    Dim LenSubStr As Long:        LenSubStr = Len(subStr)
    Dim CopyChunkSize As Long:    CopyChunkSize = 0
    Dim ConsecutiveCount As Long: ConsecutiveCount = 0
    Dim LastOccurrence As Long:   LastOccurrence = 1 - LenSubStr
    Dim OccurrenceDiff As Long

    Do Until i = 0
        OccurrenceDiff = i - LastOccurrence
        If OccurrenceDiff = LenSubStr Then
            ConsecutiveCount = ConsecutiveCount + 1
            If ConsecutiveCount <= Limit Then
                CopyChunkSize = CopyChunkSize + OccurrenceDiff
            ElseIf ConsecutiveCount = Limit + 1 Then
                Mid$(LimitConsecutiveSubstringRepetition, j, CopyChunkSize) = _
                    Mid$(Str, i - CopyChunkSize, CopyChunkSize)
                j = j + CopyChunkSize
                CopyChunkSize = 0
            End If
        Else
            CopyChunkSize = CopyChunkSize + OccurrenceDiff
            ConsecutiveCount = 1
        End If
        LastOccurrence = i
        i = InStr(i + LenSubStr, Str, subStr, Compare)
    Loop

    CopyChunkSize = CopyChunkSize + Len(Str) - LastOccurrence - LenSubStr + 1
    Mid$(LimitConsecutiveSubstringRepetition, j, CopyChunkSize) = _
        Mid$(Str, Len(Str) - CopyChunkSize + 1)

    LimitConsecutiveSubstringRepetition = _
        Left$(LimitConsecutiveSubstringRepetition, j + CopyChunkSize - 1)
End Function

'Same as LimitConsecutiveSubstringRepetition, but scans the string bytewise.
'Example illustrating the difference:
'Dim bytes As String: bytes = HexToString("0x006100610061")
'Dim subStr As String: subStr = HexToString("0x6100")
'StringToHex(LimitConsecutiveSubstringRepetition(bytes, subStr, 1) _
'    -> "0x006100610061"
'StringToHex(LimitConsecutiveSubstringRepetitionB(bytes, subStr, 1) _
'    -> "0x00610061"
Public Function LimitConsecutiveSubstringRepetitionB( _
                                           ByRef Bytes As String, _
                                  Optional ByRef subStr As String = vbNewLine, _
                                  Optional ByVal Limit As Long = 1, _
                                  Optional ByVal Compare As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                           As String
    Const methodName As String = "LimitConsecutiveSubstringRepetitionB"

    If Limit < 0 Then err.Raise 5, methodName, _
        "Argument 'limit' = " & Limit & " < 0, invalid"
    If Limit = 0 Then
        LimitConsecutiveSubstringRepetitionB = ReplaceB(Bytes, subStr, _
                                                      vbNullString, , , Compare)
        Exit Function
    Else
        LimitConsecutiveSubstringRepetitionB = Bytes
    End If
    If LenB(Bytes) = 0 Then Exit Function
    If LenB(subStr) = 0 Then Exit Function

    Dim i As Long:                i = InStrB(1, Bytes, subStr, Compare)
    Dim j As Long:                j = 1
    Dim LenBSubStr As Long:       LenBSubStr = LenB(subStr)
    Dim CopyChunkSize As Long:    CopyChunkSize = 0
    Dim ConsecutiveCount As Long: ConsecutiveCount = 0
    Dim LastOccurrence As Long:   LastOccurrence = 1 - LenBSubStr
    Dim OccurrenceDiff As Long

    Do Until i = 0
        OccurrenceDiff = i - LastOccurrence
        If OccurrenceDiff = LenBSubStr Then
            ConsecutiveCount = ConsecutiveCount + 1
            If ConsecutiveCount <= Limit Then
                CopyChunkSize = CopyChunkSize + OccurrenceDiff
            ElseIf ConsecutiveCount = Limit + 1 Then
                MidB$(LimitConsecutiveSubstringRepetitionB, j, CopyChunkSize) = _
                    MidB$(Bytes, i - CopyChunkSize, CopyChunkSize)
                j = j + CopyChunkSize
                CopyChunkSize = 0
            End If
        Else
            CopyChunkSize = CopyChunkSize + OccurrenceDiff
            ConsecutiveCount = 1
        End If
        LastOccurrence = i
        i = InStrB(i + LenBSubStr, Bytes, subStr, Compare)
    Loop

    CopyChunkSize = CopyChunkSize + LenB(Bytes) - LastOccurrence - LenBSubStr + 1
    MidB$(LimitConsecutiveSubstringRepetitionB, j, CopyChunkSize) = _
        MidB$(Bytes, LenB(Bytes) - CopyChunkSize + 1)

    LimitConsecutiveSubstringRepetitionB = _
        LeftB$(LimitConsecutiveSubstringRepetitionB, j + CopyChunkSize - 1)
End Function

'Repeats the string str, repeatTimes times.
'Works with byte strings of uneven LenB
'E.g.: RepeatString("a", 3) -> "aaa"
'      StrConv(RepeatString(MidB("a", 1, 1), 3), vbUnicode) -> "aaa"
Public Function RepeatString(ByRef Str As String, _
                    Optional ByVal repeatTimes As Long = 2) As String
    If repeatTimes = 0 Then Exit Function
    If LenB(Str) = 2 Then
        RepeatString = String$(repeatTimes, Str)
        Exit Function
    End If

    Dim NewLength As Long: NewLength = LenB(Str) * repeatTimes
    RepeatString = Space$((NewLength + 1) \ 2)
    If NewLength Mod 2 = 1 Then RepeatString = MidB$(RepeatString, 2)
    
    MidB$(RepeatString, 1) = Str
    If repeatTimes > 1 Then MidB$(RepeatString, LenB(Str) + 1) = RepeatString
End Function

'Adds fillerStr to the right side of a string repeatedly until the resulting
'string reaches length 'Length'
'E.g.: PadRight("asd", 11, "xyz") -> "asdxyzxyzxy"
Public Function PadRight(ByRef Str As String, _
                         ByVal Length As Long, _
                Optional ByVal fillerStr As String = " ") As String
    PadRight = PadRightB(Str, Length * 2, fillerStr)
End Function

'Adds fillerStr to the left side of a string repeatedly until the resulting
'string reaches length 'Length'
'E.g.: PadLeft("asd", 11, "xyz") -> "yzxyzxyzasd"
Public Function PadLeft(ByRef Str As String, _
                        ByVal Length As Long, _
               Optional ByVal fillerStr As String = " ") As String
    PadLeft = PadLeftB(Str, Length * 2, fillerStr)
End Function

'Adds fillerStr to the right side of a string repeatedly until the resulting
'string reaches length 'Length' in bytes!
'E.g.: PadRightB("asd", 16, "xyz") -> "asdxyzxy"
Public Function PadRightB(ByRef Str As String, _
                          ByVal Length As Long, _
                 Optional ByVal fillerStr As String = " ") As String
    Const methodName As String = "PadRightB"
    If Length < 0 Then err.Raise 5, methodName, _
        "Argument 'Length' = " & Length & " < 0, invalid"
    If LenB(fillerStr) = 0 Then err.Raise 5, methodName, _
        "Argument 'fillerStr' = vbNullString, invalid"

    If Length > LenB(Str) Then
        If LenB(fillerStr) = 2 Then
            PadRightB = Str & String((Length - LenB(Str) + 1) \ 2, fillerStr)
            If Length Mod 2 = 1 Then _
                PadRightB = LeftB$(PadRightB, LenB(PadRightB) - 1)
        Else
            PadRightB = Str & LeftB$(RepeatString(fillerStr, (((Length - _
                LenB(Str))) + 1) \ LenB(fillerStr) + 1), Length - LenB(Str))
        End If
    Else
        PadRightB = LeftB$(Str, Length)
    End If
End Function

'Adds fillerStr to the left side of a string repeatedly until the resulting
'string reaches length 'Length' in bytes!
'Note that this can result in an invalid UTF-16 output for uneven lengths!
'E.g.: PadLeftB("asd", 16, "xyz") -> "yzxyzasd"
'      PadLeftB("asd", 11, "xyz") -> "?????"
Public Function PadLeftB(ByRef Str As String, _
                         ByVal Length As Long, _
                Optional ByVal fillerStr As String = " ") As String
    Const methodName As String = "PadLeftB"
    If Length < 0 Then err.Raise 5, methodName, _
        "Argument 'Length' = " & Length & " < 0, invalid"
    If LenB(fillerStr) = 0 Then err.Raise 5, methodName, _
        "Argument 'fillerStr' = vbNullString, invalid"

    If Length > LenB(Str) Then
        If LenB(fillerStr) = 2 Then
            PadLeftB = String((Length - LenB(Str) + 1) \ 2, fillerStr) & Str
            If Length Mod 2 = 1 Then _
                PadLeftB = RightB$(PadLeftB, LenB(PadLeftB) - 1)
        Else
            PadLeftB = RightB$(RepeatString(fillerStr, (((Length - LenB(Str))) _
                          + 1) \ LenB(fillerStr) + 1), Length - LenB(Str)) & Str
        End If
    Else
        PadLeftB = RightB$(Str, Length)
    End If
End Function

'Works like the inbuilt 'Split', but parses string bytewise, so it splits at
'all occurrences of 'Delimiter', even at uneven byte-index positions.
'Example illustrating the difference:
'bytes = HexToString("0x00610061")
'sDelim = HexToString("0x6100")
'SplitB(bytes, sDelim)) -> "0x00", "0x61"
'Split(bytes, sDelim, "")) -> "0x00610061"
Public Function SplitB(ByRef Bytes As String, _
              Optional ByRef sDelimiter As String = " ", _
              Optional ByVal lLimit As Long = -1, _
              Optional ByVal lCompare As VbCompareMethod = vbBinaryCompare) _
                       As Variant
    Const methodName As String = "SplitB"
    If lLimit < -1 Then err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    lLimit = lLimit And &H7FFFFFFF
    
    If lLimit = 0 Or Bytes = vbNullString Then
        SplitB = Split("", , 0) 'Return empty but allocated string array
        Exit Function
    ElseIf LenB(Bytes) = 0 Or LenB(sDelimiter) = 0 Or lLimit < 2 Then
        Dim Arr() As String: ReDim Arr(0 To 0)
        Arr(0) = Bytes
        SplitB = Arr
        Exit Function
    End If

    Dim LenBDelim As Long:  LenBDelim = LenB(sDelimiter)
    ReDim Arr(0 To CountSubstringB(Bytes, sDelimiter, 1, lLimit, lCompare))
    Dim i As Long:              i = InStrB(1, Bytes, sDelimiter, lCompare)
    Dim LastOccurrence As Long: LastOccurrence = 1
    Dim Count As Long:          Count = 0

    Do Until i = 0 Or Count + 1 >= lLimit
        Dim Diff As Long: Diff = i - LastOccurrence
        Arr(Count) = MidB$(Bytes, LastOccurrence, Diff)
        Count = Count + 1
        LastOccurrence = i + LenBDelim
        i = InStrB(LastOccurrence, Bytes, sDelimiter, lCompare)
    Loop
    Arr(Count) = MidB$(Bytes, LastOccurrence)
    SplitB = Arr
End Function

'Works like the inbuilt 'Split', but if delimiter is escaped (appears twice in
'a row) the string will not be split at that position and instead the double
'delimiter will be replaced by a single one
Public Function SplitUnlessEscaped(ByRef Str As String, _
                          Optional ByRef sDelimiter As String = " ", _
                          Optional ByVal lLimit As Long = -1, _
                          Optional ByVal lCompare As VbCompareMethod = _
                                                     vbBinaryCompare) _
                                   As Variant
    Const methodName As String = "SplitUnlessEscaped"
    If lLimit < -1 Then err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    lLimit = lLimit And &H7FFFFFFF

    If lLimit = 0 Then
        SplitUnlessEscaped = Split("", , 0) 'Return empty but allocated str arr
        Exit Function
    ElseIf Len(Str) = 0 Or Len(sDelimiter) = 0 Or lLimit < 2 Then
        Dim Arr() As String:  ReDim Arr(0 To 0)
        Arr(0) = Str
        SplitUnlessEscaped = Arr
        Exit Function
    End If
    
    Dim LenDelim As Long:   LenDelim = Len(sDelimiter)
    ReDim Arr(0 To CountSubstringUnlessEscaped(Str, sDelimiter, 1, lLimit, _
                                               lCompare))
    Dim PartStart As Long:      PartStart = 1
    Dim Count As Long:          Count = 0
    Dim LastOccurrence As Long: LastOccurrence = 1
    Dim i As Long:          i = InStr(LastOccurrence, Str, sDelimiter, lCompare)
    
    Do Until i = 0 Or Count + 1 >= lLimit
        If Mid(Str, i + LenDelim, LenDelim) = sDelimiter Then
            LastOccurrence = i + 2 * LenDelim
        Else
            Arr(Count) = Replace(Mid(Str, PartStart, i - PartStart), _
                                 sDelimiter & sDelimiter, sDelimiter)
            Count = Count + 1
            PartStart = i + LenDelim
            LastOccurrence = PartStart
        End If
        i = InStr(LastOccurrence, Str, sDelimiter, lCompare)
    Loop
    
    If Count < lLimit Then Arr(Count) = Replace(Mid(Str, PartStart), _
                                            sDelimiter & sDelimiter, sDelimiter)
    SplitUnlessEscaped = Arr
End Function

'Works like 'SplitB', but if delimiter is escaped (appears twice in
'a row) the string will not be split at that position and instead the double
'delimiter will be replaced by a single one
Public Function SplitUnlessEscapedB(ByRef Bytes As String, _
                           Optional ByRef sDelimiter As String = " ", _
                           Optional ByVal lLimit As Long = -1, _
                           Optional ByVal lCompare As VbCompareMethod = _
                                                      vbBinaryCompare) _
                                    As Variant
    Const methodName As String = "SplitUnlessEscapedB"
    If lLimit < -1 Then err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    lLimit = lLimit And &H7FFFFFFF

    If lLimit = 0 Then
        SplitUnlessEscapedB = Split("", , 0) 'Return empty but allocated str arr
        Exit Function
    ElseIf LenB(Bytes) = 0 Or LenB(sDelimiter) = 0 Or lLimit < 2 Then
        Dim Arr() As String:  ReDim Arr(0 To 0)
        Arr(0) = Bytes
        SplitUnlessEscapedB = Arr
        Exit Function
    End If
    
    Dim LenBDelim As Long:   LenBDelim = LenB(sDelimiter)
    ReDim Arr(0 To CountSubstringUnlessEscapedB(Bytes, sDelimiter, 1, lLimit, _
                                                lCompare))
    Dim PartStart As Long:      PartStart = 1
    Dim Count As Long:          Count = 0
    Dim LastOccurrence As Long: LastOccurrence = 1
    Dim i As Long:       i = InStrB(LastOccurrence, Bytes, sDelimiter, lCompare)
    
    Do Until i = 0 Or Count + 1 >= lLimit
        If MidB(Bytes, i + LenBDelim, LenBDelim) = sDelimiter Then
            LastOccurrence = i + 2 * LenBDelim
        Else
            Arr(Count) = ReplaceB(MidB(Bytes, PartStart, i - PartStart), _
                                  sDelimiter & sDelimiter, sDelimiter)
            Count = Count + 1
            PartStart = i + LenBDelim
            LastOccurrence = PartStart
        End If
        i = InStrB(LastOccurrence, Bytes, sDelimiter, lCompare)
    Loop
    
    If Count < lLimit Then Arr(Count) = ReplaceB(MidB(Bytes, PartStart), _
                                            sDelimiter & sDelimiter, sDelimiter)
    SplitUnlessEscapedB = Arr
End Function

'Splits a string at every occurrence of the specified delimiter "delim", unless
'that delimiter occurs between non-escaped quotes. e.g. (" asf delim asdf ")
'will not be split. Quotes will not be removed.
'Quotes can be escaped by repetition.
'E.g.: SplitUnlessInQuotes("Hello "" ""World" "Goodbye World") returns
'      "Hello "" "" World", and "Goodbye World"
'If " is chosen as delimiter, splits at the outermost two occurrences of ", or
'if only one " exists in the string, splits the string into two parts.
'E.g. SplitUnlessInQuotes("asdf""asdf""asdf""asdf", """") returns
'    "asdf", "asdf""asdf", and "asdf"
Public Function SplitUnlessInQuotes(ByRef Str As String, _
                           Optional ByRef delim As String = " ", _
                           Optional Limit As Long = -1) As Variant
    Dim i As Long
    Dim s As String
    Dim ub As Long:         ub = -1
    Dim Parts As Variant:   ReDim Parts(0 To 0)
    Dim DoSplit As Boolean: DoSplit = True

    If delim = """" Then 'Handle this special case
        i = InStr(1, Str, """", vbBinaryCompare)
        If i <> 0 Then
            Dim j As Long: j = InStrRev(Str, """", , vbBinaryCompare)
            If i = j Then
                SplitUnlessInQuotes = Split(Str, """", , vbBinaryCompare)
                Exit Function
            Else
                ReDim Parts(0 To 2)
                Parts(0) = Left$(Str, i - 1)
                Parts(1) = Mid$(Str, i + 1, j)
                Parts(2) = Mid$(Str, j + 1)
            End If
        Else
            Parts(0) = Str
        End If
        SplitUnlessInQuotes = Parts
        Exit Function
    End If

    For i = 1 To Len(Str)
        If ub = Limit - 2 Then
            ub = ub + 1
            ReDim Preserve Parts(0 To ub)
            Parts(ub) = Mid$(Str, i)
            Exit For
        End If

        If Mid$(Str, i, 1) = """" Then
            DoSplit = Not DoSplit
            If Not DoSplit Then _
                DoSplit = InStr(i + 1, Str, """", vbBinaryCompare) = 0
        End If

        If Mid$(Str, i, Len(delim)) = delim And DoSplit Or i = Len(Str) Then
            If i = Len(Str) Then s = s & Mid$(Str, i, 1)
            ub = ub + 1
            ReDim Preserve Parts(0 To ub)
            Parts(ub) = s
            s = vbNullString
            i = i + Len(delim) - 1
        Else
            s = s & Mid$(Str, i, 1)
        End If
    Next i
    SplitUnlessInQuotes = Parts
End Function

'Reads the memory of a String to an Array of Integers
'Notes:
'   - Ignores the last byte if input has an odd number of bytes
'   - If 'outLength' is -1 (default) then the remaining length is returned
'   - Excess length is ignored
Public Function StringToIntegers(ByRef s As String, _
                        Optional ByVal startIndex As Long = 1, _
                        Optional ByVal outLength As Long = -1, _
                        Optional ByVal outLowBound As Long = 0) As Integer()
    Static sArr As SAFEARRAY_1D
    Static memValue As Variant
    Static remoteVT As Variant
    Const methodName As String = "StringToIntegers"
    Dim CLen As Long: CLen = Len(s)

    If startIndex < 1 Or startIndex > CLen Then
        err.Raise 9, methodName, "Invalid Start Index"
    ElseIf outLength < -1 Then
        err.Raise 5, methodName, "Invalid Length for output"
    ElseIf outLength = -1 Or startIndex + outLength - 1 > CLen Then
        outLength = CLen - startIndex + 1
    End If
    If IsEmpty(memValue) Then
        remoteVT = VarPtr(memValue)
        CopyMemory remoteVT, vbInteger + VT_BYREF, 2
        With sArr
            .cDims = 1
            .fFeatures = FADF_HAVEVARTYPE
            .cbElements = INT_SIZE
        End With
        memValue = VarPtr(sArr)
    End If
    With sArr
        .pvData = StrPtr(s) + (startIndex - 1) * INT_SIZE
        .rgsabound0.lLbound = outLowBound
        .rgsabound0.cElements = outLength
    End With
    RemoteAssign remoteVT, vbArray + vbInteger, StringToIntegers, memValue
End Function

'This method assures the required redirection for both the remote varType and
'   the remote value at the same time thus removing any additional stack frames
'It can be used to both read from and write to memory by swapping the order of
'   the last 2 parameters
Private Sub RemoteAssign(ByRef remoteVT As Variant, _
                         ByVal newVT As VbVarType, _
                         ByRef targetVariable As Variant, _
                         ByRef newValue As Variant)
    remoteVT = newVT
    targetVariable = newValue
    remoteVT = vbLongPtr 'Stop linking to remote address, for safety
End Sub

'Reads the memory of an Array of Integers into a String
'Notes:
'   - If 'outLength' is -1 (default) then the remaining length is returned
'   - Excess length is ignored
Public Function IntegersToString(ByRef ints() As Integer, _
                        Optional ByVal startIndex As Long = 0, _
                        Optional ByVal outLength As Long = -1) As String
    Static sArr As SAFEARRAY_1D
    Static memValue As Variant
    Static remoteVT As Variant
    Const methodName As String = "IntegersToString"

    If GetArrayDimsCount(ints) <> 1 Then
        err.Raise 5, methodName, "Expected 1D Array of Integers"
    ElseIf startIndex < LBound(ints) Or startIndex > UBound(ints) Then
        err.Raise 9, methodName, "Invalid Start Index"
    ElseIf outLength < -1 Then
        err.Raise 5, methodName, "Invalid Length for output"
    ElseIf outLength = -1 Or startIndex + outLength - 1 > UBound(ints) Then
        outLength = UBound(ints) - startIndex + 1
    End If
    If IsEmpty(memValue) Then
        remoteVT = VarPtr(memValue)
        CopyMemory remoteVT, vbInteger + VT_BYREF, 2
        With sArr
            .cDims = 1
            .fFeatures = FADF_HAVEVARTYPE
            .cbElements = BYTE_SIZE
            .rgsabound0.lLbound = 0
        End With
        memValue = VarPtr(sArr)
    End If
    With sArr
        .pvData = VarPtr(ints(startIndex))
        .rgsabound0.cElements = outLength * INT_SIZE
    End With
    RemoteAssign remoteVT, vbArray + vbByte, IntegersToString, memValue
End Function

'Returns the Number of dimensions for an input array
'Returns 0 if array is uninitialized or input not an array
'Note that a zero-length array has 1 dimension! Ex. Array() bounds are (0 to -1)
Private Function GetArrayDimsCount(ByRef Arr As Variant) As Long
    Const MAX_DIMENSION As Long = 60 'VB limit
    Dim Dimension As Long
    Dim TempBound As Long

    On Error GoTo FinalDimension
    For Dimension = 1 To MAX_DIMENSION
        TempBound = LBound(Arr, Dimension)
    Next Dimension
FinalDimension:
    GetArrayDimsCount = Dimension - 1
End Function

'Works like the inbuilt `Replace` but also accepts arrays as 'sFindOrFinds' or
''sReplaceOrReplaces'. The arrays for the find and replace values do not need
'to have the same number of elements, if they differ in the number of elements
'the following logic will be used for the replacing:
'Let sFindOrFinds contain n elements, and sReplaceOrReplaces contain m elements
'and be 0 based arrays.
'(The function converts non 0 based input arrays to 0 based arrays)
'Case 1: n > m
'   sFindOrFinds(i) gets replaced by sReplaceOrReplaces(i Mod m)
'   E.g.: sFindOrFinds = ("1", "2", "3"), sReplaceOrReplaces = ("4", "5")
'         ReplaceMultiple("123123123", ...) returns "454454454"
'         -> Each "1" and "3" get replaced by "4" while "2" gets replaced by "5"
'Case 2: n < m
'
'   E.g.: sFindOrFinds = ("1", "2"), sReplaceOrReplaces = ("3", "4", "5")
'         Every odd numbered occurrence of "1" gets replaced by "3", every even
'         numbered occurrence of "1" gets replaced by "5" and every occurrence
'         of "2" gets replaced by "4"
'         ReplaceMultiple("123123123", ...) returns "343543343"
'Notes:
''ReplaceMultipleMultiPass' is almost always faster than this function, but
'   does not support the n < m case and also the behavior can differ and be less
'   predictable, if for example finds appear in the string through prior
'   replacements that can then get replaced again in the next iteration.
'This function should not be used with more than a few thousand finds at once
'   because the runtime is proportional to n^2 because of an important check
'   as commented in the code!
Public Function ReplaceMultiple(ByRef Str As String, _
                                ByRef sFindOrFinds As Variant, _
                                ByRef sReplaceOrReplaces As Variant, _
                       Optional ByVal lStart As Long = 1, _
                       Optional ByVal lCount As Long = -1, _
                       Optional ByVal lCompare As VbCompareMethod _
                                                = vbBinaryCompare) As String
    Const methodName As String = "ReplaceMultiple"
    If lStart < 1 Then err.Raise 5, methodName, _
                               "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then err.Raise 5, methodName, _
                              "Argument 'lCount' = " & lCount & " < -1, invalid"
                              
    If lStart > Len(Str) Then
        ReplaceMultiple = Mid$(Str, lStart) 'In case LenB(str) = lStart * 2 + 1
        'Note, the inbuilt Replace function actually doesn't take this case into
        'account so for logical consistency with the inbuilt Replace, the
        'ReplaceMultiple = Mid$(str, lStart) should actually be omitted, however
        'this behavior of the inbuilt Replace could also be interpreted as a bug
        Exit Function
    End If
    
    Dim Finds As Variant
    If IsArray(sFindOrFinds) Then Finds = sFindOrFinds _
                             Else Finds = VBA.Array(sFindOrFinds)
    If Not LBound(Finds) = 0 Then _
        ReDim Preserve Finds(0 To UBound(Finds) - LBound(Finds))
    Dim Replaces As Variant
    If IsArray(sReplaceOrReplaces) Then Replaces = sReplaceOrReplaces _
                                   Else Replaces = VBA.Array(sReplaceOrReplaces)
    If Not LBound(Replaces) = 0 Then _
        ReDim Preserve Replaces(0 To UBound(Replaces) - LBound(Replaces))
    lCount = lCount And &H7FFFFFFF
    
    If UBound(Finds) = 0 And Len(Finds(0)) = 0 _
    And UBound(Replaces) = 0 And Len(Replaces(0)) = 0 Then
        ReplaceMultiple = Mid$(Str, lStart)
        Exit Function
    End If
    
    'Clean input arrays to deal with cases where one "find" contains another one
    
    'Unfortunately, this part of the algorithm introduces an O(n^2 * m)
    'complexity (n = number of finds, m average length of finds) which can make
    'this function very slow for more than a few 1000 finds. This can
    'theoretically be improved by using a Trie instead in certain cases to avoid
    'n^2 runtime complexity: https://en.wikipedia.org/wiki/Trie
    'A simple implementation of the trie algorithm has been tested, the
    'procedure is available in the test module ('ProcessFindsUsingTrie')
    'Unfortunately it performs at least 20 times slower than the naïve approach
    'implemented here:
    Dim i As Long, j As Long
    For i = 0 To UBound(Finds)
        If Len(Finds(i)) <> 0 Then
            For j = i + 1 To UBound(Finds)
                If InStr(1, Finds(j), Finds(i), lCompare) <> 0 Then
                    Finds(j) = vbNullString
                End If
            Next j
        End If
    Next i
    
    'ProcessFindsUsingTrie finds, lCompare '<-- at least 20 times slower

    'Allocate buffer
    'Buffer calculation doesn't really take into account the parameter lCount.
    'We'll take it into account at the end of the function instead
    Dim n As Long:          n = UBound(Finds) + 1
    Dim m As Long:          m = UBound(Replaces) + 1
    Dim LenBBuffer As Long: LenBBuffer = LenB(Str) - (lStart - 1) * 2
    If m > n Then
        For i = 0 To UBound(Finds)
            Dim NumReplPerFind As Long
            NumReplPerFind = IIf(i < m Mod n, (m \ n) + 1, (m \ n))
            Dim NumOcc As Long
            NumOcc = CountSubstring(Str, CStr(Finds(i)), lStart, lCount, _
                                    lCompare)
            For j = i To m - 1 Step n
                LenBBuffer = LenBBuffer + (LenB(Replaces(j)) - LenB(Finds(i))) _
                             * IIf((j - i) \ n < NumOcc Mod NumReplPerFind _
                         , NumOcc \ NumReplPerFind + 1, NumOcc \ NumReplPerFind)
            Next j
        Next i
    Else
        For i = 0 To UBound(Finds)
            NumOcc = CountSubstring(Str, CStr(Finds(i)), lStart, lCount, _
                                    lCompare)
            LenBBuffer = LenBBuffer + _
                         (LenB(Replaces(i Mod m)) - LenB(Finds(i))) * NumOcc
        Next i
    End If
    
    Dim Buffer() As Byte: ReDim Buffer(1 To LenBBuffer)
    ReplaceMultiple = Buffer
    
    'Keep track of the next position for replacing by using a min-heap
    Dim NextOccsHeap() As Long: ReDim NextOccsHeap(0 To UBound(Finds), 0 To 2)
    Dim Index1 As Long
    Dim Index2 As Long
    Dim InsertElement(0 To 2) As Long
    Dim HeapSize As Long
    
    Dim LenBReplaces() As Long: ReDim LenBReplaces(0 To UBound(Replaces))
    For i = 0 To UBound(Replaces)
        LenBReplaces(i) = LenB(Replaces(i))
    Next i
    Dim LenBFinds() As Long: ReDim LenBFinds(0 To UBound(Finds))
    For i = 0 To UBound(Finds)
        LenBFinds(i) = LenB(Finds(i))
    Next i
    
    For i = 0 To UBound(Finds)
        Dim NextOcc As Long
        If LenBFinds(i) < 2 Then
            NextOcc = 0
        Else
            NextOcc = InStr(lStart, Str, Finds(i), lCompare) * 2 - 1
        End If
        If NextOcc > 0 Then
            InsertElement(0) = NextOcc
            InsertElement(1) = i
            InsertElement(2) = i Mod m
            GoSub HeapInsert
        End If
    Next i
    
    Dim LastOccurrence As Long: LastOccurrence = lStart * 2 - 1
    Dim CurrOccurrence As Long: CurrOccurrence = NextOccsHeap(0, 0)
    Dim CurrReplaceIdx As Long: CurrReplaceIdx = NextOccsHeap(0, 2)
    Dim BuiltStrPos As Long:    BuiltStrPos = 1
    Dim Count As Long:          Count = 1
    
    'Do the replacing using InStr() and Mid$ in a loop
    Do Until HeapSize = 0 Or Count > lCount
        Dim Diff As Long: Diff = CurrOccurrence - LastOccurrence
        If Diff > 0 Then _
            MidB$(ReplaceMultiple, BuiltStrPos, Diff) = _
                MidB$(Str, LastOccurrence, Diff)
        BuiltStrPos = BuiltStrPos + Diff
        If LenBReplaces(CurrReplaceIdx) <> 0 Then
            MidB$(ReplaceMultiple, BuiltStrPos, LenBReplaces(CurrReplaceIdx)) = _
                Replaces(CurrReplaceIdx)
            BuiltStrPos = BuiltStrPos + LenBReplaces(CurrReplaceIdx)
        End If
        Count = Count + 1
        LastOccurrence = CurrOccurrence + LenBFinds(NextOccsHeap(0, 1))
        If LenBFinds(NextOccsHeap(0, 1)) < 2 Then
            InsertElement(0) = 0
        Else
            InsertElement(0) = InStr((LastOccurrence + 1) \ 2, Str, _
                                    Finds(NextOccsHeap(0, 1)), lCompare) * 2 - 1
        End If
        InsertElement(1) = NextOccsHeap(0, 1)
        If m > n Then
            CurrReplaceIdx = CurrReplaceIdx + n
            If CurrReplaceIdx >= m Then CurrReplaceIdx = NextOccsHeap(0, 1) Mod m
        End If
        InsertElement(2) = CurrReplaceIdx
        GoSub HeapRemoveMin
        If InsertElement(0) > 0 Then GoSub HeapInsert
        CurrOccurrence = NextOccsHeap(0, 0)
        CurrReplaceIdx = NextOccsHeap(0, 2)
    Loop
    
    Dim RemainderStr As String: RemainderStr = MidB$(Str, LastOccurrence)

    If BuiltStrPos <= LenB(ReplaceMultiple) Then _
        MidB$(ReplaceMultiple, BuiltStrPos, LenB(RemainderStr)) = RemainderStr
        
    'Because we didn't take lCount into account when calculating the buffer:
    If Count > lCount Then _
        ReplaceMultiple = _
            LeftB$(ReplaceMultiple, BuiltStrPos + LenB(RemainderStr))

    Exit Function
HeapSwapElements:
    Dim Temp(0 To 2) As Long
    Temp(0) = NextOccsHeap(Index1, 0)
    Temp(1) = NextOccsHeap(Index1, 1)
    Temp(2) = NextOccsHeap(Index1, 2)
    NextOccsHeap(Index1, 0) = NextOccsHeap(Index2, 0)
    NextOccsHeap(Index1, 1) = NextOccsHeap(Index2, 1)
    NextOccsHeap(Index1, 2) = NextOccsHeap(Index2, 2)
    NextOccsHeap(Index2, 0) = Temp(0)
    NextOccsHeap(Index2, 1) = Temp(1)
    NextOccsHeap(Index2, 2) = Temp(2)
    Return
HeapInsert:
    NextOccsHeap(HeapSize, 0) = InsertElement(0)
    NextOccsHeap(HeapSize, 1) = InsertElement(1)
    NextOccsHeap(HeapSize, 2) = InsertElement(2)
    Dim CurrentIndex As Long: CurrentIndex = HeapSize
    Do While CurrentIndex > 0 _
         And NextOccsHeap(CurrentIndex, 0) < _
             NextOccsHeap((CurrentIndex - 1) \ 2, 0) 'ParentNode
        Index1 = CurrentIndex
        Index2 = (CurrentIndex - 1) \ 2 'ParentIndex
        GoSub HeapSwapElements
        CurrentIndex = (CurrentIndex - 1) \ 2 'ParentIndex
    Loop
    HeapSize = HeapSize + 1
    Return
HeapRemoveMin:
    NextOccsHeap(0, 0) = NextOccsHeap(HeapSize - 1, 0)
    NextOccsHeap(0, 1) = NextOccsHeap(HeapSize - 1, 1)
    NextOccsHeap(0, 2) = NextOccsHeap(HeapSize - 1, 2)
    HeapSize = HeapSize - 1
    
    CurrentIndex = 0
    Do
        Dim ILeft As Long: ILeft = 2 * CurrentIndex + 1 'LeftChildIndex
        Dim IRight As Long: IRight = 2 * CurrentIndex + 2 'RightChildIndex
        Dim Smallest As Long: Smallest = CurrentIndex
        
        If ILeft < HeapSize Then
            If NextOccsHeap(ILeft, 0) < NextOccsHeap(Smallest, 0) Then _
                Smallest = ILeft
        End If
        If IRight < HeapSize Then
            If NextOccsHeap(IRight, 0) < NextOccsHeap(Smallest, 0) Then _
                Smallest = IRight
        End If
        
        If Smallest <> CurrentIndex Then
            Index1 = CurrentIndex
            Index2 = Smallest
            GoSub HeapSwapElements
            CurrentIndex = Smallest
        Else
            Exit Do
        End If
    Loop
    Return
End Function

'Same as ReplaceMultiple but parses the string bytewise
Public Function ReplaceMultipleB(ByRef Bytes As String, _
                                 ByRef sFindOrFinds As Variant, _
                                 ByRef sReplaceOrReplaces As Variant, _
                        Optional ByVal lStart As Long = 1, _
                        Optional ByVal lCount As Long = -1, _
                        Optional ByVal lCompare As VbCompareMethod _
                                                 = vbBinaryCompare) As String
    Const methodName As String = "ReplaceMultipleB"
    If lStart < 1 Then err.Raise 5, methodName, _
                               "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then err.Raise 5, methodName, _
                              "Argument 'lCount' = " & lCount & " < -1, invalid"
                              
    If lStart > LenB(Bytes) Then Exit Function
    
    Dim Finds As Variant
    If IsArray(sFindOrFinds) Then Finds = sFindOrFinds _
                             Else Finds = VBA.Array(sFindOrFinds)
    If Not LBound(Finds) = 0 Then _
        ReDim Preserve Finds(0 To UBound(Finds) - LBound(Finds))
    Dim Replaces As Variant
    If IsArray(sReplaceOrReplaces) Then Replaces = sReplaceOrReplaces _
                                   Else Replaces = VBA.Array(sReplaceOrReplaces)
    If Not LBound(Replaces) = 0 Then _
        ReDim Preserve Replaces(0 To UBound(Replaces) - LBound(Replaces))
    lCount = lCount And &H7FFFFFFF
    
    If UBound(Finds) = 0 And Len(Finds(0)) = 0 _
    And UBound(Replaces) = 0 And Len(Replaces(0)) = 0 Then
        ReplaceMultipleB = MidB$(Bytes, lStart)
        Exit Function
    End If
    
    'Clean input arrays to deal with cases where one "find" contains another one
    Dim i As Long, j As Long
    For i = 0 To UBound(Finds)
        If LenB(Finds(i)) <> 0 Then
            For j = i + 1 To UBound(Finds)
                If InStrB(1, Finds(j), Finds(i), lCompare) <> 0 Then
                    Finds(j) = vbNullString
                End If
            Next j
        End If
    Next i
    
    'Allocate buffer
    'Buffer calculation doesn't really take into account the parameter lCount.
    'We'll take it into account at the end of the function instead
    Dim n As Long:          n = UBound(Finds) + 1
    Dim m As Long:          m = UBound(Replaces) + 1
    Dim LenBBuffer As Long: LenBBuffer = LenB(Bytes) - (lStart - 1)
    If m > n Then
        For i = 0 To UBound(Finds)
            Dim NumReplPerFind As Long
            NumReplPerFind = IIf(i < m Mod n, (m \ n) + 1, (m \ n))
            Dim NumOcc As Long
            NumOcc = CountSubstringB(Bytes, CStr(Finds(i)), lStart, lCount, _
                                     lCompare)
            For j = i To m - 1 Step n
                LenBBuffer = LenBBuffer + (LenB(Replaces(j)) - LenB(Finds(i))) _
                             * IIf((j - i) \ n < NumOcc Mod NumReplPerFind _
                         , NumOcc \ NumReplPerFind + 1, NumOcc \ NumReplPerFind)
            Next j
        Next i
    Else
        For i = 0 To UBound(Finds)
            NumOcc = CountSubstringB(Bytes, CStr(Finds(i)), lStart, lCount, _
                                     lCompare)
            LenBBuffer = LenBBuffer + _
                         (LenB(Replaces(i Mod m)) - LenB(Finds(i))) * NumOcc
        Next i
    End If
    
    Dim Buffer() As Byte: ReDim Buffer(1 To LenBBuffer)
    ReplaceMultipleB = Buffer
    
    'Keep track of the next position for replacing by using a min-heap
    Dim NextOccsHeap() As Long: ReDim NextOccsHeap(0 To UBound(Finds), 0 To 2)
    Dim Index1 As Long
    Dim Index2 As Long
    Dim InsertElement(0 To 2) As Long
    Dim HeapSize As Long
    
    Dim LenBReplaces() As Long: ReDim LenBReplaces(0 To UBound(Replaces))
    For i = 0 To UBound(Replaces)
        LenBReplaces(i) = LenB(Replaces(i))
    Next i
    Dim LenBFinds() As Long: ReDim LenBFinds(0 To UBound(Finds))
    For i = 0 To UBound(Finds)
        LenBFinds(i) = LenB(Finds(i))
    Next i
    
    For i = 0 To UBound(Finds)
        Dim NextOcc As Long
        If LenBFinds(i) = 0 Then
            NextOcc = 0
        Else
            NextOcc = InStr(lStart, Bytes, Finds(i), lCompare) * 2 - 1
        End If
        If NextOcc > 0 Then
            InsertElement(0) = NextOcc
            InsertElement(1) = i
            InsertElement(2) = i Mod m
            GoSub HeapInsert
        End If
    Next i
    
    Dim LastOccurrence As Long: LastOccurrence = lStart
    Dim CurrOccurrence As Long: CurrOccurrence = NextOccsHeap(0, 0)
    Dim CurrReplaceIdx As Long: CurrReplaceIdx = NextOccsHeap(0, 2)
    Dim BuiltStrPos As Long:    BuiltStrPos = 1
    Dim Count As Long:          Count = 1
    
    'Do the replacing using InStrB() and MidB$ in a loop
    Do Until HeapSize = 0 Or Count > lCount
        Dim Diff As Long: Diff = CurrOccurrence - LastOccurrence
        If Diff > 0 Then _
            MidB$(ReplaceMultipleB, BuiltStrPos, Diff) = _
                MidB$(Bytes, LastOccurrence, Diff)
        BuiltStrPos = BuiltStrPos + Diff
        If LenBReplaces(CurrReplaceIdx) <> 0 Then
            MidB$(ReplaceMultipleB, BuiltStrPos, LenBReplaces(CurrReplaceIdx)) = _
                Replaces(CurrReplaceIdx)
            BuiltStrPos = BuiltStrPos + LenBReplaces(CurrReplaceIdx)
        End If
        Count = Count + 1
        LastOccurrence = CurrOccurrence + LenBFinds(NextOccsHeap(0, 1))
        If LenBFinds(NextOccsHeap(0, 1)) = 0 Then
            InsertElement(0) = 0
        Else
            InsertElement(0) = InStrB(LastOccurrence, Bytes, _
                                      Finds(NextOccsHeap(0, 1)), lCompare)
        End If
        InsertElement(1) = NextOccsHeap(0, 1)
        If m > n Then
            CurrReplaceIdx = CurrReplaceIdx + n
            If CurrReplaceIdx >= m Then CurrReplaceIdx = NextOccsHeap(0, 1) Mod m
        End If
        InsertElement(2) = CurrReplaceIdx
        GoSub HeapRemoveMin
        If InsertElement(0) > 0 Then GoSub HeapInsert
        CurrOccurrence = NextOccsHeap(0, 0)
        CurrReplaceIdx = NextOccsHeap(0, 2)
    Loop

    Dim RemainderStr As String: RemainderStr = MidB$(Bytes, LastOccurrence)
    
    If BuiltStrPos <= LenB(ReplaceMultipleB) Then _
        MidB$(ReplaceMultipleB, BuiltStrPos, LenB(RemainderStr)) = RemainderStr
        
    'Because we didn't take lCount into account when calculating the buffer:
    If Count > lCount Then _
        ReplaceMultipleB = _
            LeftB$(ReplaceMultipleB, BuiltStrPos + LenB(RemainderStr))
       
    Exit Function
HeapSwapElements:
    Dim Temp(0 To 2) As Long
    Temp(0) = NextOccsHeap(Index1, 0)
    Temp(1) = NextOccsHeap(Index1, 1)
    Temp(2) = NextOccsHeap(Index1, 2)
    NextOccsHeap(Index1, 0) = NextOccsHeap(Index2, 0)
    NextOccsHeap(Index1, 1) = NextOccsHeap(Index2, 1)
    NextOccsHeap(Index1, 2) = NextOccsHeap(Index2, 2)
    NextOccsHeap(Index2, 0) = Temp(0)
    NextOccsHeap(Index2, 1) = Temp(1)
    NextOccsHeap(Index2, 2) = Temp(2)
    Return
HeapInsert:
    NextOccsHeap(HeapSize, 0) = InsertElement(0)
    NextOccsHeap(HeapSize, 1) = InsertElement(1)
    NextOccsHeap(HeapSize, 2) = InsertElement(2)
    Dim CurrentIndex As Long: CurrentIndex = HeapSize
    Do While CurrentIndex > 0 _
         And NextOccsHeap(CurrentIndex, 0) < _
             NextOccsHeap((CurrentIndex - 1) \ 2, 0) 'ParentNode
        Index1 = CurrentIndex
        Index2 = (CurrentIndex - 1) \ 2 'ParentIndex
        GoSub HeapSwapElements
        CurrentIndex = (CurrentIndex - 1) \ 2 'ParentIndex
    Loop
    HeapSize = HeapSize + 1
    Return
HeapRemoveMin:
    NextOccsHeap(0, 0) = NextOccsHeap(HeapSize - 1, 0)
    NextOccsHeap(0, 1) = NextOccsHeap(HeapSize - 1, 1)
    NextOccsHeap(0, 2) = NextOccsHeap(HeapSize - 1, 2)
    HeapSize = HeapSize - 1
    
    CurrentIndex = 0
    Do
        Dim ILeft As Long: ILeft = 2 * CurrentIndex + 1 'LeftChildIndex
        Dim IRight As Long: IRight = 2 * CurrentIndex + 2 'RightChildIndex
        Dim Smallest As Long: Smallest = CurrentIndex
        
        If ILeft < HeapSize Then
            If NextOccsHeap(ILeft, 0) < NextOccsHeap(Smallest, 0) Then _
                Smallest = ILeft
        End If
        If IRight < HeapSize Then
            If NextOccsHeap(IRight, 0) < NextOccsHeap(Smallest, 0) Then _
                Smallest = IRight
        End If
        
        If Smallest <> CurrentIndex Then
            Index1 = CurrentIndex
            Index2 = Smallest
            GoSub HeapSwapElements
            CurrentIndex = Smallest
        Else
            Exit Do
        End If
    Loop
    Return
End Function

'Similar to ReplaceMultiple, but here, replacements will be performed one after
'another in the order of the input arrays and not all in a single pass like in
'the regular ReplaceMultiple.
'In this version lCount specifies the maximum number of replacements per find
'value, and not in total like in the regular ReplaceMultiple.
'Also, this version does not support multiple different replace values for a
'given find value, this means numFinds must be >= numReplaces.
Public Function ReplaceMultipleMultiPass(ByRef Str As String, _
                                         ByRef sFindOrFinds As Variant, _
                                         ByRef sReplaceOrReplaces As Variant, _
                                Optional ByVal lStart As Long = 1, _
                                Optional ByVal lCount As Long = -1, _
                                Optional ByVal lCompare As VbCompareMethod _
                                                         = vbBinaryCompare) _
                                         As String
    Const methodName As String = "ReplaceMultipleMultiPass"
    If lStart < 1 Then err.Raise 5, methodName, _
                               "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then err.Raise 5, methodName, _
                              "Argument 'lCount' = " & lCount & " < -1, invalid"
    
    Dim Finds As Variant
    If IsArray(sFindOrFinds) Then Finds = sFindOrFinds _
                             Else Finds = VBA.Array(sFindOrFinds)
    If Not LBound(Finds) = 0 Then _
        ReDim Preserve Finds(0 To UBound(Finds) - LBound(Finds))
    Dim Replaces As Variant
    If IsArray(sReplaceOrReplaces) Then Replaces = sReplaceOrReplaces _
                                   Else Replaces = VBA.Array(sReplaceOrReplaces)
    If Not LBound(Replaces) = 0 Then _
        ReDim Preserve Replaces(0 To UBound(Replaces) - LBound(Replaces))
        
    If UBound(Replaces) > UBound(Finds) Then err.Raise 5, methodName, "'sFind" _
        & "OrFinds' must have at least as many elements as 'sReplaceOrReplaces'"
    
    lCount = lCount And &H7FFFFFFF
    
    If lStart > Len(Str) Then Exit Function
    
    ReplaceMultipleMultiPass = Mid$(Str, lStart)
    If UBound(Finds) = 0 And Len(Finds(0)) = 0 _
    And UBound(Replaces) = 0 And Len(Replaces(0)) = 0 Then Exit Function
    
    Dim NumReplaces As Long: NumReplaces = UBound(Replaces) + 1
    
    Dim i As Long
    For i = 0 To UBound(Finds)
        ReplaceMultipleMultiPass = Replace(ReplaceMultipleMultiPass, Finds(i), _
                                Replaces(i Mod NumReplaces), , lCount, lCompare)
    Next i
End Function

'This function can replace multiple values with multiple different replace
'values in each element of an array or just in a regular string.
'E.g.: ArrayReplaceMultiple("ab", Array("a", "b"), Array("c", "d")) returns "cd"
'Or: ArrayReplaceMultiple(Array("ab", "ab"), Array("a", "b"), Array("c", "d"))
'returns an array with two elements: ("cd", "cd")
Public Function ArrayReplaceMultiple(ByVal strOrStrArr As Variant, _
                                     ByVal findOrFinds As Variant, _
                                     ByVal replaceOrReplaces As Variant, _
                            Optional ByVal compareMethod As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                     As Variant
    Const methodName As String = "ReplaceMultiple"
    If IsArray(findOrFinds) Then If Not IsArray(replaceOrReplaces) Then _
        err.Raise 5, methodName, "Finds and Replaces must both be array or not."
    If IsArray(findOrFinds) Then
        If Not UBound(findOrFinds) - LBound(findOrFinds) = _
               UBound(replaceOrReplaces) - LBound(replaceOrReplaces) Then
            err.Raise 5, methodName, _
                "There must be the same number of find and replace values"
        End If
    Else
        Dim TmpArr As Variant: ReDim TmpArr(0 To 0)
        TmpArr(0) = findOrFinds
        findOrFinds = TmpArr
        TmpArr(0) = replaceOrReplaces
        replaceOrReplaces = TmpArr
    End If
    Dim i As Long, j As Long, k As Long
    If IsArray(strOrStrArr) Then
        For i = LBound(strOrStrArr) To UBound(strOrStrArr)
            k = LBound(replaceOrReplaces)
            For j = LBound(findOrFinds) To UBound(findOrFinds)
                strOrStrArr(i) = Replace(strOrStrArr(i), findOrFinds(j), _
                                        replaceOrReplaces(k), , , compareMethod)
                k = k + 1
            Next j
        Next i
    Else
        k = LBound(replaceOrReplaces)
        For j = LBound(findOrFinds) To UBound(findOrFinds)
            strOrStrArr = Replace(strOrStrArr, findOrFinds(j), _
                                  replaceOrReplaces(k), , , compareMethod)
            k = k + 1
        Next j
    End If
    ArrayReplaceMultiple = strOrStrArr
End Function

'This function splits a string into a given number of chunks of a given length.
'If a number of chunks and a chunkLength is specified, it creates chunks of the
'given length until either the number of chunks is reached, or the string ends.
'If no number of chunks is specified, the entire string will chunkified into
'chunks of length chunkLength.
'If no chunkLength is specified, the chunk length will be automatically
'calculated as Len(str) / numberOfChunks.
'If no chunkLength and no number of chunks is specified, the string will be
'split into parts of length 1.
'If discardIncompleteChunks = False, tha last chunk can have a different length
'than specified if Len(str) is not divisible by chunkLength
'If splitUTF16Surrogates = False, UTF-16 surrogate pairs will always remain
'together, this means that Len(chunk) can differ from the specified chunkLength
'by up to 1 in that case.
Public Function ChunkifyString(ByRef Str As String, _
                      Optional ByVal numberOfChunks As Long = 0, _
                      Optional ByVal chunkLength As Long = 0, _
                      Optional ByVal discardIncompleteChunks As Boolean = False, _
                      Optional ByVal splitUTF16Surrogates As Boolean = True) _
                               As String()
    Dim LenStr As Long: LenStr = Len(Str)

    If chunkLength = 0 And numberOfChunks = 0 Then
        chunkLength = 1
    ElseIf chunkLength = 0 And numberOfChunks > 0 Then
        chunkLength = Len(Str) \ numberOfChunks
        If chunkLength = 0 Then chunkLength = 1
    End If
    
    If numberOfChunks = 0 Then
        numberOfChunks = Len(Str) \ chunkLength
        If Len(Str) Mod chunkLength > 0 Then numberOfChunks = numberOfChunks + 1
    End If

    Dim Chunks() As String: ReDim Chunks(0 To numberOfChunks - 1)
    
    Dim CurrChunkLength As Long
    Dim ChunkIndex As Long
    Dim Position As Long:           Position = 1
    
    For ChunkIndex = 0 To numberOfChunks - 1
        If Position > LenStr Then Exit For
        
        CurrChunkLength = chunkLength
        If Not splitUTF16Surrogates _
        And Position + CurrChunkLength - 1 < LenStr Then
            If AscU(Mid$(Str, Position + CurrChunkLength - 1, 2)) > &HFFFF& Then
                CurrChunkLength = CurrChunkLength - 1
            End If
        End If
        
        If Position + CurrChunkLength - 1 > LenStr Then
            If discardIncompleteChunks Then Exit For
            CurrChunkLength = LenStr - Position + 1
        End If
        
        Chunks(ChunkIndex) = Mid$(Str, Position, CurrChunkLength)
        Position = Position + CurrChunkLength
    Next ChunkIndex
    
    'If the last chunk was not used, shrink the array
    If Position >= LenStr And discardIncompleteChunks Then
        ReDim Preserve Chunks(0 To ChunkIndex - 1)
    End If
    
    ChunkifyString = Chunks
End Function

'This function can convert any variable into a convenient string for printing.
'The last 6 parameters are only relevant for arrays, the last 2 only for
'two dimensional arrays.
''maxChars' is the maximum length of the returned value
''escapeNonPrintable = True' will convert all non ANSI chars to escape sequences
''delimiter' can be used to manually define a delimiter for the returned arrays.
''maxCharsPerLine' will make the output contain newLine characters at least
'                  every 'maxCharsPerLine' characters.
''inklColIndices = True' will print column indices above the columns if the
'                        input is a two dimensional array.
''inklRowIndices = True' does the same for the row indices.
'TODO: `ToString` doesn't comply to its settings with the desired rigour.
'Problem Example:
'Multiple recursion of nested 1d arrays will ignor the line length limit.
Public Function ToString(ByVal Value As Variant, _
                Optional ByVal maxChars As Long = 0, _
                Optional ByVal escapeNonPrintable As Boolean = True, _
                Optional ByRef Delimiter As String = vbNullString, _
                Optional ByVal maxCharsPerElement As Long = 25, _
                Optional ByVal maxCharsPerLine As Long = 80, _
                Optional ByVal maxLines As Long = 10, _
                Optional ByVal inklColIndices As Boolean = True, _
                Optional ByVal inklRowIndices As Boolean = True)
    Const methodName As String = "ToString"
    
    If maxChars < 0 Then _
        err.Raise 5, methodName, "'maxChars' can't be < 0"
    If maxCharsPerElement < 0 Then _
        err.Raise 5, methodName, "'maxCharsPerElement' can't be < 0"
    If maxCharsPerLine < 0 Then _
        err.Raise 5, methodName, "'maxCharsPerLine' can't be < 0"
    If maxLines < 0 Then _
        err.Raise 5, methodName, "'maxLines' can't be < 0"
    
    If maxChars = 0 Then maxChars = &H7FFFFFFF
    If maxCharsPerElement = 0 Then maxCharsPerElement = &H7FFFFFFF
    If maxCharsPerLine = 0 Then maxCharsPerLine = &H7FFFFFFF
    If maxLines = 0 Then maxLines = &H7FFFFFFF
    
    Dim Settings As StringificationSettings
    With Settings
        .maxChars = maxChars
        .escapeNonPrintable = escapeNonPrintable
        .Delimiter = Delimiter
        .maxCharsPerElement = maxCharsPerElement
        .maxCharsPerLine = maxCharsPerLine
        .maxLines = maxLines
        .inklColIndices = inklColIndices
        .inklRowIndices = inklRowIndices
    End With
    ToString = BToString(Value, Settings)
End Function

'Recursive "Backend" function for 'ToString'
Private Function BToString(ByVal Value As Variant, _
                           ByRef Settings As StringificationSettings) As String
    'Don't use exit function in this Function! Instead use: GoTo CleanExit
    Static isRecursiveCall As Boolean
    Dim WasRecursiveCall As Boolean: WasRecursiveCall = isRecursiveCall
    isRecursiveCall = True

    Dim s As String
    Dim i As Long
    If IsArray(Value) Then
        Select Case GetArrayDimsCount(Value)
            Case 0 'Array uninitialized
                s = TypeName(Value)
            Case 1
                If WasRecursiveCall Then Settings.maxCharsPerLine = &H7FFFFFFF
                s = ToString1dArray(Value, Settings)
            Case 2
                If WasRecursiveCall Then
                    s = ToStringMultiDimArray(Value) 'No settings required
                Else
                    s = ToString2dimArray(Value, Settings)
                End If
            Case Else
                s = ToStringMultiDimArray(Value) 'No settings required
        End Select
    ElseIf IsObject(Value) Then
        Select Case True
            'Can add custom logic to ToString any object here
            'Case TypeOf value Is Collection
                's = ToStringCollection(...
            Case Else
                s = TypeName(Value)
        End Select
    ElseIf IsEmpty(Value) Then
        s = "Empty"
    Else
        s = CStr(Value)
    End If
    
    With Settings
        If Len(s) > .maxChars Then
            BToString = Left(s, Max(.maxChars - 3, 0)) & Left("...", .maxChars)
        Else
            BToString = s
        End If
        
        If .escapeNonPrintable Then BToString = EscapeUnicode(BToString, 255)
    End With
    If VarType(Value) = vbString Then BToString = "'" & BToString & "'"
    
CleanExit:
    If Not WasRecursiveCall Then isRecursiveCall = False
End Function

'Utility function for 'BToString'
'Note:
''maxChars' is only passed to exit the loop sooner in some cases
Private Function ToString1dArray(ByRef Arr As Variant, _
                                 ByRef Settings As StringificationSettings) _
                                 As String
    Dim s As String
    Dim Delimiter As String: Delimiter = Settings.Delimiter
    If StrPtr(Delimiter) = 0 Then Delimiter = ", "
    
    If UBound(Arr) - LBound(Arr) = -1 Then
        ToString1dArray = "[]"
        Exit Function
    ElseIf UBound(Arr) - LBound(Arr) = 0 Then
        ToString1dArray = "[" & BToString(Arr(UBound(Arr)), Settings) & "]"
        Exit Function
    End If

    Dim Line As String: Line = "["
    Dim i As Long
    For i = LBound(Arr) To UBound(Arr) - 1
        Line = Line & BToString(Arr(i), Settings) & Delimiter

       'Check if max characters per line would be exceeded with the next element
        If Len(Line & BToString(Arr(i + 1), Settings) & _
               Delimiter) >= Settings.maxCharsPerLine Then
            s = s & Line & vbNewLine
            Line = ""
            Dim LineCount As Long: LineCount = LineCount + 1
            If LineCount >= Settings.maxLines _
            Or Len(s) > Settings.maxChars Then Exit For
        End If
    Next i
    ToString1dArray = s & Line & BToString(Arr(UBound(Arr)), Settings) & "]"
End Function

'Utility function for 'BToString'
'Note: Will only be called in non-recursive calls of 'BToString'
Private Function ToString2dimArray(ByRef Arr As Variant, _
                                   ByRef Settings As StringificationSettings) _
                                   As String
    Dim Delimiter As String: Delimiter = Settings.Delimiter
    If StrPtr(Settings.Delimiter) = 0 Then Delimiter = "  "
    
    With Settings
        Dim colWidths() As Long
        colWidths = CalculateColumnWidths(Arr, Settings)
        Dim NumCols As Long
        NumCols = CalculateNumColumnsToFit(colWidths, .maxCharsPerLine, _
                                           Len(Delimiter))
    
        Dim NumRows As Long: NumRows = UBound(Arr, 1) - LBound(Arr, 1) + 1
        Dim FirstRows As Long: FirstRows = Min(.maxLines \ 2, NumRows)
        Dim LastRows As Long
        LastRows = Min(.maxLines - FirstRows, NumRows - FirstRows)
    End With
    Dim s As String

    If Settings.inklColIndices Then _
        s = BuildColHeadersLine(Arr, colWidths, NumCols, Delimiter, Settings) _
            & vbNewLine
                          
    Dim i As Long
    For i = LBound(Arr, 1) To LBound(Arr, 1) + FirstRows - 1
        If Len(s) > Settings.maxChars Then Exit For
        s = s & BuildLine(Arr, colWidths, NumCols, Delimiter, i, Settings) _
            & vbNewLine
    Next i
    
    If NumRows > Settings.maxLines And Len(s) < Settings.maxChars Then _
        s = s & BuildDotsLine(Arr, colWidths, NumCols, Delimiter, Settings) _
            & vbNewLine
    
    For i = UBound(Arr, 1) - LastRows + 1 To UBound(Arr, 1)
        If Len(s) > Settings.maxChars Then Exit For
        s = s & BuildLine(Arr, colWidths, NumCols, Delimiter, i, Settings) _
            & vbNewLine
    Next i
    ToString2dimArray = s & "(" & NumRows & "*" & _
                         UBound(Arr, 2) - LBound(Arr, 2) + 1 & ", " & _
                         " " & Min(NumRows, Settings.maxLines) & _
                         "*" & NumCols & " output)"
End Function

Private Function ToStringMultiDimArray(ByRef Arr As Variant) As String
    Dim s As String
    s = "Array("
    Dim i As Long
    For i = 1 To GetArrayDimsCount(Arr)
        s = s & LBound(Arr, i) & " to " & UBound(Arr, i) & ", "
    Next i
    ToStringMultiDimArray = Left(s, Len(s) - 2) & ")"
End Function

'Utility function for 'ToString2dimArray'
Private Function BuildColHeadersLine(ByRef Arr As Variant, _
                                     ByRef colWidths() As Long, _
                                     ByVal NumCols As Long, _
                                     ByVal Delimiter As String, _
                                     ByRef Settings As StringificationSettings) _
                                     As String
    Dim RowNumPadding As Long
    RowNumPadding = Max(Len(CStr(UBound(Arr, 1))), Len(CStr(LBound(Arr, 1)))) + 2
    
    Dim LenDelim As Long: LenDelim = Len(Delimiter)
    Dim j As Long
    If NumCols = UBound(Arr, 2) - LBound(Arr, 2) + 1 Then
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            BuildColHeadersLine = BuildColHeadersLine & _
                                  PadLeft(CStr(j), colWidths(j)) _
                                  & Delimiter
        Next j
        If Settings.inklRowIndices Then _
            BuildColHeadersLine = Space(RowNumPadding) & BuildColHeadersLine

        BuildColHeadersLine = Left(BuildColHeadersLine, _
                                   Len(BuildColHeadersLine) - LenDelim)
        Exit Function
    End If
    
    Dim LeftPart As String, RightPart As String
    For j = 0 To NumCols \ 2 - 1 'numCols is always even
        LeftPart = LeftPart & PadLeft(CStr(LBound(Arr, 2) + j), _
            colWidths(LBound(Arr, 2) + j)) & Space(LenDelim)
        RightPart = Space(LenDelim) & PadLeft(CStr(UBound(Arr, 2) - j), _
                    colWidths(UBound(Arr, 2) - j)) & RightPart
    Next j

    If Settings.inklRowIndices Then _
        BuildColHeadersLine = Space(RowNumPadding)

    BuildColHeadersLine = BuildColHeadersLine & _
                          Left(LeftPart, Len(LeftPart) - LenDelim) & _
                          " ... " & Right(RightPart, Len(RightPart) - LenDelim)
End Function

'Utility function for 'ToString2dimArray'
Private Function BuildDotsLine(ByRef Arr As Variant, _
                               ByRef colWidths() As Long, _
                               ByVal NumCols As Long, _
                               ByVal Delimiter As String, _
                               ByRef Settings As StringificationSettings) _
                               As String
    Dim RowNumPadding As Long
    RowNumPadding = Max(Len(CStr(UBound(Arr, 1))), Len(CStr(LBound(Arr, 1)))) + 2
    Dim LenDelim As Long: LenDelim = Len(Delimiter)
    Dim j As Long
    If NumCols = UBound(Arr, 2) - LBound(Arr, 2) + 1 Then
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            BuildDotsLine = BuildDotsLine & PadLeft("...", colWidths(j)) _
                                          & Delimiter
        Next j
        If Settings.inklRowIndices Then _
                BuildDotsLine = PadRight("..", RowNumPadding) & BuildDotsLine

        BuildDotsLine = Left(BuildDotsLine, Len(BuildDotsLine) - LenDelim)
        Exit Function
    End If
    
    Dim LeftPart As String, RightPart As String
    For j = 0 To NumCols \ 2 - 1 'numCols is always even
        LeftPart = LeftPart & PadLeft("...", colWidths(LBound(Arr, 2) + j)) _
                   & Space(LenDelim)
        RightPart = Space(LenDelim) & PadLeft("...", _
                    colWidths(UBound(Arr, 2) - j)) & RightPart
    Next j

    If Settings.inklRowIndices Then _
        BuildDotsLine = PadRight("..", RowNumPadding)

    BuildDotsLine = BuildDotsLine & Left(LeftPart, Len(LeftPart) - LenDelim) & _
                    " ... " & Right(RightPart, Len(RightPart) - LenDelim)
End Function

'Utility function for 'ToString2dimArray'
Private Function BuildLine(ByRef Arr As Variant, _
                           ByRef colWidths() As Long, _
                           ByVal NumCols As Long, _
                           ByVal Delimiter As String, _
                           ByVal RowIndex As Long, _
                           ByRef Settings As StringificationSettings) As String
    Dim RowNumPadding As Long
    RowNumPadding = Max(Len(CStr(UBound(Arr, 1))), Len(CStr(LBound(Arr, 1)))) + 2
    
    Dim LenDelim As Long: LenDelim = Len(Delimiter)
    Dim j As Long
    If NumCols = UBound(Arr, 2) - LBound(Arr, 2) + 1 Then
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            BuildLine = BuildLine & _
                        PadLeft(BToString(Arr(RowIndex, j), Settings), _
                                colWidths(j)) & Delimiter
        Next j
        If Settings.inklRowIndices Then _
            BuildLine = PadRight(CStr(RowIndex), RowNumPadding) & BuildLine

        BuildLine = Left(BuildLine, Len(BuildLine) - LenDelim)
        Exit Function
    End If
    
    Dim LeftPart As String, RightPart As String
    For j = 0 To NumCols \ 2 - 1 'numCols is always even
        LeftPart = LeftPart & PadLeft(BToString(Arr(RowIndex, _
                   LBound(Arr, 2) + j), Settings), _
                   colWidths(LBound(Arr, 2) + j)) & Delimiter
        RightPart = Delimiter & PadLeft(BToString(Arr(RowIndex, _
                    UBound(Arr, 2) - j), Settings), _
                    colWidths(UBound(Arr, 2) - j)) & RightPart
    Next j

    If Settings.inklRowIndices Then _
        BuildLine = PadRight(CStr(RowIndex), RowNumPadding)

    BuildLine = BuildLine & Left(LeftPart, Len(LeftPart) - LenDelim) & _
                " ... " & Right(RightPart, Len(RightPart) - LenDelim)
End Function

'Utility function for 'ToString2dimArray'
Private Function CalculateColumnWidths(ByRef Arr As Variant, _
                                       ByRef Settings As StringificationSettings) _
                                       As Long()
    Dim colWidths() As Long: ReDim colWidths(LBound(Arr, 2) To UBound(Arr, 2))
    Dim i As Long, j As Long
    
    ' Calculate how many rows will be printed before and after the dots
    Dim NumRows As Long: NumRows = UBound(Arr, 1) - LBound(Arr, 1) + 1
    Dim NumCols As Long: NumCols = UBound(Arr, 2) - LBound(Arr, 2) + 1
    Dim FirstRows As Long: FirstRows = Min(Settings.maxLines \ 2, NumRows)
    Dim LastRows As Long
    LastRows = Min(Settings.maxLines - FirstRows, NumRows - FirstRows)
    
    Dim SumWidths As Long: SumWidths = 0
    For j = 0 To NumCols \ 2
        Dim Col1 As Long: Col1 = LBound(Arr, 2) + j
        Dim Col2 As Long: Col2 = UBound(Arr, 2) - j
        'if col1>
        colWidths(Col1) = Max(colWidths(Col1), Len(CStr(Col1))) 'Column labels
        colWidths(Col2) = Max(colWidths(Col2), Len(CStr(Col2))) 'Column labels
        For i = LBound(Arr, 1) To LBound(Arr, 1) + FirstRows - 1
            colWidths(Col1) = Max(Len(BToString(Arr(i, Col1), Settings)), _
                                  colWidths(Col1))
            colWidths(Col2) = Max(Len(BToString(Arr(i, Col2), Settings)), _
                                  colWidths(Col2))
        Next i
        SumWidths = SumWidths + colWidths(Col1) + colWidths(Col2)
        If SumWidths > Settings.maxCharsPerLine Then Exit For
    Next j
    
    If Settings.maxLines < NumRows Then
        SumWidths = 0
        For j = 0 To NumCols \ 2
            Col1 = LBound(Arr, 2) + j
            Col2 = UBound(Arr, 2) - j
            For i = UBound(Arr, 1) - LastRows + 1 To UBound(Arr, 1)
                colWidths(Col1) = Max(Len(BToString(Arr(i, Col1), Settings)), _
                                      colWidths(Col1))
                colWidths(Col2) = Max(Len(BToString(Arr(i, Col2), Settings)), _
                                      colWidths(Col2))
            Next i
            SumWidths = SumWidths + colWidths(Col1) + colWidths(Col2)
            If SumWidths > Settings.maxCharsPerLine Then Exit For
        Next j
    End If

    CalculateColumnWidths = colWidths
End Function

'Utility function for 'ToString2dimArray'
Private Function CalculateNumColumnsToFit(ByRef colWidths() As Long, _
                                          ByVal maxCharsPerLine As Long, _
                                          ByVal delimLength As Long) As Long
    Dim TotalWidth As Long
    Dim NumCols As Long:     NumCols = UBound(colWidths) - LBound(colWidths) + 1
    Dim ExtraWidthForDots As Long: ExtraWidthForDots = 3 ' Width for the "..."
    
    If Sum(colWidths) + delimLength * (NumCols - 1) <= maxCharsPerLine Then
        CalculateNumColumnsToFit = NumCols
        Exit Function
    End If
    
    'Else, omitt at least one column:
    Dim j As Long
    For j = LBound(colWidths) To UBound(colWidths) \ 2
        TotalWidth = TotalWidth + colWidths(j) + _
                     colWidths(UBound(colWidths) - (j - LBound(colWidths))) + _
                     delimLength * 2
        If TotalWidth + ExtraWidthForDots <= maxCharsPerLine Then
            CalculateNumColumnsToFit = CalculateNumColumnsToFit + 2
        Else
            Exit For
        End If
    Next j
End Function

Private Function Sum(ParamArray p() As Variant) As Variant
    Dim v As Variant
    For Each v In p
        If IsArray(v) Then
            Dim E As Variant
            For Each E In v
                If IsArray(E) Then
                    Sum = Sum + Sum(E)
                Else
                    If IsNumeric(E) Then Sum = Sum + E
                End If
            Next E
        End If
        If IsNumeric(v) Then Sum = Sum + v
    Next v
End Function

Private Function Max(a As Long, b As Long) As Long
    If a > b Then Max = a Else Max = b
End Function

Private Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function

'Sets the formatting rules adhered to by 'Printf'
Public Sub SetPrintfSettings(Optional ByVal maxChars As Long = 0, _
                          Optional ByVal escapeNonPrintable As Boolean = True, _
                            Optional ByRef Delimiter As String = vbNullString, _
                             Optional ByVal maxCharsPerElement As Long = 25, _
                             Optional ByVal maxCharsPerLine As Long = 80, _
                             Optional ByVal maxLines As Long = 10, _
                             Optional ByVal inklColIndices As Boolean = True, _
                             Optional ByVal inklRowIndices As Boolean = True)
    Const methodName As String = "SetPrintfSettings"
    
    If maxChars < 0 Then _
        err.Raise 5, methodName, "'maxChars' can't be < 0"
    If maxCharsPerElement < 0 Then _
        err.Raise 5, methodName, "'maxCharsPerElement' can't be < 0"
    If maxCharsPerLine < 0 Then _
        err.Raise 5, methodName, "'maxCharsPerLine' can't be < 0"
    If maxLines < 0 Then _
        err.Raise 5, methodName, "'maxLines' can't be < 0"
    
    If maxChars = 0 Then maxChars = &H7FFFFFFF
    If maxCharsPerElement = 0 Then maxCharsPerElement = &H7FFFFFFF
    If maxCharsPerLine = 0 Then maxCharsPerLine = &H7FFFFFFF
    If maxLines = 0 Then maxLines = &H7FFFFFFF
    
    With PrintfSettings
        .maxChars = maxChars
        .escapeNonPrintable = escapeNonPrintable
        .Delimiter = Delimiter
        .maxCharsPerElement = maxCharsPerElement
        .maxCharsPerLine = maxCharsPerLine
        .maxLines = maxLines
        .inklColIndices = inklColIndices
        .inklRowIndices = inklRowIndices
    End With
End Sub

'Utility function for 'Printf'
Private Function GetPrintfSettings() As StringificationSettings
    If Not PrintfSettingsAreInitialized Then
        SetPrintfSettings 'With only default arguments
        PrintfSettingsAreInitialized = True
    End If
    GetPrintfSettings = PrintfSettings
End Function

'Prints any variables passed to this function. Uses formatting rules previously
'set with 'SetPrintfSettings'
Public Function Printf(ParamArray args() As Variant) As String
    Dim Arg As Variant
    Dim SArg As Variant
    Dim s As String
    Dim Settings As StringificationSettings: Settings = GetPrintfSettings
    For Each Arg In args
        If VarType(Arg) = vbString Then
            SArg = Arg
        Else
            SArg = BToString(Arg, Settings)
        End If
        If InStr(1, SArg, vbNewLine, vbBinaryCompare) <> 0 Then
            s = s & vbNewLine & SArg & vbNewLine
        ElseIf Len(s) - InStrRev(s, vbNewLine, , vbBinaryCompare) + Len(SArg) _
               > Settings.maxCharsPerLine - 2 Then
            s = s & vbNewLine & SArg & "  "
        Else
            s = s & SArg & "  "
        End If
        If Len(s) > Settings.maxChars Then
            s = s & "..."
            Exit For
        End If
    Next Arg
    Printf = TrimX(s, vbCrLf)
    Debug.Print Printf
End Function

'Prints an one or two dimensional array to the immediate window.
Public Sub PrintVar(ByRef Arr As Variant, _
           Optional ByRef Delimiter As String = vbNullString, _
           Optional ByVal maxCharsPerElement As Long = 25, _
           Optional ByVal maxCharsPerLine As Long = 80, _
           Optional ByVal maxLines As Long = 10, _
           Optional ByVal escapeNonPrintable As Boolean = True, _
           Optional ByVal printColIndices As Boolean = True, _
           Optional ByVal printRowIndices As Boolean = True)
    Debug.Print ToString(Arr, , escapeNonPrintable, Delimiter, _
                          maxCharsPerElement, maxCharsPerLine, maxLines, _
                          printColIndices, printRowIndices)
End Sub

'Works like the inbuilt trim but instead of just spaces, it will trim any
'characters occurring in 'charactersToTrim' from the edges of 'str'
Public Function TrimX(ByRef Str As String, _
             Optional ByRef charactersToTrim As String = " " & vbCrLf & vbTab) _
                      As String
    If Len(Str) = 0 Then Exit Function
    Dim StrLen As Long:   StrLen = Len(Str)
    Dim StartIdx As Long: StartIdx = 1
    Dim EndIdx As Long:   EndIdx = StrLen

    Do While StartIdx <= StrLen _
         And InStr(charactersToTrim, Mid(Str, StartIdx, 1)) > 0
        StartIdx = StartIdx + 1
    Loop

    Do While EndIdx >= 1
        If InStr(charactersToTrim, Mid(Str, EndIdx, 1)) > 0 Then
            EndIdx = EndIdx - 1
        Else
            Exit Do
        End If
    Loop

    If StartIdx <= EndIdx Then
        TrimX = Mid(Str, StartIdx, EndIdx - StartIdx + 1)
    Else
        TrimX = ""
    End If
End Function

'Returns the index of a given column letter in excel
Public Function ColLetterToNumber(ByRef colLetterOrNumber As Variant) As Long
    Const methodName As String = "ColLetterToNumber"
    
    If IsNumeric(colLetterOrNumber) Then
        ColLetterToNumber = CLng(colLetterOrNumber)
        If ColLetterToNumber > 16384 Then _
            err.Raise 5, methodName, _
                "Only values up to 'XFD', or 16384 are allowed."
    ElseIf Len(colLetterOrNumber) <= 3 _
    And (CStr(UCase(colLetterOrNumber)) Like "[A-Z]" _
    Or CStr(UCase(colLetterOrNumber)) Like "[A-Z][A-Z]" _
    Or CStr(UCase(colLetterOrNumber)) Like "[A-Z][A-Z][A-Z]") Then
        Dim i As Long
        For i = 1 To Len(colLetterOrNumber)
            ColLetterToNumber = ColLetterToNumber * 26 + _
                (Asc(UCase(Mid(colLetterOrNumber, i, 1))) - 65 + 1)
        Next i
        If ColLetterToNumber > 16384 Then _
            err.Raise 5, methodName, _
                "Only values up to 'XFD', or 16384 are allowed."
    Else
        err.Raise 5, methodName, "The input must be up to 3 letters or 5 " & _
            "digits long. Only values up to 'XFD', or 16384 are allowed and" & _
            " no mixture of letters and digits is allowed."
    End If
End Function

'Converts a string to date based on a format specified in 'dateFormat'
'E.g. DateParser("27.04.1993", "DD.MM.YYYY") = DateParser("042793", "MMDDYY")
'Follows the idea by Scott Craner: https://stackoverflow.com/a/64813581/12287457
Public Function DateParser(ByRef Str As String, _
                           ByRef dateFormat As String) As Date
    Const methodName As String = "DateParser"
    
    If Len(Str) <> Len(dateFormat) Then err.Raise 5, methodName, _
        "The input string must be of same length as the format string."
    
    Dim i As Long
    For i = 1 To Len(Str)
        If UCase(Mid(dateFormat, i, 1)) = "D" Then
            Dim LDay As Long: LDay = LDay * 10 + CLng(Mid(Str, i, 1))
        ElseIf UCase(Mid(dateFormat, i, 1)) = "Y" Then
            Dim LYear As Long: LYear = LYear * 10 + CLng(Mid(Str, i, 1))
        ElseIf UCase(Mid(dateFormat, i, 1)) = "M" Then
            Dim SMonth As String: SMonth = SMonth & Mid(Str, i, 1)
        End If
    Next i
    
    If IsNumeric(SMonth) Then
        Dim LMonth As Long: LMonth = CLng(SMonth)
    Else
        LMonth = Month(CDate("01 " & SMonth & " 2023"))
    End If
    
    DateParser = DateSerial(LYear, LMonth, LDay)
End Function
