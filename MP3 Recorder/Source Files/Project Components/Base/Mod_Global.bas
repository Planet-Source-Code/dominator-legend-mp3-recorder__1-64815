Attribute VB_Name = "Mod_Global"
Option Explicit
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __              ___            ___          |###'
'###|        | ____ \              |  |             \  \          /  /          |###'
'###|        | |   \ \             |  |              \  \        /  /           |###'
'###|        | |    \ \            |  |               \  \      /  /            |###'
'###|        | |    / /            |  |                \  \    /  /             |###'
'###|        | |___/ /     __      |  |______    __     \  \__/  /              |###'
'###|        |______/     (__)     |_________|  (__)     \______/               |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|  Global Variables   |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Varaiables used for visual components of the main form
        Global DisplayType                  As Integer
        Global DigialVolume                 As Double
        Global FileNum                      As Integer
        Global Duration                     As Long
        Global RecordSize                   As Long
'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                   ________               |###'
'###|        | ____ \              |  |                 |  ______|              |###'
'###|        | |   \ \             |  |                 | |                     |###'
'###|        | |    \ \            |  |                 | |______               |###'
'###|        | |    / /            |  |                 |  ______|              |###'
'###|        | |___/ /     __      |  |______    __     | |______               |###'
'###|        |______/     (__)     |_________|  (__)    |________|              |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|    Global Enum      |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> This Enum are used by BladeEncoder Library
        Enum BE_MP3_MODE
            MP3_MODE_STEREO = 0                     'BE_MP3_MODE_STEREO constatnt
            MP3_MODE_DUALCHANNEL = 2                'BE_MP3_MODE_DUALCHANNEL constatnt
            MP3_MODE_MONO = 3                       'BE_MP3_MODE_MONO constatnt
        End Enum
'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                   ________               |###'
'###|        | ____ \              |  |                 |  ______)              |###'
'###|        | |   \ \             |  |                 | |                     |###'
'###|        | |    \ \            |  |                 | |                     |###'
'###|        | |    / /            |  |                 | |                     |###'
'###|        | |___/ /     __      |  |______    __     | |______               |###'
'###|        |______/     (__)     |_________|  (__)    |________)              |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|    Global Const     |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Constants used by memory allocation API
        Global Const GMEM_DDESHARE          As String = &H2000
        Global Const GMEM_DISCARDABLE       As String = &H100
        Global Const GMEM_DISCARDED         As String = &H4000
        Global Const GMEM_FIXED             As String = &H0
        Global Const GMEM_INVALID_HANDLE    As String = &H8000
        Global Const GMEM_LOCKCOUNT         As String = &HFF
        Global Const GMEM_MODIFY            As String = &H80
        Global Const GMEM_MOVEABLE          As String = &H2
        Global Const GMEM_NOCOMPACT         As String = &H10
        Global Const GMEM_NODISCARD         As String = &H20
        Global Const GMEM_NOT_BANKED        As String = &H1000
        Global Const GMEM_NOTIFY            As String = &H4000
        Global Const GMEM_SHARE             As String = &H2000
        Global Const GMEM_VALID_FLAGS       As String = &H7F72
        Global Const GMEM_ZEROINIT          As String = &H40
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> This constant used to draw wav data
        Global Const PI                     As Double = 3.14159265358979
'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                  ______________          |###'
'###|        | ____ \              |  |                |_____    _____|         |###'
'###|        | |   \ \             |  |                      |  |               |###'
'###|        | |    \ \            |  |                      |  |               |###'
'###|        | |    / /            |  |                      |  |               |###'
'###|        | |___/ /     __      |  |______    __          |  |               |###'
'###|        |______/     (__)     |_________|  (__)         |__|               |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|     Global Type     |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Type used by CommonControls Library
        Type TagInitCommonControlsEx
            LngSize                         As Long
            LngICC                          As Long
        End Type
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> These 2 Structure used to draw wav data
        Type Bytes2
            Byte1                           As Byte
            Byte2                           As Byte
        End Type
        Type IntType
            IntVal                          As Integer
        End Type
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> These Structure are used by BladeEncoder Library
        Type BE_Config_MP3
            DwConfig                        As Long             ' Must be = BE_CONFIG_MP3 ( 0 )
            DwSampleRate                    As Long             ' 48000, 44100 and 32000 allowed
            ByMode                          As Byte             ' BE_MP3_MODE_STEREO, BE_MP3_MODE_DUALCHANNEL, BE_MP3_MODE_MONO
            WBitrate                        As Integer          ' 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256 and 320 allowed
            BPrivate                        As Long
            BCRC                            As Long
            BCopyright                      As Long
            BOriginal                       As Long
        End Type
        Type BE_VERSION
            ByDLLMajorVersion               As Byte             ' BladeEnc DLL Version number
            ByDLLMinorVersion               As Byte             ' BladeEnc DLL Version number
            ByMajorVersion                  As Byte             ' BladeEnc Engine Version Number
            ByMinorVersion                  As Byte             ' BladeEnc Engine Version Number
            ByDay                           As Byte             ' DLL Release date
            ByMonth                         As Byte             ' DLL Release date
            WYear                           As Integer          ' DLL Release date
            ZHomepage(0 To 256)             As Byte             ' BladeEnc Homepage URL
        End Type
'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                      ______              |###'
'###|        | ____ \              |  |                    /  __  \             |###'
'###|        | |   \ \             |  |                   /  /  \  \            |###'
'###|        | |    \ \            |  |                  /  /____\  \           |###'
'###|        | |    / /            |  |                 /  /______\  \          |###'
'###|        | |___/ /     __      |  |______    __    /  /        \  \         |###'
'###|        |______/     (__)     |_________|  (__)  /__/          \__\        |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|    Global API       |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Used for initializing CommonControls which will provide to us XP style
        Declare Function InitCommonControlsEx Lib _
                                "comctl32.dll" ( _
                                iccex As TagInitCommonControlsEx) As _
                                Boolean
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Declaration of the global WBladeEncoder API, All functions are 16 bit
        Declare Function InitializeDLL Lib _
                                "WBladeEncoder.Dll" _
                                () As Long
        Declare Function blInitStream Lib _
                                "WBladeEncoder.Dll" _
                                (ByVal SampleRate As Long, _
                                ByVal Mode As Byte, _
                                ByVal Bitrate As Integer, _
                                dwSamples As Long, _
                                dwMP3BufferSize As Long, _
                                HBEStream As Long) _
                                As Long
        Declare Function blEncodeChunk Lib _
                                "WBladeEncoder.Dll" _
                                (ByVal HBEStream As Long, _
                                ByVal nSamples As Long, _
                                pSamples As Integer, _
                                pOutput As Byte, _
                                pdwOutput As Long) _
                                As Long
        Declare Function blDeinitStream Lib _
                                "WBladeEncoder.Dll" _
                                (ByVal HBEStream As Long, _
                                pOutput As Byte, _
                                pdwOutput As Long) _
                                As Long
        Declare Function blCloseStream Lib _
                                "WBladeEncoder.Dll" _
                                (ByVal HBEStream As Long) _
                                As Long
        Declare Function blGetVersion Lib _
                                "WBladeEncoder.Dll" _
                                (Version As BE_VERSION) _
                                As Long
        Declare Function ChangeSoundVolume Lib _
                                "WBladeEncoder.Dll" _
                                (pInOut As Integer, _
                                ByVal nSamples As Long, _
                                ByVal nChannels As Integer, _
                                ByVal LVolume As Double, _
                                ByVal RVolume As Double) _
                                As Long
        Declare Function ChangeSoundVolumeB Lib _
                                "WBladeEncoder.Dll" _
                                (pInOut As Byte, _
                                ByVal nSamples As Long, _
                                ByVal nChannels As Integer, _
                                ByVal LVolume As Double, _
                                ByVal RVolume As Double) _
                                As Long
        Declare Function ChangeSoundVolumeC Lib _
                                "WBladeEncoder.Dll" _
                                Alias "ChangeSoundVolume" _
                                (ByVal pInOut As Long, _
                                ByVal nSamples As Long, _
                                ByVal nChannels As Integer, _
                                ByVal LVolume As Double, _
                                ByVal RVolume As Double) _
                                As Long
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Declaration of memory function API
        Declare Function GlobalAlloc Lib _
                                "kernel32" _
                                (ByVal wFlags As Long, _
                                ByVal dwBytes As Long) _
                                As Long
        Declare Function GlobalFree Lib _
                                "kernel32" _
                                (ByVal hMem As Long) _
                                As Long
        Declare Function GlobalLock Lib _
                                "kernel32" _
                                (ByVal hMem As Long) _
                                As Long
        Declare Function GlobalSize Lib _
                                "kernel32" _
                                (ByVal hMem As Long) _
                                As Long
        Declare Function GlobalUnlock Lib _
                                "kernel32" _
                                (ByVal hMem As Long) _
                                As Long
        Declare Function GlobalReAlloc Lib _
                                "kernel32" _
                                (ByVal hMem As Long, _
                                ByVal dwBytes As Long, _
                                ByVal wFlags As Long) _
                                As Long
        Declare Sub CopyMemory Lib _
                                "kernel32" _
                                Alias "RtlMoveMemory" _
                                (pDest As Any, _
                                pSrc As Any, _
                                ByVal ByteLen As Long)
        Declare Sub CopyMemoryB Lib _
                                "kernel32" _
                                Alias "RtlMoveMemory" _
                                (ByVal pDest As Long, _
                                ByVal pSrc As Long, _
                                ByVal ByteLen As Long)
        Declare Sub ExitProcess Lib "kernel32" _
                                (ByVal uExitCode As Long)
