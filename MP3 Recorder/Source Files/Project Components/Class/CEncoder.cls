VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem -> ###|##################|#|########################|#|######################|###
Rem -> ###|#|                                                                    |###
Rem -> ###|#|                     Class Identification                           |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Rem -> This class module involve to convert a stream of WAV into MP3 stream.
Rem -> And delever that data in event called GotMP3Data
Rem -> ******************************************************************************
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|##################|#|########################|#|######################|###
Rem -> ###|#|                                                                    |###
Rem -> ###|#|                         Developer List                             |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Rem -> Created by                   : Michael Ciurescu
Rem -> Redeveloped & Commented by   : John Fawzy (Dominator Legand)
Rem -> Email                        : Dominator_Legand@Yahoo.com
Rem -> Date                         : 6/20/2005
Rem -> Version                      : 2.0.1.*
Rem -> ******************************************************************************
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|##################|#|########################|#|######################|###
Rem -> ###|#|                                                                    |###
Rem -> ###|#|                     Class Main Varaibles                           |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Rem -> Declaration of the class variables
Private VSampleRate                     As Long             'Default value = 44100
Private VMode                           As BE_MP3_MODE      'Default value = MP3_MODE_STEREO
Private VBitrate                        As Integer          'Default value = 128
Private BufferSamples                   As Long             'Number of sampel in buffer, determined by the Encoder
Private MP3Buffer                       As Long             'The MP3Buffer length, Endoer determine that length
Private HBEStream                       As Long             'Indicate if Encoder Initialized or not (0|1)
Private MP3_Buffer()                    As Byte             'Buffer in which we will pass to user
Rem -> ******************************************************************************
Rem -> Declaration of the class event
Public Event GotMP3Data(Buffer() As Byte)
Rem -> ******************************************************************************
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|##################|#|########################|#|######################|###
Rem -> ###|#|                                                                    |###
Rem -> ###|#|                         Class Events                               |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Rem -> Event occuer when the class Initialized
Private Sub Class_Initialize()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Initializing encoder library
    Call InitializeDLL
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Reset quality parameters
    VSampleRate = -1        'Default value = 44100
    VMode = -1              'Default value = MP3_MODE_STEREO
    VBitrate = -1           'Default value = 128
End Sub
Rem -> ******************************************************************************
Rem -> Event occure when the object of the class destroyed
Private Sub Class_Terminate()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Close the stream opened by the encoder
    If HBEStream <> 0 Then blCloseStream HBEStream
End Sub
Rem -> ******************************************************************************
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|##################|#|########################|#|######################|###
Rem -> ###|#|                                                                    |###
Rem -> ###|#|                         Class Property                             |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Rem -> Property Samples (Get)
Public Property Get SamplesPerBuffer() As Long
    SamplesPerBuffer = BufferSamples
End Property
Rem -> ********************************************************************************
Rem -> Property SampleRate (Get)
Public Property Get SampleRate() As Long
    SampleRate = VSampleRate
End Property
Rem -> ********************************************************************************
Rem -> Property Channels (Get)
Public Property Get Channels() As BE_MP3_MODE
    Channels = VMode
End Property
Rem -> ********************************************************************************
Rem -> Property BitRate (Get)
Public Property Get Bitrate() As Integer
    Bitrate = VBitrate
End Property
Rem -> ******************************************************************************
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|##################|#|########################|#|######################|###
Rem -> ###|#|                                                                    |###
Rem -> ###|#|                         Class Method                               |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Rem -> Public function which initialize the stream
Public Function InitStream(Optional SampleRate As Long = 44100, _
                           Optional Mode As BE_MP3_MODE = MP3_MODE_STEREO, _
                           Optional Bitrate As Integer = 128) As Boolean
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Call blInitStream which is an API function, In WBladeEncoder library
    Rem -> Which initialize the mp3 stream, by giving to it the quality parameters
    Rem -> This function return zero when success.
    If blInitStream(SampleRate, Mode, Bitrate, BufferSamples, MP3Buffer, HBEStream) <> 0 Then Exit Function
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Update the property value with the new value
    VSampleRate = SampleRate
    VMode = Mode
    VBitrate = Bitrate
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Inform the caller that every thing is fine
    InitStream = True
End Function
Rem -> ******************************************************************************
Rem -> Public function used to encdoe WAV to MP3
Public Function EncodeChunk(WAV_Buffer() As Integer) As Boolean
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Declaring the requried varaibles
    Dim ToWrite             As Long     'What is this ???
    Dim DSamples            As Long     'What is this ???
    Dim TSamples            As Long     'What is this ???
    Dim WSamples            As Long     'What is this ???
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Get the number of the upper index of WAV_Buffer array and add to it 1
    DSamples = UBound(WAV_Buffer) + 1
    If DSamples < BufferSamples Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Resize the mp3 array with the MP3Buffer value, and erase that array
        Rem -> MP3Buffer is determined by the Encoder
        ReDim MP3_Buffer(0 To MP3Buffer - 1)
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Request from the encoder to encode our Chunk, HBEStream is handel
        Rem -> Generated by the encoder (0|1), DSamples is the length of wav buffer
        Rem -> WAV_Buffer -> is the source, MP3_Buffer -> is the destination
        If blEncodeChunk(HBEStream, DSamples, WAV_Buffer(0), MP3_Buffer(0), ToWrite) <> 0 Then Exit Function
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Resize the mp3 array with the ToWrite value, but keep old data intact
        ReDim Preserve MP3_Buffer(0 To ToWrite - 1)
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> deliver the chunk to Event and it delever to user.
        RaiseEvent GotMP3Data(MP3_Buffer)
        WSamples = BufferSamples
    Else
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Below i don't got it, I don't understand it, if any body do, let me know
        Rem -> ?????????????????
        Do
            TSamples = IIf(DSamples > BufferSamples, BufferSamples, DSamples)
            ReDim MP3_Buffer(0 To MP3Buffer - 1)
            If blEncodeChunk(HBEStream, TSamples, WAV_Buffer(WSamples), MP3_Buffer(0), ToWrite) <> 0 Then Exit Function
            If ToWrite > 0 Then
                ReDim Preserve MP3_Buffer(0 To ToWrite - 1)
                RaiseEvent GotMP3Data(MP3_Buffer)
                WSamples = WSamples + TSamples
                DSamples = DSamples - TSamples
            Else
                'Debug.Print "ToWrite = 0"
            End If
        Loop Until DSamples <= 0
    End If
    EncodeChunk = True
End Function
Rem -> ******************************************************************************
Rem -> Public function used to close the stream of the encoder
Public Function CloseStream() As Boolean
    Dim ToWrite As Long
    If MP3Buffer > 0 Then
        ReDim MP3_Buffer(0 To MP3Buffer - 1)
        If blDeinitStream(HBEStream, MP3_Buffer(0), ToWrite) <> 0 Then Exit Function
        If ToWrite > 0 Then
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> I think that this is the end of file ?!
            ReDim Preserve MP3_Buffer(0 To ToWrite - 1)
            RaiseEvent GotMP3Data(MP3_Buffer)
        End If
    End If
    Rem -> ~~~~~~~~~~~~~~~~
    Rem -> Reset varaiables
    VSampleRate = -1
    VMode = -1
    VBitrate = -1
    If blCloseStream(HBEStream) <> 0 Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Error occure while closing the stream
        Exit Function
    Else
        HBEStream = 0
    End If
    CloseStream = True
End Function
