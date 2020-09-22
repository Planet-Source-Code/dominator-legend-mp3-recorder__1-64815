VERSION 5.00
Begin VB.Form CDirectX 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   Icon            =   "CDirectX.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   345
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Label Lbl_Interface 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "DirectX Audio Capture Module"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2295
   End
End
Attribute VB_Name = "CDirectX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Rem -> ###|##################|#|########################|#|######################|###
Rem -> ###|#|                                                                    |###
Rem -> ###|#|                     Module Identification                          |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Rem -> This form module involve to capture a stream of WAV data from Soundcard.
Rem -> And delever that data in event called GotWavData, the resone on which this
Rem -> Module is form module instead of class module, that we need a Form object,
Rem -> To pass it to CreateEvent function which create to use 3 Events, so we can
Rem -> Use the buffer, if you try to give to this function (CreateEvent), the Frm_Main
Rem -> It will cause Atutomation Error exception, cause the CallBack event is
Rem -> Implemented in CDirectx form not in the Main form, you may cut all code in this
Rem -> Form and past it in Frm_Main, but its will result poor code, *my opinion*.
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
Rem -> Redeveloped & Commented by   : John Fawzy (Dominator Legend)
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
Rem -> ###|#|                    Module Main Varaibles                           |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Rem -> Declaration of Primary DirectX Objects
Private DX                      As New DirectX8
Private SEnum                   As DirectSoundEnum8
Private DIC                     As DirectSoundCapture8
Rem ->
Rem ->
Rem ->
Rem -> ******************************************************************************
Rem -> Declaration of Buffer stuff...
Private Buff                    As DirectSoundCaptureBuffer8
Private BuffWave                As WAVEFORMATEX
Private BuffDesc                As DSCBUFFERDESC
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem -> Declaration of The buffer sizes for the DirectSound
Private BuffLen                 As Long
Private HalfBuffLen             As Long
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem -> Lets say the buffer is like this, Below is the buffer divided into two parts
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem ->
Rem ->      |---------------------------|---------------------------|
Rem ->      |    Left Side Of Buffer    |    Right Side Of Buffer   |
Rem ->      |---------------------------|---------------------------|
Rem ->
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem -> Where you see a "|" character, an event will occur
Rem -> Whene you read the left side of the buffer, the DirectX writes on the right
Rem -> Side of the buffer, And vice versa, if you read the data on the same side
Rem -> As DirectX writes, the sound will be messed up
Rem ->
Rem ->
Rem ->
Rem -> ******************************************************************************
Rem -> Events stuff, This will allow us to use callback method, i.e.
Rem -> An events DirectXEvent8_DXCallback will Implemented by DirectX.
Implements DirectXEvent8
Rem ->
Rem ->
Rem ->
Rem -> ******************************************************************************
Rem -> The declaration of EventsNotify which is array of structure from type
Rem -> DSBPOSITIONNOTIFY which work as EventsID container,
Rem -> As descriped above, when we operate with the buffer we have 3 events, I.e.
Rem -> 3 '|' characters, The first "|" mean StartEvent, The second "|" mean MidEvent
Rem -> The Third "|" mean EndEvent, By calling DX.CreateEvent method, directx will
Rem -> Create for us 3 ID for the 3 events, so we declare the following varaibles
Private EventsNotify()              As DSBPOSITIONNOTIFY
Private StartEvent                  As Long
Private MidEvent                    As Long
Private EndEvent                    As Long
Rem ->
Rem ->
Rem ->
Rem -> ******************************************************************************
Rem -> If SoundVolume = 1, ie 100%, DigitalVolume variable declaration.
Public SoundVolume                  As Double
Rem ->
Rem ->
Rem ->
Rem -> ******************************************************************************
Rem -> Declaration of Memory Buffer Allocation
Public MemPtr                       As Long            ' Allocated memory handle
Public Ptr_WavBuffer                As Long            ' Pointer for allocated buffer
Const M_BufferLength                As Long = 529200   ' Total length of the buffer
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem -> M_BufferPos is the position where to write next in the buffer If this
Rem -> Variable Is not decremented regularly, the buffer will get VERY big!
Public M_BufferPos As Long
Rem ->
Rem ->
Rem ->
Rem -> ******************************************************************************
Rem -> Declaration Of Event GotWavData which will occur when we got data
Public Event GotWavData(Buffer() As Byte)
Rem -> ******************************************************************************
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|##################|#|########################|#|######################|###
Rem -> ###|#|                                                                    |###
Rem -> ###|#|                        Module Events                               |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Private Sub Class_Initialize()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Assign the volum with 100%
    SoundVolume = 1
End Sub
Private Sub Class_Terminate()
    Rem -> ~~~~~~~~~~~~~~~~~~~
    Rem -> Destroy Every thing
    UninitializeSound
End Sub
Rem -> ******************************************************************************
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|#|
Rem -> ###|##################|#|########################|#|######################|###
Rem -> ###|#|                                                                    |###
Rem -> ###|#|                        Module Method                               |###
Rem -> ###|#|                                                                    |###
Rem -> ******************************************************************************
Rem -> Declaration Of CallBack Event, This events create a buffer on memory.
Rem -> In this event, we check if we allocate a memory size to hold our captured data
Rem -> or not, if we don't allocate any location on memory and we don't got its address
Rem -> then we allocate it and got a pointer to its address.
Private Sub DirectXEvent8_DXCallback(ByVal EventID As Long)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> If its the first time, then create a buffer to hold the wav data
    If ((Ptr_WavBuffer = 0) Or (MemPtr = 0)) Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> The GlobalAlloc function allocates the specified number of
        Rem -> bytes from the heap, which take as a parameters
        Rem ->
        Rem -> Parameter #1
        Rem -> The type of allocated space, which will be GMEM_SHARE which mean
        Rem -> Allocates memory to be used by the dynamic data exchange (DDE)
        Rem -> functions for a DDE conversation, this memory is not shared globally.
        Rem -> However, this flag is available for compatibility purposes.
        Rem -> It may be used by some applications to enhance the performance of DDE
        Rem -> operations and should, therefore, be specified if the memory is to be
        Rem -> used for DDE. Only processes that use DDE or the clipboard for
        Rem -> interprocess communications should specify this flag.
        Rem -> In addition to this flag we use GMEM_ZEROINIT
        Rem -> To Initializes memory contents to zero
        Rem ->
        Rem -> Parameter #2
        Rem -> 2- Size of allocated space on the heap.
        Rem ->
        Rem -> This function is return to us a pointer to the allocated space which
        Rem -> Assigned to MemPtr varaiable.
        MemPtr = GlobalAlloc(GMEM_SHARE Or GMEM_ZEROINIT, M_BufferLength)
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> The GlobalLock function locks a global memory object and returns a
        Rem -> Pointer to the first byte of the object’s memory block.
        Rem -> The memory block associated with a locked memory object cannot be
        Rem -> moved or discarded.
        Rem ->
        Rem -> Parameter #1
        Rem -> hMem which Identifies the global memory object.
        Rem -> This handle is returned by the GlobalAlloc function.
        Rem -> If the function succeeds, the return value is a pointer to the
        Rem -> first byte of the memory block.
        Rem -> If the function fails, the return value is NULL.
        Ptr_WavBuffer = GlobalLock(MemPtr)
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> We check if block is allocated or not and if its then we
        If Ptr_WavBuffer = 0 Then _
            MsgBox "Can't allocate the required space on heap.", vbCritical, "Allocation Error"
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Finally we reset the position of buffer to zero, so DX start writing
        Rem -> The captured data from the begining of buffer.
        M_BufferPos = 0
    End If
    
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Its supposed that compiler don't enter this area cause if its happen
    Rem -> It mean that allocated space is not enough to hold the returned data
    Rem -> But if compile enter this area, we will make the buffer size bigger.
    Rem -> M_BufferLength which is declared above as const is = 500K, which
    Rem -> Is enough to hold =~ 3 seconds of captured sound, so if compiler is
    Rem -> gets here, this mean that that something wrong occure while processing
    Rem -> WAVE data
    If M_BufferPos + HalfBuffLen >= M_BufferLength Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> The GlobalUnlock function decrements the lock count associated
        Rem -> with a memory object that was allocated with the GMEM_MOVEABLE flag,
        Rem -> Which mean its Unblock the blocked area of heap to reallocate it again
        Rem -> Or to to allow us to delete this area, its parameter is handel of
        Rem -> Allocated block.
        GlobalUnlock MemPtr
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Make the allocated buffer bigger, by using GlobalReAlloc function which
        Rem -> Will ReAllocate our buffer with bigger size, and use GMEM_ZEROINIT to
        Rem -> Reset the contents of allocated block to zero.
        MemPtr = GlobalReAlloc(MemPtr, (M_BufferPos + (HalfBuffLen * 2)) * 1.2, GMEM_ZEROINIT)
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Lock the allocated area as we done above
        Rem -> Please refare to first IF statment for detailes.
        Ptr_WavBuffer = GlobalLock(MemPtr)
    End If
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> We check if required space is allocated or not, and if its we goining on
    If Ptr_WavBuffer <> 0 Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> When this events is fired we got the eventid, then we select the eventid
        'Debug.Print EventID
        Select Case EventID
            Case StartEvent
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem -> Reading the right side of buffer [I Think so]
                Buff.ReadBuffer HalfBuffLen, HalfBuffLen, ByVal (Ptr_WavBuffer + M_BufferPos), DSCBLOCK_DEFAULT
                M_BufferPos = M_BufferPos + HalfBuffLen
            Case MidEvent
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem -> Reading the left side of buffer [I Think so]
                Buff.ReadBuffer 0, HalfBuffLen, ByVal (Ptr_WavBuffer + M_BufferPos), DSCBLOCK_DEFAULT
                M_BufferPos = M_BufferPos + HalfBuffLen
            Case EndEvent
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem -> Finishing From recording [I Think so]
                'Debug.Print "EndEvent Occure"
            Case Else
                Debug.Print "EventID: " & EventID
        End Select

        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Let the parent control know that we have more WAVE data
        Rem -> By firing the GotWavData event, and aslo we check if the evnt id is
        Rem -> Either from start or from midlle, cause if its the Endevent, we will
        Rem -> Not deliver any data
        If EventID = StartEvent Or EventID = MidEvent Then
            If BuffWave.nBitsPerSample = 16 And SoundVolume <> 1 Then
                Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Rem -> Change the sound volume... (Digital volume), Of the encoder
                ChangeSoundVolumeC (Ptr_WavBuffer + M_BufferPos) - HalfBuffLen, _
                                     HalfBuffLen \ (BuffWave.nBitsPerSample \ 8), _
                                    BuffWave.nChannels, SoundVolume, SoundVolume
            End If
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Here we Request from CopyMemory API function to copy
            Rem -> That buffer from a location In memory, that location is allocated,
            Rem -> Above, Copy the memory to a byte array, so we can use it, this
            Rem -> Function take as parameters the array we want to fill, a pointer
            Rem -> To the source, the position in which we read from it.
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Resize the buffer array with the length of M_BufferPos
            Dim WAVBuffer() As Byte
            ReDim WAVBuffer(M_BufferPos - 1)
            CopyMemory WAVBuffer(0), ByVal Ptr_WavBuffer, M_BufferPos
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Raising the GotWavData event and give it as parameters the buffer
            Rem -> We fill it above, and that to inform the user that we got the chunk
            RaiseEvent GotWavData(WAVBuffer)
        End If
    End If
End Sub
Rem -> ******************************************************************************
Rem -> This is the function which Initialize the directx with function
Rem -> Parameters and force it to capture sound
Public Function Initialize(ByRef Owner As Form, _
                           Optional ByVal SamplesPerSec As Long = 44100, _
                           Optional ByVal BitsPerSample As Integer = 16, _
                           Optional ByVal Channels As Integer = 2, _
                           Optional ByVal HalfBufferLen As Long = 0) As String
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Handel Any Error may occur
    On Error GoTo ReturnError
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Initialize the Directsound Enumuration object -> ???
    Set SEnum = DX.GetDSEnum
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Initialize the Directsound Capture object -> ???
    Set DIC = DX.DirectSoundCaptureCreate(SEnum.GetGuid(1))
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Fill WAVEFORMATEX structure with selected format and paremeters
    Rem -> This structure will be passed to fxFormat property which is the format
    Rem -> Of captured sound of BuffDesc object which is Buffer Descryption
    Rem -> I.e. this structure contain the parameters which decide the sound quality
    BuffWave.nFormatTag = WAVE_FORMAT_PCM
    BuffWave.nChannels = Channels
    BuffWave.nBitsPerSample = BitsPerSample
    BuffWave.lSamplesPerSec = SamplesPerSec
    BuffWave.nBlockAlign = (BuffWave.nBitsPerSample * BuffWave.nChannels) \ 8
    BuffWave.lAvgBytesPerSec = BuffWave.lSamplesPerSec * BuffWave.nBlockAlign
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Determine the half length of the buffer
    If HalfBufferLen <= 0 Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> It's the first time, we determine the half of the buffer
        Rem -> by the fallwing equation.
        HalfBuffLen = BuffWave.lAvgBytesPerSec / 10
        HalfBuffLen = HalfBuffLen - (HalfBuffLen Mod BuffWave.nBlockAlign)
    Else
        HalfBuffLen = HalfBufferLen
    End If
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Determine the length of the buffer, which is the double of half
    Rem -> Of the buffer determined above
    BuffLen = HalfBuffLen * 2
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> After we fill the BuffWave object, now we pass it to
    Rem -> Buffer descryption object, and also pass the buffer length
    BuffDesc.fxFormat = BuffWave        'Rem -> The structure we fill above
    BuffDesc.lBufferBytes = BuffLen     'Rem -> The Buffer length we got above
    BuffDesc.lFlags = DSCBCAPS_DEFAULT
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Finally we pass the buffer descryption to CreateCaptureBuffer
    Rem -> Function which create the buffer and return to us an handle for
    Rem -> DirectSoundCaptureBuffer8 object.
    Set Buff = DIC.CreateCaptureBuffer(BuffDesc)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Reinitialize EventsNotify which is array of structure from type
    Rem -> DSBPOSITIONNOTIFY which work as Events container, with 3 indexes
    Rem -> EventsNotify(0) will be contain the EventID of StartEvent
    Rem -> EventsNotify(1) will be contain the EventID of MidlleEvent
    Rem -> EventsNotify(2) will be contain the EventID of EndEvent
    Rem -> I.e. we Reinitialize an array which will contain 3 indexes
    Rem -> for the 3 events that we have, in which StartEvent Event will occure
    Rem -> when the DirectX write on the left part of buffer, the MidEvent Event
    Rem -> will occure when the DirectX write on the right side, the EndEvent Event
    Rem -> will occure when the DirectX stop capturing process.
    Rem ->
    Rem -> Please Refare to the above figure which explain the shape of buffer
    Rem -> for mor detailes.
    ReDim EventsNotify(0 To 2) As DSBPOSITIONNOTIFY
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Here we call the CreateEvent function and give to it the Hwnd of this form
    Rem -> Which its return to us the EventID of StartEvent, which we used in the
    Rem -> Prev Select Case statment, Refare to above Event "DirectXEvent8_DXCallback"
    StartEvent = DX.CreateEvent(Me)
    EventsNotify(0).hEventNotify = StartEvent
    EventsNotify(0).lOffset = 1
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Here we call the CreateEvent function and give to it the Hwnd of this form
    Rem -> Which its return to us the EventID of MidEvent.
    MidEvent = DX.CreateEvent(Me)
    EventsNotify(1).hEventNotify = MidEvent
    EventsNotify(1).lOffset = HalfBuffLen
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Here we call the CreateEvent function and give to it the Hwnd of this form
    Rem -> Which its return to us the EventID of EndEvent.
    EndEvent = DX.CreateEvent(Me)
    EventsNotify(2).hEventNotify = EndEvent
    EventsNotify(2).lOffset = DSBPN_OFFSETSTOP
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Finally we inform the compiler that we finished from assigning the events
    Rem -> By giving to "SetNotificationPositions" function the number of events and
    Rem -> an array of structure which contain the eventsID descryption.
    Buff.SetNotificationPositions 3, EventsNotify()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Tell the caller which call this function that no error occure
    Initialize = ""
Rem -> ~~~~~~~~~~~~~~~~~~~~~~
Rem -> Error Handelar Section
Exit Function
ReturnError:
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem -> An error has been occured, inform the caller and UninitializeSound
Initialize = "Error: " & Err.Number & vbNewLine & _
             "Desription: " & Err.Description & vbNewLine & _
             "Source: " & Err.Source
Err.Clear
UninitializeSound
Exit Function
End Function
Rem -> ******************************************************************************
Rem -> Function used to start sound
Public Function SoundPlay() As Boolean
Rem -> ~~~~~~~~~~~~~~~~
Rem -> Handel Any Error
On Error GoTo ReturnError
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Check if there is a buffer, so we can start it.
    If Not Buff Is Nothing Then Buff.Start DSCBSTART_LOOPING
    SoundPlay = True
Rem -> ~~~~~~~~~~~~~~~~~~~~~~
Rem -> Error Handelar Section
Exit Function
ReturnError:
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem -> There is no buffer, inform the caller
SoundPlay = False
Err.Clear
End Function
Rem -> ******************************************************************************
Rem -> Function used to stop sound
Public Function SoundStop() As Boolean
Rem -> ~~~~~~~~~~~~~~~~
Rem -> Handel Any Error
On Error GoTo ReturnError
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Check if there is a buffer, so we can stop it.
    If Not Buff Is Nothing Then
        Buff.Stop
    End If
    SoundStop = True
Rem -> ~~~~~~~~~~~~~~~~~~~~~~
Rem -> Error Handelar Section
Exit Function
ReturnError:
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem -> There is no buffer, inform the caller
SoundStop = False
Err.Clear
End Function
Rem -> ******************************************************************************
Rem -> Function used to destroy all object in this class
Public Sub UninitializeSound()
Rem -> ~~~~~~~~~~~~~~~
Rem -> Ignor any error
On Error Resume Next
    Rem -> ~~~~~~~~~~~~~~
    Rem -> Stop the sound
    If Not SoundStop Then MsgBox "Unable to stop the sound, maybe there is no buffer created.", vbCritical, "Error stoping sound"
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Execute any process in process Queu.
    DoEvents
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Destroy allocated buffer created above.
    GlobalUnlock MemPtr
    GlobalFree MemPtr
    MemPtr = 0
    HalfBuffLen = 0
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Destroy the events we created above.
    DX.DestroyEvent EventsNotify(0).hEventNotify
    DX.DestroyEvent EventsNotify(1).hEventNotify
    DX.DestroyEvent EventsNotify(2).hEventNotify
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Destroy array, which was hold the events handel
    Erase EventsNotify
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Destroy DirectX Objects
    Set Buff = Nothing
    Set DIC = Nothing
    Set SEnum = Nothing
End Sub
