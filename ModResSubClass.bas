Attribute VB_Name = "ModResSubClass"
Option Explicit

'ModResSubClass
'For use with BASeROOS- can be used early-bound or late-bound.

'requirements: include a reference to the ROOSIFACE typelib.

'purpose:
'subclasses the Visual Basic Run-time LoadRes... constants,
'replacing them with equivalent calls to the ROOS.
'in order to access the actual Resource file of the application (assuming you have a resource file)
'the client will now have to qualify with the VB library, as is done here.

'the reason this is a module that must be included is because
'it must be in the project namespace to stand above the
'Visual Basic Run-time procedures.
Private mGrabber As IResGrabber



Public Function InitResGrabber(ByVal ProgID As String) As Boolean
    'attempts to initialize the resource grabber object to the given ProgID.
    On Error Resume Next
    Set mGrabber = CreateObject(ProgID)
    If Err <> 0 Then
        InitResGrabber = False
        MsgBox "Failed to load Resource Grabber Module:" & ProgID
    Else
        InitResGrabber = True
        CDebug.Post "Success loading Resource grabber:" & ProgID
    End If
End Function


Public Function LoadResData(pId As Variant, pLoadType As Variant) As Byte()
    LoadResData = mGrabber.Advanced.LoadResData(pId, pLoadType)
End Function
Public Function LoadResString(ByVal pId As Long) As String
    LoadResString = mGrabber.Advanced.LoadResString(pId)
    ' LoadresPicture(id,restype as Integer) as IPictureDisp
End Function
Public Function LoadResPicture(ByVal Id, ByVal ResType As Integer) As IPictureDisp
    Set LoadResPicture = mGrabber.Advanced.LoadResPicture(Id, ResType)
End Function

