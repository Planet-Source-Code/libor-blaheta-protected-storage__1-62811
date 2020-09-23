VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00AFAF00&
   Caption         =   "VB6 PStore enumeration"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000A0&
   Icon            =   "vbPStoreEnum.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H00F0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      IntegralHeight  =   0   'False
      Left            =   660
      TabIndex        =   0
      Top             =   660
      Width           =   5955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' --- Protected Storage ---
'
' Well, one question comes immediately - what is the Protected Storage (PS)?
'
' MS description of it, is
'
' --------------------------------------------------------------------------------------
' Microsoft Service Description - Provides protected storage for sensitive data,
' such as private keys, to prevent access by unauthorized services, processes, or users.
' Only on Win2000 and WinXP.
' --------------------------------------------------------------------------------------
'
' Interesting, isn't it?
'
' Eg. have you known that Outlook Express or Internet Explorer heavily use PS?
' Yes, your OE and IE passwords and are stored right in the PS.
'
' I'm not gonna comment the code very much, I think advanced VB programmer could understand it quite well
' and I do not want any VB lamers abuse this (and to tell a true I'm lazy to write any commnets :-) ).
' I have not found any VB source code that can work with PS so I think this can be useful for curious programmers who want to know what is behind the scene.
'
' If you want to read something about PS interface, go here
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/devnotes/winprog/ipstore.asp
'
' Again, sorry for the poor code commenting ;-)
'
' Libor Blaheta
'

Private Type PST_TYPEINFO
    cbSize As Long
    pDisplayName As Long
End Type

Private Type GUID_
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4 As Integer
    Data5 As Integer
    Data6 As Long
End Type

Private Type PST_PROVIDERINFO
    cbSize As Long
    GUID As GUID_
    Capabilities As Long
    pProviderName As Long
End Type

Private Type PST_PROMPTINFO
    cbSize As Long
    dwPromptFlags As Long
    hwndApp As Long
    szPrompt As String
End Type

Dim Pstore           As CPStore      'PS object
Dim EnumPStoreTypes  As CEnumTypes
Dim EnumPSubTypes    As CEnumTypes
Dim EnumPSubItems    As CEnumItems

Dim hPstore          As Long
Dim hResult          As Long

Private Declare Sub MyCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function PStoreCreateInstance Lib "pstorec.dll" (pThis As Any, ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Sub Form_Load()

Dim pstGUID As GUID_     ' PStoreType, current enumeration
Dim subGUID As GUID_     ' PStoreSubType
Dim pItem   As Long      ' PStoreItem (pointer to a null-terminated Wide string)
   
Dim abData() As Byte, sData As String, i As Long
Dim tPrompt As PST_PROMPTINFO
Dim pData As Long         'pointer to byte array
Dim lDataLen As Long      'size of byte array
    
    
    Show
    
    'get an interface pointer to a storage provider
    hResult = PStoreCreateInstance(Pstore, 0, 0, 0)
    List1.AddItem "  &PStore = " & Hex$(ObjPtr(Pstore))
    
    'get info about provider
    Pstore.GetInfo VarPtr(pData)
    List1.AddItem "   Provider Info = """ & GetStr_Provider(pData) & """"
   
    'let's enum all item in PS
    Pstore.EnumTypes 0, 0, EnumPStoreTypes
    If EnumPStoreTypes Is Nothing Then Exit Sub

   '
   ' we don't get HRESULTS returned in VB COM-mode, but VB will raise
   ' an error when it sees a non-zero result, so we just trap all
   ' errors and exit when Err.Number is set
   '
   '
   
    On Error Resume Next
   
    Do
        
        'get a next item
        EnumPStoreTypes.Next 1, VarPtr(pstGUID), 0
        If Err.Number Then Exit Do
        
        List1.AddItem ""
        List1.AddItem "   StoreType = " & Guid2String(pstGUID)
        
        'get its subitems
        Pstore.EnumSubtypes 0, VarPtr(pstGUID), 0, EnumPSubTypes
        
        'get info about item
        Pstore.GetTypeInfo 0&, VarPtr(pstGUID), VarPtr(pData), 0&
        List1.AddItem "   Item Info= """ & GetStr_Info(pData) & """"
        List1.AddItem ""
        
        Do
            
            'enum subitems
            EnumPSubTypes.Next 1, VarPtr(subGUID), 0
            If Err.Number Then Err.Clear: Exit Do
            
            List1.AddItem "     SubType = " & Guid2String(subGUID)
            
            'get info about sub-item
            Pstore.GetSubtypeInfo 0&, VarPtr(pstGUID), VarPtr(subGUID), VarPtr(pData), 0&
            List1.AddItem "     SubItem Info= """ & GetStr_Info(pData) & """"
            List1.AddItem ""
            
            'enum subitems
            Pstore.EnumItems 0, VarPtr(pstGUID), VarPtr(subGUID), 0, EnumPSubItems
                       
            Do
                
                EnumPSubItems.Next 1, pItem, 0
                If Err.Number Then
                    Err.Clear
                    Exit Do
                End If
                    
                List1.AddItem "       pItem = " & Hex$(pItem) & " => " & GetBSTR(pItem)
                
                'init
                sData = ""
                tPrompt.cbSize = Len(tPrompt)
                
                'read item
                Pstore.ReadItem 0, VarPtr(pstGUID), VarPtr(subGUID), pItem, lDataLen, pData, VarPtr(tPrompt), 0
                If lDataLen = 0 Then GoTo 10
                
                'prepare array
                ReDim abData(1 To lDataLen)
                'copy data
                CopyMemory ByVal VarPtr(abData(1)), ByVal pData, lDataLen
                'free memory
                CoTaskMemFree pData
                
                'to unicode
                For i = 1 To lDataLen
                    If abData(i) = 0 Then sData = sData & "(0)" Else sData = sData & Chr(abData(i))
                Next i

10:
                List1.AddItem "       ReadItem = " & lDataLen & " (bytes)"
                List1.AddItem "       Data = """ & Replace(sData, "(0)", "") & """"
                List1.AddItem ""
                
            Loop
        Loop
    Loop
    
End Sub

'get BSTR from pointer
Private Function GetBSTR(ByVal pData As Long) As String
Dim iChar As Integer
   
    Do
        CopyMemory iChar, ByVal pData, 2
        If iChar = 0 Then Exit Do
        GetBSTR = GetBSTR & Chr$(iChar)
        pData = pData + 2
    Loop

End Function

Private Function Guid2String(ID As GUID_) As String
    Guid2String = "{" & Right$("00000000" & Hex$(Rev4(ID.Data1)), 8) & _
             "-" & Right$("0000" & Hex$(Rev2(ID.Data2)), 4) & _
             "-" & Right$("0000" & Hex$(Rev2(ID.Data3)), 4) & _
             "-" & Right$("0000" & Hex$(Rev2(ID.Data4)), 4) & _
             "-" & Right$("0000" & Hex$(Rev2(ID.Data5)), 4) & _
                   Right$("00000000" & Hex$(Rev4(ID.Data6)), 8) & "}"
End Function

Private Function Rev4(ByVal Val As Long) As Long
Dim b(3) As Byte, v(3) As Byte
    CopyMemory v(0), Val, 4
    b(3) = v(0): b(2) = v(1): b(1) = v(2): b(0) = v(3)
    CopyMemory Rev4, b(0), 4
End Function

Private Function Rev2(ByVal Val As Integer) As Integer
Dim b(1) As Byte, v(1) As Byte
   CopyMemory v(0), Val, 2
   b(1) = v(0): b(0) = v(1)
   CopyMemory Rev2, b(0), 2
End Function

'get info about item
Private Function GetStr_Info(ByVal pItemInfo As Long, Optional bRelease As Boolean = True) As String
Dim tItemInfo As PST_TYPEINFO

    With tItemInfo
        If pItemInfo > 0 Then
            
            MyCopyMemory VarPtr(tItemInfo), pItemInfo, Len(tItemInfo)
            
            If .cbSize > 0 And .pDisplayName > 0 Then
                GetStr_Info = GetBSTR(tItemInfo.pDisplayName)
                If bRelease = True Then CoTaskMemFree pItemInfo
            End If
            
        End If
    End With
    
End Function

'get info about provider
Private Function GetStr_Provider(ByVal pProviderInfo As Long, Optional bRelease As Boolean = True) As String
Dim tProvider As PST_PROVIDERINFO

    With tProvider
        If pProviderInfo > 0 Then
            
            MyCopyMemory VarPtr(tProvider), pProviderInfo, Len(tProvider)
            
            If .cbSize > 0 And .pProviderName > 0 Then
                GetStr_Provider = GetBSTR(tProvider.pProviderName)
                If bRelease = True Then CoTaskMemFree pProviderInfo
            End If
            
        End If
    End With
    
End Function

Private Sub Form_Resize()
   If WindowState = 1 Then Exit Sub
   List1.Move 30, 30, Me.ScaleWidth - 60, Me.ScaleHeight - 60
End Sub

