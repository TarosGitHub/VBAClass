VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Child00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private m_parent As Parent

' Methods
Private Sub Class_Initialize()
    Set m_parent = New Parent
    m_parent.SetChild Me
End Sub

Private Sub Class_Terminate()
    Set m_parent = Nothing
End Sub

Public Property Get Parent() As Parent
    Set Parent = m_parent
End Property

Public Sub FuncA()
    MsgBox "CChild00"
End Sub

