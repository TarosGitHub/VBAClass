VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 子クラスへの参照
Private m_child As Object

' メソッド
Private Sub Class_Initialize()
    ' DO NOTHING
End Sub

Private Sub Class_Terminate()
    ' DO NOTHING
End Sub

Public Sub SetChild(a_child As Object)
    Set m_child = a_child
End Sub

' 仮想関数
Public Sub FuncA()
    m_child.FuncA
End Sub
