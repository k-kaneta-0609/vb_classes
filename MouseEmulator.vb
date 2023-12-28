Imports System.Runtime.InteropServices

#Region _
"=========================　　クラス宣言部　　========================="
''' <summary>
''' マウス操作をシミュレート(擬似的に操作する)
''' </summary>
''' <remarks></remarks>
''' ==============================修正履歴==============================
''' 管理№{TAB}障害№{TAB}ﾊﾞｰｼﾞｮﾝ{TAB}YYYY/MM/DD{TAB}氏名{TAB}対応内容
''' 001                   4.0.0.0     2017/08/17     K.Kaneta 新規作成
''' ====================================================================
Public NotInheritable Class MouseEmulator

    Private Const INPUT_MOUSE = 0                   ' マウスイベント

    Private Const MOUSEEVENTF_MOVE = &H1            ' マウスを移動する
    Private Const MOUSEEVENTF_ABSOLUTE = &H8000     ' 絶対座標指定
    Private Const MOUSEEVENTF_LEFTDOWN = &H2        ' 左　ボタンを押す
    Private Const MOUSEEVENTF_LEFTUP = &H4          ' 左　ボタンを離す
    Private Const MOUSEEVENTF_RIGHTDOWN = &H8       ' 右　ボタンを押す
    Private Const MOUSEEVENTF_RIGHTUP = &H10        ' 右　ボタンを離す
    Private Const MOUSEEVENTF_MIDDLEDOWN = &H20     ' 中央ボタンを押す
    Private Const MOUSEEVENTF_MIDDLEUP = &H40       ' 中央ボタンを離す
    Private Const MOUSEEVENTF_WHEEL = &H800         ' ホイールを回転する

    ' マウスイベント(mouse_eventの引数と同様のデータ)
    <StructLayout(LayoutKind.Sequential)>
    Private Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    ' キーボードイベント(keybd_eventの引数と同様のデータ)
    <StructLayout(LayoutKind.Sequential)>
    Private Structure KEYBDINPUT
        Public wVk As Short
        Public wScan As Short
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    ' ハードウェアイベント
    <StructLayout(LayoutKind.Sequential)>
    Private Structure HARDWAREINPUT
        Public uMsg As Integer
        Public wParamL As Short
        Public wParamH As Short
    End Structure

    ' 各種イベント(SendInputの引数データ)
    <StructLayout(LayoutKind.Explicit)>
    Private Structure INPUT
        <FieldOffset(0)> Public type As Integer
        <FieldOffset(4)> Public mi As MOUSEINPUT
        <FieldOffset(4)> Public ki As KEYBDINPUT
        <FieldOffset(4)> Public hi As HARDWAREINPUT
    End Structure

    ' キー操作、マウス操作をシミュレート(擬似的に操作する)
    <DllImport("user32.dll")>
    Private Shared Sub SendInput(ByVal nInputs As Integer, ByRef pInputs As INPUT, ByVal cbsize As Integer)
    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks>インスタンス化を不可とする。</remarks>
    ''' ==============================修正履歴==============================
    ''' 管理№{TAB}修正内容
    ''' 001        新規作成
    ''' ====================================================================
    Private Sub New()
    End Sub

    ''' <summary>
    ''' マウスの左クリックをシミュレート
    ''' </summary>
    ''' <remarks></remarks>
    ''' ==============================修正履歴==============================
    ''' 管理№{TAB}修正内容
    ''' 001        新規作成
    ''' ====================================================================
    Public Shared Sub ClickLeft()

        Dim inp As INPUT() = New INPUT(1) {}

        ' マウスの左ボタンを押す
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        inp(0).mi.dx = 0
        inp(0).mi.dy = 0
        inp(0).mi.mouseData = 0
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        ' マウスの左ボタンを離す
        inp(1).type = INPUT_MOUSE
        inp(1).mi.dwFlags = MOUSEEVENTF_LEFTUP
        inp(1).mi.dx = 0
        inp(1).mi.dy = 0
        inp(1).mi.mouseData = 0
        inp(1).mi.dwExtraInfo = 0
        inp(1).mi.time = 0

        ' マウス操作実行
        Send(inp)

    End Sub

    ''' <summary>
    ''' マウスの中央クリックをシミュレート
    ''' </summary>
    ''' <remarks></remarks>
    ''' ==============================修正履歴==============================
    ''' 管理№{TAB}修正内容
    ''' 001        新規作成
    ''' ====================================================================
    Public Shared Sub ClickMiddle()

        Dim inp As INPUT() = New INPUT(1) {}

        ' マウスの中央ボタンを押す
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_MIDDLEDOWN
        inp(0).mi.dx = 0
        inp(0).mi.dy = 0
        inp(0).mi.mouseData = 0
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        ' マウスの中央ボタンを離す
        inp(1).type = INPUT_MOUSE
        inp(1).mi.dwFlags = MOUSEEVENTF_MIDDLEUP
        inp(1).mi.dx = 0
        inp(1).mi.dy = 0
        inp(1).mi.mouseData = 0
        inp(1).mi.dwExtraInfo = 0
        inp(1).mi.time = 0

        ' マウス操作実行
        Send(inp)

    End Sub

    ''' <summary>
    ''' マウスの右クリックをシミュレート
    ''' </summary>
    ''' <remarks></remarks>
    ''' ==============================修正履歴==============================
    ''' 管理№{TAB}修正内容
    ''' 001        新規作成
    ''' ====================================================================
    Public Shared Sub ClickRight()

        Dim inp As INPUT() = New INPUT(1) {}

        ' マウスの右ボタンを押す
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_RIGHTDOWN
        inp(0).mi.dx = 0
        inp(0).mi.dy = 0
        inp(0).mi.mouseData = 0
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        ' マウスの右ボタンを離す
        inp(1).type = INPUT_MOUSE
        inp(1).mi.dwFlags = MOUSEEVENTF_RIGHTUP
        inp(1).mi.dx = 0
        inp(1).mi.dy = 0
        inp(1).mi.mouseData = 0
        inp(1).mi.dwExtraInfo = 0
        inp(1).mi.time = 0

        ' マウス操作実行
        Send(inp)

    End Sub

    ''' <summary>
    ''' マウスの左ダブルクリックをシミュレート
    ''' </summary>
    ''' <remarks></remarks>
    ''' ==============================修正履歴==============================
    ''' 管理№{TAB}修正内容
    ''' 001        新規作成
    ''' ====================================================================
    Public Shared Sub DoubleClickLeft()

        Dim inp As INPUT() = New INPUT(3) {}

        ' マウスの右ボタンを押す
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        inp(0).mi.dx = 0
        inp(0).mi.dy = 0
        inp(0).mi.mouseData = 0
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        ' マウスの右ボタンを離す
        inp(1).type = INPUT_MOUSE
        inp(1).mi.dwFlags = MOUSEEVENTF_LEFTUP
        inp(1).mi.dx = 0
        inp(1).mi.dy = 0
        inp(1).mi.mouseData = 0
        inp(1).mi.dwExtraInfo = 0
        inp(1).mi.time = 0

        ' マウスの右ボタンを押す
        inp(2).type = INPUT_MOUSE
        inp(2).mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        inp(2).mi.dx = 0
        inp(2).mi.dy = 0
        inp(2).mi.mouseData = 0
        inp(2).mi.dwExtraInfo = 0
        inp(2).mi.time = 0

        ' マウスの右ボタンを離す
        inp(3).type = INPUT_MOUSE
        inp(3).mi.dwFlags = MOUSEEVENTF_LEFTUP
        inp(3).mi.dx = 0
        inp(3).mi.dy = 0
        inp(3).mi.mouseData = 0
        inp(3).mi.dwExtraInfo = 0
        inp(3).mi.time = 0

        ' マウス操作実行
        Send(inp)

    End Sub

    ''' <summary>
    ''' マウスの右ダブルクリックをシミュレート
    ''' </summary>
    ''' <remarks></remarks>
    ''' ==============================修正履歴==============================
    ''' 管理№{TAB}修正内容
    ''' 001        新規作成
    ''' ====================================================================
    Public Shared Sub DoubleClickRight()

        Dim inp As INPUT() = New INPUT(3) {}

        ' マウスの右ボタンを押す
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_RIGHTDOWN
        inp(0).mi.dx = 0
        inp(0).mi.dy = 0
        inp(0).mi.mouseData = 0
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        ' マウスの右ボタンを離す
        inp(1).type = INPUT_MOUSE
        inp(1).mi.dwFlags = MOUSEEVENTF_RIGHTUP
        inp(1).mi.dx = 0
        inp(1).mi.dy = 0
        inp(1).mi.mouseData = 0
        inp(1).mi.dwExtraInfo = 0
        inp(1).mi.time = 0

        ' マウスの右ボタンを押す
        inp(2).type = INPUT_MOUSE
        inp(2).mi.dwFlags = MOUSEEVENTF_RIGHTDOWN
        inp(2).mi.dx = 0
        inp(2).mi.dy = 0
        inp(2).mi.mouseData = 0
        inp(2).mi.dwExtraInfo = 0
        inp(2).mi.time = 0

        ' マウスの右ボタンを離す
        inp(3).type = INPUT_MOUSE
        inp(3).mi.dwFlags = MOUSEEVENTF_RIGHTUP
        inp(3).mi.dx = 0
        inp(3).mi.dy = 0
        inp(3).mi.mouseData = 0
        inp(3).mi.dwExtraInfo = 0
        inp(3).mi.time = 0

        ' マウス操作実行
        Send(inp)

    End Sub

    ''' <summary>
    ''' マウスホイールのシミュレート
    ''' </summary>
    ''' <param name="move">移動数(正数:上移動、負数:下移動)</param>
    ''' <remarks></remarks>
    ''' ==============================修正履歴==============================
    ''' 管理№{TAB}修正内容
    ''' 001        新規作成
    ''' ====================================================================
    Public Shared Sub Wheel(ByVal move As Integer)

        Dim inp As INPUT() = New INPUT(0) {}

        ' マウスの右ボタンを押す
        inp(0).type = INPUT_MOUSE
        inp(0).mi.dwFlags = MOUSEEVENTF_WHEEL
        inp(0).mi.dx = 0
        inp(0).mi.dy = 0
        inp(0).mi.mouseData = move
        inp(0).mi.dwExtraInfo = 0
        inp(0).mi.time = 0

        ' マウス操作実行
        Send(inp)

    End Sub

    ''' <summary>
    ''' マウス操作実行
    ''' </summary>
    ''' <param name="inp">操作内容</param>
    ''' <remarks></remarks>
    ''' ==============================修正履歴==============================
    ''' 管理№{TAB}修正内容
    ''' 001        新規作成
    ''' ====================================================================
    Private Shared Sub Send(ByVal inp As INPUT())

        ' マウス操作実行
        SendInput(inp.Length, inp(0), Marshal.SizeOf(inp(0)))

    End Sub

End Class

#End Region
