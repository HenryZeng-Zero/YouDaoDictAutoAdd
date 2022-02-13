#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.

GetPos(key){
   ; 0 : 单词框
   ; 1 : 继续添加按钮
   ; 2 : 确定按钮
   ; 3 : 取消按钮
   ; 4 : 加入复习计划

   If (key = 0){
      Return "x68 y79"
   }Else If (key = 1){
      Return "x332 y407"
   }Else If (key = 2){
      Return "x243 y408"
   }Else If (key = 3){
      Return "x415 y406"
   }Else If (key = 4){
      Return "x47 y367"
   }
}

AutoClick(key){
   poaNow := GetPos(key)
   ControlClick, %poaNow%, ahk_class YdCommonCefWrapperWnd
}

::/sync::
   Sleep, 300

   If (Not WinExist("网易有道词典")){
      MsgBox, 0, 错误, 有道词典没有打开
      Return
   }

   Clipboard := ""

   Send, {CtrlDown}a{CtrlUp}
   Send, {CtrlDown}c{CtrlUp}

   ClipWait

   Words := Clipboard

   WinSet, AlwaysOnTop,On,,网易有道词典
   WinSet, AlwaysOnTop,Off,,网易有道词典

   MsgBox, 4, Tips, 是否需要操作提示?
   IfMsgBox Yes
   {
      MsgBox, 0, Tips:请完成步骤后点击确定, 请点击左侧的 [单词本]
      MsgBox, 0, Tips:请完成步骤后点击确定, 点击 [新旧排序] 右边的设置图标?选择添加单词
      MsgBox, 0, Tips:请完成步骤后点击确定, 选择一个分组或创建一个空的分组
      MsgBox, 0, Tips:请完成步骤后点击确定, 选择[语种]
   }

   AutoClick(4)
   ; 取消 [加入复习计划]

   StringReplace, Words, Words, `r`n ,`n , All

   

   Loop, Parse, Words, `n,
   {
      if not StrLen(A_LoopField) = 0
      {
         AutoClick(0)
         Sleep, 10
         SendInput, %A_LoopField%

         MsgBox, 3, Tips, 确认正常后点击 [是] ,跳过本单词点击 [否] 退出导入流程 [取消]
         IfMsgBox, Yes
            AutoClick(1)
         IfMsgBox, No
         {
            AutoClick(0)
            Sleep, 10
            Send, {CtrlDown}a{CtrlUp}
            Send, {BackSpace}
         }
         IfMsgBox, Cancel
         {
            Return
         }

         Sleep, 1000
      }
   }

   AutoClick(3)

Return