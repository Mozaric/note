----

# Visual Basic .NET Note

----

# 目錄

+ [Visual Studio 與 .NET Framework]()
+ [VB 程式的編譯]()
+ [變數與常數 (Variable and Constant)]()
	+ [變數的宣告 (Variable Declaration)]()
		+ [Suffixes]()
	+ [靜態修飾子 (Static)]()
	+ [常數 (Constant)]()
	+ [列舉 (Enum)]()
	+ [日期 (Date)]()
	+ [字串 (String)]()
		+ [字串的分割 (Split)]()
		+ [字串的結合 (Join)]()
		+ [PadRight & PadLeft]()
		+ [換行字元]()
		+ [格式化輸出]()
	+ [變數型態轉換]()
		+ [Parse]()
		+ [TryParse]()
+ [運算子 (Operator)]()
	+ [關係運算子 (Relation Operators)]()
	+ [邏輯運算子 (Logical Operators)]()
	+ [位元運算子 (Bitwise Operators)]()
	+ [位元位移運算子]()
	+ [Like Operator]()
+ [決策 (Decision Making)]()
	+ [If Then Else]()
	+ [IIf(Inline If)]()
	+ [If 運算子]()
	+ [Select Case]()
	+ [Microsoft.VisualBasic.Switch]()
	+ [Choose]()
+ [迴圈 (Loop)]()
	+ [For Next]()
	+ [For Each Next]()
	+ [While]()
	+ [Do While Loop]()
	+ [Do Until Loop]()
	+ [Do Loop While]()
	+ [Do Loop Until]()
+ [陣列 (Array)]()
	+ [多維陣列 (Multi-Dimension Array)]()
	+ [改變陣列大小 (Change Size of Array)]()
	+ [LBound & UBound]()
	+ [不規則陣列 (Jagged Array)]()
	+ [System.Array 類別]()
+ [副程式與函式 (Sub and Function)]()
	+ [副程式 (Sub)]()
	+ [函式 (Function)]()
	+ [Sub 與 Function 的不同]()
	+ [參數傳遞機制]()
	+ [選擇性參數 (Optional)]()
	+ [當參數是陣列時]()
	+ [當輸入的參數數目不確定時]()
+ [偵錯與例外處理 (Exception)]()
	+ [偵錯模式 (Debug Mode)]()
	+ [結構化例外狀況]()
	+ [Try Catch Finally]()
	+ [自訂例外處理]()
	+ [On Error]()
		+ [On Error Goto <label>]()
		+ [On Error Resume Next]()
+ [類別、模組與結構 (Class, Module and Structure)]()
	+ [類別 (Class)]()
		+ [Shared]()
	+ [模組 (Module)]()
	+ [類別與模組的差異]()
	+ [結構 (Structure)]()
	+ [變數的使用範圍]()
+ [常用的資料結構]()
	+ [List]()
	+ [ArrayList]()
	+ [Queue]()
+ [型別系統的實值、參考型別 (Value Type and Reference Type)]()
	+ [實值型別(Value Type)]()
	+ [參考型別(Reference Type)]()
+ [Thread]()
	+ [Background Thread 與 Foreground Thread 的差異]()
	+ [SyncLock]()
	+ [ManualResetEvent]()
	+ [委派 (Delegate)]()
		+ [跨執行緒執行]()
+ [其他]()
	+ [System.Math 類別]()
	+ [亂數 (Random Number)]()
	+ [DataAndTime]()
	+ [檢查資料型別]()
	+ [IsArray]()
	+ [CurDir]()
	+ [With ... End With]()
	+ [非同步作業 (Asynchronous Programming Model)]()
	+ [Region]()
	+ [Line Continuation]()
	+ [Comments]()
+ [視窗程式設計]()
	+ [Windows Form]()
		+ [當 Project 內有多個 Form 時]()
		+ [Sub Main()]()
		+ [Application 類別]()
	+ [控制項]()
		+ [InputBox]()
		+ [MsgBox]()
		+ [Label]()
		+ [TextBox]()
		+ [Button]()
		+ [RadioButton]()
		+ [CheckBox]()
		+ [DateTimePicker]()
		+ [Timer]()
		+ [RitchTextBox]()
		+ [MaskedTextBox]()
		+ [ToolTip]()
		+ [Help 類別]()
		+ [日期時間控制項]()
			+ [MonthCalendar]()
			+ [DateTimePicker]()
		+ [版面控制]()
			+ [GroupBox]()
			+ [FlowLayoutPanel]()
			+ [TableLayoutPanel]()
			+ [TabControl]()
		+ [清單]()
			+ [ListBox]()
			+ [ComboBox]()
			+ [CheckedListBox]()
		+ [提供檢視的控制項]()
			+ [ImageList]()
			+ [ListView]()
			+ [TreeView]()
		+ [OpenFileDialog]()
		+ [MenuStrip]()
		+ [DataTable]()
		+ [DataGridView]()
		+ [Font & ForeColor]()
		+ [BeginUpdate]()
		+ [設定按 Tab 時的跳的順序]()
	+ [控制項的事件]()
		+ [Event Handler 的參數列]()
		+ [共用事件處理程序]()
		+ [新增/移除事件處理程序]()
		+ [滑鼠事件]()
		+ [鍵盤事件]()
		+ [Custom Event]()

----

# Visual Studio 與 .NET Framework

.NET Framework 提供 Visual Studio 一個安全性高、整合性高的工作環境。使用者可以使用 Visual Basic 、 C# 、 Visual C++ 等程式語言做開發。

.NET Framework 的特色：
> 一致性的設計環境。使用 Visual Studio 撰寫程式碼，**不同的程式語言能夠互相參照**。
> 提高程式碼的安全執行環境。
> 提供特定應用程式開發領域所需的程式庫，讓 Windows 與 Web 應用程式開發時有一致的開發設計環境。

# VB 程式的編譯

VB 程式碼 -> 使用 VB 編譯器 -> 編譯成 MSIL 語言 -> 再使用 JIT 編譯器 -> 編譯成機械碼 -> 執行程式
MSIL ： Microsoft Intermediate Language
JIT ： Just-in-Time

# 變數與常數 (Variable and Constant)

## 變數的宣告 (Variable Declaration)

Dim | Public | Private <變數名稱> As <變數型態>
eg. `Dim num As Integer`

變數型態： Byte 、 Boolean 、 Char 、 Double 、 Integer 、 String 。[Or Here](https://www.tutorialspoint.com/vb.net/vb.net_data_types.htm)

如果想用 VB 內的識別字來當變數名稱時，需在兩側加上中括號，例：
```vb
Dim [TextBox] As Integer = 0

Public Sub InvokeButtonText(ByVal [Button] As Button, ByVal [Text] As String)
    If Button.InvokeRequired Then
        Dim t As New DlgButtonText(AddressOf InvokeButtonText)
        Button.Invoke(t, New Object() {Button, Text})
    Else
        Button.Text = Text
    End If
End Sub
```

### Suffixes

使用此符號可讓變數宣告變得較簡短。

```vb
%-Integer
$-String
@-Decimal
&-Long
#-Double
!-Single
```
eg. `Dim num As Integer` 可寫成 `Dim num%`

## 靜態修飾子 (Static)

一般的變數的生命週期會隨著程序的停止而消失，如果想要保留區域變數的值，可使用 Static 宣告變數。
eg. `Static count As Integer = 0`

## 常數 (Constant)

Const <變數名稱> As <變數型態> = <變數數值>
eg. `Const PI As Double = 3.1415926`

## 列舉 (Enum)

Enum <enumration-name> (As <type>)
	<member1> (= <init-value>)
	<member2>
	<member3>
End Enum

eg.
```vb
Enum Tints
	Red = 1
	Orange
	Yellow
End Enum

label1.Color = Tints.Orange
```

eg.
```vb
Enum Money
	Thousand = 1000
	Hundred = 100
	Ten = 10
	Dollar = 1
End Enum

Dim coin As Money
coin = Money.Thousand + Money.Hundred
```

## 日期 (Date)

Data 型別可用來放日期/時間的資料，但前後需以 # 符號做區隔。
eg. `Dim d As Date = #11/25/2017 15:00 PM#`

## 字串 (String)

Dim <str-name> As String

對VB來說，字串是不變的，當你給字串指定新值時，他其實是放棄原來字串，改為指向新字串。**字串是參考型別**。
eg.
```vb
Dim str As String = "NewYork"
For index = 0 to 6
	Console.WriteLine("Index[{0}] Char'{1}'", index, str.Chars(index)) '印出 str 字串中 index 0 到 6 的每個字元
Next

int length = str.Length '取得字串長度
```

常用方法： CopyTo() 、 Concat() 、 Join() 、 Insert() 、 Replace() 、 Split() 。

[String](https://www.tutorialspoint.com/vb.net/vb.net_strings.htm)
[String Method in VB](https://msdn.microsoft.com/zh-tw/library/dd789093.aspx)

### 字串的分割 (Split)

根據陣列中的字元分割字串成子字串。

eg. 

### 字串的結合 (Join)

串連字串陣列的所有項目，並在每個項目之間使用指定的分隔符號。

eg. `String.Join(separator, strArray)`

### PadRight & PadLeft

設定輸出字元靠右或靠左對齊。

eg.
```vb
Dim str1 As String = "blah".PadLeft(8, '_');
'output: "____blah"

Dim str2 As String = "blah".PadRight(8, '_');
'output: "blah____"
```

### 換行字元

Chr(10) 或 vbCrLf

### 格式化輸出

輸出資料時可透過 Format 函數來進行格式化的輸出，包含字型、色彩、日期和數值。
[String.Format](https://msdn.microsoft.com/zh-tw/library/microsoft.visualbasic.strings.format(v=vs.110).aspx)

## 變數型態轉換

CInt(var)：將 var 轉成 Integer
CDbl(var)：將 var 轉成 Double
CStr(var)：將 var 轉成 String
CByte(var)：將 var 轉成 Byte

ToString ，任何型態的變數都有此方法，將變數轉換成 String。
eg. `MsgBox(num.ToString())`

### Parse

將 String 轉成數值、Date。
eg. `Dim number As Integer = Integer.Parse(Label1.Text)`

### TryParse

有時會有遇到意料外的輸入，輸入的字串不是數字，但若仍將其轉成數字時，Parse() 會失敗並出 bug，此時可用 TryParse()，即使轉換失敗也不會卡住。
eg. `Integer.TryParse(str, num)`

# 運算子 (Operator)

`+` (加)、 `-` (減)、 `*` (乘)、 `/` (除)、 `\ ` (取商數)、 `Mod` (取餘數)、 `^` (次方運算)

## 關係運算子 (Relation Operators)

`>` (大於)、 `>=` (大於等於)、 `<` (小於)、 `<=` (小於等於)、 `=` (等於)、 `<>` (不等於)

## 邏輯運算子 (Logical Operators)

`OR` 、 `Not` 、 `And` 、 `Xor` 、 `AndAlso` 、 `OrElse`

[Logical Operators](https://msdn.microsoft.com/zh-tw/library/wz3k228a.aspx)

## 位元運算子 (Bitwise Operators)

`And` 、 `Or` 、 `Xor`

## 位元位移運算子

`>>` 、 `<<`

## Like Operator

根據模式比較字串。

[Like Operator](https://msdn.microsoft.com/zh-tw/library/swf8kaxw.aspx)

# 決策 (Decision Making)

## If Then Else

If <判斷式> Then
    <判斷式為 true 執行此處>
Else
    <判斷式為 false 執行此處>
End If

或是另外一種較簡短的寫法：

If <判斷式> Then <執行此處第一行> : <執行此處第二行> : <執行此處第三行>

## IIf(Inline If)

IIf(<判斷式>, <判斷式為true時回傳>, <判斷式為false時回傳>)
但要注意，IIf的運作是TruePart與FalsePart都會先進行運算，萬一我們裡面的運算有可能會有除以零的情況，就會出錯。所以有些判斷還是以If Else寫會比較好。

eg. `num = IIf(n1 > n2, n1, n2)`

[IIf的陷阱](http://trufflepenne.blogspot.tw/2013/11/vbnetiif.html)

## If 運算子

If(<判斷式>, <判斷式為true時回傳>, <判斷式為false時回傳>)

如果改用 If 函數，由於它具有短路（short circuit）的特性，只會評估其中一個引數值，因此相對保險一些。

eg. `num = If(n1 > n2, n1, n2)`

[If 運算子](https://msdn.microsoft.com/zh-tw/library/bb513985.aspx)

## Select Case

Select Case <var>
	Case <value1>
		<do-something1>
	Case <value2>
		<do-something2>
	Else
		<do-something3>
End Select

eg.
```vb
Select Case age
	Case Is >= 18
		MsgBox("限制級")
	Case 12 To 17
		MsgBox("輔導級")
	Case 6, 7, 8, 9, 10, 11
		MsgBox("保護級")
	Case Is < 6
		MsgBox("普遍級")
	Case Else
		MsgBox("蛤?")
End Select
```

功能與 C++ 中的 switch case 相似，但其 Case 後可輸入的東西較多樣，且不需輸入 break 來防止其跑到下一個 Case 內。

另外，

Select Case True
	Case <value1>
		<do-something1>
	Case <value2>
		<do-something2>
	Else
		<do-something3>
End Select

可用來取代 If Else 的結構，如：
```vb
Select Case True
	Case testVariable < 0
		Console.Write("You must supply a positive value.")
	Case testVariable > 10
		Console.Write("Please enter a number from 0-10.")
	Case True
		Call DoWork(testVariable)
End Select
```

## Microsoft.VisualBasic.Switch

Microsoft.VisualBasic.Switch(<判斷式1>, <值1>, <判斷式2>, <值2>, ...)
如果<判斷式1>為 true ，就會回傳<值1>，若不成立，則看<判斷式2>，如果<判斷式2>為 true ，就會回傳<值2>，以此類推。

eg.
`str = CStr(Microsoft.VisualBasic.Switch(age >= 18, "限制級", age >= 12, "輔導級", age >= 6, "保護級", age > 0, "普遍級"))`

## Choose

Choose(index, <值1>, <值2>, <值3>, ...)
根據 index 的數值來回傳值， index 值為 1 時，回傳<值1>， index 值為 2 時，回傳<值2>， index 值為 3 時，回傳<值3>，以此類推。

eg.
`size = Choose(i, "S", "M", "L", "XL")`

# 迴圈 (Loop)

## For Next

For <variable> (As <type>) = <start-value> To <end-value> Step <step-value>
    <do-something>
Next <variable>

eg.
```vb
For i As Integer = 0 To 10 Step 2
	<do-something>
Next i
```

另外：
`Continue`等同於C++的`continue;`，馬上進行下一次的迴圈，用法為`Continue For | Do | While`
`Exit`等同於C++的`break;`，結束並離開目前執行的迴圈，用法為`Exit For | Do | While`

## For Each Next

For Each <element> (As <type>) In <group>
    <do-something>
Next <element>

對<group>內的每個<element>做<do-something>

eg.
```vb
Dim i() As Integer = {1, 4, 7}
Dim s() As Char = {'a', 'b', 'c'}
Dim str As String = ""

For Each num As Integer In i
    For Each ch As Char In s
        str = str & num & ch
    Next ch
Next num

MsgBox(str)
```
output: 1a1b1c4a4b4c7a7b7c

## While

While <判斷式>
	<do-something>
End While

當判斷式為 true 時，進入迴圈，直到判斷式為 false 時，離開迴圈。**與 Do While Loop 功能相同**

## Do While Loop

Do While <判斷式>
	<do-something>
Loop

當判斷式為 true 時，進入迴圈，直到判斷式為 false 時，離開迴圈。**與 While 功能相同**

## Do Until Loop

Do Until <判斷式>
	<do-something>
Loop

當判斷式為 false 時，進入迴圈，直到判斷式為 true 時，離開迴圈。

## Do Loop While

Do
	<do-something>
Loop While <判斷式>

**先進入回圈內執行一次**，然後再判斷，當判斷式為 true 時，進入迴圈，判斷式為 false 時，離開。

## Do Loop Until

Do
	<do-something>
Loop Until <判斷式>

**先進入回圈內執行一次**，然後再判斷，當判斷式為 false 時，進入迴圈，判斷式為 true 時，離開。

# 陣列 (Array)

Dim <array-name>(<max-index-val>) As <type> **包含了 max-index-val + 1 個項目，index範圍從 0 到 max-index-val**
Dim <array-name>(<lower-bound> To <uppder-bound>) As <type>
Dim <array-name>() As <type> = { <value1>, <value2>, <value3>, ..., <valueN> }
<array-name>(<index>) = <value>

eg. `Dim n(10) As Integer`
eg. `Dim n(0 To 10) As Integer`
eg. `Dim n() As Integer = {5, 10, 22, 39}`
eg. `n(2) = 50`

## 多維陣列 (Multi-Dimension Array)

Dim <array-name>(row, col) As <type>

eg.
```vb
Dim n(2, 3) As Integer
n(0, 0) = 1
n(2, 3) = 5
```

eg. `Dim num(,) As Integer = {{1, 2, 3}, {4, 5, 6}}`

## 改變陣列大小 (Change Size of Array)

ReDim (Preserve) <array-name>(<new-max-index-value>)
重設陣列大小後，陣列內的元素會被完全清空，除非加上 Preserve 關鍵字來進行保留內容。

eg.
```vb
Dim n(10) As Integer
ReDim n(20)
ReDim Preserve n(30)
```

## LBound & UBound

LBound(<arr>(, <arrRank>))

取得陣列 <arr> 的**最小**索引值 (index)， <arrRank> 代表陣列的維度，一維陣列輸入 1 ，二維陣列輸入 2 。

UBound(<arr>(, <arrRank>))

取得陣列 <arr> 的**最大**索引值 (index)， <arrRank> 代表陣列的維度，一維陣列輸入 1 ，二維陣列輸入 2 。

## 不規則陣列 (Jagged Array)

eg.
```vb
Dim num1() As Integer = {11, 22, 33}
Dim num2() As Integer = {20, 50}
Dim num3(1)() As Integer
num3(0) = num1
num3(1) = num2
```

eg.
```vb
Dim num3()() As Integer = New Integer(1)() {...}
num3(0) = New Integer() {11, 22, 33}
num3(1) = New Integer() {20, 50}
```

eg.
```vb
Dim num3()() As Integer = {New Integer() {11, 22, 33}, New Integer() {20, 50}}
```

## System.Array 類別

每個 array 本身有提供些 method ，或是 Array 關鍵字內也有 method 可使用。

eg.
```vb
Dim std() As String = {"Apple", "October", "Monday", "Frank"}

Dim len As Integer = std.Length() '取得 std 的長度
Array.Sort(std) '將 std 進行排列
Dim idx As Integer = Array.BinarySearch(std, "Monday") '在 std 內搜尋 "Mondy" 並回傳其 index
Array.Reverse(std) '將 std 反轉
```

eg.
```vb
Dim num(,) As Integer = {{1, 2, 3}, {4, 5, 6}}

Dim rk As Integer = num.Rank '取得 num 的維度數
Dim len As Integer = num.GetLength(1) '取得 num(1) 的長度
num.SetValue(30, 1, 1) '將 num(1, 1) 的值改為 30
```

[System.Array](https://msdn.microsoft.com/zh-tw/library/system.array(v=vs.110).aspx)

# 副程式與函式 (Sub and Function)

## 副程式 (Sub)

(Public | Private) Sub <sub-name>(<parameters>)
	<program-region>
End Sub

呼叫 Sub 時使用 call ，例如`call dispaly()`，亦可將 call 省略。

eg.
```vb
Private Sub Display(ByVal name As String)
	txtBox.Text = "Hi" & name
End Sub

Display("Frank")
```

新增模組檔案方式為：專案右鍵 -> 加入新增項目 -> 模組。

可用`Exit Sub`來離開 Sub 。

## 函式 (Function)

(Public | Private) Function <function-name>(<parameters>) As <return-var-type>
	<program-region>
	return <var> | <function-name> = <value>
End Function

eg.
```vb
Public Function totalScore1(ByVal grade1 As Integer, ByVal grade2 As Integer, ByVal grade3 As Integer) As Integer
	return (grade1 + grade2 + grade3)
End Function

Public Function totalScore2(ByVal grade1 As Integer, ByVal grade2 As Integer, ByVal grade3 As Integer) As Integer
	totalScore2 = (grade1 + grade2 + grade3)
End Function
```

## Sub 與 Function 的不同

Sub 可以傳入參數，但不會有回傳值。 Event 屬於 Sub ，如 mouse click 之類的 event 。
Function 可以傳入參數，但要有回傳值。

## 參數傳遞機制

ByVal ，傳值。
ByRef ，傳址。

## 選擇性參數 (Optional)

可用 Optional 關鍵字將某個參數設定為選擇項，但同時要給其參數預設值。
選擇性參數只能放在參數 list 的最右側，可放多個。

eg.
```vb
Sub Test(ByRef n1 As Integer, ByVal n2 As Integer = 10, ByVal n3 As Integer = 30)
	n1 = n2 + n3
End Sub

'呼叫時
Test(num1)
Test(num2, , 20)
Test(n1 := num, n3 := 40)
```

## 當參數是陣列時

eg.
```vb
Sub Test(ByVal str As String, ByVal score() As String)
	<do-something>
End Sub
```

要注意陣列與字串是**參考型態**的變數，所以即使傳入時是用 ByVal ，但若在 Sub/Function 內改變其數值，仍會直接改到在外部該變數的值。

## 當輸入的參數數目不確定時

一般的選擇性參數是在參數前方加上 Optional 關鍵字，選擇性的陣列參數是在前方加上 ParamArray 關鍵字。
ParamArray 關鍵字也讓你可以使用多個引數，但不可以和 ByVal 、 ByRef 或 Optional 共用。
ParamArray 是用 ByVal 的方式傳遞。

eg. 
```vb
Public Sub display(ByVal name As String, ParamArray score() As String)
	Dim str As String = ""

	str &= name & vbCrLf
	For index = 0 To UBound(score)
		str &= score(index) & vbCrLf
	Next

	MsgBox(str)
End Sub
```

呼叫時

```vb
Dim name As String = "Frank"
Dim score() As String = {"11", "22", "33", "44"}

display(name, score)
```

或是

```vb
Dim name As String = "Frank"
Dim grade1 As String = "55"
Dim grade2 As String = "66"
Dim grade3 As String = "66"

display(name, grade1, grade2, grade3)
```

經過 ParamArray 修飾的 sub 可以丟入一個陣列，或是丟入多個參數。

# 偵錯與例外處理 (Exception)

在預防程式出錯時，我們通常會先設想幾種例外狀況(例如做除法時除以零)，然後當其狀況快要發生時，我們要避免程式去執行該錯誤，而是搶先一步中斷程式並跳出通知告知使用者，此種預防錯誤的功能通常可以用 If Else 來寫，但若程式內到處都是相同的語法， If Else ，看的人就無法一眼看出這程式重要的運算是哪個區塊，預防錯誤的程式碼又是哪塊，導致閱讀的困難，所以如果在寫預防錯誤的程式碼時，也盡量使用語言內提供的語法去寫(Try Catch 、 On Error)，做個區隔。

一般的程式錯誤可以分成三種，語法錯誤、邏輯錯誤、執行階段錯誤。

> 語法錯誤，一般情況下IDE會自動幫我們抓出來並告知。
> 邏輯錯誤，此種錯誤就只能使用 Debug 模式設定中斷點，逐步去看程式運行是否如同自己所想，或是自己重新檢查程式碼慢慢抓。
> 執行階段錯誤，在執行過程中發生不可預期的錯誤，例如要讀取光碟機內的光碟，結果光碟是壞掉的，此時就會擲出例外狀況。

## 偵錯模式 (Debug Mode)

逐步執行：一行一行執行。
不進入函數：一樣為一行一行執行，但若下一行是函數，則會進入函數並一次執行直到函數結束並跳出再暫停。
跳離函數：若目前在函數內，則函數內剩下的行數會一次執行完並跳出函數再暫停。
在 Debug 模式下還能看到，目前的所有變數的值。

## 結構化例外狀況

發生的原因並非應用程式發生問題，而是使用者不當的操作。例如：光碟機忘記放光碟片，或是操作環境問題，例如：想下載檔案卻沒有與網路連線。由於都是不可預期發生的狀況，稱為結構化例外狀況。此時就要用 Try 來預防錯誤/偵測。

## Try Catch Finally

Try
	<statements that may cause an error>
Catch
	<statements to use when an error occurs>
Finally
	<statements to use no matter what happens>
End Try

eg.
```vb
Dim str() As String = {"5", "6", "7", "8"}
Dim str2 As String = ""

For index = 0 To 4
	Try
		str2 &= str(index) & vbCrLf
	Catch ex As Exception
		MsgBox(ex.Message, vbOKOnly + vbCritical, "Error")
	Finally
		MsgBox("陣列元素" & index & vbCrLf, vbOKOnly + vbInformation, "Finally區塊")
	End Try
Next
```

[Try Catch Finally](https://msdn.microsoft.com/zh-tw/library/fk6t46tz.aspx)

Exception 常用的屬性： ex.Message 、 ex.GetType.ToString()

## 自訂例外處理

Throw <exception>

當我們遇到些不是 System 內的 Exception ，我們可以自定義 Exception 並將他擲出 (Throw) 。例如：今天我們寫了一個每個月分天數顯示的程式，使用者只要輸入月份，就會顯示該月份有幾天，可是今天使用者輸入了 15 ，而根本沒有 15 月這種東西，此時就可以用 Try Catch 加上 Throw 來擲出 Exception 。

eg.
```vb
Public Sub checkMonth(ByVal m As Integer)
	Dim str As String = ""

	If m > 12 Or m < 1 Then
		Throw New ArgumentOutOfRangeException
	Else
		str &= m & " 月有 "
		Select Case m
			Case 2
				str &= 28
			Case 4, 6, 9, 11
				str &= 30
			Case Else
				str &= 31
		End Select
		str &= " 天"

		MsgBox(str)
	End If
End Sub
```

使用時：
```vb
Dim str As String = InputBox("請輸入月份")

Try
	checkMonth(CInt(str))
Catch ex1 As ArgumentOutOfRangeException
	MsgBox("月份錯誤")
Catch ex2 As Exception
	MsgBox(ex2.Message & vbCrLf & ex2.GetType.ToString())
End Try
```

Throw 大多是程式內部會自己 Throw 例外出來，但我們也可以自己加入想要 Throw 例外的時機與內容。

## On Error

常用屬性： Err.Number (看錯誤代碼)、 Err.Description (看錯誤訊息)

### On Error Goto <label>

鍵入 On Error Goto <label> 來啟動錯誤偵測，當發生錯誤時，會跳到 <label> 那一行

eg.
```vb
Private Sub errorButton_Click(sender As Object, e As EventArgs) Handles errorButton.Click
	Dim n1 As Integer = 17
	Dim n2 As Integer = 0

	On Error GoTo ErrHandler
	Dim n3 As Integer = n1 / n2
	MsgBox(n1 & " / " & n2 & " = " & n3)

ErrHandler:
	If Err.Number <> 0 Then
		MsgBox(Err.Number & " " & Err.Description, , "Error")
		Err.Clear()
	End If
End Sub
```

### On Error Resume Next

鍵入 On Error Resume Next 來啟動錯誤偵測，但當發生錯誤時，會跳過錯誤繼續執行。

# 類別、模組與結構 (Class, Module and Structure)

## 類別 (Class)

<access-modifier> Class <class-name>
	<member-declaration>
End Class

類別只提供物件的樣板，建立類別後，必須以 **New** 關鍵字來實體化物件。

eg. `Dim <obj> As New <class-name>`

Class 的建構式：
```vb
Sub New() {}
Sub New(ByVal s As String) {}
Sub New(ByVal s As String, i As Integer) {}
```

### Shared



## 模組 (Module)

傳統的 VB 會使用 **Module** 來設立應用程式的**公用變數及程序**。

Module <module-name>
	<member-declaration>
End Module

## 類別與模組的差異

Class 與 Module 看起來很相似，但其實兩者差異相當大。

**Module 底下的函式可以直接呼叫使用**，而 Class 則是需要先以 New 關鍵字將物件實體化才能使用底下的函式。
**Class 可以依序需求建立多個物件**，而 Module 則無法以實體化方式產生多份。

## 結構 (Structure)

(<access-modifier>) Structure <structure-name>
	Dim <member-name> As <data-type>
	Dim <member-name> As <data-type>
End Structure

## 變數的使用範圍

在 Class 、 Module 、 Structure 中使用 **Private** 修飾的變數、函式，只能使用於該 Class 、 Module 、 Sturcture ，若是 **Public** 、 **Dim** ，可用於整個 Project 。

# 常用的資料結構

## List

List 只能儲存同一類型的資料。提供搜尋、排序和管理清單的方法。

宣告時要告知是要存何種類型的資料，例`Dim L As New List(Of T)`。

常用屬性： Count 、 Item()
常用方法： Add() 、 AddRange() 、 Clear() 、 IndexOf() 、 Insert() 、 Remove() 、 RemoveAt() 、 Sort()

[List](https://msdn.microsoft.com/zh-tw/library/6sh2ey19(v=vs.110).aspx)

## ArrayList

ArrayList 的元素型別為 Object ，也就是你丟什麼東西給 ArrayList 它都會變成 Object ，然後要使用時再轉換成你要的型別。

常用屬性： Count 、 Item()
常用方法： Add() 、 AddRange() 、 Clear() 、 IndexOf() 、 Insert() 、 Remove() 、 RemoveAt() 、 Sort()

[ArrayList](https://msdn.microsoft.com/zh-tw/library/system.collections.arraylist(v=vs.110).aspx)
[ArrayList 與 List 執行效能比較](https://dotblogs.com.tw/yc421206/archive/2009/10/22/11213.aspx)

## Queue

Queue 的元素型別為 Object ， Queue 是用先進先出的方式處理物件的集合，例如到銀行排隊，先排的人先處理。

常用屬性： Count
常用方法： Enqueue() 、 Dequeue() 、 Peek()

[Queue](https://msdn.microsoft.com/zh-tw/library/system.collections.queue(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1)
[一般集合 - 佇列 Queue 類別 / 堆疊 Stack 類別](https://dotblogs.com.tw/yc421206/archive/2009/01/23/6930.aspx)

# 型別系統的實值、參考型別 (Value Type and Reference Type)

## 實值型別(Value Type)

實值型別會將變數直接存放於記憶體的 Stack 區塊中，而資料大小是固定的。由於記憶體採用堆疊的方式儲存，所以變數會有生命週期，執行程序停止時，堆疊也會消失。

eg.
```vb
Dim x As Integer = 25
Dim y As Integer = 30
y = x
```

## 參考型別(Reference Type)

宣告參考型別時，會使用記憶體 Stack 、 Managed Heap 區塊， Stack 的變數名稱紀錄著 Managed Heap 配置的記憶體位址，而變數值會存放於 Managed Heap 內，所以資料大小並非固定。

eg.
```vb
Dim myForm As Form1
myForm = New Form1()
myForm.ForeColor = Color.Blue
```

以 New 關鍵字將 myForm 初始化時， Managed Heap 才會配置記憶體位址，因此 Stack 區塊存的是一個指向 Managed Heap 記憶體位址的指標。

# Thread

執行緒

```vb
Sub Test()
	Dim thd_test As New Thread(AddressOf task_test)
	Dim i As Integer = 10
	thd_test.IsBackground = True
	thd_test.Name = "test"
	thd_test.Start(i)
End Sub

Sub task_test(num As Integer)
	<do-something>
End Sub
```

## Background Thread 與 Foreground Thread 的差異

若主程序已下達中止工作命令了，有任一前景執行緒 (Foreground Thread) 尚未完成工作，**程序不會立即中止**，需待前景執行緒完成工作後才會終止。
反之，背景執行緒 (Background Thread) 不管工作有沒有完成，一但收到中止命令，**馬上就停下手邊的工作中止工作**。

## SyncLock

使用執行緒時常會共用一些資源(方法)，為了不讓執行緒同時間進入同一個資源，避免造成資源在演算過程中發生錯亂，可以使用 `SyncLock` 來鎖定資源，等待執行緒工作完成後才會自動解除鎖定，讓其他的執行緒來使用資源，


```vb
SyncLock DataGridViewww
	<do-something>
End SyncLock
```

上述程式碼代表，鎖定 DataGridViewww 物件，直到 SyncLock 區塊結束才讓其他執行緒來用 DataGridViewww 物件。

## ManualResetEvent

用來管理多執行緒同步的問題

Public mre As New ManualResetEvent(False)
後方為 False 才能讓 WaitOne 有阻塞效果

常用方法： WaitOne() 、 Set() 、 Reset()

使用 WaitOne() 來使執行緒暫停，使用 Set() 讓執行緒繼續運作，通常 WaitOne() 會放在執行緒內，然後把 Set() 放在執行緒外面。
使用過一次 WaitOne() 、 Set() 後， WaitOne() 會暫時失去阻塞功能，必須使用 Reset() 來讓其恢復功能。

## 委派 (Delegate)

委派最常見的用處，就是將我們自己的 Function 當成參數，傳到另一個 Function 來跑。

舉例來說，今天我們用 Bubble Sort 來排列一整數陣列，如果我們要遞增排列，就要寫一個 Function 來做，如果要遞減排列，又要在寫一個 Function ，但是遞增與遞減兩個程式其實只會在判斷式那裡差一個運算符號，此時我們就可以使用委派，把我們的判斷式另外寫成一個 Function ，然後把它當成參數，丟入執行 Bubble Sort 的 Function 內。

首先宣告我們的 Delegate ：
```vb
Delegate Function dlgOperator(ByVal x As Integer, ByVal y As Integer) As Boolean
```

然後是 Bubble Sort 的演算法：
```vb
Public Sub DoSort(ByRef data() As Integer, ByVal op As dlgOperator)
	Dim i, j, temp As Integer

	For i = 0 To UBound(data)
		For j = i + 1 To UBound(data)
			If op.Invoke(data(i), data(j)) Then
				temp = data(i)
				data(i) = data(j)
				data(j) = temp
			End If
		Next
	Next
End Sub
```
我們把自定義的判斷式 dlgOperator 當成參數，然後在程式內使用 Invoke 來呼叫他。

然後是我們自定義的判斷式：
```vb
Public Function isLarger(ByVal x As Integer, ByVal y As Integer) As Boolean
	If x > y Then
		Return True
	Else
		Return False
	End If
End Function

Public Function isSmaller(ByVal x As Integer, ByVal y As Integer) As Boolean
	If x < y Then
		Return True
	Else
		Return False
	End If
End Function
```

今天當我們想要遞增排序排列時，就把 isLarger 丟進去 DoSort ：
```vb
DoSort(data, AddressOf isLarger)
```

想要遞減排序排列時，就把 isSmaller 丟進去 DoSort ：
```vb
DoSort(data, AddressOf isSmaller)
```

如果我們今天又有不同的判斷式，我們原本的 DoSort 也完全不用動，因為演算法完全沒變，只是判斷式改變而已，所以只需要在寫一個判斷式 Function ，然後在呼叫 DoSort 時當成參數丟進去即可。

### 跨執行緒執行

今天我們開啟一個執行緒，然後讓他去不斷的(無窮迴圈)判斷某項事物，並且要把判斷的結果顯示在 TextBox 內，如果你在執行緒內使用 `TextBox.Text = ...` ，此時會發生錯誤，他會告知你「跨執行緒作業無效: 存取控制項 'TextBox' 時所使用的執行緒與建立控制項的執行緒不同」，此時就要使用 Delegate 。

```vb
Private Delegate Sub DlgControlText(ByVal [Control] As Control, ByVal [Text] As String)

Public Sub InvokeControlText(ByVal [Control] As Control, ByVal Text As String)
	If Control.InvokeRequired Then										'判斷 Control 物件是否在同一個執行緒上
		Dim t As New DlgControlText(AddressOf InvokeControlText)		'當 InvokeRequired 為 True 時，表示在不同的執行緒上，進行委派
		Control.Invoke(t, New Object() {Control, Text})
	Else																'當 InvokeRequired 為 False 時，表示在同一個執行緒上，
		Control.Text = Text                                             '所以可以正常的呼叫到這個 Control 物件
	End If
End Sub
```

[HOW TO：進行對 Windows Form 控制項的安全執行緒呼叫](https://msdn.microsoft.com/zh-tw/library/ms171728(v=vs.100).aspx?appId=Dev14IDEF1&l=ZH-TW&k=k(EHInvalidOperation.WinForms.IllegalCrossThreadCall)%3bk(TargetFrameworkMoniker-.NETFramework,Version%3dv4.5.2)%3bk(DevLang-VB)&rd=true&cs-save-lang=1&cs-lang=vb#code-snippet-1)

# 其他

## System.Math 類別

Abs() 、 Asin() 、 Acos() 、 Atan() 、 Atn() 、 Exp() 、 Log() 、 Round() 、 Pow() 、 Sqr()

[Math](https://msdn.microsoft.com/zh-tw/library/thc0a116.aspx)

## 亂數 (Random Number)

Rnd() 、 Randomize() 、 Int()

[亂數](http://blog.yljh.mlc.edu.tw/blog/blog_form.asp?blog=paguma&view=writings&writing_id=76)
[亂數](https://msdn.microsoft.com/zh-tw/library/f7s023d2(v=vs.90).aspx)
[System.Random](https://msdn.microsoft.com/zh-tw/library/system.random(v=vs.110).aspx)

## DataAndTime

DateAdd() 、 DateDiff()

[DateAndTime](https://msdn.microsoft.com/zh-tw/library/microsoft.visualbasic.dateandtime(v=vs.110).aspx)

## 檢查資料型別

IsArray() 、 IsDate() 、 IsNumeric() 、 IsNothing()

## IsArray

IsArray(<var>)

判斷 <var> 是不是一個陣列，是的話回傳 true ，反之回傳 false 。

## CurDir

FileSystem.CurDir() ，取得目前位置

## With ... End With

當我們要修改**同一個物件**下的變數時，可使用 With ... End With 來減少重複的程式碼

eg.
```vb
With txtReward
	.BackColor = Color.Darked
	.ForeColor = Color.White
	.ReadOnly = True
End With
```

## 非同步作業 (Asynchronous Programming Model)

非同步作業

使用 IAsyncResult 設計模式的非同步作業會實作為兩種方法，名為 **BeginOperationName** 和 **EndOperationName** ，分別開始和結束非同步作業 OperationName 。
呼叫 BeginOperationName 之後，應用程式可以繼續對進行呼叫的執行緒執行指令，而**非同步作業會在不同的執行緒上執行**。 每次呼叫 BeginOperationName 後，應用程式也應該**呼叫 EndOperationName 才能取得作業結果**。

[非同步作業](https://msdn.microsoft.com/zh-tw/library/ms228963(v=vs.110).aspx)

## Region

可以將某區塊的程式碼變成可折疊

eg.
```vb
#Region "<section-title>"
	<program>
#End Region
```

## Line Continuation

當一行程式碼太長，想要分成兩行時，在要換行的地方加上空白字元( )與底線(_)即可。

eg.
```
Private Sub helloButton_Click(sender As Object, e As _
EventArgs) Handles helloButton.Click
{
	<do-something>
}
```

## Comments

單行註解使用單引號(')，或是 REM ，目前沒有多行註解方式。

# 視窗程式設計

## Windows Form

Form 是我們的視窗程式主體，所有的控制項都會放在 Form 上面，如： Button 、 Label 、 TextBox ，在 Project 建立時， IDE 就自動幫我們建立了一個 Form 。
如果我們想再新增一個 Form ，可以在程式碼內輸入 `Dim newForm As New Form` ，若要讓 Form 顯示記得打上 `newForm.Show()` 。

Form 在設計時與實際執行時的視窗大小不同，可把 Form 屬性中的 AutoScaleMode 改成 none。

[Windows Form Properties and Methods and Events](https://www.tutorialspoint.com/vb.net/vb.net_forms.htm)
[不同電腦，FORM和字型的大小不一樣，AutoScaleMode有設定](http://www.blueshop.com.tw/board/FUM20050124191756KKC/BRD201509071652066OR.html)

### 當 Project 內有多個 Form 時

如果想確認第二個 Form 是否有開啟，有的話顯示已經開啟，沒有的話則開啟第二個 Form ，可以用以下的寫法。

```vb
If Application.OpenForms().OfType(Of Form2).Any Then
  MessageBox.Show("Form2 Opened")
Else
  Dim f2 As New Form2
  f2.Text = "Form2"
  f2.Show()
End If
```

### Sub Main()

Module <module-name>
	Sub Main()
		<program-region>
	End Sub
End Module

在一般的情況下，應用程式的入口會是 Form1 。
但我們可以透過雙擊方案總管內的 My Project ，取消勾選 *啟用應用程式架構* ，然後將起始物件改成 Sub Main ，這樣應用程式的入口就會變成 Sub Main() 。
此時我們可以在 Sub Main() 內寫登入系統，唯有登入成功，才會開始跑 Form1 ，登入失敗，程式就直接結束。

[Sub Main](https://www.ptt.cc/bbs/Visual_Basic/M.1112516567.A.C39.html)

### Application 類別

在 .NET Framework 類別庫底下， System.Windows.Forms 名稱空間下提供 Application 類別來管理應用程式及 Windows 的相關訊息。常用的方法有 Run() 與 Exit() 。

[Application.Run](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.application.run(v=vs.110).aspx)
[Application.Exit](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.application.exit(v=vs.110).aspx)

## 控制項

### InputBox

InputBox(<提示訊息>, <標題>, <預設回覆>, <視窗 x 座標>, <視窗 y 座標>)

eg. `InputBox("Please Enter Your Name: ")`
eg. `InputBox("Please Enter Your Name: ", "Test", "none", 50, 50)`

### MsgBox

eg. `MsgBox("num = " & num, vbOKOnly + vbInformation, "Window Title")`
[Here](http://yes.nctu.edu.tw/VB/6_Func/MsgBox.htm)

### Label

常用屬性：Text 、 AutoSize 、 Size 、 TextAlign 、 BorderStyle 、 BackgroudImage 、 BackColor 、 ForeColor 、 Font

### TextBox

常用屬性： MultiLine 、 CharacterCasing 、 MaxLength 、 PasswordChar 、 ReadOnly 、 ScrollBar 、 TextAlign 、 WordWrap
常用方法： Clear() 、 Copy() 、 Cut() 、 Paste() 、 Undo() 、 ClearUndo() 、 Focus()

### Button

常用屬性： Enabled
常用事件： Click()

### RadioButton

**單選**選項。

常用屬性： Text 、 Apperance 、 Checked 、 TextAlign 、 AutoCheck
常用事件： CheckedChanged() 、 Click()

[RadioButton](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.radiobutton(v=vs.110).aspx)

### CheckBox

可**多選**選項。

常用屬性： Text 、 Checked 、 ThreeState 、 ChecState

[CheckBox](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.checkbox(v=vs.110).aspx)

### DateTimePicker

eg. `Dim d As Date = DateTimePicker.Value.Date`

### Timer

可用來執行週期性的動作。

當 Timer 啟動 (Timer.Enabled = True) 後，每隔一段時間 (Timer.Interval) 會執行 Timer_Tick() 內的程式碼，亦可將 Timer 停下 (Timer.Enabled = False) 。

[Timer](https://msdn.microsoft.com/zh-tw/library/system.timers.timer(v=vs.110).aspx)

### RitchTextBox

配合檔案運作，建立 RTF(Rich Text File) 格式檔案

[RitchTextBox](http://vbplaying.blogspot.tw/2013/04/vbrichtextbox.html)
[RitchTextBox類別](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.richtextbox(v=vs.110).aspx)

### MaskedTextBox

具有遮罩功能，使用者依據格式輸入字元。

常用屬性： BeepOrError 、 MaskFull 、 ValidatingType
常用方法： Clear()
常用事件： MaskInputRejected() 、 TypeValidationCompleted()

[MaskedTextBox](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.maskedtextbox(v=vs.110).aspx)
[MaskedTextBox.Mask](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.maskedtextbox.mask(v=vs.110).aspx)

### ToolTip

ToolTip 提供**工具提示**。一般我們在用軟體時，把滑鼠移到工具列的某一個圖示按鈕，工具提示會顯示其作用或用途，這就是 ToolTip 。

常用屬性： AutomaticDelay 、 InitialDelay 、 ShowAlways 、 ToolTipTitle
常用方法： SetToolTip(<顯示提示的控制項>, <提示訊息>) 、 Show(<提示訊息>, <顯示提示的控制項>, x相對座標, y相對座標, <持續時間>)

[ToolTip](http://www.visual-basic-tutorials.com/Tutorials/Controls/ToolTip.html)
[ToolTip](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.tooltip(v=vs.110).aspx)

### Help 類別

Help.ShowPopup(<parent>, <caption>, <location>)

<parent> : 指定欲顯示說明交談窗的**控制項**
<caption> : 顯示於說明交談窗的訊息
<location> : 顯示交談窗的位置

[Help Class](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.help(v=vs.110).aspx)

### 日期時間控制項

#### MonthCalendar

MonthCalendar 提供一個視覺化介面來選取日期。

常用屬性： AnnuallyBoldedDates 、 BoldedDates 、 MonthlyBoldeDates 、 TitleBackColor 、 TitleForeColor 、 TrailingForeColor 、 CalendarDimensions
常用事件： DateSelected() 、 DateChange()

#### DateTimePicker

若要顯示特定日期和時間，可使用 DateTimePicker 。其提供一個下拉式清單來讓使用者選擇日期。

常用屬性： ShowUpDown 、 ShowCheckBox 、 Format

[DateTimePicker.CustomFormat](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.datetimepicker.customformat(v=vs.110).aspx)
[如何：使用 DateTimePicker 控制項顯示時間](https://msdn.microsoft.com/zh-tw/library/ms229631(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-3)

### 版面控制

#### GroupBox

選項**容器**。

有時我們會在 GroupBox 內放很多控制項，有可能都是相同種類的，若我們想要一次性處理裡面的控制項，可以使用 For Each 與 GroupBox.Controls 來達成， GroupBox.Controls 是一個集合，裡面存的是 GroupBox 內的所有控制項。
如果裡面放的是不同類型的控制項，我們可以在 For Each 內先使用 Control 型態的變數，在使用 TypeOf 去確認每個 Control 的型態，在對其做處理。

[一次處理容器中所有的控制項](https://dotblogs.com.tw/piercejhuang/2012/01/07/64818)

[GroupBox](https://msdn.microsoft.com/zh-tw/library/system.windows.forms.groupbox(v=vs.110).aspx)

#### FlowLayoutPanel

以水平或垂直**流向**來排列控制項。

常用屬性： WrapContents 、 FlowDirection 、 FlowBreak

#### TableLayoutPanel

以**格線**排列，版面大小由欄、列做決定。

一個格子只能放一個控制項，若要調整其的對齊方式，可由控制項的 Anchor 與 Dock 屬性調整。

[HOW TO：在 TableLayoutPanel 控制項中對齊和縮放控制項](https://msdn.microsoft.com/zh-tw/library/ms171688(v=vs.100).aspx)

#### TabControl

以**索引標籤頁**來製作多個頁面管理控制項。

要修改 Tab 的名稱，要到 TabControl 的屬性中的 TabPages 旁的小按鈕。

常用屬性： Alignment 、 Appearance
常用方法： Add() 、 Insert() 、 Remove() 、 Hide() 、 Show() 
常用事件： SelectedIndexChanged() 、 Selected() 、 DrawItem()

### 清單

#### ListBOx

ListBox 會將項目以清單顯示出來，提供使用者從中選取一個或多個項目。

常用屬性： SelectionMode 、 MultiColumn 、 Sorted 、 **Items** 、 Items.Conut 、 SelectedIndex 、 SelectedItem 、 SelectedItmes 、 Text
常用方法： SetSelected() 、 GetSelected() 、 ClearSelected()

要設定 ListBox 內的 Item ，從屬性的 Items 內去設定。
設定選取某個項目， `ListBox.Items.SetSelected(<index>, True)` 。
新增項目， `ListBox.Items.Add(<新項目>)` 或是 `ListBox.Items.AddRange(<新項目陣列>)` 。
移除項目， `ListBox.Items.Remove(<項目>)` 或是 `ListBox.Items.RemoveAt(<index>)` 。
清除所有項目， `ListBox.Items.Clear()` 。
查看是否已有此項目， `ListBox.Items.Contains(<項目>)` 。

#### ComboBox

ComboBox 提供下拉式清單，當清單內沒有想要的項目，還可以自行輸入。
ComboBox 主要分成兩個區塊，一個是上方的可供使用者輸入的 *文字區塊* ，另一個是下方顯示項目的 *清單方塊* 。

常用屬性： SelectedIndex 、 SelectedItem 、 DropDownStyle 、 Text 、 DropDownWidth 、 MaxLength 、 MaxDropDownItems
常用事件： SelectedIndexChanged()

新增項目， `ComboxBox.Items.Add(<新項目>)` 或是 `ComboxBox.Items.AddRange(<新項目陣列>)` 或是 `ComboxBox.Items.Insert(<index>, <新項目>)` 。
移除項目， `ComboxBox.Items.Remove(<項目>)` 或是 `ComboxBox.Items.RemoveAt(<index>)` 。
清除所有項目， `ComboxBox.Items.Clear()` 。

#### CheckedListBox

CheckedListBox 可視為 CheckBox 與 ListBox 的組合。

常用屬性： CheckState 、 CheckOnClick 、 Items(<index>)
常用方法： SetItemChecked() 、 GetItemChecked()
常用事件： SelectedIndexChanged()

### 提供檢視的控制項

#### ImageList

提供多種圖片格式來存放多張圖檔。

常用屬性： Images 、 ColorDepth 、 ImageSize 、 ImageStream 、 TransparentColor
常用方法： Draw()

從檔案名稱讀取圖片進來， `ImageList.Images.Add(Image.FromFile(<file-name>))` 。
從 ImageList 讀取圖片， `PictureBox.Image = ImageList.Images.Item(<index>)` 。
從 ImageList 移除圖片， `ImageList.Images.RemoveAt(<index>)` 。

[ImageList](http://blog.xuite.net/alwaysfuturevision/liminzhang/10441373-Visual+Basic+2005+-+%E8%AE%80%E8%80%85%E8%A9%A2%E5%95%8F+ImageList+%E5%95%8F%E9%A1%8C)

#### ListView

提供四種檢視項目，包含**清單**、**圖示**、**縮圖**和**詳細資料**。

常用屬性： View 、 

#### TreeView

透過節點顯示階層式資料，節點含有選擇性核取方塊或圖示組成。

### OpenFileDialog

`OpenFileDialog1.ShowDialog()` ，打開讀取檔案的視窗，如果有選取檔案會回傳 `DialogResult.OK` 。
開啟可選取多個檔案， `OpenFileDialog.Multiselect = True`
檔案副檔名過濾器， `OpenFileDialog.Filter = "Image Files(*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif|All Files (*.*)|*.*"`
讀取到的檔案數目， `OpenFileDialog.FileNames.Length`
讀取到的檔案名稱， `OpenFileDialog.FileName` 或是 `OpenFileDialog.FileNames(<index>)`

### MenuStrip

MenuStrip -> 插入標準項目 -> 插入檔案、編輯、工具、說明
直接輸入項目名稱，如 `字型(&F)` ， &F 代表加入對應按鍵 `F`。
輸入 `-` 加入分隔線。
選取項目後按下 `Delete` 來刪除該項目。

[MenuStrip](https://home.gamer.com.tw/creationDetail.php?sn=2104038)

### DataTable

DataTable 通常用來儲存從資料庫的資料。
但有時候我們也會想把陣列的資料變成 DataTable 來使用，或者將從資料庫中的資料取出後，再加入幾筆資料，這時就需要 DIY DataTable 了。

完成後可以使用 DataGridview 來呈現表格。

### DataGridView



### Font & ForeColor

Font(<字體>, <字體大小>, <字體style>)
eg. `Label1.Font = New Font("新細明體", 9, FontStyle.Bold)`
eg. `Label1.Color = Color.Red`
eg. `Label2.Color = Color.FromArgb(255, 0, 0)`

[Here](https://home.gamer.com.tw/creationDetail.php?sn=1823197)

### BeginUpdate

控制項每一次更新的時候，都要進行重新繪製，如果我們不得不在一個迴圈內不斷更新控制項時，為了提高效率，避免他每次更新都重繪，就可以在迴圈外前後加上 控制項.BeginUpdate() 與 控制項.EndUpdate() ，讓其在這期間更新完不會進行重繪，而是直到呼叫 EndUpdate() 才會重繪。

[使用控制項的BeginUpdate功能](https://dotblogs.com.tw/jeff-yeh/2008/11/17/6014)

### 設定按 Tab 時的跳的順序

功能表 -> 檢視 -> 定位順序
也可從每一個物件的屬性欄位去看 TabIndex 、 TabStop

## 控制項的事件

### Event Handler 的參數列

eg.
```vb
Private Sub helloButton_Click(sender As Object, e As EventArgs) Handles helloButton.Click

End Sub
```

sender 提供引發事件的物件，以上面範例來說，其就是指Button物件。
e 提供事件的資訊。不過不同的事件使用不同的物件來傳遞，如滑鼠事件是使用 MouseEventArgs ，鍵盤事件是使用 KeyEventArgs 。
Handles helloButton.Click 代表這個 Sub 是處理 helloButton 物件的 Click 事件。

### 共用事件處理程序

eg.
```vb
Private Sub confirmButton_Click(sender As Object, e As EventArgs) Handles confirmButton.Click, resetButton.Click
	Dim btn As Button = CType(sender, Button)

	If btn.Name = "confirmButton" Then
		<dosomething1>
	ElseIf btn.Name = "resetButton" Then
		<dosomething2>
	End If
End Sub
```

confirmButton_Click 這個 function 同時處理 confirmButton.Click 與 resetButton.Click 兩個事件，
把 sender 這個 Object 使用 CType() 轉換成 Button 後，再去看其 Button.Name 屬性來判斷是按下哪個 Button。

### 新增/移除事件處理程序

AddHandler | RemoveHandler <event>, AddressOf <event-handler>

eg.
```vb
AddHandler errorButton.MouseHover, AddressOf showNO

Public Sub showNO(ByVal sender As Object, ByVal e As EventArgs)
	MsgBox("nonono")
End Sub

RemoveHandler errorButton.MouseHover, AddressOf showNO
```

使用 AddHandler 來新增事件處理。新增 errorButton 這個物件的 MouseHover 事件，然後處理事件的 Function 為 showNO 。
這樣之後把滑鼠移到 errorButton 上方時，就會跳出訊息說 nonono 。
不需要該事件處理時再使用 RemoveHandler 來停止。

### 滑鼠事件

滑鼠的點擊事件
MouseClick() 、 MouseDoubleClick()

滑鼠的移動事件
MouseDown() 、 MouseUp() 、 MouseMove()

透過 e 取得滑鼠的訊息

滑鼠的拖曳事件
DragEnter() 、 DragOver() 、 DragDrop() 、 DragLeave()

[Mouse Event](https://msdn.microsoft.com/zh-tw/library/ms171542(v=vs.110).aspx)
[Mouse Drag Event](https://msdn.microsoft.com/zh-tw/library/ms171546(v=vs.110).aspx)

### 鍵盤事件

VB為了執行效率，其**預設為不會引發鍵盤事件**，必須去表單的 KeyPreview 更改為 True ，才能開啟觸發鍵盤事件。

鍵盤事件： KeyDown() 、 KeyPress() 、 KeyUp()

KeyPress() 可用來偵測組合按鍵。

在 KeyPress() 事件中， e.Handled 原本是用來取得 KeyPress() 是否已經處理完事件的一個 flag ， True 代表已處理， False 代表未處理，但我們可以令 e.Handled = True 來告知系統其事件已經處理，而實際上卻是未處理。

[KeyEvent](http://tsuozoe.pixnet.net/blog/post/19733703-vb.net-%E9%8D%B5%E7%9B%A4%E4%BA%8B%E4%BB%B6%E4%BB%8B%E7%B4%B9-(keypress%E3%80%81keydown-%E5%92%8C-keyup-%E4%BA%8B))
[KeyCode](https://msdn.microsoft.com/en-us/library/aa243025(v=vs.60).aspx)

### Custom Event

1. declare an Event in a class
```vb
Public Class Class1
	Public Event MyEvent()
End Class
```

2. RaiseEvent in class method
```vb
Public Sub RaiseMyEvent
	RaiseEvent MyEvent()
End Sub
```

3. declare a object with event
```vb
Public WithEvents obj As New Class1()
```

4. handle the event
```vb
Private Sub HandleMyEvent() Handles obj.MyEvevnt
	<do-something>
End SUb
```

----
