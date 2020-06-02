# excel vba 编写宏

因为 hr 工作需求，帮着写了一个处理员工财务数据的宏。第一次接触宏的编写，用到了以下的知识点，记录一下。

1. 开启 vba，偏好设置 -> 通用 -> 工具栏 -> 主tab 下勾选开发者
2. 开发者 -> 打开vb 编辑器 -> 应用导航 -> 插入 -> 模块
3. 选中模块，开始开发
4. 如何比较两个单元格的日期大小：DateDiff("D", DateValue(Cells(row1, column1)), DateValue(row2, column2)) > 0
5. xlUP xlDown 沿文档方向向上，向下，可以用来取 row 的最小最大值
6. IsNumeric(value) 判断是否数字, IsDate(value) 判断是否日期
7. If 多个条件组合 使用 And 或者 Or，不建议使用 & |，判断有问题，没深究原因
8. 保存时需要选择文件为开启宏的文件，并制定后缀 .xlsm


```vba
Sub calcQQ()
    Dim xRow As Long
    Dim xColumn As Long
    Dim activeTotal As Long
    Dim unactiveTotal As Long
    Dim offTotal As Long

    For xRow = Application.Cells(Rows.Count, 1).End(xlUp).Row To 3 Step -1
        activeTotal = 0
        unactiveTotal = 0
        offTotal = 0

        For xColumn = Application.Cells(xRow, Columns.Count).End(xlToLeft).Column To 19 Step -1
            Dim dateCell
            Dim countCell
            dateCell = Application.Cells(xRow, xColumn)
            countCell = Application.Cells(xRow, xColumn + 1)

            If IsDate(dateCell) Then
              If DateDiff("D", DateValue(dateCell), DateValue(Application.Cells(xRow, 4))) >= 0 & IsNumeric(countCell) Then
                activeTotal = activeTotal + countCell
              End If
              If DateDiff("D", DateValue(dateCell), DateValue(Application.Cells(xRow, 4))) < 0 & IsNumeric(countCell) Then
                unactiveTotal = unactiveTotal + countCell
              End If
              If DateDiff("D", DateValue(dateCell), DateValue(Application.Cells(xRow, 7))) >= 0 & IsNumeric(countCell) Then
                offTotal = offTotal + countCell
              End If
            End If
        Next xColumn

        Application.Cells(xRow, 5).Value = activeTotal
        Application.Cells(xRow, 6).Value = unactiveTotal
        Application.Cells(xRow, 8).Value = offTotal

    Next xRow

End Sub
```
