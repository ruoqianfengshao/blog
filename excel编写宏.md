# excel vba 编写宏

因为 hr 工作需求，帮着写了一个处理员工财务数据的宏。第一次接触宏的编写，用到了以下的知识点，记录一下。

1. 开启 vba，偏好设置 -> 通用 -> 工具栏 -> 主tab 下勾选开发者
2. 开发者 -> 打开vb 编辑器 -> 应用导航 -> 插入 -> 模块
3. 选中模块，开始开发
4. 如何比较两个单元格的日期大小：DateDiff("D", DateValue(Cells(row1, column1)), DateValue(row2, column2)) > 0
5. xlUP xlDown 沿文档方向向上，向下，可以用来取 row 的最小最大值
6. IsNumeric(value) 判断是否数字, IsDate(value) 判断是否日期
