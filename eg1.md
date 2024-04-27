```js
function copyDataFromTest2ToTest1() {

    
    // 打开 TEST1 和 test2 文件
    var test1Book = Workbooks.Open("C:\\Users\\Administrator\\Desktop\\test1.xlsx");
    var test2Book = Workbooks.Open("C:\\Users\\Administrator\\Desktop\\test2.xlsx");
    
    // 获取 TEST1 和 test2 的工作表
    var test1Sheet = test1Book.Sheets("Sheet1"); // 假设 TEST1 的数据在 Sheet1
    var test2Sheet = test2Book.Sheets("Sheet1"); // 假设 test2 的数据在 Sheet1
    
    // 获取 TEST1 和 test2 的数据范围
    var test1Range = test1Sheet.Range("A1").CurrentRegion; // 假设数据从 A1 开始，并连续
    var test2Range = test2Sheet.Range("A1").CurrentRegion; // 假设数据从 A1 开始，并连续
    
    // 遍历 TEST1 的数据，查找匹配的学号，并复制姓名和性别
    var test1Data = test1Range.Value2;
    var test2Data = test2Range.Value2;
    
    for (var i = 1; i <= test1Data.length-1; i++) {
        var test1StudentId = test1Data[i][0];
		Console.log("test2Data.length="+test2Data.length);
        
        for (var j = 1; j <= test2Data.length-1; j++) {
        	console.log("i="+i+"    j="+j);
            var test2StudentId = test2Data[j][0];
            
            if (test1StudentId === test2StudentId) {
                // 找到匹配的学号，复制姓名和性别到 TEST1 的 D 和 E 列
                var name = test2Data[j][1];
                var gender = test2Data[j][2];
                
                // 假设 TEST1 的 D 和 E 列是空的，直接赋值
                test1Sheet.Cells(i + 1, 4).Value2 = name; // D 列
                test1Sheet.Cells(i + 1, 5).Value2 = gender; // E 列
                
                break; // 找到匹配项后，跳出内层循环
            }
        }
    }
    test1Sheet.Range("A1:E6").AutoFilter(5, Array("女"), xlFilterValues, undefined, undefined);
    // 保存 TEST1 文件
    test1Book.Save();
    
    // 关闭 TEST1 和 test2 文件
    test1Book.Close();
    test2Book.Close();
}
```

test1

| 学号 | 科目 | 成绩 | 姓名 | 性别 |
| ---- | ---- | ---- | ---- | ---- |
| 1    | 数学 | 23   |      |      |
| 2    | 语文 | 23   |      |      |
| 3    | 英语 | 43   |      |      |
| 4    | 历史 | 54   |      |      |
| 5    | 政治 | 45   |      |      |

test2

| 学号 | 姓名 | 性别 |
| ---- | ---- | ---- |
| 5    | 张三 | 男   |
| 4    | 网红 | 女   |
| 33   | 粘贴 | 男   |
| 2    | 燕子 | 女   |
| 1    | 废物 | 男   |

out

| 学号 | 科目 | 成绩 | 姓名 | 性别 |
| ---- | ---- | ---- | ---- | ---- |
| 2    | 语文 | 23   | 燕子 | 女   |
| 4    | 历史 | 54   | 网红 | 女   |