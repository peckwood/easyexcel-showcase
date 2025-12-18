### 前端项目

是easyexcel-showcase-web

# WriteDemo01(不创建对象的写)

不创建对象的写

# WriteDemo02(不创建对象的写 - 综合性)

写的方式

- 不创建对象的写

### 样式

- 字体
- 水平对齐
- 设置所有列的宽度
- 动态设置样式(设置数据后的第一行的样式)

### 其它

- CellWriteHandler
- 获取cell内容(获取不到head内容)
- 占据多列的大标题
- 合并了2个单元格的'合计'
- 支持不同列数的sheet的整行合并
- 插入数据后插入一条合并多列的行
- 

# WebDemo01Controller (不创建对象的写)

web下的写(不创建对象)

# WebDemo02Controller (填充多个sheet)

通过 http://localhost:8080/web-demo02 访问

关键点:

- easyexcel的row, column, sheet都是从0开始

- 指定resources下文件夹的方式

  ```
  String templateFileName = WebDemo02Controller.class.getResource("/").getPath() + "excel-template/WebDemo02ControllerTemplate.xlsx";
          ExcelWriter excelWriter = EasyExcel
                  .write(outputStream)
                  .withTemplate(templateFileName)
                  ...
  ```

- 一般找的resources文件夹一般是项目启动类(@SpringBootApplication)所在的模块下的resources文件夹, 在maven多模块项目下要注意

展示内容

- sheet1展示如何填充excel
- sheet2 展示即使固定的列头, 也可以使用动态填充的方式, 而且不会创建新行(forceNewRow不设置为true)
- sheet3 展示通过重写CellWriteHandler接口afterCellDispose方法方式动态合并列头
- 也展示了异常发生时, 如何返回JSON

前端展示了excel下载和返回JSON这两种情况的处理

