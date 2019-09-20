# emailgo
通过golang编写的，工资条批量邮件分发工具

工资条通过golang的html/template渲染后发送，由于前端能力太弱，所以大部分功能处理交由后端处理，前端网页只做页面风格处理

目前只支持xlsx文件，并且表格支持1-2行表头，邮箱必须为最后一列

工资条发送展示:

1.单表头的：
![image](https://github.com/tanganyu1114/emailgo/blob/master/00001.jpg)

2.多表头的：
![img](https://github.com/tanganyu1114/emailgo/blob/master/000002.jpg)
