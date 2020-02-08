# README

这是一个用于NOIP等考试的考生代码回收脚本
通过将考场代码整合到U盘，运行程序判断是否考生文件都符合要求。
并将各个考场考生文件汇总。
具体流程参见[《文件回收流程》](.//文件回收流程.md)



第一次使用github，多多指教



##　[接口.xlsx](.//接口.xlsx):

| 接口           | 接口名字    |      |      |
| -------------- | ----------- | ---- | ---- |
| 考生文件夹路径 | J1          |      |      |
| 目标文件夹路径 | all         |      |      |
| 考场名称       | all         |      |      |
| 考生名单名字   | name.xlsx   |      |      |
| 跳过第一行     | 考场        |      |      |
| 题目名字       | number      | work |      |
| 认可文件后缀   | cpp         | pas  | c    |
| 保存的位置     | sample.xlsx |      |      |





## 注意事项

1. 考生名单：默认文件名为"[name.xlsx](name.xlsx)"，第一列为考生ID，第二列为考生考场。若名单第一行需要被跳过，则设置跳过接口“skipped”，default为“考场”。

2. 接口参数：接口文件名为："接口.xlsx"，为空则为手动输入

   考场名词可以为："J1","Y1","all"等等

   保存的位置默认格式为.xlsx格式

   题目名字和认可后缀可以往后自由扩展

   目标文件夹为"0"则为判断该文件夹种文件是否合法

   若接口为空则按顺序手动输入

3.  注意查看考生程序最后一次修改时间是否变化。

