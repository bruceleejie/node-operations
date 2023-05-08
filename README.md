# node-operations

工作中使用nodejs开发的服务的一些总结

> 2023.05.08更新于北京

## 一、借鉴说明：
本demo借鉴了很多网上的例子：
例子一：https://juejin.cn/post/7044020172109119496 （基本的使用）

例子二：https://juejin.cn/post/7205843311340470329#heading-10 （合并行的操作和配置参考）

## 二、第一个demo利用nodejs将json数据转成表格数据并生成word

首先：利用到了officegen插件；
其次：根据不同的类别生成table到word中；

实现逻辑：先组成表头的行，再组成表格数据；把表头放到大数组的第0项，表格数据一次插入到大数组中；然后利用docx.createTable()生成一个表格。

demo中有两种生成方式：
一种是把同一个大的父级children里的所有子项扁平化成一个数组里，根据父级的key来确认是否合并行；
另一种是把同一个大父级的children里的每一个小的children合并判断相同项合并；

具体使用可以参考自己的需要进行选择；

### 1、文件介绍：
注意：file文件夹里是我准备的可以直接使用的json文件，虽然都脱敏了，但不可商用，如果出现纠纷，demo发布者本人不负责任，商用者负责！！！！！

> generateDoc.ts是我的初始demo；
> 
> generateDocArrFunc.ts是我领导的思路，也就是两种方法中的第二个，代码量少，适合有父子层级的json数据；
> 
> generateDocFlatFunc.ts是我的思路，也就是两种方法中的第一个，代码量多一些，适合每个子元素有父级唯一标识的json数据；



