# csv2xlsx
本程序最初是为了方便将实验数据(csv)转换成xlsx文件
所写的，现在提供给有着同样需求的人。

##使用环境
Windows10  
Python3.6(直接通过py文件执行的情况下，需要安装xlsxwriter包)

Linux等其他环境下未测试



##使用方法
 * 将需要转换的文件放入根目录下doc文件夹内
 * 将dist目录下csv2xlsx.exe放到根目录下
 * 运行csv2xlsx.exe  

程序会在根目录下生成一个名为output的文件夹，并把doc文件夹内所有的扩展名为csv文件
转换并同名保存至output内。

##注意点
 * 程序默认把所有扩展名为csv的文件转换成xlsx文件，
 后续可以增加文件选择功能，只转换需要的文件

##License
MIT 

