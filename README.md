# createSQLfile
读取Excel文件中的内容生成SQL文件

如下图所示
![image](https://github.com/sunny0826/createSQLfile/blob/master/example/exampleimage.jpg)

生成update语句
在Excel文件中，第一行行需要update的表名，第二行为字段名，每一列为该字段的对应值，这里可以进行where过滤，只需要修改main函数下where数组内的数字即可，需要update的字段同理，执行后就会生成名字为update#表名.sql的SQL文件。

生成insert语句
生成insert语句SQL文件的Excel格式与update的相同，但是传入参数方面，因为不需要过滤条件，所以只需要往数组中写入需要插入字段的列数就好。