POI读取Excel有两种模式，一种是用户模式，一种是SAX事件驱动模式，将xlsx格式的文档转换成CSV格式后进行读取。<br/>

用户模式API接口丰富，使用POI的API可以很容易读取Excel，但用户模式消耗的内存很大，当遇到很大sheet、大数据网格，假空行、公式等问题时，很容易导致内存溢出。<br/>

POI官方推荐解决内存溢出的方式使用CVS格式解析，即SAX事件驱动模式。<br/>

下面主要是讲解如何读取大批量数据：<br/>

POI以SAX解析excel2007文件：<br/>
解决思路：通过继承DefaultHandler类，重写process()，startElement()，characters()，endElement()这四个方法。
process()方式主要是遍历所有的sheet，并依次调用startElement()、characters()方法、endElement()这三个方法。startElement()用于设定单元格的数字类型（如日期、数字、字符串等等）。
characters()用于获取该单元格对应的索引值或是内容值（如果单元格类型是字符串、INLINESTR、数字、日期则获取的是索引值；其他如布尔值、错误、公式则获取的是内容值）。
endElement()根据startElement()的单元格数字类型和characters()的索引值或内容值，最终得出单元格的内容值，并打印出来。<br/>

POI通过继承HSSFListener类来解决Excel2003文件：<br/>
解决思路：重写process()，processRecord()两个方法，其中processRecord是核心方法，用于处理sheetName和各种单元格数字类型。<br/>

参考资料：

https://www.cnblogs.com/huangjian2/p/6238237.html

https://www.cnblogs.com/yfrs/p/5689347.html

http://blog.csdn.net/lishengbo/article/details/40711769

https://www.cnblogs.com/wshsdlau/p/5643847.html

http://blog.csdn.net/lipinganq/article/details/78775195

http://blog.csdn.net/lipinganq/article/details/53389501

http://blog.csdn.net/zmx729618/article/details/72639037

http://blog.csdn.net/daiyutage/article/details/53010491
