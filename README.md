# easyexcel-plus-spring-boot-starter

## 介绍

为了简化easyexcel的使用，进行一些扩展，只需要通过一个注解做到导出excel

## 设想

基于easyexcel的基础上扩展，

第一版：想通过在controller方法上添加一个注解就叫responseexcel把，表明excel的名称，sheet名称等信息，可以直接将方法返回的list响应excel

第二版：简化一些属性字典的映射，比如性别的字段，数据里存储0，1，excel直接导出男女，通过实体类字段上增加注解实现

第三版：简化合并的开发，通过在实体类字段上标注注解，增加responseexcel的属性标注excel需不需要合并内容

第四版：easyexcel没有图表的功能，在想要不要增加这样的功能，考虑下

## 使用说明
等待补充


## 参与贡献

1.  Fork 本仓库
2.  新建 Feat_xxx 分支
3.  提交代码
4.  新建 Pull Request
