# DelExcelPassword

解密被编辑保护的Excel文件(xls, xlsx)

原理：
1. 将Excel文件重命名为zip/rar压缩文件，并解压
2. 通过正则匹配，替换掉 xl/worksheets 目录下所有.xml文件中密码保护的配置信息<sheetProtection[^<>]*>
3. 重新打包文件为Excel文件类型
