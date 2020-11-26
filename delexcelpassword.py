# -*- coding:utf-8 -*-
"""
    解密被编辑保护的Excel文件(xls, xlsx)
    原理：
        1. 将Excel文件重命名为zip/rar压缩文件，并解压
        2. 通过正则匹配，替换掉 xl/worksheets 目录下所有.xml文件中密码保护的配置信息<sheetProtection[^<>]*>
        3. 重新打包文件为Excel文件类型
    @Created on 2020/11/26
    @Author: lgg
"""

import io
import os
import re
import shutil
import zipfile


class DelExcelPassword:
    temp_path = "./dist/"
    support_file_type = ['xls', 'xlsx']
    temp_extract_zip_path = ''
    source_path = ''
    source_full_name = ''
    source_name = ''
    source_file_type = ''
    zip_suffix = 'zip'
    result_start_flag = '[已解密]'

    def __init__(self, source_path):
        self.source_path = source_path
        self.source_full_name = os.path.basename(source_path)
        name, suffix = os.path.splitext(self.source_full_name)
        self.source_name = name
        self.source_file_type = suffix[1:]
        self.mk_temp_dir()

    def process(self):
        if not self.is_support_file_type():
            # print(self.source_full_name, '不是支持的文件类型！')
            return

        if self.source_name.startswith(self.result_start_flag):
            # print(self.source_full_name, '已处理过的文件，忽略！')
            return

        print('开始处理:', self.source_full_name, end='\t')
        zip_file_path = self.copy2zip()
        self.extract_zip(zip_file_path)
        os.remove(zip_file_path)

        self.remove_sheet_protection()

        self.zip_result_files()
        print('处理完成！')

    def mk_temp_dir(self):
        if not os.path.exists(self.temp_path):
            os.mkdir(self.temp_path)

    def is_support_file_type(self):
        return self.source_file_type in self.support_file_type

    def copy2zip(self):
        """
            拷贝并重命名为zip文件，存放到 temp_path 目录中
        :return: 返回文件路径
        """
        zip_filename = '%s.%s' % (self.source_name, self.zip_suffix)
        if os.path.exists(zip_filename):
            os.remove(zip_filename)
        dist_path = os.path.join(self.temp_path, zip_filename)
        shutil.copy(self.source_path, dist_path)
        return dist_path

    def extract_zip(self, zip_filename):
        """
            将压缩文件解压到 temp_path 下以源文件名命名的文件夹
        :param zip_filename: 压缩文件
        :return: None
        """
        self.temp_extract_zip_path = os.path.join(self.temp_path, self.source_name)
        with zipfile.ZipFile(zip_filename) as zip_file:
            file_list = zip_file.namelist()
            for file in file_list:
                # 解压到以源文件名命名的文件夹
                zip_file.extract(file, path=self.temp_extract_zip_path)

    def remove_sheet_protection(self):
        """
            通过正则替换删除 <sheetProtection /> 密码保护信息
        :return:
        """
        for root, directory, files in os.walk(self.temp_extract_zip_path):
            for filename in files:
                name, suf = os.path.splitext(filename)
                if suf == '.xml':
                    xml_file = os.path.join(root, filename)
                    # 删除密码保护信息
                    with io.open(xml_file, "r", encoding="utf-8") as f1, \
                            io.open("%s.bak" % xml_file, "w", encoding="utf-8") as f2:
                        for line in f1:
                            f2.write(re.sub('<sheetProtection[^<>]*>', '', line))
                    os.remove(xml_file)
                    os.rename("%s.bak" % xml_file, xml_file)

    def zip_result_files(self):
        """
            压缩已源文件命名的已处理完的文件夹，并删除文件夹
        :return: None
        """
        result_file = os.path.join(self.temp_path, self.result_start_flag + self.source_full_name)
        with zipfile.ZipFile(result_file, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for dirpath, dirNames, fileNames in os.walk(self.temp_extract_zip_path):
                fpath = dirpath.replace(self.temp_extract_zip_path, '')
                fpath = fpath and fpath + os.sep or ''
                for name in fileNames:
                    zip_file.write(os.path.join(dirpath, name), os.path.join(fpath, name))

        shutil.rmtree(self.temp_extract_zip_path)


def main():
    """
        遍历处理当前目录下所有Excel文件
    :return: None
    """
    base_path = './'
    xls_files = []
    for root, directory, files in os.walk(base_path):
        for file in files:
            xls_files.append(file)
    for xls_file in xls_files:
        DelExcelPassword(xls_file).process()


if __name__ == '__main__':
    main()
