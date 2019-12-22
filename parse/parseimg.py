import json
import os, shutil
import pdb
import sys
import zipfile
import xml.etree.cElementTree as ET


# 判断是否是文件和判断文件是否存在
def isfile_exist(file_path):
    if not os.path.isfile(file_path):
        print("It's not a file or no such file exist ! %s" % file_path)
        return False
    else:
        return True


# 修改指定目录下的文件类型名，将excel后缀名修改为.zip
def change_file_name(file_path, new_type='.zip'):
    #
    if not isfile_exist(file_path):
        return ''

    extend = os.path.splitext(file_path)[1]  # 获取文件拓展名
    if extend != '.xlsx' and extend != '.xls':
        print("It's not a excel file! %s" % file_path)
        return False

    file_name = os.path.basename(file_path)  # 获取文件名
    new_name = str(file_name.split('.')[0]) + new_type  # 新的文件名，命名为：xxx.zip

    dir_path = os.path.dirname(file_path)  # 获取文件所在目录
    new_path = os.path.join(dir_path, new_name)  # 新的文件路径
    if os.path.exists(new_path):
        os.remove(new_path)

    os.rename(file_path, new_path)  # 保存新文件，旧文件会替换掉

    return new_path  # 返回新的文件路径，压缩包


# 解压文件
def unzip_file(zipfile_path):
    if not isfile_exist(zipfile_path):
        return False

    if os.path.splitext(zipfile_path)[1] != '.zip':
        print("It's not a zip file! %s" % zipfile_path)
        return False

    file_zip = zipfile.ZipFile(zipfile_path, 'r')
    file_name = os.path.basename(zipfile_path)  # 获取文件名
    # 获取文件所在目录
    zipdir = os.path.join(os.path.dirname(zipfile_path), str(file_name.split('.')[0]))
    for files in file_zip.namelist():
        file_zip.extract(files, os.path.join(zipfile_path, zipdir))  # 解压到指定文件目录

    file_zip.close()
    return True

# 读取解压后的文件夹，输出xml路径
def read_img(zipfile_path, img_path=None):
    if not isfile_exist(zipfile_path):
        return False
    dir_path = os.path.dirname(zipfile_path)  # 获取文件所在目录
    file_name = os.path.basename(zipfile_path)  # 获取文件名
    unzip_dir = os.path.join(dir_path, str(file_name.split('.')[0]))
    # excel变成压缩包解压后，excel中的图片在media目录
    if img_path:
        pic_dir = 'xl' + os.sep + 'media'
        pic_path = os.path.join(dir_path, str(file_name.split('.')[0]), pic_dir)
        file_list = os.listdir(pic_path)
        for file in file_list:
            filepath = os.path.join(pic_path, file)
            # print(filepath,img_path)
            # 复制文件
            if os.path.isfile(os.path.join(img_path, os.path.basename(filepath))):
                os.remove(os.path.join(img_path, os.path.basename(filepath)))
            shutil.move(filepath, img_path)
    drawings_path = os.path.join(unzip_dir, 'xl', 'drawings', 'drawing1.xml')
    drawings_rels_path = os.path.join(unzip_dir, 'xl', 'drawings', '_rels', 'drawing1.xml.rels')
    image_list = img_info(drawings_path, drawings_rels_path)
    shutil.rmtree(unzip_dir)
    return image_list


# 还原文件名
def revert_dir(zipfile_path):
    extend = os.path.splitext(zipfile_path)[1]  # 获取文件拓展名
    if extend != '.zip':
        print("It's not a zip file! %s" % zipfile_path)
        return False

    file_name = os.path.basename(zipfile_path)  # 获取文件名
    new_name = str(file_name.split('.')[0]) + '.xlsx'  # 新的文件名，命名为：xxx.xlsx

    dir_path = os.path.dirname(zipfile_path)  # 获取文件所在目录
    new_path = os.path.join(dir_path, new_name)  # 新的文件路径
    if os.path.exists(new_path):
        os.remove(new_path)

    os.rename(zipfile_path, new_path)  # 保存新文件，旧文件会替换掉

    return new_path  # 返回新的文件路径，压缩包

def parseimg(excel_file_path, img_path=None):
    """
    返回图片信息
    :param excel_file_path: excel地址
    :param img_path: 图片地址：
            1. img_path默认为空，可以不传，只返回图片位置信息
            2. 指定img_path，则将excel图片解析到指定路径
    """
    result = {}
    zip_file_path = change_file_name(excel_file_path)
    if img_path != None and not os.path.exists(img_path):
        os.makedirs(img_path)
    if zip_file_path != '':
        unzip_msg = unzip_file(zip_file_path)
        if unzip_msg:
            image_list = read_img(zip_file_path, img_path)
        else:
            result['msg'] = 'unzip file failed'
            return result
    revert_dir(zip_file_path)
    return image_list

# 返回图片信息
def img_info(drawings_path, drawings_rels_path):
    try:
        tree = ET.parse(drawings_path)
        rels_tree = ET.parse(drawings_rels_path)
        # 获得根节点
        root = tree.getroot()
        rels_root = rels_tree.getroot()
    except Exception as e:  # 捕获除与程序退出sys.exit()相关之外的所有异常
        print("parse drawing1.xml or drawing1.xml.rels fail!")
        sys.exit()
    rels_xmlns = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    rel_dict = {}
    for rel in rels_root.iterfind(rels_xmlns + 'Relationship'):  # 在文件中查找Value的节点，生成器
        rel_dict[rel.get('Id')] = rel.get('Target').replace('../media/','')
    ns = {'xmlns_xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
          'xmlns_a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'xmlns_r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    xmlns_r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

    img_list = []
    for twoCellAnchor in root.findall('xmlns_xdr:twoCellAnchor', ns):  # 在文件中查找Value的节点，生成器
        value_list = []
        fplace_dic = {}
        tplace_dic = {}
        fplace = twoCellAnchor.find('xmlns_xdr:from', ns)
        fplace_dic['col'] = int(fplace[0].text)
        fplace_dic['row'] = int(fplace[2].text)
        value_list.append(fplace_dic)
        tplace = twoCellAnchor.find('xmlns_xdr:to', ns)
        tplace_dic['col'] = int(tplace[0].text)
        tplace_dic['row'] = int(tplace[2].text)
        value_list.append(tplace_dic)
        pic = twoCellAnchor.find('xmlns_xdr:pic', ns)
        blipFill = pic.find('xmlns_xdr:blipFill', ns)
        blip = blipFill.find('xmlns_a:blip', ns)
        target = blip.get(xmlns_r + 'embed')
        img_name = get_value(target, rel_dict)
        value_list.append(img_name)
        img_list.append(value_list)
    return img_list

def get_value(target, rels_list):
    """
    通过key找到字典列表对应的value
    :param target: 目标key
    :param rels_list: 字典列表
    :return: 对应的value
    """
    for key,value in rels_list.items():
        if target == key:
            return value

# if __name__ == '__main__':
#     #excel地址
#     excel_path = 'C:\\Users\\luche\\Desktop\\EP13.xlsx'
#     img = 'C:\\Users\\luche\\Desktop\\4444'
#     data_json = parseimg(excel_path,img)
#     print(data_json)