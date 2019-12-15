- 安装卸载
  - 将dist下的安装包parseimg-1.0.tar.gz放到项目根目录下
  - **pip install parseimg-1.0.tar.gz**安装成功
  - 卸载**pip uninstall parseimg**即可
- 使用方法
  - 安装成功后，导入相关模块方法**from parse.parseimg import parseimg**
  - 调用parseimg方法，传入excel文件路径即可，返回excel图片信息
  - 返回信息json格式：[[{"col": 2, "row": 0}, {"col": 2, "row": 0}, "image1.png"], [{"col": 2, "row": 3}, {"col": 2, "row": 4}, "image2.png"], [{"col": 2, "row": 4}, {"col": 2, "row": 5}, "image3.png"], [{"col": 2, "row": 2}, {"col": 2, "row": 2}, "image2.png"]]

