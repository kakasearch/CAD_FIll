# CAD_FIll
## 功能
 识别CAD中由多段线或直线围成的闭合多边形，按要求的面积分割多边形并填充指定图案
## 使用
 ![](https://github.com/kakasearch/CAD_FIll/blob/master/%E6%BC%94%E7%A4%BA.gif)
 1. 先在AutoCAD中打开要处理的图纸
 2. 运行本程序后会在cad指令行中提示选择边界，此时选择需要处理的多段线或直线
 3. 程序成功边界后会提示用户输入面积，输入面积后程序通过二分法自动在多边形中填充对应面积的图案
## 注意事项
 1. 多边形只能由多段线或直线组成，多边形只能有1个
 2. 如多边形生成失败，可先在cad中通过`boundary`命令拾取边界
 3. 填充图案的图案样式、颜色、比例等可通过`config.json`中修改
 4. `config.json`中的`is_visual`用于过程调试，使用时默认关闭
 5. 使用`pyinstaller`打包时，需要使用`pyinstaller --hidden-import=scipy._lib.array_api_compat.numpy.fft main.py`或`pyinstaller main.spec`打包，以避免出现打包后依赖库缺失
