# TEM 自动处理工具

自动旋转、裁剪 TEM 图片，生成 PowerPoint 报告。

---

## 功能说明

1. 从两个文件夹读取 TEM 图片（`.tif`）：**Standard**（截面）和 **Planar**（平面）
2. 用 FFT（快速傅里叶变换）自动检测图片倾斜角度
3. 自动旋转，让层状结构变水平（Standard）或竖直（Planar）
4. 按标准尺寸裁剪和缩放
5. 生成 PPT，每页多张图，按类型分组

---

## 手把手教程

### 第一步：安装 Python

如果电脑上没有 Python：
- 打开 https://www.python.org/downloads/
- 下载安装，**安装时一定要勾选 "Add Python to PATH"**

验证安装成功：打开 PowerShell，输入：
```
python --version
```
如果显示版本号（如 `Python 3.11.x`），说明安装成功。

### 第二步：安装依赖库

在 PowerShell 中运行：

```
pip install python-pptx Pillow opencv-python-headless numpy
```

等它跑完，看到 `Successfully installed ...` 就行了。

### 第三步：准备图片

建一个文件夹，里面放两个子文件夹：

```
我的TEM数据/
├── standard/       ← 放 Standard（截面）TEM 的 .tif 文件
└── planar/         ← 放 Planar（平面）TEM 的 .tif 文件
```

> **说明：** 子文件夹名字只要以 "standard" 或 "planar" 开头就行（不区分大小写）。
> 比如 `Standard TEM Data/` 和 `Planar TEM Data/` 都可以识别。
>
> 可以只放一种类型（比如只有 standard 文件夹也行）。

### 第四步：运行

把 `tem_process.py` 放到任意位置，然后在 PowerShell 中运行：

```
python tem_process.py 我的TEM数据
```

搞定！PPT 会自动保存为 `我的TEM数据/output.pptx`。

---

## 更多用法

```
# 处理当前目录（当前目录下有 standard/ 和 planar/ 文件夹）
python tem_process.py

# 处理指定路径
python tem_process.py "C:\Users\xxx\Desktop\TEM数据"

# 指定输出文件名
python tem_process.py 我的TEM数据 -o 报告.pptx
```

---

## 输出效果

- PPT 尺寸：13.333 × 7.5 英寸（宽屏）
- Standard 图片：缩放到 3.8×3.8"，裁剪到 2.8×2.5"
- Planar 图片：缩放到 2.8×2.8"，裁剪到 1.81×2.5"
- 每页自动排列多张图，按类型分组
- 每张图下方显示文件名

---

## 旋转原理

工具用 **FFT（快速傅里叶变换）** 分析图片的频谱。TEM 图中的平行层状结构会在频域中形成明显的方向性峰值，测量这个峰值偏离水平/竖直的角度，就是需要旋转的角度。

- **Standard TEM**：层状结构应该水平 → 检测偏离水平多少度
- **Planar TEM**：结构应该竖直 → 检测偏离竖直多少度

精度：与手动旋转相比，误差通常 **< 1°**。

---

## 常见问题

**Q: 提示 `pip` 不是命令？**
A: 用 `python -m pip install ...` 代替 `pip install ...`

**Q: 提示找不到 standard 或 planar 文件夹？**
A: 检查子文件夹名称是否以 "standard" 或 "planar" 开头

**Q: 图片旋转角度不对？**
A: 确认图片放对了文件夹（截面图放 standard，平面图放 planar）
