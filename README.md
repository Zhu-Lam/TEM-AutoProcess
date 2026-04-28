# TEM Auto-Processor

Automatically rotate and crop TEM images, then generate a PowerPoint report.

---

## What It Does

1. Reads TEM images (`.tif`) from two folders: **Standard** and **Planar**
2. Detects tilt angle using FFT (Fast Fourier Transform)
3. Rotates each image to make layers perfectly horizontal (Standard) or vertical (Planar)
4. Crops and resizes to standard PPT dimensions
5. Generates a PowerPoint file with all processed images

---

## Step-by-Step Setup

### 1. Install Python

If you don't have Python installed:
- Go to https://www.python.org/downloads/
- Download and install (check **"Add Python to PATH"** during installation)

### 2. Install Dependencies

Open a terminal (PowerShell or Command Prompt) and run:

```
pip install python-pptx Pillow opencv-python-headless numpy
```

### 3. Prepare Your Images

Create a folder with this structure:

```
my_tem_data/
├── standard/       ← Put Standard (cross-section) TEM .tif files here
└── planar/         ← Put Planar TEM .tif files here
```

> **Note:** Folder names just need to start with "standard" or "planar" (case-insensitive).
> For example, `Standard TEM Data/` and `Planar TEM Data/` both work.

### 4. Run

```
python tem_process.py my_tem_data
```

That's it! The output PPT will be saved as `my_tem_data/output.pptx`.

---

## More Options

```
# Process current directory
python tem_process.py

# Process a specific folder
python tem_process.py "C:\path\to\my_data"

# Custom output file name
python tem_process.py my_tem_data -o my_report.pptx
```

---

## Output

- PowerPoint file (13.333 × 7.5 inches, widescreen)
- Standard images: resized to 3.8×3.8", cropped to 2.8×2.5"
- Planar images: resized to 2.8×2.8", cropped to 1.81×2.5"
- Multiple images per slide, grouped by type
- Filename shown below each image

---

## How Rotation Detection Works

The tool uses **FFT (Fast Fourier Transform)** to analyze the frequency spectrum of each image. TEM images have strong parallel layer structures that create a directional peak in the frequency domain. By measuring the angle of this peak, the tool determines exactly how much to rotate the image to straighten the layers.

- **Standard TEM**: layers should be horizontal → detects deviation from horizontal
- **Planar TEM**: structures should be vertical → detects deviation from vertical

Typical accuracy: **< 1° error** compared to manual rotation.
