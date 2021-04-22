# -*- coding: utf-8 -*-
"""
Created on Sun Apr 18 12:49:07 2021

@author: Greg
"""
import math
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
from shapely.geometry import Polygon
import imageio
from scipy import ndimage
from skimage import filters
from skimage import measure

#%% Test directory is current directory
data_path = Path.cwd()
folder = Path.cwd()
template_dir = Path.cwd()

#%% File Paths
template_file_name = 'CutOut Size Test.xlsx'
output_file_name = 'CutOut Size Check.xlsx'
image_file_name = 'image2021-04-16-095423-1.jpg'
template_path = template_dir / template_file_name
save_file = data_path / output_file_name
image_file = data_path / image_file_name

#%% read image
cutout_image = imageio.imread(image_file_name)
plt.imshow(cutout_image, cmap=plt.cm.gray)

#%% Get image size
dpi = np.array(cutout_image.meta['dpi'])
image_size = cutout_image.shape / dpi

#%% cross section
plt.plot(cutout_image[1500])
plt.plot(cutout_image[1000])
#%% median filter
med_denoise = ndimage.median_filter(cutout_image, 10)
#%%
plt.plot(med_denoise[2000])
plt.axis([0, 5000, 0, 20])
#%% find contours
contours = measure.find_contours(med_denoise, 20)
contours = sorted(contours, key=len, reverse=True)
a = [len(c) for c in contours]
a.sort(reverse=True)
#%% Shape Searching
fig, ax = plt.subplots()
#ax.imshow(med_denoise, cmap=plt.cm.gray)
for contour in contours[:3]:
    x = contour[:,1] / dpi[0]
    y = contour[:,0] / dpi[1]
    ax.plot(x, y, linewidth=2)


#%% image area
page = Polygon(contours[0]/dpi)
page.bounds
page.area

#%% Add Figure
in_scale = 72.0 # inches to Pixles conversion
cm_scale = in_scale/2.54 # cm to Pixles conversion
encoder = [9 / 25.4, 0, 0, 0]  # Encoder is 9 mm height

insert = Polygon(contours[1]/dpi)
insert_size = np.array(insert.bounds) + encoder # Removed encoder to get just uinsert remaining

workbook = xw.Book()
image_sheet = workbook.sheets.add('CutOut Image')

height = image_size[0]*in_scale
width = image_size[1]*in_scale
pic = image_sheet.pictures.add(image_file, name="Cutout", width=width, height=height)

#%% Crop with margin
margin = np.array([-1, -1, 1, 1]) * 0.5 # 1/2" margin
margin = 0 # 1/2" margin
pic_location = np.array([0, 0])  # Top, Left in pixles
crop_size = insert_size + margin
crop_dim = np.int_(crop_size * in_scale)

# horiz = int(pic.left)
# virt = int(pic.top)

top, left, bottom, right = crop_dim
pic.api.ShapeRange.PictureFormat.CropTop = top
pic.api.ShapeRange.PictureFormat.CropLeft = left
pic.api.ShapeRange.PictureFormat.CropBottom = height - bottom
pic.api.ShapeRange.PictureFormat.CropRight = width - right
pic.top = pic_location[0]
pic.left = pic_location[1]


#%% Add cross hair
edges = (crop_size + [0.5, 0, 0, 0]) * in_scale
#mid_point = ((crop_dim[1::2] - crop_dim[0::2]) / 2)  #+ crop_dim[0:1]

mid_point = np.array([pic.height, pic.width]) / 2 + pic_location
circle = image_sheet.api.Shapes.AddShape(4, mid_point[1], mid_point[0], 10, 10)

virt_line = image_sheet.api.Shapes.AddConnector(1, mid_point[0], pic_location[0], mid_point[0], pic.height)
horz_line = image_sheet.api.Shapes.AddConnector(1, pic_location[1], mid_point[1], pic.width, mid_point[1])
half_line = image_sheet.api.Shapes.AddConnector(1, pic_location[1], mid_point[1], pic.width/2, mid_point[1])


c = int('{:#x}{:x}{:x}'.format(0,0,255),16)  # RGB to Int
half_line.Line.ForeColor.RGB = c
virt_line = image_sheet.api.Shapes.AddConnector(1, mid_point[0], pic_location[0], mid_point[0], pic.height/2)
virt_line.Line.ForeColor.RGB = c
# Encoder is 9 mm height
# Wide encoder is 9 cm wide
# Narrow encoder is 52 mm wide
# Wide encoder distance to first hole (plugged) 14mm
# Narrow encoder distance to first hole (plugged) 8 mm
#%% VB Code

#     ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 132.75, 58.5, 133.5, _
#         223.5).Select
#         Selection.ShapeRange.Line.BeginArrowheadStyle = msoArrowheadTriangle
#         Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
#     Selection.ShapeRange.IncrementLeft 6.75
#     Selection.ShapeRange.IncrementTop 2.25
#     Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOpen
#     With Selection.ShapeRange.Line
#         .Visible = msoTrue
#         .Weight = 2
#     End With
#     With Selection.ShapeRange.Line
#         .Visible = msoTrue
#         .ForeColor.RGB = RGB(255, 0, 0)
#         .Transparency = 0
#     End With
#     ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 52.5, 144, 219, 146.25). _
#         Select
#         Selection.ShapeRange.Line.BeginArrowheadStyle = msoArrowheadTriangle
#         Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
#     With Selection.ShapeRange.Line
#         .Visible = msoTrue
#         .Weight = 3
#     End With
#     Selection.ShapeRange.Line.BeginArrowheadStyle = msoArrowheadOpen
#     Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadOpen
#     With Selection.ShapeRange.Line
#         .EndArrowheadLength = msoArrowheadShort
#         .EndArrowheadWidth = msoArrowheadWide
#     End With
#     With Selection.ShapeRange.Line
#         .BeginArrowheadLength = msoArrowheadLong
#         .BeginArrowheadWidth = msoArrowheadWidthMedium
#     End With
#     With Selection.ShapeRange.Line
#         .Visible = msoTrue
#         .ForeColor.RGB = RGB(0, 176, 80)
#         .Transparency = 0
#     End With
#     Selection.ShapeRange.IncrementLeft 8.25
#     Selection.ShapeRange.IncrementLeft -8.25
#     Selection.ShapeRange.IncrementTop -5.25


# Const msoConnectorStraight = 1
# Const msoArrowheadTriangle = 2
# Const msoArrowheadOpen = 3
# Const msoArrowheadLong = 3
# Const msoArrowheadShort = 1
# Const msoArrowheadLengthMedium = 2
# Const msoArrowheadNarrow = 1
# Const msoArrowheadWide = 3
# Const msoArrowheadWidthMedium = 2


#%% Insert shape
a = contours[1]/dpi
insert = Polygon(a)
b = insert.bounds
x = a[:,0]
y = a[:,1]
plt.plot(x, y, linewidth=2)
c = (a[:,0] > b[0] + 0.5) * (a[:,0] < b[2] - 0.5) * (a[:,1] < b[1] + 0.5)
l1 = a[c,:]
x = l1[:,0]
y = l1[:,1]
p_fit = np.polyfit(x, y, deg=1)
fit_x = np.arange(b[0],b[2],0.5)
fit_y = p_fit[1] + fit_x * p_fit[0]
plt.plot(x, y, 'b+', fit_x, fit_y, 'r-', linewidth=2, markevery=5)
a = math.degrees(math.atan(p_fit[0]))
pic.api.ShapeRange.Rotation = a
#%%
c = (a[:,0] > b[0] + 0.5) * (a[:,0] < b[2] - 0.5) * (a[:,1] > b[3] - 0.5)
l1 = a[c,:]
x = l1[:,0]
y = l1[:,1]
p_fit = np.polyfit(x, y, deg=1)
fit_x = np.arange(b[0],b[2],0.5)
fit_y = p_fit[1] + fit_x * p_fit[0]
plt.plot(x, y, 'b+', fit_x, fit_y, 'r-', linewidth=2, markevery=5)

#%%
c = (a[:,1] > b[1] + 0.5) * (a[:,1] < b[3] - 0.5) * (a[:,0] > b[2] - 0.5)
l1 = a[c,:]
x = l1[:,1]
y = l1[:,0]
p_fit = np.polyfit(x, y, deg=1)
fit_x = np.arange(b[1],b[3],0.5)
fit_y = p_fit[1] + fit_x * p_fit[0]
plt.plot(x, y, 'b+', fit_x, fit_y, 'r-', linewidth=2, markevery=5)

plt.plot(x, y, 'b+', markevery=5)
plt.plot(fit_x, fit_y, 'r-', linewidth=2)
plt.axis([0, 5000, 0, 20])


#%%
apperature_points = contours[2]/dpi*2.54
apperature = Polygon(apperature_points)
apperature.bounds
cutout_area = apperature.area
cutout_perim = apperature.length
cutout_eq_sq = 4 * cutout_area / cutout_perim
cutout_size = list(apperature.bounds)
cutout_extent = max([abs(x) for x in apperature.bounds])
x = apperature_points[:,0]
y = apperature_points[:,1]
plt.plot(x, y, linewidth=2)

#%% Shape Searching
fig, ax = plt.subplots()
#ax.imshow(med_denoise, cmap=plt.cm.gray)
for contour in contours:
    x = contour[:,1] / dpi[0]
    y = contour[:,0] / dpi[1]
    ax.plot(x, y, linewidth=2)
