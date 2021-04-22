# -*- coding: utf-8 -*-
"""
Created on Fri Apr 16 18:43:12 2021

@author: gsalomon
"""


import imageio
from scipy import ndimage
import matplotlib.pyplot as plt
import numpy as np
from skimage import filters
from skimage import measure

circle = imageio.imread('image2021-04-16-095423-1.jpg')
plt.imshow(circle, cmap=plt.cm.gray)
plt.axis('off')

med_denoise = ndimage.median_filter(circle, 10)
plt.imshow(med_denoise, cmap=plt.cm.gray)
plt.axis('off')

val = filters.threshold_otsu(med_denoise)
mask = med_denoise < val
plt.imshow(mask, cmap=plt.cm.gray)
plt.axis('off')

plt.contour(mask, levels=2)


a = med_denoise*mask
med_denoise = ndimage.median_filter(a, 3)
plt.imshow(med_denoise, cmap=plt.cm.gray)

b = med_denoise[:-1] - med_denoise[1:]
plt.imshow(b, cmap=plt.cm.gray)
plt.axis('off')

filter_bl_f = ndimage.gaussian_filter(med_denoise, 1)
alpha = 30
sharpened = med_denoise + alpha * (med_denoise - filter_bl_f)
plt.imshow(sharpened, cmap=plt.cm.gray)
plt.axis('off')

sx = ndimage.sobel(circle,axis=0, mode='constant')
sy = ndimage.sobel(circle,axis=1, mode='constant')
sob = np.hypot(sx, sy).astype(np.int)
plt.imshow(sob, cmap=plt.cm.gray)
plt.axis('off')
med_denoise = ndimage.median_filter(sob, 50)
plt.imshow(med_denoise, cmap=plt.cm.gray)
plt.axis('off')

hist, bin_edges = np.histogram(circle, bins=60)
#hist, bin_edges, patches = plt.hist(np.ravel(circle), bins=60)
bin_centers = 0.5*(bin_edges[:-1] + bin_edges[1:])
plt.plot(bin_centers, hist)

x = sum(circle)
hist, bin_edges, patches = plt.hist(x, bins=60)

b = np.ravel(circle)
c = b[b < 50]
hist, bin_edges, patches = plt.hist(c, bins=60)

bb_img = (circle < 128) * circle
plt.imshow(bb_img, cmap=plt.cm.gray)
plt.axis('off')

binary_img = (circle > 75) * (circle < 125)
binary_img = binary_img.astype(np.int)
plt.imshow(binary_img, cmap=plt.cm.gray)
open_img = ndimage.binary_opening(binary_img)
plt.imshow(open_img, cmap=plt.cm.gray)
close_img = ndimage.binary_closing(open_img)
plt.imshow(close_img, cmap=plt.cm.gray)



sx = ndimage.sobel(mask,axis=0, mode='constant')
sy = ndimage.sobel(mask,axis=1, mode='constant')
sob = np.hypot(sx, sy).astype(np.int)
plt.imshow(sob, cmap=plt.cm.gray)
plt.axis('off')

filter_bl_f = ndimage.gaussian_filter(sob, 1)
alpha = 30
sharpened = sob + alpha * (sob - filter_bl_f)
plt.imshow(sharpened, cmap=plt.cm.gray)
plt.axis('off')

#%% Shape Searching
level = (np.max(cutout_image) + np.min(cutout_image)) / 2
contours = measure.find_contours(cutout_image, level)


fig, ax = plt.subplots()
ax.imshow(cutout_image, cmap=plt.cm.gray)
for contour in contours:
    ax.plot(contour[:, 1], contour[:, 0], linewidth=2)
contours = measure.find_contours(cutout_image, 20)
