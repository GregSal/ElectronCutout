# -*- coding: utf-8 -*-
"""
Created on Sun Apr 18 12:03:11 2021

@author: Greg
"""
#%% Imports
from pathlib import Path
import pandas as pd
import numpy as np
import xlwings as xw
import pydicom
import imageio
import shapely

import matplotlib.pyplot as plt
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
#%% Load DICOM Data
plan_files = [file for file in folder.glob('**/RP*.dcm')]
plan_file = plan_files[0]
ds = pydicom.dcmread(plan_file)
field_ds = ds.BeamSequence[0]
block_seq = getattr(field_ds,'BlockSequence', None)
block = block_seq[0]
block_coord_data = getattr(block,'BlockData', None)
block_coordinates = np.array(block_coord_data).reshape((-1,2))/10
block_coordinates = np.row_stack([block_coordinates,block_coordinates[0]])

#%%
from shapely.geometry import Polygon
apperature = Polygon(block_coordinates)
cutout_area = apperature.area
cutout_perim = apperature.length
cutout_size = list(apperature.bounds)
cutout_extent = max([abs(x) for x in apperature.bounds])
