# -*- coding: utf-8 -*-
"""
Created on Wed Feb 10 21:24:34 2021

@author: Greg
"""
#%% Imports
from pathlib import Path
import pandas as pd
import numpy as np
import xlwings as xw
import pydicom
import imageio
from shapely.geometry import Polygon
from scipy import ndimage
from skimage import filters
from skimage import measure
from load_dicom_e_plan import get_plan_data
from load_dicom_e_plan import get_block_coord


#%% Scale Factors
in_scale = 72.0 # inches to Pixles conversion
cm_scale = in_scale/2.54 # cm to Pixles conversion

#%% functions
def get_block_info(block_coords, workbook):
    def insert_block_coordinates(field_groups, selected_field, workbook):    
    # Add Block Coordinates for first field to CutOut Coordinates table for plotting
        coords = field_groups.get_group(selected_field)
        coords_sheet = workbook.sheets['CutOut Coordinates']
        coords_sheet.range('A3').options(pd.DataFrame, header=False, 
                                         index=False).value = coords
        return coords

    def insert_ssd(plan_df, selected_field, workbook):
        # Set SSD from selected_field
        ssd = plan_df.at['Actual SSD',selected_field]
        ssd_range = workbook.names['SSD'].refers_to_range
        ssd_range.value = ssd

    def insert_applicator_size(plan_df, selected_field, workbook):
        # Set applicator size from selected_field
        insert_size_range = workbook.names['Insert_Size'].refers_to_range
        insert_size = plan_df.at['ApplicatorOpening',selected_field]
        insert_size_range.value = insert_size

    # Select field to use for cutout dimensions
    field_groups = block_coords.groupby(['PlanId', 'FieldId'], axis='columns')
    selected_field = list(field_groups.groups)[0]

    coords = insert_block_coordinates(field_groups, selected_field, workbook)
    insert_ssd(plan_df, selected_field, workbook)
    insert_applicator_size(plan_df, selected_field, workbook)
    return coords


def insert_cutout_dimensions(coords, workbook):
    # Cutout shape info
    apperature = Polygon(np.array(coords))
    cutout_area = apperature.area
    cutout_perim = apperature.length
    cutout_eq_sq = 4 * cutout_area / cutout_perim
    cutout_size = list(apperature.bounds)
    cutout_extent = max([abs(x) for x in apperature.bounds])
    
    cutout_area_range = workbook.names['Cutout_Area'].refers_to_range
    cutout_area_range.value = cutout_area
    cutout_perim_range = workbook.names['Cutout_Perimeter'].refers_to_range
    cutout_perim_range.value = cutout_perim
    cutout_eq_sq_range = workbook.names['Cutout_Eq._Sq.'].refers_to_range
    cutout_eq_sq_range.value = cutout_eq_sq
    cutout_extent_range = workbook.names['Cutout_Extent'].refers_to_range
    cutout_extent_range.value = cutout_extent


def scale_cutout_graph(workbook):
    # Set the size and scale of the graph to match the applicator
    image_sheet = workbook.sheets['CutOut Image']
    outline = image_sheet.charts['Outline']
    # Graph Size
    outline.width = cm_scale*insert_size
    outline.height = cm_scale*insert_size
    # Graph max and min limits
    outline.api[1].Axes().Item(1).MinimumScale = -insert_size/2
    outline.api[1].Axes().Item(1).MaximumScale = insert_size/2
    outline.api[1].Axes().Item(2).MinimumScale = -insert_size/2
    outline.api[1].Axes().Item(2).MaximumScale = insert_size/2
    return image_sheet, outline



#%% Directory Paths
# data_path = Path(r'L:\temp\Plan Checking Temp')
# folder = data_path / 'DICOM'
# template_dir = Path(r'\\dkphysicspv1\e$\Gregs_Work\Plan Checking\Plan Check Tools\Templates')

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
block_coords, plan_df = get_plan_data(plan_files)

#%% Save Data
workbook = xw.Book(template_path)
workbook.save(save_file)
plan_data_sheet = workbook.sheets.add('Plan Data')
plan_data_sheet.range('A1').value = plan_df
plan_data_sheet.autofit()
block_coords_sheet = workbook.sheets.add('Block Coordinates')
block_coords_sheet.range('A1').value = block_coords
workbook.save(save_file)


coords = get_block_info(workbook)
insert_cutout_dimensions(coords, workbook)

image_sheet, outline = scale_cutout_graph(workbook)



#%% Load cutout image
cutout_image = imageio.imread(image_file)
# Get image size
dpi = np.array(cutout_image.meta['dpi'])
image_size = cutout_image.shape / dpi
# median filter
med_denoise = ndimage.median_filter(cutout_image, 10)
# find contours
contours = measure.find_contours(med_denoise, 20)
contours = sorted(contours, key=len, reverse=True)
encoder = [9 / 25.4, 0, 0, 0]  # Encoder is 9 mm height
insert_outline = Polygon(contours[1]/dpi)
insert_limits = np.array(insert_outline.bounds) + encoder # Removed encoder to get just insert

#%% Shape Searching
# Insert image

height = image_size[0]*in_scale
width = image_size[1]*in_scale
cutout_shape = image_sheet.pictures.add(image_file, name="Cutout")
# cutout_shape.api.ZOrder = 1  # Send to back  ### FIXME Not Working ####

cutout_shape.width = width
cutout_shape.height = height
#%% Crop with margin
margin = np.array([-1, -1, 1, 1]) * 0.5 # 1/2" margin
#margin = 0 # 1/2" margin
pic_location = np.array([0, 0])  # Top, Left in pixles
crop_size = insert_limits + margin
crop_dim = np.int_(crop_size * in_scale)

top, left, bottom, right = crop_dim
cutout_shape.api.ShapeRange.PictureFormat.CropTop = top
cutout_shape.api.ShapeRange.PictureFormat.CropLeft = left
cutout_shape.api.ShapeRange.PictureFormat.CropBottom = height - bottom
cutout_shape.api.ShapeRange.PictureFormat.CropRight = width - right
cutout_shape.top = pic_location[0]
cutout_shape.left = pic_location[1]

image_sheet.activate()
#%% Set the crosshair arrows to match the applicator
mid_point = np.array([cutout_shape.height, cutout_shape.width]) / 2 + pic_location

up_arrow = image_sheet.shapes['UpArrow']
up_arrow.height = cm_scale*insert_size
up_arrow.width = 0
up_arrow.left = mid_point[1]
up_arrow.top = mid_point[0] - cm_scale*insert_size/2

horz_arrow = image_sheet.shapes['HorzArrow']
horz_arrow.width = cm_scale*insert_size
horz_arrow.height = 0
horz_arrow.top = mid_point[0]
horz_arrow.left = mid_point[1] - cm_scale*insert_size/2

# combine the arrows to form a cross hair
image_sheet.shapes.api.Range(['UpArrow', 'HorzArrow']).Align(1,0) # Align Center
image_sheet.shapes.api.Range(['UpArrow', 'HorzArrow']).Align(4,0) # Align middle
image_sheet.shapes.api.Range(['UpArrow', 'HorzArrow']).Group()

#%% Other field Parameters
parameter_sheet = workbook.sheets['CutOut Parameters']
params = ['RadiationType', 'SetupTechnique', 'ToleranceTable', 'Linac', 
          'Energy', 'GantryAngle', 'ApplicatorID', 'AccessoryCode', 
          'BlockTrayID', 'InsertCode', 'BlockType', 'MaterialID', 
          'BlockDivergence', 'BlockName', 'SourceToBlockTrayDistance']
parameter_sheet.range('B1').options(pd.DataFrame, header=True, 
                                 index=False).value = plan_df.loc[params,:]

