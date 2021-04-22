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

#%% Tolerance table reference
def get_tolerance_tables(ds):
    tolerance_ref = pd.Series(
        {tol.ToleranceTableNumber: tol.ToleranceTableLabel
         for tol in ds.ToleranceTableSequence},
         name='ToleranceTable')
    return tolerance_ref


#%% plan setup
def get_plan_setup(ds):
    setup_ref = pd.DataFrame(
        {setup.PatientSetupNumber: {
            'SetupTechnique': setup.SetupTechnique,
            'PatientOrientation': setup.PatientPosition
            }
        for setup in ds.PatientSetupSequence}
        ).T
    return setup_ref


#%% MUs
def get_mus(ds):
    MU_data = dict()
    dose_data = dict()
    field_mu_seq = getattr(ds,'FractionGroupSequence', None)
    if field_mu_seq:
        beams = field_mu_seq[0]
        beam_seq = getattr(beams,'ReferencedBeamSequence', None)
        if beam_seq:
            MU_data = dict()
            dose_data = dict()
            for field in beam_seq:
                field_MU = getattr(field,'BeamMeterset', None)
                if field_MU:
                    MU_data[field.ReferencedBeamNumber] = field_MU
                field_Dose = getattr(field,'BeamDose', None)
                if field_Dose:
                    dose_data[field.ReferencedBeamNumber] = field_Dose
    if MU_data:
        plan_mus = pd.Series(MU_data, name='MUs')
        if dose_data:
            plan_dose = pd.Series(dose_data, name='Beam Dose')
            plan_mus = pd.concat([plan_mus, plan_dose],axis='columns')
    else:
        plan_mus = pd.DataFrame()
    return plan_mus


#%% Fields
def get_block_info(field_ds):
    block_seq = getattr(field_ds,'BlockSequence', None)
    if not block_seq:
        return {}
    block = block_seq[0]
    block_data = {
        'BlockName': getattr(block,'BlockName', None),
        'BlockTrayID': getattr(block,'BlockTrayID', None),
        'MaterialID': getattr(block,'MaterialID', None),
        'BlockType': getattr(block,'BlockType', None),
        'InsertCode': getattr(block,'AccessoryCode', None),
        'BlockDivergence': getattr(block,'BlockDivergence', None)
        }
    position = getattr(block,'BlockMountingPosition', None)
    distance = getattr(block,'SourceToBlockTrayDistance', None)
    thickness = getattr(block,'BlockThickness', None)
    if distance:
        block_data['SourceToBlockTrayDistance'] = distance/10
        if position:
            block_data['BlockMountingPosition'] = position
            if thickness:
                block_data['BlockThickness'] = thickness/10
                if position in 'PATIENT_SIDE':
                    block_data['SourceToBlockDistance'] = (distance-thickness)/10
                else:
                    block_data['SourceToBlockDistance'] = (distance+thickness)/10
    block_coord_data = getattr(block,'BlockData', None)
    # Convert into (x,y) paris in units of cm
    block_coordinates = np.array(block_coord_data).reshape((-1,2))/10
    # Add the first point on the end as the last point to close the loop
    block_coordinates = np.row_stack([block_coordinates,block_coordinates[0]])
    block_data['Coordinates'] = block_coordinates
    return block_data


def get_applicator_info(field_ds):
    appl_seq = getattr(field_ds,'ApplicatorSequence', None)
    if not appl_seq:
        return {}
    appl = appl_seq[0]
    appl_geom = appl.ApplicatorGeometrySequence[0]
    appl_data = {
        'AccessoryCode': getattr(appl,'AccessoryCode', None),
        'ApplicatorID': getattr(appl,'ApplicatorID', None),
        'ApplicatorType': getattr(appl,'ApplicatorType', None),
        'ApplicatorApertureShape': getattr(appl_geom,'ApplicatorApertureShape', None),
        'ApplicatorOpening': getattr(appl_geom,'ApplicatorOpening', None)
        }
    size = getattr(appl_geom,'ApplicatorOpening', None)
    if size:
        appl_data['ApplicatorOpening'] = size/10
    return appl_data


def get_control_point_data(field_ds):
    control_point_seq = field_ds.ControlPointSequence
    control_point = control_point_seq[0]
    initial_field_data = {
        'CollimatorAngle': getattr(control_point,'BeamLimitingDeviceAngle', None),
        'DoseRate': getattr(control_point,'DoseRateSet', None),
        'GantryAngle': getattr(control_point,'GantryAngle', None),
        'Energy': getattr(control_point,'NominalBeamEnergy', None),
        'CouchAngle': getattr(control_point,'PatientSupportAngle', None),
        'Actual SSD': getattr(control_point,'SourceToSurfaceDistance', None),
        'Isocentre': getattr(control_point,'IsocenterPosition', None)
        }
    return initial_field_data


def get_field_data(ds):
    fields = list()
    for field_ds in ds.BeamSequence:
        field_data = {
                'BeamNumber': getattr(field_ds,'BeamNumber', None),
                'ToleranceTableNumber': getattr(field_ds,'ReferencedToleranceTableNumber', None),
                'PatientSetupNumber': getattr(field_ds,'ReferencedPatientSetupNumber', None),
                'FieldId': getattr(field_ds,'BeamName', None),
                'FieldType': getattr(field_ds,'BeamType', None),
                'RadiationType': getattr(field_ds,'RadiationType', None),
                'Weight': getattr(field_ds,'FinalCumulativeMetersetWeight', None),
                'Linac': getattr(field_ds,'TreatmentMachineName', None),
                'SetupField': getattr(field_ds,'TreatmentDeliveryType', None),
                'SAD': getattr(field_ds,'SourceAxisDistance', None),
                'NumberOfBlocks': getattr(field_ds,'NumberOfBlocks', None),
                'NumberOfBoli': getattr(field_ds,'NumberOfBoli', None),
                'NumberOfControlPoints': getattr(field_ds,'NumberOfControlPoints', None),
                'NumberOfWedges': getattr(field_ds,'NumberOfWedges', None)
                }
        initial_field_data = get_control_point_data(field_ds)
        field_data.update(initial_field_data)
        appl_data = get_applicator_info(field_ds)
        field_data.update(appl_data)
        block_data = get_block_info(field_ds)
        field_data.update(block_data)
        fields.append(field_data)
    field_df = pd.DataFrame(fields)
    return field_df


#%% Clean and Combine
def combine_field_tables(field_df, plan_mus, setup_ref, tolerance_ref):
    field_df['Actual SSD'] = field_df['Actual SSD']/10
    field_df['SAD'] = field_df['SAD']/10
    field_df = field_df.merge(plan_mus, how="left", left_on='BeamNumber',
                              right_index=True)
    field_df = field_df.merge(setup_ref, how="left", left_on='PatientSetupNumber',
                              right_index=True)
    field_df = field_df.merge(tolerance_ref, how="left", left_on='ToleranceTableNumber',
                              right_index=True)
    field_df.drop(columns=['BeamNumber', 'ToleranceTableNumber',
                           'PatientSetupNumber'], inplace=True)
    return field_df


def get_merged_field_data(ds):
    tolerance_ref = get_tolerance_tables(ds)
    setup_ref = get_plan_setup(ds)
    plan_mus = get_mus(ds)
    field_df = get_field_data(ds)
    field_df = combine_field_tables(field_df, plan_mus, setup_ref, tolerance_ref)
    return field_df


def get_block_coord(plan_df):
    block_coord_df = plan_df.loc['Coordinates',:]
    field_groups = block_coord_df.groupby(['PlanId', 'FieldId'])
    blk_grps = list()
    for name, group in field_groups:
        blk_grps.append(group.apply({'X': lambda x: x[:,0], 'Y': lambda x: x[:,1]}))
    blk_grps_df = pd.concat(blk_grps)
    block_coords = blk_grps_df.apply(pd.Series).T
    block_coords = block_coords.reorder_levels([1,2,0], axis='columns')
    return block_coords


#%% Load File
def get_plan_data(plan_files):
    plan_data = list()
    patient_ids = list()
    for plan_file in plan_files:
        ds = pydicom.dcmread(plan_file)
        plan_name = ds.RTPlanLabel
        patient_id = ds.PatientID
        patient_ids.append(patient_id)
        #SOPInstanceUID
        field_df = get_merged_field_data(ds)
        field_df['PlanId'] = plan_name
        field_df['PatientID'] = patient_id
        plan_data.append(field_df)
    plan_df = pd.concat(plan_data)
    plan_df.set_index(['PlanId', 'FieldId'], inplace=True)
    plan_df = plan_df.T
    block_coords = get_block_coord(plan_df)
    plan_df.drop(index=['Coordinates'], inplace=True)
    return block_coords, plan_df

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

#%% Add Block Coordinates for first field to CutOut Coordinates table for plotting
field_groups = block_coords.groupby(['PlanId', 'FieldId'], axis='columns')
first_field = list(field_groups.groups)[0]

coords = field_groups.get_group(first_field)
coords_sheet = workbook.sheets['CutOut Coordinates']
coords_sheet.range('A3').options(pd.DataFrame, header=False, 
                                 index=False).value = coords
#%% Cutout shape info
apperature = Polygon(coords)
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

#%% Set SSD ad applicator size from first field
ssd = plan_df.at['Actual SSD',first_field]
ssd_range = workbook.names['SSD'].refers_to_range
ssd_range.value = ssd

insert_size_range = workbook.names['Insert_Size'].refers_to_range
insert_size = plan_df.at['ApplicatorOpening',first_field]
insert_size_range.value = insert_size

#%% Set the size and scale of the graph to match the applicator
image_sheet = workbook.sheets['CutOut Image']
outline = image_sheet.charts['Outline']

in_scale = 72.0 # inches to Pixles conversion
cm_scale = in_scale/2.54 # cm to Pixles conversion

# Graph Size
outline.width = cm_scale*insert_size
outline.height = cm_scale*insert_size

# Graph max and min limits
outline.api[1].Axes().Item(1).MinimumScale = -insert_size/2
outline.api[1].Axes().Item(1).MaximumScale = insert_size/2
outline.api[1].Axes().Item(2).MinimumScale = -insert_size/2
outline.api[1].Axes().Item(2).MaximumScale = insert_size/2

#%% Set the crosshair arrows to match the applicator
up_arrow = image_sheet.shapes['UpArrow']
up_arrow.height = cm_scale*insert_size
up_arrow.width = 0
horz_arrow = image_sheet.shapes['HorzArrow']
horz_arrow.width = cm_scale*insert_size
horz_arrow.height = 0

# combine the arrows to form a cross hair
image_sheet.shapes.api.Range(['UpArrow', 'HorzArrow']).Align(1,0) # Align Center
image_sheet.shapes.api.Range(['UpArrow', 'HorzArrow']).Align(4,0) # Align middle
image_sheet.shapes.api.Range(['UpArrow', 'HorzArrow']).Group()


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
cutout_shape = image_sheet.pictures.add(image_file, name="Cutout", 
                                        width=width, height=height)
cutout_shape.api.ZOrder = 1  # Send to back

#%% Crop with margin
margin = np.array([-1, -1, 1, 1]) * 0.5 # 1/2" margin
pic_location = np.array([0, 0])  # Top, Left in pixles
crop_size = insert_size + margin
crop_dim = np.int_(crop_size * in_scale)

top, left, bottom, right = crop_dim
cutout_shape.api.ShapeRange.PictureFormat.CropTop = top
cutout_shape.api.ShapeRange.PictureFormat.CropLeft = left
cutout_shape.api.ShapeRange.PictureFormat.CropBottom = height - bottom
cutout_shape.api.ShapeRange.PictureFormat.CropRight = width - right
cutout_shape.top = pic_location[0]
cutout_shape.left = pic_location[1]

image_sheet.activate()
mid_point = np.array([cutout_shape.height, cutout_shape.width]) / 2 + pic_location

#%% Other field Parameters
parameter_sheet = workbook.sheets['CutOut Parameters']
params = ['RadiationType', 'SetupTechnique', 'ToleranceTable', 'Linac', 
          'Energy', 'GantryAngle', 'ApplicatorID', 'AccessoryCode', 
          'BlockTrayID', 'InsertCode', 'BlockType', 'MaterialID', 
          'BlockDivergence', 'BlockName', 'SourceToBlockTrayDistance']
parameter_sheet.range('B1').options(pd.DataFrame, header=True, 
                                 index=False).value = plan_df.loc[params,:]

