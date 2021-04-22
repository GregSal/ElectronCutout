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

#%% DICOM Plan Sections
# Tolerance table reference
def get_tolerance_tables(ds):
    tolerance_ref = pd.Series(
        {tol.ToleranceTableNumber: tol.ToleranceTableLabel
         for tol in ds.ToleranceTableSequence},
         name='ToleranceTable')
    return tolerance_ref


# plan setup
def get_plan_setup(ds):
    setup_ref = pd.DataFrame(
        {setup.PatientSetupNumber: {
            'SetupTechnique': setup.SetupTechnique,
            'PatientOrientation': setup.PatientPosition
            }
        for setup in ds.PatientSetupSequence}
        ).T
    return setup_ref


# MUs
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


# Fields
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

# Applicator
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

# Control Point
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

#%% All DICOM Field Data
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


#%% Clean and Combine Field Data
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

# Merge Field Data
def get_merged_field_data(ds):
    tolerance_ref = get_tolerance_tables(ds)
    setup_ref = get_plan_setup(ds)
    plan_mus = get_mus(ds)
    field_df = get_field_data(ds)
    field_df = combine_field_tables(field_df, plan_mus, setup_ref, tolerance_ref)
    return field_df

#%% Block Coordinates Table
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


#%% Load Files
def read_dicom_plan(plan_file):
    ds = pydicom.dcmread(plan_file)
    plan_name = ds.RTPlanLabel
    patient_id = ds.PatientID
    #SOPInstanceUID
    field_df = get_merged_field_data(ds)
    field_df['PlanId'] = plan_name
    field_df['PatientID'] = patient_id
    return field_df


# Read all files
def get_plan_data(plan_files):
    plan_data = list()
    for plan_file in plan_files:
        field_df = read_dicom_plan(plan_file)
        plan_data.append(field_df)
    plan_df = pd.concat(plan_data)
    plan_df.set_index(['PlanId', 'FieldId'], inplace=True)
    plan_df = plan_df.T
    return plan_df

#%% Main
def main():
    # File Paths
    folder = Path.cwd()
    output_file_name = 'Electron Plan DICOM Info.xlsx'
    save_file = folder / output_file_name
    # Load DICOM Data
    plan_files = [file for file in folder.glob('**/RP*.dcm')]
    plan_df = get_plan_data(plan_files)
    block_coords = get_block_coord(plan_df)
    plan_df.drop(index=['Coordinates'], inplace=True)
    # Save Data
    workbook = xw.Book()
    workbook.save(save_file)
    plan_data_sheet = workbook.sheets.add('Plan Data')
    plan_data_sheet.range('A1').value = plan_df
    plan_data_sheet.autofit()
    block_coords_sheet = workbook.sheets.add('Block Coordinates')
    block_coords_sheet.range('A1').value = block_coords
    block_coords_sheet.autofit()
    workbook.save(save_file)


if __name__ == '__Main__':
    main()
