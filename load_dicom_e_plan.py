"""Load DICOM data for electron plans.

Created on Wed Feb 10 21:24:34 2021

@author: Greg
"""
#%% Imports
from pathlib import Path
from typing import Dict, List, Any
import numpy as np
import pandas as pd
import xlwings as xw
import pydicom


#%% DICOM Plan Sections
def get_tolerance_tables(ds: pydicom.Dataset) -> pd.Series:
    """Extract the tolerance table name from the dataset.

    Args:
        ds (pydicom.Dataset): The DICOM dataset for a plan.
    Returns:
        tolerance_ref (pd.Series): The tolerance table name for each field in
            the plan.
    """
    tolerance_ref = pd.Series(
        {tol.ToleranceTableNumber: tol.ToleranceTableLabel
         for tol in ds.ToleranceTableSequence},
        name='ToleranceTable')
    return tolerance_ref


def get_plan_setup(ds: pydicom.Dataset) -> pd.DataFrame:
    """Extract the setup technique and patient orientation from the dataset.

    Args:
        ds (pydicom.Dataset): The DICOM dataset for a plan.
    Returns:
        setup_ref (pd.DataFrame): The setup technique and patient orientation
            for a plan.
    """
    setup_ref = pd.DataFrame(
        {
            setup.PatientSetupNumber: {
                'SetupTechnique': setup.SetupTechnique,
                'PatientOrientation': setup.PatientPosition
            }
            for setup in ds.PatientSetupSequence
        }
    ).T
    return setup_ref


def get_mus(ds: pydicom.Dataset) -> pd.DataFrame:
    """Extract the MU settings from the dataset.

    Args:
        ds (pydicom.Dataset): The DICOM dataset for a plan.
    Returns:
        plan_mus (pd.DataFrame): The MUs and field dose for each field in
            the plan.  An empty DataFrame is returned if none of the fields
            contain MUs.
    """
    MU_data = dict()
    dose_data = dict()
    field_mu_seq = ds.get('FractionGroupSequence')
    if field_mu_seq:
        beams = field_mu_seq[0]
        beam_seq = beams.get('ReferencedBeamSequence')
        if beam_seq:
            for field in beam_seq:
                field_MU = field.get('BeamMeterset')
                if field_MU:
                    MU_data[field.ReferencedBeamNumber] = field_MU
                field_Dose = field.get('BeamDose')
                if field_Dose:
                    dose_data[field.ReferencedBeamNumber] = field_Dose
    if MU_data:
        plan_mus = pd.Series(MU_data, name='MUs')
        if dose_data:
            plan_dose = pd.Series(dose_data, name='Beam Dose')
            plan_mus = pd.concat([plan_mus, plan_dose], axis='columns')
    else:
        plan_mus = pd.DataFrame()
    return plan_mus


#%% DICOM Field Subsection Data
def get_block_info(field_ds: pydicom.Dataset) -> Dict[str, Any]:
    """Extract insert and cutout DICOM data for a given field.

    All insert related parameters are re-scaled from mm to cm.
    Args:
        field_ds (pydicom.Dataset): The DICOM BeamSequence sub-dataset for a
            field within a plan DICOM dataset.
    Returns:
        block_data (Dict[str, Any]): A dictionary containing insert and cutout
            attributes for the field.
    """
    block_seq = field_ds.get('BlockSequence')
    if not block_seq:
        return {}
    block = block_seq[0]
    block_data = {
        'BlockName': block.get('BlockName'),
        'BlockTrayID': block.get('BlockTrayID'),
        'MaterialID': block.get('MaterialID'),
        'BlockType': block.get('BlockType'),
        'InsertCode': block.get('AccessoryCode'),
        'BlockDivergence': block.get('BlockDivergence')
    }
    position = block.get('BlockMountingPosition')
    distance = block.get('SourceToBlockTrayDistance')
    thickness = block.get('BlockThickness')
    if distance:
        block_data['SourceToBlockTrayDistance'] = distance / 10
        if position:
            block_data['BlockMountingPosition'] = position
            if thickness:
                block_data['BlockThickness'] = thickness / 10
                if position in 'PATIENT_SIDE':
                    block_data['SourceToBlockDistance'] = (
                        (distance - thickness) / 10)
                else:
                    block_data['SourceToBlockDistance'] = (
                        (distance + thickness) / 10)
    # Extract the coordinates for the cutout as an np.array
    block_coord_data = block.get('BlockData')
    # Convert into (x,y) parts in units of cm
    block_coordinates = np.array(block_coord_data).reshape((-1, 2)) / 10
    # Add the first point on the end as the last point to close the loop
    block_coordinates = np.row_stack([block_coordinates, block_coordinates[0]])
    block_data['Coordinates'] = block_coordinates
    return block_data


def get_applicator_info(field_ds: pydicom.Dataset) -> Dict[str, Any]:
    """Extract Applicator DICOM data for a given field.

    ApplicatorOpening is rescaled from mm to cm.
    Args:
        field_ds (pydicom.Dataset): The DICOM BeamSequence sub-dataset for a
            field within a plan DICOM dataset.
    Returns:
        appl_data (Dict[str, Any]): A dictionary containing applicator
            attributes for the field.
    """
    def get_size(appl, appl_data):
        appl_geom_seq = appl.get('ApplicatorGeometrySequence')
        if appl_geom_seq:
            appl_geom = appl_geom_seq[0]
            shape = appl_geom.get('ApplicatorApertureShape')
            appl_data['ApplicatorApertureShape'] = shape
            size = appl_geom.get('ApplicatorOpening')
            if size:
                appl_data['ApplicatorOpening'] = size / 10
        else:
            appl_id = appl.get('ApplicatorID')
            size_match = re.match('A([0-9]+)',appl_id)
            if size_match:
                appl_data['ApplicatorOpening'] = int(size_match.group(1))
            else:
                appl_data = None
        return appl_data

    appl_seq = field_ds.get('ApplicatorSequence')
    if appl_seq:
        appl = appl_seq[0]
        appl_data = {
            'AccessoryCode': appl.get('AccessoryCode'),
            'ApplicatorID': appl.get('ApplicatorID'),
            'ApplicatorType': appl.get('ApplicatorType',),
        }
        appl_data = get_size(appl, appl_data)
    else:
        appl_data = None
    return appl_data


def get_control_point_data(field_ds: pydicom.Dataset) -> Dict[str, Any]:
    """Extract field settings from DICOM data for a given field.

    This method assumes a static field and only extracts settings from the first
        control point in the Control Point Sequence.
    Args:
        field_ds (pydicom.Dataset): The DICOM BeamSequence sub-dataset for a
            field within a plan DICOM dataset.
    Returns:
        initial_field_data (Dict[str, Any]): Field parameters.
    """
    control_point_seq = field_ds.ControlPointSequence
    control_point = control_point_seq[0]
    initial_field_data = {
        'CollimatorAngle': control_point.get('BeamLimitingDeviceAngle'),
        'DoseRate': control_point.get('DoseRateSet'),
        'GantryAngle': control_point.get('GantryAngle'),
        'Energy': control_point.get('NominalBeamEnergy'),
        'CouchAngle': control_point.get('PatientSupportAngle'),
        'Actual SSD': control_point.get('SourceToSurfaceDistance'),
        'Isocentre': control_point.get('IsocenterPosition')
    }
    return initial_field_data


#%% Combine DICOM Field Data
def get_field_data(ds: pydicom.Dataset) -> pd.DataFrame:
    """Extract all field related parameters settings for each field.

    Args:
        ds (pydicom.Dataset): The DICOM dataset for a plan.
    Returns:
        field_df (pd.DataFrame): Field parameters for all fields in the plan.
    """
    fields = list()
    for field_ds in ds.BeamSequence:
        # Get general field parameters and references for linking with
        # tolerance tables, setup, and MUs.
        field_data = {
            'BeamNumber': field_ds.get('BeamNumber'),
            'ToleranceTableNumber': field_ds.get('ReferencedToleranceTableNumber'),
            'PatientSetupNumber': field_ds.get('ReferencedPatientSetupNumber'),
            'FieldId': field_ds.get('BeamName'),
            'FieldType': field_ds.get('BeamType'),
            'RadiationType': field_ds.get('RadiationType'),
            'Weight': field_ds.get('FinalCumulativeMetersetWeight'),
            'Linac': field_ds.get('TreatmentMachineName'),
            'SetupField': field_ds.get('TreatmentDeliveryType'),
            'SAD': field_ds.get('SourceAxisDistance'),
            'NumberOfBlocks': field_ds.get('NumberOfBlocks'),
            'NumberOfBoli': field_ds.get('NumberOfBoli'),
            'NumberOfControlPoints': field_ds.get('NumberOfControlPoints'),
            'NumberOfWedges': field_ds.get('NumberOfWedges')
        }
        # Step through the field sub-datasets to extract relevant information
        initial_field_data = get_control_point_data(field_ds)
        field_data.update(initial_field_data)
        # Get applicator data. Assumes only one applicator per field.
        appl_data = get_applicator_info(field_ds)
        if appl_data:
            field_data.update(appl_data)
            # Get insert data. Assumes only one insert per field.
            block_data = get_block_info(field_ds)
            field_data.update(block_data)
            # Add the field dictionary to the list of fields
            fields.append(field_data)
    #convert list of dictionaries to a data frame
    field_df = pd.DataFrame(fields)
    return field_df


#%% Clean and Combine Field Data
def combine_field_tables(field_df, plan_mus, setup_ref, tolerance_ref):
    """Marge all field related parameters settings.

    SAD and Actual SSD are also re-scaled from mm to cm.
    Args:
        field_df (pd.DataFrame): Field parameters for all fields in the plan.
        plan_mus (pd.DataFrame): The MUs and field dose for each field in
            the plan.
        setup_ref (pd.DataFrame): The setup technique and patient orientation
            for a plan.
        tolerance_ref (pd.Series): The tolerance table name for each field in
            the plan.
    Returns:
        field_df (pd.DataFrame): Field parameters for all fields in the plan
            after merging MUs, Setup and Tolerance data.
    """
    field_df['Actual SSD'] = field_df['Actual SSD'] / 10
    field_df['SAD'] = field_df['SAD'] / 10
    field_df = field_df.merge(plan_mus, how="left", left_on='BeamNumber',
                              right_index=True)
    field_df = field_df.merge(setup_ref, how="left",
                              left_on='PatientSetupNumber',
                              right_index=True)
    field_df = field_df.merge(tolerance_ref, how="left",
                              left_on='ToleranceTableNumber',
                              right_index=True)
    field_df.drop(columns=['BeamNumber', 'ToleranceTableNumber',
                           'PatientSetupNumber'], inplace=True)
    return field_df

# Merge Field Data


def get_merged_field_data(ds: pydicom.Dataset) -> pd.DataFrame:
    """Extract and merge field related info from the dataset.

    Args:
        ds (pydicom.Dataset): The DICOM dataset for a plan.
    Returns:
        field_df (pd.DataFrame): Field parameters for all fields in the plan
            after merging MUs, Setup and Tolerance data.
    """
    tolerance_ref = get_tolerance_tables(ds)
    setup_ref = get_plan_setup(ds)
    plan_mus = get_mus(ds)
    field_df = get_field_data(ds)
    field_df = combine_field_tables(
        field_df, plan_mus, setup_ref, tolerance_ref)
    return field_df


#%% Block Coordinates Table
def get_block_coord(plan_df: pd.DataFrame) -> pd.DataFrame:
    """Extract the cutout coordinates for each field.

    Extracts the np.array block coordinates for each field, returning them as
        a separate DataFrame.  The returned dataFrame has a multi-level column
        index.  The top level is the plan ID, the second level is the field ID,
        and the lowest level is X or Y, the coordinate data pairs for the
        cutout.
    Args:
        plan_df (pd.DataFrame): Field parameters for all fields in all plans.
    Returns:
        block_coords (pd.DataFrame): The X & Y coordinate pairs for each insert
        in each plan.
    """
    block_coord_df = plan_df.loc['Coordinates', :]
    field_groups = block_coord_df.groupby(['PatientReference', 'PlanId', 'FieldId'])
    blk_grps = list()
    for name, group in field_groups:
        blk_grps.append(group.apply(
            {'X': lambda x: x[:, 0], 'Y': lambda x: x[:, 1]}))
    blk_grps_df = pd.concat(blk_grps)
    block_coords = blk_grps_df.apply(pd.Series).T
    block_coords = block_coords.reorder_levels([1, 2, 3, 0], axis='columns')
    block_coords.columns.names = ['PatientReference', 'PlanId', 'FieldId', 'Axis']
    return block_coords


#%% Load DICOM Files
def read_dicom_plan(plan_file: Path) -> pd.DataFrame:
    """Load a DICOM plan file and extract field data from it.

    Args:
        plan_file (Path): Full path to the DICOM Plan file.
    Returns:
        field_df (pd.DataFrame): Field parameters for all fields in the plan.
    """
    ds = pydicom.dcmread(plan_file)
    dicom_type = ds.Modality
    if 'RTPLAN' not in dicom_type:
        return None
    plan_name = ds.RTPlanLabel
    patient_id = ds.PatientID
    patient_name = str(ds.PatientName)
    patient_birth_date = ds.PatientBirthDate
    field_df = get_merged_field_data(ds)
    field_df['PlanId'] = plan_name
    field_df['PatientId'] = patient_id
    field_df['PatientName'] = patient_name
    field_df['PatientBirthDate'] = patient_birth_date
    field_df['PatientReference'] = (field_df.PatientName + " (" +
                                    field_df.PatientId + ")")
    return field_df


# Read all files
def get_plan_data(dicom_folder: Path) -> pd.DataFrame:
    """Load field data from all DICOM plan files in a directory.

    Args:
        dicom_folder (Path): Full path to a folder containing DICOM Plan files.
    Returns:
        plan_df (pd.DataFrame): Field parameters for all fields in all plans.
    """
    plan_files = [file for file in dicom_folder.glob('**/RP*.dcm')]
    plan_data = list()
    for plan_file in plan_files:
        field_df = read_dicom_plan(plan_file)
        plan_data.append(field_df)
    plan_df = pd.concat(plan_data)
    plan_df.set_index(['PatientReference', 'PlanId', 'FieldId'], inplace=True)
    plan_df = plan_df.T
    return plan_df


#%% Main
def main():
    """Run test with sample files.

    Returns:
        None.
    """
    # File Paths
    # TODO turn this into Unit Tests
    dicom_folder = Path.cwd()
    output_file_name = 'Electron Plan DICOM Info.xlsx'
    save_file = folder / output_file_name
    # Load DICOM Data
    plan_df = get_plan_data(dicom_folder)
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
