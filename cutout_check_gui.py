"""Generate the CutoutCheck GUI.

Created on Sun Apr 25 11:51:36 2021

@author: Greg
"""
#%% Imports etc.
from pathlib import Path
from functools import partial
from typing import Tuple, List

import PySimpleGUI as sg
import pandas as pd
from load_dicom_e_plan import get_plan_data
from load_dicom_e_plan import get_block_coord
from Cutout_Analysis import show_cutout_info, add_block_info, save_data


#%% File Selection
def set_file_paths(default_file_paths, parameters):
    file_paths = dict()
    if parameters:
        for path_name, default_path in default_file_paths.items():
            selected_path = parameters.get(path_name, default_path)
            file_paths[path_name] = Path(selected_path)
    return file_paths


#%% Field Selection
def load_dicom_plans(selected_file_paths):
    dicom_folder = selected_file_paths['dicom_folder']
    plan_df = get_plan_data(dicom_folder)
    block_coords = get_block_coord(plan_df)
    field_options = build_field_options(plan_df, block_coords)
    return block_coords, plan_df, field_options


def build_field_options(plan_df, block_coords):
    field_options = block_coords.columns.to_frame(index=True)
    field_options = field_options.droplevel('Axis')
    field_options.drop(columns='Axis', inplace=True)
    field_options.drop_duplicates(inplace=True)
    patient_info = plan_df.T[['PatientId','PatientName', 'PatientBirthDate']]
    field_options = pd.concat([field_options, patient_info], axis='columns')
    return field_options


#%% Build GUI
def make_window(**file_paths):
    def set_action_buttons():
        action_buttons = {
            'Back': dict(
                button_text = 'Back',
                disabled = True,
                enable_events = True,
                key = 'Back'
                ),
            'Next': dict(
                button_text = 'Next',
                disabled = False,
                enable_events = True,
                key = 'Next'
                ),
            'Cancel': dict(
                button_text = 'Cancel',
                disabled = False,
                enable_events = True,
                key = 'Cancel'
                )
            }
        actions_list = [[sg.Button(**btn)
                        for btn in action_buttons.values()]]
        return actions_list

    def set_field_selection():
        patient_text = [[
            sg.Text('Name\nID\nBirth Date'),
            sg.Text(key='PatientText',size=(16,3))
            ]]
        selector_set = [
            [sg.Combo([], key='PatientSelector', size=(30,1),
                      enable_events=True, disabled=True)],
            [sg.Combo([], key='PlanSelector', size=(30,1),
                      enable_events=True, disabled=True)],
            [sg.Combo([], key='FieldSelector', size=(30,1),
                      enable_events=True, disabled=True)]
            ]
        v_bar = sg.HorizontalSeparator(key='V_Bar')
        field_selection_frame = [[
            sg.Column(patient_text, key='Patient Info'),
             v_bar,
             sg.Column(selector_set, key='Selectors')
             ]]
        return field_selection_frame

    def file_selection_frame(**default_file_paths):
        def make_file_selection_list(dicom_folder,
                                     image_file,
                                     template_path,
                                     save_data_file):
            file_selection_list = [
                dict(frame_title='Select DICOM directory to scan',
                     file_k='dicom_folder',
                     selection='dir',
                     starting_path=dicom_folder
                     ),
                dict(frame_title='Select Cutout Image File to Load',
                     file_k='image_file',
                     selection='read file',
                     starting_path=image_file,
                     file_type=(('Image Files', '*.jpg'),)
                     ),
                dict(frame_title='CutOut Check Template File',
                     file_k='template_path',
                     starting_path=template_path,
                     selection='read file',
                     file_type=(('Excel Files', '*.xlsx'),)
                     ),
                dict(frame_title='Save CutOut Check As:',
                     file_k='save_data_file',
                     starting_path=save_data_file,
                     selection='save file',
                     file_type=(('Excel Files', '*.xlsx'),)
                     )
            ]
            return file_selection_list

        def file_selector(selection, file_k, frame_title,
                            starting_path=Path.cwd(), file_type=None):
            try:
                starting_path = Path(starting_path)
            except TypeError:
                starting_path=Path.cwd()
            if starting_path.is_dir():
                initial_dir = str(starting_path)
                initial_file = ''
            else:
                initial_dir = str(starting_path.parent)
                initial_file = str(starting_path)

            if 'read file' in selection:
                browse = sg.FileBrowse(initial_folder=initial_dir,
                                        file_types=file_type)
            elif 'save file' in selection:
                browse = sg.FileSaveAs(initial_folder=initial_dir,
                                        file_types=file_type)
            elif 'read files' in selection:
                browse = sg.FilesBrowse(initial_folder=initial_dir,
                                        file_types=file_type)
            elif 'dir' in selection:
                browse = sg.FolderBrowse(initial_folder=initial_dir)
                initial_file = initial_dir
            else:
                raise ValueError(f'{selection} is not a valid browser type')
            frame_k = file_k + '_frame'
            file_selector_frame = sg.Frame(
                title=frame_title, key=frame_k, layout=[
                [sg.InputText(key=file_k, default_text=initial_file), browse]]
                )
            return file_selector_frame

        file_selection_list = make_file_selection_list(**default_file_paths)
        file_frame_list = [[file_selector(**selection)]
                           for selection in file_selection_list]
        return file_frame_list

    field_selection_frame = set_field_selection()
    file_frame_list = file_selection_frame(**file_paths)
    actions_list = set_action_buttons()
    window = sg.Window('Electron Cutout Check',
                       finalize=True, resizable=True,
                       layout=[
        [sg.Column(field_selection_frame, key='Field Selection')],
        [sg.Column(file_frame_list, key='File Selection')],
        [sg.Column(actions_list, key='Actions')]
        ])
    for elm in window.element_list():
        elm.expand(expand_x=True)
    window['V_Bar'].expand(expand_y=True)
        # TODO don't expand button elements
    return window


#%% GUI Methods
def cancel_action(*args, **kwargs):
    sg.popup_error('Cutout Analysis Canceled')
    return None


def update_widgets(window, elm_updates):
    for btn, updates in elm_updates.items():
        window[btn].update(**updates)


def update_field_selection(window, selection_options, selector=None, selection=None):
    def update_selection(selection_options, source_col, selector, default):
        selection_list = list(set(selection_options[source_col]))
        window[selector].update(values=selection_list,
                                value=default,
                                disabled=False)

    def update_patient(window, selection_options, default):
        template_rows = ['{PatientName:<16s}',
                            '{PatientId:<16s}',
                            '{PatientBirthDate:<16s}']
        pt_template = '\n'.join(template_rows)
        pt_text = pt_template.format(**default)
        window['PatientText'].update(value=pt_text)
        window.refresh()

    source_index = {
        'PatientSelector': 'PatientReference',
        'PlanSelector': 'PlanId',
        'FieldSelector': 'FieldId'
        }
    if selection:
        selection_options = selection_options.xs(
            selection, level=source_index[selector])
    else:
        # # This is intended to initialize the patients selector but not
        # update it with a reduced list based on other selections.
        # This keeps all patients available for selection.
        first_field = selection_options.iloc[0].to_dict()
        default_patient = first_field['PatientReference']
        update_selection(selection_options, 'PatientReference', 'PatientSelector', default_patient)

    first_field = selection_options.iloc[0].to_dict()
    update_selection(selection_options, 'PlanId', 'PlanSelector', first_field['PlanId'])
    update_selection(selection_options, 'FieldId', 'FieldSelector', first_field['FieldId'])
    update_patient(window, selection_options, first_field)
    window.refresh()
    return selection_options


def main_actions(window, default_file_paths):
    """Contour Analysis steps:

    1) Select DICOM folder
        Read all DICOM Plan files in the folder
    2) Select the field for Aperture from list of available fields.
        Choose from Patient -> Plan -> Field
        Resetting Patient resets Plan and Field Options
        Defaults to first field found.
    3) Select Cutout Image
        Must be JPEG format
    4) Set Report File Name
        Must be .xlsx type
        Will overwrite existing file
    5) Generate Analysis Report

    window = make_window(**default_file_paths)
    selected_file_paths = select_file_paths(window, default_file_paths)
    # TODO Start with simple linear flow with disabled Back
    # Perhaps have 1 main method that responds to Next, Back, Cancel, Finish
    # Use status element to indicate where in work flow
    #if selected_file_paths:

    block_coords, plan_df, field_options = load_dicom_plans(selected_file_paths)
    selected_field = test_select_field(window, field_options)

    Args:
        block_coords (pd.DataFrame): Table with all fields containing
            apertures Column levels expected are: ['PlanId', 'FieldId', 'Axis'].
    Returns:
        selected_field (Tuple[str]): The PlanId and FieldId of the selected
            field as a tuple.
    """
    #%% 1) Select DICOM folder
    dcm_fldr_updates = {
        #'dicom_folder_frame': dict(visible = True),
        'dicom_folder': dict(disabled = False),
        #'image_file_frame': dict(visible = True),
        'image_file': dict(disabled = False),
        #'template_path_frame': dict(visible = True),
        'template_path': dict(disabled = True),
        #'save_data_file_frame': dict(visible = True),
        'save_data_file': dict(disabled = False),
        'PatientSelector': dict(disabled=True, values=[], value=''),
        'PlanSelector': dict(disabled=True, values=[], value=''),
        'FieldSelector': dict(disabled=True, values=[], value=''),
        'Back': dict(disabled = True),
        'Next': dict(disabled = False, text = 'Next')
        }
    update_widgets(window, dcm_fldr_updates)
    done = False
    while not(done):
        event, parameters = window.read(timeout=200)
        if event == sg.TIMEOUT_KEY:
            continue
        if event in 'Next':
            selected_file_paths = set_file_paths(default_file_paths, parameters)
            done = True
        elif event in 'Cancel':
            cancel_action()
            selected_file_paths = None
            done = True
    if not selected_file_paths:
        return None

    ########################
    #%% Load DICOM Plans
    block_coords, plan_df, field_options = load_dicom_plans(selected_file_paths)


    #%% 2) Select the field for Aperture from list of available fields.
    #    Choose from Patient -> Plan -> Field
    #    Resetting Patient resets Plan and Field Options
    #    Defaults to first field found.
    # Currently selects first field.  The aperture coordinates in the Selected
    #     field are used for the cutout dimensions.
    # After selecting patient build restricted plan and field list from original list
    # After selecting plan, build restricted patient and field list starting with current patient list
    # After selecting field, build restricted patient and plan list starting with current patient and plan list
    # After building new patient list populate patient text and default patient from first in list
    # After building new plan list populate default plan from first in list
    # After building new field list populate default field from first in list

    update_field_selection(window, field_options)
    fld_updates = {
        'dicom_folder_frame': dict(visible = False),
        'dicom_folder': dict(disabled = True),
        'Back': dict(disabled = False),
        'Next': dict(disabled = False, text = 'Next')
        }
    update_widgets(window, fld_updates)
    done=False
    parameters = None
    selected_field = None
    selection_options = field_options
    while not(done):
        event, parameters = window.read(timeout=200)
        if event == sg.TIMEOUT_KEY:
            continue
        if event in ['PatientSelector', 'PlanSelector', 'FieldSelector']:
            selection = parameters[event]
            if event in 'PatientSelector':
                selection_options = field_options  # Reset selections
            selection_options = update_field_selection(
                window, selection_options,
                selector=event, selection=parameters[event])
        elif event in 'Cancel':
            done = True
        elif event in 'Next':
            selected_field = (parameters['PatientSelector'],
                              parameters['PlanSelector'],
                              parameters['FieldSelector'])
            done = True
    if not selected_field:
        return None

    #%% 4) Set Report File Name
    selected_file_paths = set_file_paths(default_file_paths, parameters)
    save_data_file= selected_file_paths['save_data_file']
    template_path= selected_file_paths['template_path']

    #%% Save Cutout Info
    insert_size = plan_df.at['ApplicatorOpening', selected_field]
    workbook = save_data(block_coords, plan_df, save_data_file, template_path)
    add_block_info(plan_df, block_coords, selected_field, workbook)


    #%% 3) Select Cutout Image
    selected_file_paths = set_file_paths(default_file_paths, parameters)
    image_file = selected_file_paths['image_file']

    #%% 5) Generate Analysis Report
    show_cutout_info(image_file, insert_size, workbook)
    return None


################################################################
#%% Main


def main():
    # Directory Paths
    # data_path = Path(r'L:\temp\Plan Checking Temp')
    # dicom_folder = data_path / 'DICOM'
    # template_dir = Path(r'\\dkphysicspv1\e$\Gregs_Work\Plan Checking\Plan Check Tools\Templates')
    # template_dir = Path.cwd() / 'Electron Insert QA'

    # Test directory is current directory
    data_path = Path.cwd() / 'Test Files'
    dicom_folder = Path.cwd() / 'Test Files'
    template_dir = Path.cwd() / 'Template Files'
    output_path = Path.cwd() / 'Output'
    output_file_name = 'CutOut Size Check Test.xlsx'

    # Default File Paths
    template_file_name = 'CutOut Size Check.xlsx'
    image_file_name = 'CutoutTest3.jpg'
    default_file_paths = dict(
        dicom_folder = dicom_folder,
        image_file = data_path / image_file_name,
        template_path = template_dir / template_file_name,
        save_data_file = output_path / output_file_name
        )
    window = make_window(**default_file_paths)
    main_actions(window, default_file_paths)
    window.close()


if __name__ == '__main__':
    main()
