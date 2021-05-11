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
from Cutout_Analysis import analyze_cutout


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
        field_selection_frame = [
            sg.Column(patient_text, key='Patient Info'),
             v_bar,
             sg.Column(selector_set, key='Selectors')
             ]
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

        path_list = list(default_file_paths.keys())
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
# TODO Perhaps have 1 main method that responds to Next, Back, Cancel, Finish
# Use status element to indicate where in workflow
def next_action(window, btn_updates, btn_actions):
    for btn, updates in btn_updates.items():
        window[btn].update(**updates)
    done = False
    while not(done):
        event, parameters = window.read(timeout=200)
        if event == sg.TIMEOUT_KEY:
            continue
        if event in btn_actions:
            action_method = btn_actions[event]
            action_results = action_method(parameters)
            done = True
    #window.close()
    return action_results


def cancel_action(*args, **kwargs):
    sg.popup_error('Cutout Analysis Canceled')
    return None

#%% File Selection
def set_file_paths(default_file_paths, parameters, event):
    path_list = list(default_file_paths.keys())
    file_paths = dict()
    if parameters:
        for path_name, default_path in default_file_paths.items():
            selected_path = parameters.get(path_name, default_path)
            file_paths[path_name] = Path(selected_path)
    return file_paths


def select_file_paths(window, default_file_paths):
    set_paths = partial(set_file_paths, default_file_paths)
    btn_updates = {
        'Back': dict(disabled = True),
        'Next': dict(disabled = False,
                     text = 'Next')
        }
    btn_actions = {'Next': set_paths,
                   'Cancel': cancel_action}
    file_paths = next_action(window, btn_updates, btn_actions)
    return file_paths


def build_field_options(plan_df, block_coords):
    field_options = block_coords.columns.to_frame(index=True)
    field_options = field_options.droplevel('Axis')
    field_options.drop(columns='Axis', inplace=True)
    field_options.drop_duplicates(inplace=True)
    patient_info = plan_df.T[['PatientId','PatientName', 'PatientBirthDate']]
    field_options = pd.concat([field_options, patient_info], axis='columns')
    return field_options


def load_dicom_plans(selected_file_paths):
    dicom_folder = selected_file_paths['dicom_folder']
    plan_df = get_plan_data(dicom_folder)
    block_coords = get_block_coord(plan_df)
    field_options = build_field_options(plan_df, block_coords)
    return block_coords, plan_df, field_options

#%% Field Selection
def select_field(window, field_options) -> Tuple[str]:
    """Select field for Aperture from list of available fields.

    Currently selects first field.  The aperture coordinates in the Selected
        field are used for the cutout dimensions.
    After selecting patient build restricted plan and field list from original list
    After selecting plan, build restricted patient and field list starting with current patient list
    After selecting field, build restricted patient and plan list starting with current patient and plan list
    After building new patient list populate patient text and default patient from first in list
    After building new plan list populate default plan from first in list
    After building new field list populate default field from first in list
    Args:
        block_coords (pd.DataFrame): Table with all fields containing
            apertures Column levels expected are: ['PlanId', 'FieldId', 'Axis'].
    Returns:
        selected_field (Tuple[str]): The PlanId and FieldId of the selected
            field as a tuple.
    """
    def update_field_selection(window, selection_options, selection=None, selector=None):
        def update_selection(selection_options, source_col, selector):
            selection_list = list(set(selection_options[source_col]))
            window[selector].update(values=selection_list,
                                    value=selection_list[0],
                                    disabled=False)
            return selection_list

        def update_patient(window, selection_options):
            template_rows = ['{PatientName:<16s}',
                                '{PatientId:<16s}',
                                '{PatientBirthDate:<16s}']
            pt_template = '\n'.join(template_rows)

            first_field = selection_options.iloc[0].to_dict()
            pt_text = pt_template.format(**first_field)
            window['PatientText'].update(value=pt_text)
            window.refresh()

        source_index = {
            'PatientSelector': 'PatientReference',
            'PlanSelector': 'PlanId',
            'FieldSelector': 'FieldId'
            }
        if selection:
            selection_options = selection_options.xs(selection,
                                                     level=source_index[selector])
        else:
            # Always keep all patients available for selection
            update_selection(selection_options, 'PatientReference', 'PatientSelector')
        update_selection(selection_options, 'PlanId', 'PlanSelector')
        update_selection(selection_options, 'FieldId', 'FieldSelector')
        update_patient(window, selection_options)
        window.refresh()
        return selection_options

    update_field_selection(window, field_options)

    set_patient = partial(update_field_selection, window, 'PatientSelector', field_options)
    set_plan = partial(update_field_selection, window, 'PlanSelector', selection_options)
    btn_updates = {
        'Back': dict(disabled = True),
        'Next': dict(disabled = False,
                     text = 'Next')
        }
    done=False
    full_selection = None
    selection_options = field_options
    while not(done):
        event, parameters = window.read(timeout=200)
        if event in ['PatientSelector', 'PlanSelector', 'FieldSelector']:
            selection = parameters[event]
            if event in 'PatientSelector':
                selection_options = field_options  # Reset selections
            selection_options = update_field_selection(window, event, selection_options,
                                           selection)
        elif event in 'Cancel':
            done = True
        elif event in 'Next':
            full_selection = (parameters['PatientSelector'],
                              parameters['PlanSelector'],
                              parameters['FieldSelector'])
            done = True
    return full_selection
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
    selected_file_paths = select_file_paths(window, default_file_paths)

    #if selected_file_paths:

    block_coords, plan_df, field_options = load_dicom_plans(selected_file_paths)
    selected_field = test_select_field(window, field_options)


    #### SElect plan and field
    selected_field = select_field(window, plan_df, block_coords)
    ###

    insert_size = plan_df.at['ApplicatorOpening', selected_field]

    ## In Cutout Analysis
        #    analyze_cutout(**selected_file_paths)
    workbook = save_data(block_coords, plan_df, save_data_file, template_path)
    add_block_info(plan_df, block_coords, selected_field, workbook)
    show_cutout_info(image_file, insert_size, workbook)




if __name__ == '__main__':
    main()