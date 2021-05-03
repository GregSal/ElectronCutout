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


#%%  Select File
def file_selection_frame(**default_file_paths):
    """Generate a

    Args:
        **file_paths (TYPE): DESCRIPTION.
    Raises:
        ValueError: DESCRIPTION.
    Returns:
        TYPE: DESCRIPTION.
    """

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
    # Dummy Invisible Widget containing the file selector keys as default_values
    #Listbox(path_list,
    #    default_values = path_list,
    #    select_mode = sg.LISTBOX_SELECT_MODE_MULTIPLE,
    #    enable_events = False,
    #    bind_return_key = False,
    #    size = (None, None),
    #    disabled = False,
    #    auto_size_text = None,
    #    font = None,
    #    no_scrollbar = False,
    #    background_color = None,
    #    text_color = None,
    #    highlight_background_color = None,
    #    highlight_text_color = None,
    #    key = None,
    #    pad = None,
    #    tooltip = None,
    #    right_click_menu = None,
    #    visible = False,
    #    metadata = None)
    file_frame_list = [[file_selector(**selection)]
                       for selection in file_selection_list]
    return file_frame_list


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
    actions_list = [[sg.Button(**btn)]
                    for btn in action_buttons.values()]
    return actions_list


def next_action(window, btn_updates, btn_actions):
    for btn, updates in btn_updates.items():
        window[btn].update(**updates)
    done = False
    while not(done):
        event, parameters = window.read(timeout=200)
        if event == sg.TIMEOUT_KEY:
            continue
        if event in action_dict:
            action_method = action_dict[event]
            action_results = action_method(parameters)
            done = True
    window.close()
    return action_results


def set_file_paths(default_file_paths, parameters):
    path_list = list(default_file_paths.keys())
    file_paths = dict()
    if parameters:
        for path_name, default_path in default_file_paths.items():
            selected_path = parameters.get(path_name, default_path)
            file_paths[path_name] = selected_path
    return file_paths


def cancel_action(parameters):
    sg.popup_error('Cutout Analysis Canceled')
    return None


def select_file_paths(window, default_file_paths):
    btn_updates = {
        'Back': dict(disabled = True),
        'Next': dict(disabled = False,
                     button_text = 'Next')
        }
    btn_actions = {'Next': set_file_paths,
                   'Cancel': cancel_action}
    file_paths = next_action(window, btn_updates, btn_actions)
    return file_paths


def get_file_paths(window, default_file_paths):
    while True:
        event, parameters = window.read(timeout=200)
        if event == sg.TIMEOUT_KEY:
            continue
        if 'Submit' in event:
            break
        if 'Cancel' in event:
            sg.popup_error('Operation canceled')
            parameters = None
            break
    window.close()
    path_list = list(default_file_paths.keys())
    if parameters:
        file_paths = {path_name: Path(selected_path)
                      for path_name, selected_path in parameters.items()
                      if path_name in path_list}
    else:
        file_paths = None
    return file_paths


def make_window(**file_paths):
    file_frame_list = file_selection_frame(**file_paths)
    actions_list = [[sg.Submit(), sg.Cancel()]]
    window = sg.Window('Electron Cutout Check', finalize=True, resizable=True,
                       layout=[
        [sg.Column(file_frame_list, key='File Selection')],
        [sg.Column(actions_list, key='Actions')]
        ])
    for elm in window.element_list():
        elm.expand(expand_x=True)
        # TODO Add display Patient info
        # TODO add field selector to GUI
    return window

def select_field(block_coords: pd.DataFrame) -> Tuple[str]:
    """Select field for Aperture from list of available fields.

    Currently selects first field.  The aperture coordinates in the Selected
        field are used for the cutout dimensions.
    Args:
        block_coords (pd.DataFrame): Table with all fields containing
            apertures Column levels expected are: ['PlanId', 'FieldId', 'Axis'].
    Returns:
        selected_field (Tuple[str]): The PlanId and FieldId of the selected
            field as a tuple.
    """
    field_list = block_coords.columns.to_frame(index=True)
    patients = list(set(field_list['PatientReference']))
    selected_patient = patients[0]
    plans = list(set(field_list.loc[selected_patient,'PlanId']))
    selected_plan = plans[0]
    fields = list(set(field_list.loc[(selected_patient, selected_plan), 'FieldId']))
    selected_field = fields[0]
    return (selected_patient, selected_plan, selected_field)

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
    selected_file_paths = get_file_paths(window, default_file_paths)

    #if selected_file_paths:

    dicom_folder = selected_file_paths['dicom_folder']
    plan_df = get_plan_data(dicom_folder)
    block_coords = get_block_coord(plan_df)



    #### SElect plan and field
    selected_field = select_field(block_coords)
    ###

    insert_size = plan_df.at['ApplicatorOpening', selected_field]

    ## In Cutout Analysis
        #    analyze_cutout(**selected_file_paths)
    workbook = save_data(block_coords, plan_df, save_data_file, template_path)
    add_block_info(plan_df, block_coords, selected_field, workbook)
    show_cutout_info(image_file, insert_size, workbook)




if __name__ == '__main__':
    main()