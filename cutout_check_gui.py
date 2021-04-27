"""Generate the CutoutCheck GUI.

Created on Sun Apr 25 11:51:36 2021

@author: Greg
"""
#%% Imports etc.
from pathlib import Path
import PySimpleGUI as sg

from Cutout_Analysis import analyze_cutout


#%%  Select File
def file_selection_frame(**file_paths):
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
        if isinstance(starting_path, str):
            initial_dir = Path.cwd()
            initial_dir = starting_path
        elif starting_path.is_dir():
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
        else:
            raise ValueError(f'{selection} is not a valid browser type')
        frame_k = file_k + '_frame'
        file_selector_frame = sg.Frame(
            title=frame_title, key=frame_k, layout=[
            [sg.InputText(key=file_k, default_text=initial_file), browse]]
            )
        return file_selector_frame

    file_selection_list = make_file_selection_list(**file_paths)
    file_frame_list = [[file_selector(**selection)]
                       for selection in file_selection_list]
    return file_frame_list


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


#%% Main
def main():
    # Directory Paths
    # data_path = Path(r'L:\temp\Plan Checking Temp')
    # dicom_folder = data_path / 'DICOM'
    # template_dir = Path(r'\\dkphysicspv1\e$\Gregs_Work\Plan Checking\Plan Check Tools\Templates')
    # template_dir = Path.cwd() / 'Electron Insert QA'

    # Test directory is current directory
    data_path = Path.cwd()
    dicom_folder = Path.cwd()
    template_dir = Path.cwd()
    output_file_name = 'CutOut Size Check Test.xlsx'

    # Default File Paths
    template_file_name = 'CutOut Size Check.xlsx'
    image_file_name = 'image2021-04-16-095423-1.jpg'
    default_file_paths = dict(
        dicom_folder = dicom_folder,
        image_file = data_path / image_file_name,
        template_path = template_dir / template_file_name,
        save_data_file = data_path / output_file_name
        )
    window = make_window(**default_file_paths)
    selected_file_paths = get_file_paths(window, default_file_paths)
    if selected_file_paths:
        analyze_cutout(**selected_file_paths)


if __name__ == '__main__':
    main()