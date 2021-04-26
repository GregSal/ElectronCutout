"""Generate the CutoutCheck GUI.

Created on Sun Apr 25 11:51:36 2021

@author: Greg
"""
#%% Imports etc.
from pathlib import Path
import PySimpleGUI as sg


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

    def make_file_selection_list(dicom_starting_path, template_file_path,
                                 image_file_starting_path,
                                 save_data_file_name='plan_check_data.xlsx',
                                 save_form_file_name='PlanCheck.xlsx'):
        file_selection_list = [
            dict(frame_title='Select DICOM directory to scan',
                 file_k='printout_file',
                 selection='dir',
                 starting_path=dicom_starting_path
                 ),
            dict(frame_title='Select Coutout Image File to Load',
                 file_k='image_file',
                 selection='read file',
                 starting_path=image_file_starting_path,
                 file_type=(('Image Files', '*.jpg'),)
                 ),
            dict(frame_title='CutOut Check Template File',
                 file_k='template_file',
                 starting_path=template_file_path,
                 selection='read file',
                 file_type=(('Excel Files', '*.xlsx'),)
                 ),
            dict(frame_title='Save CutOut Check As:',
                 file_k='save_data_file',
                 starting_path=save_data_file_name,
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
        file_selector_frame = sg.Frame(title=frame_title, layout=[
            [sg.InputText(key=file_k, default_text=initial_file), browse]])
        return file_selector_frame

    file_selection_list = make_file_selection_list(**file_paths)
    file_frame_list = [[file_selector(**selection)]
                       for selection in file_selection_list]
    return file_frame_list
