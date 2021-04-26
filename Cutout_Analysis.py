"""Analyze Electron Cutout Shape.

Created on Wed Feb 10 21:24:34 2021

@author: Greg
"""
#%%  Imports
import math
from pathlib import Path
from statistics import mean
from typing import Tuple, List
import imageio
import numpy as np
import pandas as pd
import xlwings as xw
from scipy import ndimage
from skimage import measure
from shapely.geometry import Polygon
from load_dicom_e_plan import get_plan_data
from load_dicom_e_plan import get_block_coord


#%%  Scale Factors; Used as global variables.
in_scale = 72.0  # inches to Pixels conversion
cm_scale = in_scale / 2.54  # cm to Pixels conversion


#%% This section contains functions that enter data into the spreadsheet.
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
    field_groups = block_coords.groupby(level=['PlanId', 'FieldId'],
                                        axis='columns')
    selected_field = list(field_groups.groups)[0]
    return selected_field


def save_data(block_coords: pd.DataFrame, plan_df: pd.DataFrame,
              save_file: Path, template_path: Path) -> xw.Book:
    """Load the template spreadsheet and save a copy containing plan data.

    Args:
        block_coords (pd.DataFrame): The x, y coordinates of the aperture.
        plan_df (pd.DataFrame): Plan Parameters obtained from the DICOM File.
        save_file (Path): Path to use for saving the filled template.
        template_path (Path): Path to the Excel template.
    Returns:
        workbook (xw.Book): Excel workbook containing the data.
    """
    def get_workbook(save_file: Path, template_path: Path) -> xw.Book:
        """Load the template spreadsheet and save a copy.

        Args:
            save_file (Path): Path to use for saving the filled template.
            template_path (Path): Path to the Excel template.
        Returns:
            workbook (xw.Book): Excel workbook containing the data.
        """
        workbook = xw.Book(template_path)
        workbook.save(save_file)
        return workbook

    def add_plan_data(plan_df: pd.DataFrame, workbook: xw.Book):
        """Store the plan data in the spreadsheet.

        Args:
            plan_df (pd.DataFrame): Plan Parameters obtained from DICOM File.
            workbook (xw.Book): Excel workbook containing the data.
        Returns:
            None.
        """
        plan_data_sheet = workbook.sheets.add('Plan Data')
        plan_data_sheet.range('A1').value = plan_df
        plan_data_sheet.autofit()

    def add_field_parameters(plan_df: pd.DataFrame, workbook: xw.Book):
        """Store the Field parameter data in the spreadsheet.

        Args:
            plan_df (pd.DataFrame): Plan Parameters obtained from DICOM File.
            workbook (xw.Book): Excel workbook containing the data.
        Returns:
            None.
        """
        prm_sht = workbook.sheets['CutOut Parameters']
        prms = ['RadiationType', 'SetupTechnique', 'ToleranceTable', 'Linac',
                'Energy', 'GantryAngle', 'ApplicatorID', 'AccessoryCode',
                'BlockTrayID', 'InsertCode', 'BlockType', 'MaterialID',
                'BlockDivergence', 'BlockName', 'SourceToBlockTrayDistance']
        prm_sht.range('B1').options(pd.DataFrame, header=True,
                                    index=False).value = plan_df.loc[prms, :]

    workbook = get_workbook(save_file, template_path)
    add_plan_data(plan_df, workbook)
    add_field_parameters(plan_df, workbook)
    workbook.save(save_file)
    return workbook


def add_block_info(block_coords: pd.DataFrame, selected_field: Tuple[str],
                   workbook: xw.Book):
    """Store aperture and SSD data in the spreadsheet.

    Args:
        block_coords (pd.DataFrame): Table with apertures for all fields.
        selected_field (Tuple[str]): The PlanId and FieldId index of the
            selected field.
        workbook (xw.Book): Excel workbook containing the data.
    Returns:
        None.
    """

    def add_block_coordinates(block_coords: pd.DataFrame,
                              selected_field: Tuple[str], workbook: xw.Book):
        """Store aperture coordinates.

        Add aperture coordinates for the selected field to the CutOut
            Coordinates table for plotting.
        Args:
            block_coords (pd.DataFrame): Table with apertures for all fields.
            selected_field (Tuple[str]): The PlanId and FieldId index of the
                selected field.
            workbook (xw.Book): Excel workbook containing the data.
        Returns:
            coords (pd.DataFrame): The x,y coordinates for the aperture.
        """
        coords = block_coords.loc[:, selected_field]
        coords_sheet = workbook.sheets['CutOut Coordinates']
        coords_sheet.range('A3').options(pd.DataFrame, header=False,
                                         index=False).value = coords
        return coords

    def insert_ssd(plan_df: pd.DataFrame, selected_field: Tuple[str],
                   workbook: xw.Book):
        """Store SSD from selected_field in the spreadsheet.

        Args:
            plan_df (pd.DataFrame): Plan Parameters obtained from DICOM File.
            selected_field (Tuple[str]): The PlanId and FieldId index of the
                selected field.
            workbook (xw.Book): Excel workbook containing the data.
        Returns:
            None.
        """
        ssd = plan_df.at['Actual SSD', selected_field]
        ssd_range = workbook.names['SSD'].refers_to_range
        ssd_range.value = ssd

    def insert_applicator_size(plan_df, selected_field, workbook):
        """Store applicator size from selected_field in the spreadsheet.

        Args:
            plan_df (pd.DataFrame): Plan Parameters obtained from DICOM File.
            selected_field (Tuple[str]): The PlanId and FieldId index of the
                selected field.
            workbook (xw.Book): Excel workbook containing the data.
        Returns:
            None.
        """
        insert_size_range = workbook.names['Insert_Size'].refers_to_range
        insert_size = plan_df.at['ApplicatorOpening', selected_field]
        insert_size_range.value = insert_size

    def add_cutout_dimensions(coords: pd.DataFrame, workbook: xw.Book):
        """Calculate and store applicator shape parameters in the spreadsheet.

        Calculates Equivalent Square for aperture and x, y extent of
            aperture.
        Args:
            coords (pd.DataFrame): The x,y coordinates for the aperture.
            workbook (xw.Book): Excel workbook containing the data.
        Returns:
            None.
        """
        apperature = Polygon(np.array(coords))
        cutout_area = apperature.area
        cutout_perim = apperature.length
        cutout_eq_sq = 4 * cutout_area / cutout_perim
        cutout_extent = max([abs(x) for x in apperature.bounds])

        cutout_area_range = workbook.names['Cutout_Area'].refers_to_range
        cutout_area_range.value = cutout_area
        cutout_perim_range = workbook.names['Cutout_Perimeter'].refers_to_range
        cutout_perim_range.value = cutout_perim
        cutout_eq_sq_range = workbook.names['Cutout_Eq._Sq.'].refers_to_range
        cutout_eq_sq_range.value = cutout_eq_sq
        cutout_extent_range = workbook.names['Cutout_Extent'].refers_to_range
        cutout_extent_range.value = cutout_extent

    coords = add_block_coordinates(block_coords, selected_field, workbook)
    insert_ssd(plan_df, selected_field, workbook)
    insert_applicator_size(plan_df, selected_field, workbook)
    add_cutout_dimensions(coords, workbook)


def scale_cutout_graph(insert_size: int, image_sheet: xw.Sheet) -> xw.Chart:
    """Set the size and scale of the graph to match the applicator.

    Args:
        insert_size (int): The size of the applicator used.
            Can be one of {6, 10, 15, 20, 25}
        image_sheet (xw.Sheet): The worksheet containing the plot of the cutout.
    Returns:
        outline_graph (xw.Chart): The Excel plot of the aperture opening.
    """
    outline_graph = image_sheet.charts['Outline']
    # Set the graph size to match that of the applicator.
    outline_graph.width = cm_scale * insert_size
    outline_graph.height = cm_scale * insert_size
    # Set the graph max and min limits to match that of the applicator.
    outline_graph.api[1].Axes().Item(1).MinimumScale = -insert_size / 2
    outline_graph.api[1].Axes().Item(1).MaximumScale = insert_size / 2
    outline_graph.api[1].Axes().Item(2).MinimumScale = -insert_size / 2
    outline_graph.api[1].Axes().Item(2).MaximumScale = insert_size / 2
    return outline_graph


#%% Image Manipulation Functions
def get_image_size(cutout_image):
    """Get the height, width and resolution of the Scanned cutout image.

    Args:
        cutout_image (imageio image): The image and meta-data for the scanned
            cutout image.
    Returns:
        height (float): The height of the image in inches.
        width (float): The width of the image in inches.
        dpi (int): The resolution of the image in dots per inch.
    """
    dpi = np.array(cutout_image.meta['dpi'])
    image_size = cutout_image.shape / dpi
    height = image_size[0] * in_scale
    width = image_size[1] * in_scale
    return height, width, dpi


def find_outline(cutout_image, dpi):
    """Identify the external outline of the insert.

    The external outline is the metal frame around the insert.
    Args:
        cutout_image (imageio image): The image and meta-data for the scanned
            cutout image.
        dpi (int): The resolution of the image in dots per inch.
    Returns:
        insert_outline (np.array): x,y coordinates approximating the outside
            extent of the Cerrobend.
        insert_limits (np.array of size 4): the maximum and minimum extent of
            the Cerrobend in the x and Y directions. The order of the values is
            as follows:
                    [x_min (image top), y_min (image left),
                     x_max (image bottom), y_max (image right)]
    """
    # apply a median filter to reduce the noise, but keep the edge locations.
    med_denoise = ndimage.median_filter(cutout_image, 10)
    # Generate contours at a threshold just above black (20)
    # The largest contour will be the page size
    # The second largest will be the insert outline including encoding strip
    contours = measure.find_contours(med_denoise, 20)
    contours = sorted(contours, key=len, reverse=True)  # Sort by size
    insert_outline = contours[1] / dpi  # Select the second largest contour.
    insert_shape = Polygon(insert_outline)
    # Removed encoder tab at the top to get just insert
    encoder = [9 / 25.4, 0, 0, 0]  # Encoder strip is 9 mm height
    insert_limits = np.array(insert_shape.bounds) + encoder
    return insert_outline, insert_limits


def add_cutout_image(image_file: Path, image_sheet: xw.Sheet,
                     height: float, width: float) -> xw.Picture:
    """Insert the scanned cutout image into the spreadsheet.

    Args:
        image_file (Path): Full path to the scanned cutout image file.
        image_sheet (xw.Sheet): The worksheet containing the plot of the cutout.
        height (float): The height of the image in inches.
        width (float): The width of the image in inches.
    Returns:
        cutout_shape (xw.Picture): The Picture object in the spreadsheet.
    """
    cutout_shape = image_sheet.pictures.add(image_file, name="Cutout")
    # Set the image size to 100% scale
    cutout_shape.width = width
    cutout_shape.height = height
    # Send the picture to the back layer so that the graph can appear on
    # top of it.
    cutout_shape.api.ShapeRange.ZOrder(1)
    return cutout_shape


def crop_cutout_image(insert_limits: np.array, cutout_shape: xw.Picture,
                      height: float, width: float, pic_location: List[int]):
    """Crop the cutout image with a margin around the insert.

    Args:
        insert_limits (np.array of size 4): the maximum and minimum extent of
            the Cerrobend in the x and Y directions. The order of the values is
            as follows:
                    [x_min (image top), y_min (image left),
                     x_max (image bottom), y_max (image right)]
        cutout_shape (xw.Picture): The Picture object in the spreadsheet.
        height (float): The height of the image in inches.
        width (float): The width of the image in inches.
    Returns:
        None.
    """
    # Set the margin around the cutout
    margin = np.array([-1, -1, 1, 1]) * 0.5  # 1/2" margin
    # Calculate the crop sizes
    crop_size = insert_limits + margin
    crop_dim = np.int_(crop_size * in_scale)
    top, left, bottom, right = crop_dim
    # Do the image cropping
    cutout_shape.api.ShapeRange.PictureFormat.CropTop = top
    cutout_shape.api.ShapeRange.PictureFormat.CropLeft = left
    cutout_shape.api.ShapeRange.PictureFormat.CropBottom = height - bottom
    cutout_shape.api.ShapeRange.PictureFormat.CropRight = width - right
    # Move the image to the desired spot
    cutout_shape.top = pic_location[0]
    cutout_shape.left = pic_location[1]


def set_arrows(insert_size: int, mid_point: np.array, image_sheet: xw.Sheet):
    """Set the cross-hair arrows to match the applicator.

    Args:
        insert_size (int): The size of the applicator used.
            Can be one of {6, 10, 15, 20, 25}
        mid_point (np.array of size 2): The center point of the cropped cutout
            image. Has the form [height, width]
        image_sheet (xw.Sheet): The worksheet containing the plot of the cutout.
    Returns:
        None.
    """
    # Identify the Arrow Shapes
    up_arrow = image_sheet.shapes['UpArrow']
    horz_arrow = image_sheet.shapes['HorzArrow']
    #Set the arrow size to match the applicator size
    up_arrow.height = cm_scale * insert_size
    horz_arrow.width = cm_scale * insert_size
    #Make the arrows orthogonal
    up_arrow.width = 0
    horz_arrow.height = 0
    #Center the arrows on the middle of the cutout image
    up_arrow.left = mid_point[1]
    horz_arrow.top = mid_point[0]
    up_arrow.top = mid_point[0] - cm_scale * insert_size / 2
    horz_arrow.left = mid_point[1] - cm_scale * insert_size / 2
    # combine the arrows to form a cross hair
    # Align Center
    image_sheet.shapes.api.Range(['UpArrow', 'HorzArrow']).Align(1, 0)
    # Align middle
    image_sheet.shapes.api.Range(['UpArrow', 'HorzArrow']).Align(4, 0)
    # Group the arrows
    image_sheet.shapes.api.Range(['UpArrow', 'HorzArrow']).Group()


def rotate_image(insert_outline: np.array, insert_limits, cutout_shape):
    """Rotate the cutout image so that the insert is square on the page.

    Args:
        insert_outline (np.array): x,y coordinates approximating the outside
            extent of the Cerrobend.
        insert_limits (np.array of size 4): the maximum and minimum extent of
            the Cerrobend in the x and Y directions. The order of the values is
            as follows:
                    [x_min (image top), y_min (image left),
                     x_max (image bottom), y_max (image right)]
        cutout_shape (xw.Picture): The Picture object in the spreadsheet.
    Returns:
        None.
    """
    def find_angle(x_selection: List[bool], y_selection: List[bool]) -> float:
        """Determine the angle from the slope of one of the insert edges.

        Args:
            x_selection (List[bool]): A boolean index to the desired x range of
                the insert contour.
            y_selectionw (List[bool]): A boolean index to the desired y range of
                the insert contour.
        Returns:
            angle (float): The angle of the slope in degrees.
        """
        # Extract the edge from the insert contour.
        fit_selection = x_selection * y_selection
        selected_outline = insert_outline[fit_selection, :]
        # apply a linear fit to the line of the edge.
        fit_x = selected_outline[:, 0]
        fit_y = selected_outline[:, 1]
        p_fit = np.polyfit(fit_x, fit_y, deg=1)
        # Calculate the angle of the slope from the fit.
        angle = math.degrees(math.atan(p_fit[0]))
        return angle

    # Set a margin around the insert contour to avoid corner effects
    margin = np.array([-1, -1, 1, 1]) * 0.5  # 1/2" margin
    lim_x_low, lim_y_low, lim_x_hi, lim_y_hi = (insert_limits - margin)
    # Identify the different edges
    outline_x = insert_outline[:, 0]
    x_selection = (outline_x > lim_x_low) * (outline_x < lim_x_hi)
    outline_y = insert_outline[:, 1]
    y_low = (outline_y < lim_y_low)
    y_hi = (outline_y > lim_y_hi)
    #Get the average angle
    angle = mean([find_angle(x_selection, y_low),
                  find_angle(x_selection, y_hi)])
    #Rotate the image
    cutout_shape.api.ShapeRange.Rotation = angle


def show_cutout_info(image_file: Path, insert_size: int, workbook: xw.Book):
    """Compare the insert image with the cutout shape.

    Args:
        image_file (Path): Full path to the scanned cutout image file.
        insert_size (int): The size of the applicator used.
            Can be one of {6, 10, 15, 20, 25}
        workbook (xw.Book): Excel workbook containing the data.
    Returns:
        None.
    """
    image_sheet = workbook.sheets['CutOut Image']
    image_sheet.activate()
    # Set the location for the cutout image.
    pic_location = [0, 0]  # Top, Left in pixels
    outline_graph = scale_cutout_graph(insert_size, image_sheet)
    # TODO center the graph over the cutout image
    cutout_image = imageio.imread(image_file)
    height, width, dpi = get_image_size(cutout_image)
    insert_outline, insert_limits = find_outline(cutout_image, dpi)
    cutout_shape = add_cutout_image(image_file, image_sheet, height, width)
    crop_cutout_image(insert_limits, cutout_shape, height, width, pic_location)
    rotate_image(insert_outline, insert_limits, cutout_shape)
    mid_point = np.array([cutout_shape.height,
                          cutout_shape.width]) / 2 + pic_location
    set_arrows(insert_size, mid_point, image_sheet)
    # TODO Move the arrows to the top of the shape layers

def analyze_cutout(dicom_folder=Path.cwd(),
                   template_path='CutOut Size Check.xlsx',
                   save_data_file='CutOut Size Check Test.xlsx',
                   image_file='Cutout scan.jpg'):
    plan_files = [file for file in dicom_folder.glob('**/RP*.dcm')]

    plan_df = get_plan_data(plan_files)
    block_coords = get_block_coord(plan_df)
    selected_field = select_field(block_coords)
    insert_size = plan_df.at['ApplicatorOpening', selected_field]
    workbook = save_data(block_coords, plan_df, save_data_file, template_path)
    add_block_info(block_coords, selected_field, workbook)
    show_cutout_info(image_file, insert_size, workbook)

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
    image_file_name = 'image2021-04-16-111118-1.jpg'
    selected_file_paths = dict(
        dicom_starting_path = dicom_folder,
        image_file_starting_path = data_path / image_file_name,
        template_file_path = template_dir / template_file_name,
        save_data_file_name = data_path / output_file_name
        )

    analyze_cutout(**selected_file_paths)


if __name__ == '__main__':
    main()
