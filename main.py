import os
import re
import pandas as pd
from bs4 import BeautifulSoup
from shapely.geometry import Point, Polygon
# import geopandas as gpd
# gpd.io.file.fiona.drvsupport.supported_drivers['KML'] = 'rw'

ZONES_KML_PATH = 'zones.kml'
KML_LAYERS_ABSOLUTE_PATH = 'C:\\Users\\Edu\\Desktop\\point-within-poly\\kml_layers'
XLS_LAYERS_ABSOLUTE_PATH = 'C:\\Users\\Edu\\Desktop\\point-within-poly\\xls_layers'


def _kml_points_to_df(file_name):
    """
    Given an .kml file (with points)
    Return a df

    Note: I not using the geopandas read_file for this,
    because its only return one folder per each kml file
    """
    print('_kml_points_to_df: ' + file_name)
    kml_df = pd.DataFrame(columns=['Longitude',
                                   'Latitude',
                                   'Altitude',
                                   'Name',
                                   'Description',
                                   'Folder'])

    kml_df_index = 0
    with open(file_name, 'r', encoding='utf8') as file:
        kml_soup = BeautifulSoup(file.read(), 'xml')
        folders = kml_soup.select('Folder')
        for folder in folders:
            folder_name = folder.find('name').getText()
            placemarks = folder.find_all('Placemark')
            for placemark in placemarks:
                point = placemark.find('Point')
                if point is not None:
                    name_find = placemark.find('name')
                    desc_find = placemark.find('description')
                    coord_find = point.find('coordinates')

                    name = name_find.getText() if name_find else ""
                    description = desc_find.getText() if desc_find else ""
                    coordinates = coord_find.getText().split(',') if coord_find else ""

                    kml_df.loc[kml_df_index, 'Folder'] = folder_name
                    kml_df.loc[kml_df_index, 'Name'] = name
                    kml_df.loc[kml_df_index, 'Description'] = description
                    kml_df.loc[kml_df_index,
                               'Longitude':'Altitude'] = coordinates

                    kml_df_index += 1
    return kml_df


def _kml_polys_to_df(file_name):
    """
    Given an .kml file (with polys)
    Return a df

    This function is intended to avoid the geopandas installtion (only shapely is required)
    """
    print('_kml_polys_to_df: ' + file_name)
    kml_df = pd.DataFrame(columns=['Name',
                                   'Description',
                                   'Folder',
                                   'geometry'])

    kml_df_index = 0
    with open(file_name, 'r', encoding='utf8') as file:
        kml_soup = BeautifulSoup(file.read(), 'xml')
        folders = kml_soup.select('Folder')
        for folder in folders:
            folder_name = folder.find('name').getText()
            placemarks = folder.find_all('Placemark')
            for placemark in placemarks:
                poly = placemark.find('Polygon')
                if poly is not None:
                    name_find = placemark.find('name')
                    desc_find = placemark.find('description')
                    coord_find = poly.find('coordinates')

                    name = name_find.getText() if name_find else ""
                    description = desc_find.getText() if desc_find else ""
                    coordinates = []

                    if coord_find:
                        coordinates_list = (coord_find
                                            .getText()
                                            .replace('\n', '')
                                            .replace('\t', '')
                                            .split(' ')
                                            )

                        for coord_str in coordinates_list:
                            coordinate = [float(number) for number in coord_str.split(',') if coord_str]
                            if len(coordinate) == 3:
                                coordinates.append(tuple(coordinate))

                    kml_df.loc[kml_df_index, 'Folder'] = folder_name
                    kml_df.loc[kml_df_index, 'Name'] = name
                    kml_df.loc[kml_df_index, 'Description'] = description
                    kml_df.loc[kml_df_index, 'geometry'] = Polygon(coordinates)

                    kml_df_index += 1
    return kml_df


def _get_layers(folder_path):
    """
        Given a folder absolute path
        Return a dictonary with df layers (geopandas read of .kml files)

        './kml_layers'
            filename1.kml
            filename2.kml

        return
            {
                filename1: df1,
                filename2: df2
            }
    """
    os.chdir(folder_path)
    file_names = os.listdir()
    layers = {}

    for file_name in file_names:
        if (re.match(r'^.*\.kml$', file_name)):
            layer_name = file_name.replace('.kml', '')
            # layer = gpd.read_file(file_name, driver='KML')
            layer = _kml_points_to_df(file_name)
            layers[layer_name] = layer

    return layers


def _calculate_zone(point, zones):
    """
        Given a point (x, y), and a list of Zones (df´s)
        Return in which Zone the placemark is
    """
    for index, zone in zones.iterrows():
        if point.within(zone['geometry']):
            return zone['Name']
    return 'Sin Zona'


def _get_point(Longitude, Latitude):
    return Point(Longitude, Latitude)


def _complete_layer_info(layer, zones):
    layer['Zone'] = (layer
                     .apply(lambda row: _get_point(float(row['Longitude']),
                                                   float(row['Latitude'])), axis=1)
                     .apply(lambda point: _calculate_zone(point, zones))
                     )


def _write_excel_from_df(df, file_name):
    print('_write_excel_from_df: ' + file_name)
    df.to_excel(file_name, sheet_name='Hoja1')


def main():

    # zones = gpd.read_file(ZONES_KML_PATH, driver='KML')
    zones = _kml_polys_to_df(ZONES_KML_PATH)
    layers = _get_layers(KML_LAYERS_ABSOLUTE_PATH)
    placemark_total = 0

    os.chdir(XLS_LAYERS_ABSOLUTE_PATH)  # Path to Store the .xls files
    for layer_name, layer in layers.items():
        _complete_layer_info(layer, zones)
        _write_excel_from_df(layer, layer_name + '.xls')
        placemark_total += layer.shape[0]

    print('Numero total de placemark a los que se les asigno una Zona: {}'
          .format(placemark_total))


if __name__ == '__main__':
    main()