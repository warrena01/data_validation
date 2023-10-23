import geopandas as gpd
gdf = gpd.read_file(r'J:\cms\External\TeamMembers\Adam Warren\Cresta Zones\South American Cresta Zones Low Res\South American Cresta Low Res.shp')
my_list = ['Argentina', 'Bolivia', 'Brazil', 'Chile', 'Colombia', 'Ecuador', 'Guyana', 'Paraguay', 'Peru', 'Suriname', 'Uruguay', 'Venezuela', 'Mexico', 'Cuba', 'Dominican Republic', 'Haiti', 'Jamaica', 'Puerto Rico', 'Trinidad and Tobago', 'Bahamas', 'Barbados', 'Belize', 'Costa Rica', 'El Salvador', 'Guatemala', 'Honduras', 'Nicaragua', 'Panama', 'Antigua and Barbuda', 'Dominica', 'Grenada', 'Saint Kitts and Nevis', 'Saint Lucia', 'Saint Vincent and the Grenadines']
my_countries = gdf.loc[gdf['COUNTRY'].isin(my_list)]

import os
folder_path = r'J:\cms\External\TeamMembers\Adam Warren\Cresta Zones\South American Cresta Zones Per Country Low Res'

for each_country in my_list:
    country_gdf = gdf.loc[gdf['COUNTRY'] == each_country]
    destination_folder = folder_path + rf'\{each_country}'
    destination_file = destination_folder + rf'\{each_country}.shp'
    if destination_folder.exists():
        country_gdf.to_file(destination_file)
    else:
        os.makedirs(destination_folder)
                                        
