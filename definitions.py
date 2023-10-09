
def in_country_low_res(low_resolution, data):
    ''' THIS QUICKLY CHECKS IF EACH POINT IS CLEARLY WITHIN THE COUNTRY '''
    # create a gdf for the input data
    geodata = geopandas.GeoDataFrame(data, geometry=geopandas.points_from_xy(data.lng, data.lat), crs = low_resolution.crs)
    
    # create a buffer 0.05 units inside the border to check
    buffer = low_resolution.buffer(-0.5)
    buffer = buffer.to_crs(low_resolution.crs)

    # create a new geodataframe to store our false points
    not_in_buffer = geopandas.GeoDataFrame(columns=geodata.columns, crs=low_resolution.crs)

    # loop through each lat long and check if its inside
    for id, row in geodata.iterrows():
        if buffer.contains(row.geometry).bool() == False:
            # add to the gdf that stores points outside of zone
            gdf = geopandas.GeoDataFrame([row], columns=not_in_buffer.columns, crs=not_in_buffer.crs)
            not_in_buffer = pandas.concat([not_in_buffer, gdf], ignore_index=True)
    
    return not_in_buffer

def in_country_high_res(high_resolution, low_res_output):
    '''THIS PROVIDES A SLOWER, MORE PRPECISE LOOK IF EACH POINT IS WITHIN THE COUNTRY '''
    not_in_high_res = geopandas.GeoDataFrame(columns=low_res_output.columns, crs=low_res_output.crs)
    # loop through each lat long and check if its inside
    for id, row in low_res_output.iterrows():
        if high_resolution.contains(row.geometry).bool() == False:
            gdf = geopandas.GeoDataFrame([row], columns=low_res_output.columns, crs=low_res_output.crs)
            print(f'point with reference XXXX is outside of the country')

def is_centroid(low_resolution, data):
    '''THIS CHECKS IF THE COORDINATE HAS BEEN PUT AT THE CENTROID OF THE SHAPEFILE'''
    geodata = geopandas.GeoDataFrame(data, geometry=geopandas.points_from_xy(data.lng, data.lat), crs = low_resolution.crs)
    centroid = low_resolution.centroid.iloc[0]
    centroid_count = 0
    for id, row in geodata.iterrows():
        if centroid == row.geometry:
            centroid_count += 1
    # report the details back to the user
    print(f'At lat/long location {centroid.x}, {centroid.y}, the centroid of {low_resolution.NAME_NL.iloc[0]}, there are {centroid_count} points')


guatemala_high = natural_earth_10.loc[natural_earth_10['ISO_A3'] ==  'GTM']
guatemala_low = natural_earth_50.loc[natural_earth_50['ISO_A3'] ==  'GTM']
# 'c:\Users\warre\Documents\Programming\Python\gt.csv'
data = pandas.read_csv(rf'{input("copy the file path here: ")}')

is_centroid(guatemala_low, data)


