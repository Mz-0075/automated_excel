import osmnx as ox
import simplekml


def get_roads_within_radius(lat, lon, radius):
    # Download the street network within the given radius
    graph = ox.graph_from_point((lat, lon), dist=radius, network_type='all')

    # Convert the graph into a GeoDataFrame
    gdf_nodes, gdf_edges = ox.graph_to_gdfs(graph)

    # Create a KML object using simplekml
    kml = simplekml.Kml()

    # Loop through the roads and add them to the KML
    for _, row in gdf_edges.iterrows():
        coordinates = []
        # Loop through the geometry and convert to coordinates
        for x, y in zip(row['geometry'].coords.xy[0], row['geometry'].coords.xy[1]):
            coordinates.append((x, y))

        # Add the line to KML
        line = kml.newlinestring(name="Road", coords=coordinates)

    # Save to KML file
    kml.save("EZKN15.kml")
    print("KML file saved!")


# Example: San Francisco coordinates and 1km radius
lat = 8.99534
lon = 76.71983
radius = 1500
get_roads_within_radius(lat, lon, radius)
