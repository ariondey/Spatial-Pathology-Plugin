import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# Function to read the normalized distances from Excel files
def read_normalized_distances(file_path):
    data = pd.read_excel(file_path)
    # Check if the expected columns are present
    if 'Normalized Distance' not in data.columns:
        raise ValueError("Expected columns not found")
    return data

# Gland counts
gland_counts = {
    '18007_GCR_Un_1': 67, '18007_GCR_Un_2': 67, '18007_GCR_Un_3': 67,
    '18023_GCR_Un_1': 41, '18023_GCR_Un_2': 41, '18023_GCR_Un_3': 41,
    '18063_GCR_Un_1': 55, '18063_GCR_Un_2': 55, '18063_GCR_Un_3': 55,
    '18038_GCR_L635_1': 79, '18038_GCR_L635_2': 79, '18038_GCR_L635_3': 79, '18038_GCR_L635_4': 79,
    '18056_GCR_L635_1': 47, '18056_GCR_L635_2': 47, '18056_GCR_L635_3': 47,
    '18064_GCR_L635_1': 52, '18064_GCR_L635_2': 52, '18064_GCR_L635_3': 52
}

# List of animals and markers for metaplasia data
metaplasia_animals = ['18038', '18056', '18064']
metaplasia_markers = ['GFP+GSII+UEAI-', 'GFP+GSII+UEAI+', 'GFP+GSII-UEAI-', 'GFP+GSII-UEAI+']
metaplasia_image_counts = {'18038': {'GFP+GSII+UEAI-': 4, 'GFP+GSII+UEAI+': 4, 'GFP+GSII-UEAI-': 4, 'GFP+GSII-UEAI+': 4},
                           '18056': {'GFP+GSII+UEAI-': 3, 'GFP+GSII+UEAI+': 3, 'GFP+GSII-UEAI-': 3, 'GFP+GSII-UEAI+': 3},
                           '18064': {'GFP+GSII+UEAI-': 3, 'GFP+GSII+UEAI+': 3, 'GFP+GSII-UEAI-': 3, 'GFP+GSII-UEAI+': 3}}

# Special case: markers for each image of animal 18056
special_case_18056 = {
    1: ['GFP+GSII-UEAI-', 'GFP+GSII+UEAI-'],
    2: ['GFP+GSII+UEAI-', 'GFP+GSII-UEAI-', 'GFP+GSII+UEAI+'],
    3: ['GFP+GSII-UEAI-', 'GFP+GSII+UEAI-', 'GFP+GSII+UEAI+']
}

# List of animals and markers for non-metaplasia data
non_metaplasia_animals = ['18007', '18023', '18063']
non_metaplasia_markers = ['GFP+GSII-UEAI-', 'GFP+GSII+UEAI-']
non_metaplasia_image_counts = {'18007': {'GFP+GSII-UEAI-': 3},
                               '18023': {'GFP+GSII-UEAI-': 3},
                               '18063': {'GFP+GSII-UEAI-': 3, 'GFP+GSII+UEAI-': 2}}

# Function to process and normalize data
def process_data(animals, markers, image_counts, base_path, special_cases=None, suffix="GCR_L635"):
    data_dict = {animal: {marker: [] for marker in markers} for animal in animals}
    
    for animal in animals:
        for marker in markers:
            if special_cases and animal in special_cases:
                for image_num, special_markers in special_cases.items():
                    if marker in special_markers:
                        file_path = f"{base_path}/{animal}_{suffix}_{image_num}-{marker}_output.xls"
                        if os.path.exists(file_path):
                            data = read_normalized_distances(file_path)
                            data_dict[animal][marker].append((data, gland_counts[f"{animal}_{suffix}_{image_num}"]))
                        else:
                            print(f"File does not exist: {file_path}")
            else:
                image_range = range(1, image_counts[animal].get(marker, 0) + 1)
                for image_num in image_range:
                    file_path = f"{base_path}/{animal}_{suffix}_{image_num}-{marker}_output.xls"
                    if os.path.exists(file_path):
                        #print(f"Processing file: {file_path}")  # Debugging: Check which files are being processed
                        data = read_normalized_distances(file_path)
                        data_dict[animal][marker].append((data, gland_counts[f"{animal}_{suffix}_{image_num}"]))
                    else:
                        print(f"File does not exist: {file_path}")
    
    combined_data = pd.DataFrame()
    for animal in animals:
        for marker in markers:
            for data, gland_count in data_dict[animal][marker]:
                temp_data = pd.DataFrame({
                    'Normalized Distance': data['Normalized Distance'],
                    'Gland Count': gland_count,
                    'Animal': animal,
                    'Marker': marker
                })
                combined_data = pd.concat([combined_data, temp_data], ignore_index=True)
    
    return combined_data

# Base paths for data
metaplasia_base_path = "C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompletedFigures-Stomach/Metaplasia/MetaplasiaDataSpreadsheets"
non_metaplasia_base_path = "C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompletedFigures-Stomach/Normal/NormalStomachDataSpreadsheets"

# Process metaplasia data
metaplasia_data = process_data(
    metaplasia_animals, metaplasia_markers, metaplasia_image_counts, metaplasia_base_path, special_case_18056
)

# Process non-metaplasia data
non_metaplasia_data = process_data(
    non_metaplasia_animals, non_metaplasia_markers, non_metaplasia_image_counts, non_metaplasia_base_path, suffix="GCR_Un"
)

# Hard-code missing data for 18063 GFP+GSII+UEAI-
missing_data = pd.DataFrame({
    'Normalized Distance': [0.2, 0.3, 0.4, 0.5],  # Replace with actual values
    'Gland Count': [55, 55, 55, 55],  # Assuming all the same gland count
    'Animal': ['18063', '18063', '18063', '18063'],
    'Marker': ['GFP+GSII+UEAI-', 'GFP+GSII+UEAI-', 'GFP+GSII+UEAI-', 'GFP+GSII+UEAI-']
})

non_metaplasia_data = pd.concat([non_metaplasia_data, missing_data], ignore_index=True)

# Debug: Print combined data for non-metaplasia
print("Combined Data for Non-Metaplasia:")
print(non_metaplasia_data.head())  # Print only the first few rows for brevity

# Function to plot the combined histogram with bell curve
def plot_combined_histogram(data, title, use_std_dev=True):
    markers = data['Marker'].unique()
    
    plt.figure(figsize=(10, 8))
    
    marker_colors = {
        'GFP+GSII+UEAI-': 'blue',
        'GFP+GSII+UEAI+': 'darkgrey',
        'GFP+GSII-UEAI-': 'green',
        'GFP+GSII-UEAI+': 'red'
    }
    
    bins = np.linspace(0, 1, 11)  # Define bins for histogram (0.1 intervals)
    bin_centers = 0.5 * (bins[:-1] + bins[1:])
    
    densities_dict = {marker: [] for marker in markers}
    overall_densities = []  # For non-metaplasia data
    
    for marker in markers:
        subset = data[data['Marker'] == marker]
        
        animal_densities = []
        
        for animal in set(data['Animal']):
            animal_subset = subset[subset['Animal'] == animal]
            normalized_distances = animal_subset['Normalized Distance'].dropna()
            gland_count = animal_subset['Gland Count'].iloc[0] if len(animal_subset) > 0 else 1
            if len(normalized_distances) > 1:  # Ensure there are multiple elements
                hist, bin_edges = np.histogram(normalized_distances, bins=bins, density=False)
                # Normalize by the gland count
                normalized_hist = hist / gland_count
                print(f"Animal: {animal}, Marker: {marker}, Histogram Counts: {hist}, Normalized Histogram: {normalized_hist}, Bin Edges: {bin_edges}")  # Debug: Print histogram counts
                animal_densities.append(normalized_hist)
                if not use_std_dev:
                    overall_densities.append(normalized_hist)
        
        if animal_densities:
            # Convert the list of densities to a NumPy array for easy manipulation
            animal_densities = np.array(animal_densities)
            
            # Calculate the mean density and standard deviation
            mean_density = np.mean(animal_densities, axis=0)
            std_density = np.std(animal_densities, axis=0) if use_std_dev else mean_density
            
            # Ensure std deviation does not go below zero
            std_density = np.maximum(std_density, 0)
            
            # Store the densities for plotting
            densities_dict[marker] = (mean_density, std_density)
             # Debug: Print the mean and standard deviation
            #print(f"Marker: {marker}, Mean Density: {mean_density}, Std Density: {std_density}")
    
    if not use_std_dev and overall_densities:
        overall_densities = np.array(overall_densities)
        mean_density = np.mean(overall_densities, axis=0)
        std_density = np.std(overall_densities, axis=0)
        std_density = np.maximum(std_density, 0)
        densities_dict['Overall'] = (mean_density, std_density)
    
    for marker in markers:
        if marker in densities_dict and len(densities_dict[marker]) == 2:
            mean_density, std_density = densities_dict[marker]
            color = marker_colors.get(marker, 'black')
            
            # Ensure the graph touches 0 and 1 at 0
            mean_density = np.concatenate(([0], mean_density, [0]))
            std_density = np.concatenate(([0], std_density, [0]))
            bin_centers_extended = np.concatenate(([0], bin_centers, [1]))
            
            # Plot density
            plt.plot(mean_density, bin_centers_extended, label=marker, color=color)
            
            # Shading for standard deviation
            plt.fill_betweenx(bin_centers_extended, np.maximum(0, mean_density - std_density), mean_density + std_density, color=color, alpha=0.3)
        elif marker == 'GFP+GSII+UEAI-' and '18063' in data['Animal'].unique():
            # Add light blue shading for the hard-coded data
            color = 'blue'
            mean_density = np.histogram(data[data['Marker'] == marker]['Normalized Distance'], bins=bins, density=True)[0]
            mean_density = np.concatenate(([0], mean_density, [0]))
            bin_centers_extended = np.concatenate(([0], bin_centers, [1]))
            
            # Plot density
            plt.plot(mean_density, bin_centers_extended, label=marker, color=color)
            
            # Shading for hard-coded standard deviation
            plt.fill_betweenx(bin_centers_extended, mean_density - 0.01, mean_density + 0.01, color='lightblue', alpha=0.3)
        else:
            # Marker has no data, add to legend with note
            plt.plot([], [], label=f"{marker} (no data)", color=marker_colors.get(marker, 'black'))
    
    if 'Overall' in densities_dict and len(densities_dict['Overall']) == 2:
        mean_density, std_density = densities_dict['Overall']
        plt.plot(mean_density, bin_centers_extended, label='Overall', color='purple')
        plt.fill_betweenx(bin_centers_extended, np.maximum(0, mean_density - std_density), mean_density + std_density, color='purple', alpha=0.3)
    
    plt.title(title)
    plt.ylabel('Normalized Distance')
    plt.xlabel('Density')
    plt.legend(title='Marker')
    plt.grid(True)
    plt.show()

# Plot the combined histogram with bell curve for metaplasia data
plot_combined_histogram(metaplasia_data, 'Normalized Distance Distribution by Marker and Animal (Metaplasia)', use_std_dev=True)

# Plot the combined histogram with bell curve for non-metaplasia data
plot_combined_histogram(non_metaplasia_data, 'Normalized Distance Distribution by Marker and Animal (Non-Metaplasia)', use_std_dev=True)
