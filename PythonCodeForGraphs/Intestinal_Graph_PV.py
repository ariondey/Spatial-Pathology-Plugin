import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# Function to read the normalized distances from Excel files
def read_normalized_distances(file_path):
    data = pd.read_excel(file_path)
    # Check if the expected column name is present
    if 'Normalized Distance' in data.columns:
        data.rename(columns={'Normalized Distance': 'Normalized_Distance'}, inplace=True)
    elif 'Normalized_Distance' not in data.columns:
        raise ValueError("Normalized Distance column not found")
    return data

# Animal and marker details
animals = ['BF13', 'BF14', 'BF15.2']
markers = ['EzrinCy3', 'Ki67Cy5', 'OLFM4488']

# Function to process and normalize data
def process_data(animals, markers, base_path):
    data_dict = {animal: {marker: [] for marker in markers} for animal in animals}
    
    for animal in animals:
        for marker in markers:
            for image_num in range(1, 4):  # Assuming 3 images per marker
                file_path = f"{base_path}/{animal}_OLFM4488_EzrinCy3_Ki67Cy5_20X_{image_num}_merge_done-{marker}_output.xls"
                if os.path.exists(file_path):
                    print(f"Processing file: {file_path}")  # Debugging: Check which files are being processed
                    data = read_normalized_distances(file_path)
                    data_dict[animal][marker].append(data)
                else:
                    print(f"File does not exist: {file_path}")
    
    combined_data = pd.DataFrame()
    for animal in animals:
        for marker in markers:
            for data in data_dict[animal][marker]:
                temp_data = pd.DataFrame({
                    'Normalized_Distance': data['Normalized_Distance'],
                    'Animal': animal,
                    'Marker': marker
                })
                combined_data = pd.concat([combined_data, temp_data], ignore_index=True)
    
    return combined_data

# Base path for data
base_path = "C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines"

# Process data
combined_data = process_data(animals, markers, base_path)

# Function to plot the combined histogram with bell curve
def plot_combined_histogram_with_bell_curve(data, title):
    markers = data['Marker'].unique()
    
    plt.figure(figsize=(10, 8))
    
    marker_colors = {
        'EzrinCy3': 'orange',
        'Ki67Cy5': 'lightgreen',
        'OLFM4488': 'blue'
    }
    
    bins = np.linspace(0, 1, 11)  # Define bins for histogram (0.1 intervals)
    bin_centers = 0.5 * (bins[:-1] + bins[1:])
    
    densities_dict = {marker: [] for marker in markers}
    
    for marker in markers:
        subset = data[data['Marker'] == marker]
        
        animal_densities = []
        
        for animal in set(data['Animal']):
            animal_subset = subset[subset['Animal'] == animal]
            normalized_distances = animal_subset['Normalized_Distance'].dropna()
            if len(normalized_distances) > 1:  # Ensure there are multiple elements
                hist, bin_edges = np.histogram(normalized_distances, bins=bins, density=True)
                # Normalize by the sum of hist
                normalized_hist = hist / sum(hist)
                animal_densities.append(normalized_hist)
        
        if animal_densities:
            # Convert the list of densities to a NumPy array for easy manipulation
            animal_densities = np.array(animal_densities)
            
            # Calculate the mean density and standard deviation
            mean_density = np.mean(animal_densities, axis=0)
            std_density = np.std(animal_densities, axis=0)
            
            # Ensure std deviation does not go below zero
            std_density = np.maximum(std_density, 0)
            
            # Store the densities for plotting
            densities_dict[marker] = (mean_density, std_density)
             # Debug: Print the mean and standard deviation
            print(f"Marker: {marker}, Mean Density: {mean_density}, Std Density: {std_density}")
    
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
        else:
            # Marker has no data, add to legend with note
            plt.plot([], [], label=f"{marker} (no data)", color=marker_colors.get(marker, 'black'))
    
    plt.title(title)
    plt.ylabel('Normalized Distance')
    plt.xlabel('Density')
    plt.legend(title='Marker')
    plt.grid(True)
    plt.show()

# Plot the combined histogram with bell curve for all animals
plot_combined_histogram_with_bell_curve(combined_data, 'Normalized Distance Distribution by Marker (All Animals)')

# Plot separate histograms for each animal
for animal in animals:
    animal_data = combined_data[combined_data['Animal'] == animal]
    plot_combined_histogram_with_bell_curve(animal_data, f'Normalized Distance Distribution by Marker ({animal})')
