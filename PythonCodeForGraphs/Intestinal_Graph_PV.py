import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import gaussian_kde
from scipy.interpolate import interp1d

def read_normalized_distances(file_path):
    data = pd.read_excel(file_path)
    # Check if the expected column name is present
    if 'Normalized Distance' in data.columns:
        data.rename(columns={'Normalized Distance': 'Normalized_Distance'}, inplace=True)
    elif 'Normalized_Distance' not in data.columns:
        raise ValueError("Normalized Distance column not found")
    return data

# Read the data from the separate Excel files for each animal and each fluorophore
bf13_olfm4488_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines/OLFM4488/OLFM4488_CombinedByAT/BF13_OLFM4488_EzrinCy3_ki67Cy5_20X_123_merged.xlsx")
bf13_ezrinCy3_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines/EzrinCy3/EzrinCy3_CombinedByAT/BF13_OLFM4488_EzrinCy3_Ki67Cy5_20X_123.xlsx")
bf13_ki67Cy5_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines/Ki67Cy5/Ki67Cy5_CombinedByAT/BF13_OLFM4488_EzrinCy3_Ki67Cy5_20X_123_.xlsx")

# BF14 data
bf14_olfm4488_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines/OLFM4488/OLFM4488_CombinedByAT/BF14_OLFM4488_EzrinCy3_ki67Cy5_20X_123_merged.xlsx")
bf14_ezrinCy3_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines/EzrinCy3/EzrinCy3_CombinedByAT/BF14_OLFM4488_EzrinCy3_Ki67Cy5_20X_123.xlsx")
bf14_ki67Cy5_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines/Ki67Cy5/Ki67Cy5_CombinedByAT/BF14_OLFM4488_EzrinCy3_Ki67Cy5_20X_123_.xlsx")

# BF15.2 data
bf15_2_olfm4488_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines/OLFM4488/OLFM4488_CombinedByAT/BF15.2_OLFM4488_EzrinCy3_ki67Cy5_20X_123_merged.xlsx")
bf15_2_ezrinCy3_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines/EzrinCy3/EzrinCy3_CombinedByAT/BF15.2_OLFM4488_EzrinCy3_Ki67Cy5_20X_123.xlsx")
bf15_2_ki67Cy5_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompetedFigures-Intestine/Normal-20240528T185729Z-001/Normal/CompletedDataSpreadsheetsIntestines/Ki67Cy5/Ki67Cy5_CombinedByAT/BF15.2_OLFM4488_EzrinCy3_Ki67Cy5_20X_123_.xlsx")

# Combine data for each fluorophore and add appropriate columns
combined_data = pd.concat([
    bf13_olfm4488_data.assign(Fluorophore='OLFM4488', Animal='BF13'),
    bf13_ezrinCy3_data.assign(Fluorophore='EzrinCy3', Animal='BF13'),
    bf13_ki67Cy5_data.assign(Fluorophore='Ki67Cy5', Animal='BF13'),
    bf14_olfm4488_data.assign(Fluorophore='OLFM4488', Animal='BF14'),
    bf14_ezrinCy3_data.assign(Fluorophore='EzrinCy3', Animal='BF14'),
    bf14_ki67Cy5_data.assign(Fluorophore='Ki67Cy5', Animal='BF14'),
    bf15_2_olfm4488_data.assign(Fluorophore='OLFM4488', Animal='BF15.2'),
    bf15_2_ezrinCy3_data.assign(Fluorophore='EzrinCy3', Animal='BF15.2'),
    bf15_2_ki67Cy5_data.assign(Fluorophore='Ki67Cy5', Animal='BF15.2')
])

# Calculate normalization factor
normalization_factor = sum([2807.102505, 1923.644387, 1146.726315, 1047.336197, 2107.781163, 1060.626366, 1083.302597, 1056.149944, 2188.880138]) / 686

# Standard deviation values for each bin for EzrinCy3 data
ezrin_bin_std_devs = [0.008628145, 0.028418914, 0.028721503, 0.028611132, 0.02821937, 
                      0.029179678, 0.029510586, 0.027960446, 0.028752226, 0.019878786]

# Provided standard deviation values for each bin for Ki67Cy5 data
ki67_bin_std_devs = [0.019413, 0.020298, 0.024766, 0, 0, 0, 0, 0, 0, 0]

# Standard deviation values for each bin for OLFM4488 data
olfm_bin_std_devs = [0.017273541, 0.031571691, 0, 0, 0, 0, 0, 0, 0, 0]

def plot_combined_histogram_with_bell_curve(data, normalization_factor, ezrin_bin_std_devs, ki67_bin_std_devs, olfm_bin_std_devs):
    fluorophores = data['Fluorophore'].unique()
    
    plt.figure(figsize=(10, 8))
    
    for fluorophore in fluorophores:
        subset = data[data['Fluorophore'] == fluorophore]
        density = gaussian_kde(subset['Normalized_Distance'].dropna())
        xs = np.linspace(0, 1, 1000)
        ys = density(xs) / normalization_factor
        
        # Plot density
        plt.plot(ys, xs, label=fluorophore)
        
        # Apply shading for EzrinCy3 data
        if fluorophore == 'EzrinCy3':
            bin_edges = np.linspace(0, 1, len(ezrin_bin_std_devs) + 1)
            bin_centers = 0.5 * (bin_edges[:-1] + bin_edges[1:])
            std_devs_interp = interp1d(bin_centers, ezrin_bin_std_devs, kind='linear', fill_value="extrapolate")
            std_devs_smooth = std_devs_interp(xs)
            plt.fill_betweenx(xs, np.maximum(0, ys - std_devs_smooth), ys + std_devs_smooth, color='orange', alpha=0.3)
        
        # Apply shading for OLFM4488 data
        if fluorophore == 'OLFM4488':
            bin_edges = np.linspace(0, 1, len(olfm_bin_std_devs) + 1)
            bin_centers = 0.5 * (bin_edges[:-1] + bin_edges[1:])
            std_devs_interp = interp1d(bin_centers, olfm_bin_std_devs, kind='linear', fill_value="extrapolate")
            std_devs_smooth = std_devs_interp(xs)
            plt.fill_betweenx(xs, np.maximum(0, ys - std_devs_smooth), ys + std_devs_smooth, color='blue', alpha=0.3)
        
        # Apply Shading for Ki67Cy5 Data 
        if fluorophore == 'Ki67Cy5':
            bin_edges = np.linspace(0, 1, len(ki67_bin_std_devs) + 1)
            bin_centers = 0.5 * (bin_edges[:-1] + bin_edges[1:])
            std_devs_interp = interp1d(bin_centers, ki67_bin_std_devs, kind='linear', fill_value="extrapolate")
            std_devs_smooth = std_devs_interp(xs)
            plt.fill_betweenx(xs, np.maximum(0, ys - std_devs_smooth), ys + std_devs_smooth, color='lightgreen', alpha=0.3)
            
    plt.title('Normalized Distance Distribution by Fluorophore')
    plt.ylabel('Normalized Distance')
    plt.xlabel('Density (Normalized)')
    plt.xlim(0, 1)  # Restrict x-axis to 0-1
    plt.ylim(0, 1)  # Restrict y-axis to 0-1
    plt.legend(title='Fluorophore')
    plt.grid(True)
    plt.show()

# Plot the combined histogram with bell curve
plot_combined_histogram_with_bell_curve(combined_data, normalization_factor, ezrin_bin_std_devs, ki67_bin_std_devs, olfm_bin_std_devs)
