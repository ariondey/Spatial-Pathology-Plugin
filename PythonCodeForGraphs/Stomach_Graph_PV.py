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

# Read the data from the separate Excel files
meta_pgfp_gsii_ueai_minus_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompletedFigures-Stomach/Metaplasia/MetaplasiaDataSpreadsheets/Combined_MetaplasiaSpreadsheets/MPCombined_GFP+GSII+UEAI-.xlsx")
meta_pgfp_gsii_ueai_plus_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompletedFigures-Stomach/Metaplasia/MetaplasiaDataSpreadsheets/Combined_MetaplasiaSpreadsheets/MPCombined_GFP+GSII+UEAI+.xlsx")
meta_pgfp_gsii_minus_ueai_minus_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompletedFigures-Stomach/Metaplasia/MetaplasiaDataSpreadsheets/Combined_MetaplasiaSpreadsheets/MPCombined_GFP+GSII-UEAI-.xlsx")
meta_pgfp_gsii_minus_ueai_plus_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompletedFigures-Stomach/Metaplasia/MetaplasiaDataSpreadsheets/Combined_MetaplasiaSpreadsheets/MPCombined_GFP+GSII-UEAI+.xlsx")
non_meta_plasia_gfp_gsii_ueai_minus_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompletedFigures-Stomach/Normal/NormalStomachDataSpreadsheets/NormalStomachDataCombined/NonMP_Combined_GFP+GSII+UEAI-.xlsx")
non_meta_plasia_gfp_gsii_minus_ueai_minus_data = read_normalized_distances("C:/Users/ari/Documents/Professional/WFU/SPPPaperData/CompletedFigures-Stomach/Normal/NormalStomachDataSpreadsheets/NormalStomachDataCombined/NonMP_Combined_GFP+GSII-UEAI-.xlsx")

# Combine data for Metaplasia and Normal
combined_data_meta = pd.concat([
    meta_pgfp_gsii_ueai_minus_data.assign(Marker='GFP+GSII+UEAI-'),
    meta_pgfp_gsii_ueai_plus_data.assign(Marker='GFP+GSII+UEAI+'),
    meta_pgfp_gsii_minus_ueai_minus_data.assign(Marker='GFP+GSII-UEAI-'),
    meta_pgfp_gsii_minus_ueai_plus_data.assign(Marker='GFP+GSII-UEAI+')
])

combined_data_normal = pd.concat([
    non_meta_plasia_gfp_gsii_ueai_minus_data.assign(Marker='Normal_GFP+GSII+UEAI-'),
    non_meta_plasia_gfp_gsii_minus_ueai_minus_data.assign(Marker='Normal_GFP+GSII-UEAI-')
])

# Calculate normalization factors
metaplasia_normalization_factor = sum([1384.49, 1397.054, 5605.565, 4164.964, 6232.651, 1059.408, 4254.375, 2099.684, 1034.886, 1056.22]) / 686
normal_normalization_factor = sum([1405.942, 2797.663, 4181.194, 5212.550007, 1089.870, 2151.066159, 1087.514, 1049.658, 2086.007]) / 686

def plot_histogram_with_bell_curve(data, normalization_factor, std_devs_dict, title):
    markers = data['Marker'].unique()
    
    plt.figure(figsize=(10, 8))
    
    colors = ['lightblue', 'orange', 'lightgreen', 'pink', 'purple', 'yellow']
    
    for i, marker in enumerate(markers):
        subset = data[data['Marker'] == marker]
        density = gaussian_kde(subset['Normalized_Distance'].dropna())
        xs = np.linspace(0, 1, 1000)
        ys = density(xs) / normalization_factor
        
        # Normalize ys
        ys = (ys - ys.min()) / (ys.max() - ys.min())
        
        # Plot density
        plt.plot(ys, xs, label=marker)
        
        std_devs = std_devs_dict.get(marker, [0.01] * 10)  # Placeholder values
        
        bin_edges = np.linspace(0, 1, len(std_devs) + 1)
        bin_centers = 0.5 * (bin_edges[:-1] + bin_edges[1:])
        std_devs_interp = interp1d(bin_centers, std_devs, kind='linear', fill_value="extrapolate")
        std_devs_smooth = std_devs_interp(xs)
        damping_factor = np.exp(-5 * (xs - 0.8))  # Apply damping factor to temper upper bounds
        std_devs_smooth = std_devs_smooth * damping_factor.clip(0, 1)
        
        plt.fill_betweenx(xs, ys - std_devs_smooth, ys + std_devs_smooth, color=colors[i % len(colors)], alpha=0.3)
        
    plt.title(title)
    plt.ylabel('Normalized Distance')
    plt.xlabel('Density (Normalized)')
    plt.xlim(0, 1)
    plt.ylim(0, 1)
    plt.legend(title='Marker')
    plt.grid(True)
    plt.show()

# Placeholder standard deviation values for each marker
meta_std_devs_dict = {
    'GFP+GSII+UEAI-': [0.016615, 0.026636, 0.030103, 0.027554, 0.025207, 0.025831],
    'GFP+GSII+UEAI+': [0.08939, 0.02489, 0.030392, 0.031342, 0.028058, 0.03286],
    'GFP+GSII-UEAI-': [0.019358, 0.028315, 0.026615, 0.032726, 0.035687, 0.027887],
    'GFP+GSII-UEAI+': [0.019735, 0.029035, 0.01551, 0, 0.039065, 0.025785, 0.032763, 0.032703]
}

normal_std_devs_dict = {
    'Normal_GFP+GSII+UEAI-': [0.020995, 0.027393, 0.020558],
    'Normal_GFP+GSII-UEAI-': [0.018515, 0.024981, 0.032567]
}

# Plot the combined histogram with bell curve for Metaplasia
plot_histogram_with_bell_curve(combined_data_meta, metaplasia_normalization_factor, meta_std_devs_dict, 'Normalized Distance Distribution (Metaplasia)')

# Plot the combined histogram with bell curve for Normal
plot_histogram_with_bell_curve(combined_data_normal, normal_normalization_factor, normal_std_devs_dict, 'Normalized Distance Distribution (Normal)')