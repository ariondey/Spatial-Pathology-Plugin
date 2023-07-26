# Spatial Pathology Plugin

A [ImageJ] (https://imagej.net/) plugin to quantify biomarker distribution along the gastrointestinal unit axis.
ImageJ is a Java-based image processing software that serves as a widely used platform for microscopy image analysis in the basic sciences.  

## Usage
This is a gastrointestinal-focused spatial pathology plugin to advance research in spatial analysis of GI epithelial tissues to quantify immunostaining markers in a tissue landscape through a spatial lens.

### Installing
In order to use this plugin, you need Fiji which is the latest version of ImageJ, you must use this version in order to get the remote update. This plugin is hosted on a remote webserver so you can download it using this tutorial: https://imagej.net/update-sites/following and by using the URL http://sites.imagej.net/SpatialPathology/
You only need to do this on the first install, if it was installed correctly you will see "SpatialpathologyIJMJava" in your plugins folder and if any changes are made to the plugin, all you will need to do is update your imagej. 
### Workflow 
When you launch the plugin you will be immediately asked what you want the bin interval to be. Once an image is run, the normalized distances to user selected points will range from 0 (directly on the base line) or 1.0 (directly on the top line). 
Example (assuming you choose a bin interval of 0.1): 
```md
| 0   | 0.1  | 0.2  | (This will be your interval)
| 0   | 0.5  | 1.5  | (This will be the cells / gland)


Table: This table illustrates what the table you upload must look like. The file also must be saved as a .csv (notably, ImageJ interprets the UTF-8 character marker in plaintext and it will not work, so make sure it is a regular .CSV and not a .csv UTF-8.

```
 Therefore, if you want the table at the end to count instances  Once you do that, you will then be prompted and asked if you want to draw the line through uploading a .csv file of the line's coordinates. 
The structure of the csv you could upload must be as follows: 
```md
| X     | Y    |
| ----- | ---- |
| 0     | 0    |
| 1     | 1    |
| ...   | ...  |

Table: This table illustrates what the table you upload must look like. The file also must be saved as a .csv (notably, ImageJ interprets the UTF-8 character marker in plaintext and it will not work, so make sure it is a regular .CSV and not a .csv UTF-8.

```
If you opt to not upload a .csv of the lines, you can draw them using the segmented line tool (the one selected by default) or through the freehand line tool. The first prompt will ask for the base of the gland and the second line will be the top. 
After this, you will be asked to place points through a .csv. The format for this is exactly the same as the format for adding a line, and if you opt not to use these, you will be able to draw them manually. You will also be prompted twice to store things: once for the edited image, and once for your excel sheets.


### Data Output
There will be two excel sheets as the output. One of them is all the information in the image (centroid coordinates, distances from the base to the centroid, distances from the top to the centroid and then the normalized distance). The other one the averaged out one that uses the normalized distances and the bin intervals that the user specified earlier.