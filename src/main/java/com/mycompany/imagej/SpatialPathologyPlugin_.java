import static java.lang.Math.sqrt;
import java.io.*;
import java.awt.Color;
import java.awt.Component;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashSet;
import java.util.List;

import ij.IJ;
import ij.io.Opener;
import ij.ImagePlus;
import ij.WindowManager;
import ij.gui.GenericDialog;
import ij.gui.Overlay;
import ij.gui.PointRoi;
import ij.gui.PolygonRoi;
import ij.gui.Roi;
import ij.gui.WaitForUserDialog;
import ij.io.OpenDialog;
import ij.measure.ResultsTable;
import ij.plugin.PlugIn;
import ij.plugin.frame.RoiManager;
import ij.process.FloatPolygon;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import java.util.Random;
import java.util.Set;

import ij.io.DirectoryChooser;
import ij.io.FileSaver;
import ij.plugin.tool.PlugInTool;


public class SpatialPathologyPlugin_ implements PlugIn {

    // Settings. Edit to change from default values. ///////////////////////////////
    // These are safely editable by the user

    // output file name
    // 0: prompt user for output file name & location
    // 1: output file will be given a reasonable timestamped name & location (same
    // as image directory) automatically
    private static final int outputFileAuto = 1;

    // coordinates on the image to draw the spline
    // 0: the coordinates would be obtained from a csv file
    // 1: manually drawn by the user
    // private static final int manualCoordinateFlag = 1;

    // Global Arrays
    public float[] xPointsGlobal;
    public float[] xPoints2Global;
    public float[] yPointsGlobal;
    public float[] yPoints2Global;
    public Float[] userPickedXCoordsGlobal;
    public Float[] userPickedYCoordsGlobal;
    public Float[] userPickedCSVXCoordsGlobal;
    public Float[] userPickedCSVYCoordsGlobal;
    public float[] bottomLineArrayDistancesGlobal;
    public float[] topLineArrayDistancesGlobal;
    public String[] MarkerTypeGlobal;
    public String imageFileName;
    public String imgFilePath;
    public Integer lineChartIterationCount = 0;
    public String selectedFolderPath;
    public Integer generateHistogramFlag = 0;
    public boolean notConclusionRun = true;
    public double BincountHisto = 0;
    public double BincountZeroPointOneHisto = 0;
    public double BincountZeroPointTwoHisto = 0;
    public double BincountZeroPointThreeHisto = 0;
    public double BincountZeroPointFourHisto = 0;
    public double BincountZeroPointFiveHisto = 0;
    public double BincountZeroPointSixHisto = 0;
    public double BincountZeroPointSevenHisto = 0;
    public double BincountZeroPointEightHisto = 0;
    public double BincountZeroPointNineHisto = 0;
    public double BincountOnePointZeroHisto = 0;
    public boolean lineInputIsCSV;
    public String outputFileNameTS;
    public double chosenInterval = 0;
    public double[] globalBinCountArray;
    public double lengthLine1;
    public double lengthLine2;
    public double distanceToDivide;
    public ImagePlus imp0;
    public float[] distanceToDivideArray;

    ////////////////////////////////////////////////////////////////////////////////



    @Override
    public void run(String arg) {
        
        double start = 0;
        double end = 1.0;
        boolean isValidInput = false;
       
          while (!isValidInput) {
                try {
                    String intervalInput = JOptionPane.showInputDialog(null, "Enter the chosen interval between " + start + " and " + end + ":");
                    if (intervalInput == null) {
                        // User clicked cancel or closed the dialog
                        IJ.exit();
                    }

                    chosenInterval = Double.parseDouble(intervalInput);

                    if (chosenInterval < start || chosenInterval > end) {
                        throw new IllegalArgumentException("The chosen interval must be between " + start + " and " + end + ".");
                    }

                    isValidInput = true;
                } catch (NumberFormatException e) {
                    JOptionPane.showMessageDialog(null, "Invalid input! Please enter a valid integer.", "Error", JOptionPane.ERROR_MESSAGE);
                } catch (IllegalArgumentException e) {
                    JOptionPane.showMessageDialog(null, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                }
            }

            // At this point, chosenInterval is valid, and you can use it as needed.
            //System.out.println("Chosen interval: " + chosenInterval);
          
           // System.out.println("array global size" + ((double) 1.0 / chosenInterval ));
            double epsilon = 1e-10;
            globalBinCountArray = new double[(int) Math.ceil((1.0 + epsilon) / chosenInterval) + 1];
        outputFileNameTS = makeOutputFileNameTS();

        // Control Variable for the repeat of the macro
        int repeatFlag = 1;
        while (repeatFlag == 1) {
            String outputFileName = "";
            // Prompt User to select image for measurement
             imp0 = null;
            boolean canOpen;
            // Create an OpenDialog to prompt the user for an image file
            do {
            	lengthLine1 = 0;
            	lengthLine2 = 0;
                OpenDialog openDialog = new OpenDialog("Select Image");
                imageFileName = openDialog.getFileName();
                imgFilePath = openDialog.getPath();
                String imgDirectory = openDialog.getDirectory();
                Opener o = new Opener();
                int type = o.getFileType(imgFilePath);
                if (type == Opener.UNKNOWN || type == Opener.JAVA_OR_TEXT || type == Opener.ROI || type == Opener.TEXT
                        || type == Opener.TABLE) {
                    canOpen = false;
                } else {
                    canOpen = true;
                    System.out.println("Image file: " + imgFilePath);
                    imp0 = IJ.openImage(imgFilePath);
                    imp0.show();
                }
            } while (canOpen == false);

            if (outputFileAuto != 1) {
                OpenDialog manualCSVgd = new OpenDialog("Please select .xls file to store results");
                outputFileName = manualCSVgd.getPath();

            } else if (outputFileName == "") {
              Random random = new Random();
              int rand_int2 = random.nextInt(1000);
                outputFileName = imgFilePath + "_" + outputFileNameTS + rand_int2 + "_output.xls";
            }
            System.out.println("Output file: " + outputFileName);

            int result0 = JOptionPane.showConfirmDialog((Component) null,
                    "Do you want to add points through a CSV file for the lines?",
                    "Coordinate Line Input (Cancel exits plugin)", JOptionPane.YES_NO_CANCEL_OPTION);

            switch (result0) {
            case 0:
                lineInputIsCSV = true;
                break;
            case 1:
                lineInputIsCSV = false;
                break;
            case 2:
            case -1:
                System.out.println("Do you want to add points through a CSV file for the lines? : User cancelled");
                return;
            }
            System.out.println("Do you want to add points through a CSV file for the lines? : " + lineInputIsCSV);

            // GenericDialog manualPointsBool = new GenericDialog("Coordinate Line Input");
            // Add a check-box to the dialog
            // manualPointsBool.addCheckbox("Check this box if you want to add points
            // through a CSV for the lines", false);
            // Display the dialog
            // manualPointsBool.showDialog();
            // Check if the OK button was clicked
            // if (manualPointsBool.wasOKed()) {
            imp0 = WindowManager.getCurrentImage();
            // Retrieve the value of the check-box
            // lineInputIsCSV = manualPointsBool.getNextBoolean();
            if (lineInputIsCSV) {
                // Code for CSV point input
                IJ.log("User selected CSV point input.");
                OpenDialog tablePromptLine = new OpenDialog("Choose CSV File for Line 1 (Base)");
                String csvFilePathLine = tablePromptLine.getPath();
                String csvDirectory = tablePromptLine.getDirectory();
                ResultsTable table = ResultsTable.open2(csvFilePathLine);
                xPointsGlobal = new float[table.size()];
                yPointsGlobal = new float[table.size()];
                Overlay overlay = new Overlay();
                for (int i = 0; i < table.size(); i++) {
                    float xCsvValueLine = (float) table.getValueAsDouble(table.getColumnIndex("X"), i);
                    xPointsGlobal[i] = xCsvValueLine;
                    float yCsvValueLine = (float) table.getValueAsDouble(table.getColumnIndex("Y"), i);
                    yPointsGlobal[i] = yCsvValueLine;
                }
                for (int i = 1; i < xPointsGlobal.length; i++) {
                  float dx = (xPointsGlobal[i] - xPointsGlobal[i - 1]);
                  float dy = (yPointsGlobal[i] - yPointsGlobal[i - 1]);
                  lengthLine1 = lengthLine1 + Math.sqrt(dx * dx + dy * dy);
                 System.out.println(lengthLine1 + Math.sqrt(dx * dx + dy * dy));
              }
               PolygonRoi Line1roi = new PolygonRoi(xPointsGlobal, yPointsGlobal, Roi.FREELINE);
               
                RoiManager roiManager = new RoiManager();
                roiManager.add(imp0, Line1roi, -1);
 
                roiManager.select(roiManager.getSelectedIndex());
                roiManager.runCommand("draw");
                int imageID = WindowManager.getCurrentImage().getID();
                int[] imageIDs = WindowManager.getIDList();
                if (imageIDs != null && imageIDs.length > 0) {
                    for (int id : imageIDs) {
                        if (id != imageID) {
                            IJ.selectWindow(id);
                            IJ.run("Close");
                        }
                    }
                }

            } else {
                // LINE 1 CODE
                // ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                lengthLine1 = 0;
                int lining1Success = 0;
                while (lining1Success == 0) {
                    IJ.setTool("polyline");
                    new WaitForUserDialog(
                            "Please draw line 1 (the bottom line), \n This macro will not proceed unless a segmented line or freehand line is selected")
                            .show();
                    ImagePlus imagePlus1 = WindowManager.getCurrentImage();
                    Roi result = imagePlus1.getRoi();

                    if (result != null && (result.getType() == Roi.POLYLINE || result.getType() == Roi.FREELINE)) {
                        // Get the polygon representing the line
                        FloatPolygon linePolygon = result.getInterpolatedPolygon();

                        // Retrieve the coordinates of the line
                        float[] xPoints = linePolygon.xpoints;
                        float[] yPoints = linePolygon.ypoints;

                        // Process the coordinates as desired
                        for (int i = 0; i < linePolygon.npoints; i++) {
                            float x = xPoints[i];
                            float y = yPoints[i];

                        }
                        for (int i = 1; i < xPoints.length; i++) {
                            float dx = (xPoints[i] - xPoints[i - 1]);
                            float dy = (yPoints[i] - yPoints[i - 1]);
                            lengthLine1 = lengthLine1 + Math.sqrt(dx * dx + dy * dy);
                            //System.out.println(lengthLine1 + Math.sqrt(dx * dx + dy * dy));
                        }

                        if (lengthLine1 > 0) {
                            IJ.run("Fit Spline");
                            // TODO: Main problems are, line is too thick and just black and I can't rename
                            // the selection
                            xPointsGlobal = new float[xPoints.length];
                            yPointsGlobal = new float[yPoints.length];
                            for (int i = 0; i < xPoints.length; i++) {
                                xPointsGlobal[i] = xPoints[i];
                            }
                            for (int i = 0; i < yPoints.length; i++) {
                                yPointsGlobal[i] = yPoints[i];
                            }
                            PolygonRoi line1Roi = new PolygonRoi(linePolygon, Roi.POLYLINE);
                            line1Roi.setName("Line 1");
                            PolygonRoi.setColor(Color.YELLOW);
                            PolygonRoi.setDefaultFillColor(Color.yellow);
                            int imageID = WindowManager.getCurrentImage().getID();

                            RoiManager roiManager = new RoiManager();
                            roiManager.add(imp0, line1Roi, -1);
                            roiManager.select(roiManager.getSelectedIndex());
                            roiManager.runCommand("draw");
                            roiManager.runCommand("Show All");
                            int[] imageIDs = WindowManager.getIDList();
                            if (imageIDs != null && imageIDs.length > 0) {
                                for (int id : imageIDs) {
                                    if (id != imageID) {
                                        IJ.selectWindow(id);
                                        IJ.run("Close");
                                    }
                                }
                            }
                            lining1Success = 1;
                        }
                    } else {
                        // clear the line; inform user it was not a proper type of line and ask the user
                        // to try again
                        imagePlus1.deleteRoi();
                        IJ.showMessage("Error!", "Retry: Please draw a segmented line");
                    }
                } // while (lining1Success == 0)

            }

            if (lineInputIsCSV) {
                // Code for CSV point input
                IJ.log("User selected CSV point input.");
                OpenDialog tablePromptLine = new OpenDialog("Choose CSV File for Line 2 (Top)");
                String csvFilePathLine = tablePromptLine.getPath();
                String csvDirectory = tablePromptLine.getDirectory();
                ResultsTable table = ResultsTable.open2(csvFilePathLine);
                xPoints2Global = new float[table.size()];
                yPoints2Global = new float[table.size()];
                Overlay overlay = new Overlay();
                for (int i = 0; i < table.size(); i++) {
                    float xCsvValueLine = (float) table.getValueAsDouble(table.getColumnIndex("X"), i);
                    xPoints2Global[i] = xCsvValueLine;
                    float yCsvValueLine = (float) table.getValueAsDouble(table.getColumnIndex("Y"), i);
                    yPoints2Global[i] = yCsvValueLine;
                }
                PolygonRoi Line2roi = new PolygonRoi(xPoints2Global, yPoints2Global, Roi.FREELINE);
                Line2roi.setStrokeColor(Color.yellow);
                RoiManager roiManager = RoiManager.getInstance2();
                roiManager.add(imp0, Line2roi, -1);
                roiManager.select(roiManager.getSelectedIndex());
                roiManager.runCommand("draw");
                for (int i = 1; i < xPoints2Global.length; i++) {
                  float dx = (xPoints2Global[i] - xPoints2Global[i - 1]);
                  float dy = (yPoints2Global[i] - yPoints2Global[i - 1]);
                  lengthLine2 = lengthLine2 + Math.sqrt(dx * dx + dy * dy);
                  System.out.println(lengthLine2 + Math.sqrt(dx * dx + dy * dy));
              }
                int imageID = WindowManager.getCurrentImage().getID();
                int[] imageIDs = WindowManager.getIDList();
                if (imageIDs != null && imageIDs.length > 0) {
                    for (int id : imageIDs) {
                        if (id != imageID) {
                            IJ.selectWindow(id);
                            IJ.run("Close");
                        }
                    }
                }

            } else {
                // LINE 2 CODE
                // ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                lengthLine2 = 0;
                int lining2Success = 0;
                while (lining2Success == 0) {
                    IJ.setTool("polyline");
                    new WaitForUserDialog(
                            "Please draw line 2 (the top line), \n This macro will not proceed unless a segmented line or freehand line is selected")
                            .show();
                    Roi result = IJ.getImage().getRoi();
                    if (result != null && (result.getType() == Roi.POLYLINE || result.getType() == Roi.FREELINE)) {
                        // Get the polygon representing the line
                        FloatPolygon linePolygon2 = result.getInterpolatedPolygon();

                        // Retrieve the coordinates of the line
                        float[] xPoints2 = linePolygon2.xpoints;
                        float[] yPoints2 = linePolygon2.ypoints;

                        for (int i = 1; i < xPoints2.length; i++) {
                            float dx = (xPoints2[i] - xPoints2[i - 1]);
                            float dy = (yPoints2[i] - yPoints2[i - 1]);
                            lengthLine2 = lengthLine2 + Math.sqrt(dx * dx + dy * dy);
                        }

                        if (lengthLine2 > 0) {
                            IJ.run("Fit Spline");
                            RoiManager roiManager = RoiManager.getInstance();
                            if (roiManager == null) {
                                // If ROI Manager is null, create a new instance
                                roiManager = new RoiManager();
                                roiManager.setVisible(true);
                            }
                            // TODO:
                            xPoints2Global = new float[xPoints2.length];
                            yPoints2Global = new float[yPoints2.length];
                            for (int i = 0; i < xPoints2.length; i++) {
                                xPoints2Global[i] = xPoints2[i];
                            }
                            for (int i = 0; i < yPoints2.length; i++) {
                                yPoints2Global[i] = yPoints2[i];
                            }
                            PolygonRoi line2Roi = new PolygonRoi(linePolygon2, Roi.POLYLINE);
                            line2Roi.setName("Line 2");
                            PolygonRoi.setColor(Color.YELLOW);
                            PolygonRoi.setDefaultFillColor(Color.yellow);
                            int imageID = WindowManager.getCurrentImage().getID();
                            roiManager.add(imp0, line2Roi, -1);

                            roiManager.select(roiManager.getSelectedIndex());
                            roiManager.runCommand("draw");
                            roiManager.runCommand("Show All");
                            int[] imageIDs = WindowManager.getIDList();
                            if (imageIDs != null && imageIDs.length > 0) {
                                for (int id : imageIDs) {
                                    if (id != imageID) {
                                        IJ.selectWindow(id);
                                        IJ.run("Close");
                                    }
                                }
                            }
                            lining2Success = 1;
                        }
                    } else {
                        // clear the line; inform user it was not a proper type of line and ask the user
                        // to try again
                        IJ.getImage().deleteRoi();
                        IJ.showMessage("Error!", "Retry: Please draw a segmented line");
                    }
                }
            }

            int result1 = JOptionPane.showConfirmDialog((Component) null,
                    "Do you want to add points through a CSV file?", "CSV Point Input (Cancel exits plugin)",
                    JOptionPane.YES_NO_CANCEL_OPTION);

            boolean pointInputIsCSV = true;
            switch (result1) {
            case 0:
                pointInputIsCSV = true;
                break;
            case 1:
                pointInputIsCSV = false;
                break;
            case 2:
            case -1:
                System.out.println("Do you want to add points through a CSV file? : User cancelled");
                
                return;
            }
            System.out.println("Do you want to add points through a CSV file? : " + pointInputIsCSV);

            // GenericDialog manualCoordinateBool = new GenericDialog("CSV Point Input");
            // Add a check-box to the dialog
            // manualCoordinateBool.addCheckbox("Check this box if you want to add points
            // through a CSV", false);
            // Display the dialog
            // manualCoordinateBool.showDialog();
            // Check if the OK button was clicked
            // if (manualCoordinateBool.wasOKed()) {
            // Retrieve the value of the check-box
            // manualInput = manualCoordinateBool.getNextBoolean();
            // System.out.println(manualInput);
            // Do something with the check-box value
            if (pointInputIsCSV) {
                // Code for CSV point input
                IJ.log("User selected CSV point input.");
                String markerNumber = JOptionPane.showInputDialog("How many markers will be used on this image?");
                Integer chosenMarkerNumber = Integer.parseInt(markerNumber);
               
                OpenDialog tablePrompt = new OpenDialog("Choose CSV File");
                String csvFilePath = tablePrompt.getPath();
                String csvDirectory = tablePrompt.getDirectory();
                // Initialize a list to store each marker type's coordinates
                List<Float[]> xCoordsPerMarker = new ArrayList<>();
                List<Float[]> yCoordsPerMarker = new ArrayList<>();
                List<Overlay> overlaysPerMarker = new ArrayList<>();
                ResultsTable table = ResultsTable.open2(csvFilePath);
                userPickedXCoordsGlobal = new Float[table.size()];
                userPickedYCoordsGlobal = new Float[table.size()];
                MarkerTypeGlobal = new String[table.size()];
                float[] xCoordsAuto = new float[table.size()];
                float[] yCoordsAuto = new float[table.size()];
                // Set to keep track of processed marker types
                Set<String> processedMarkers = new HashSet<>();
                // Loop through each marker type
                for (int markerIndex = 0; markerIndex < chosenMarkerNumber; markerIndex++) {
                	   String markerName = null;
					// Find the next unprocessed marker type
                    for (int i = 0; i < table.size(); i++) {
                        String currentMarkerType = table.getStringValue(table.getColumnIndex("MarkerType"), i);
                        if (!processedMarkers.contains(currentMarkerType)) {
                            markerName = currentMarkerType;
                            processedMarkers.add(markerName);
                            break;
                        }
                    }
                    
                    // Initialize arrays for this marker type
                    Float[] xCoords = new Float[table.size()];
                    Float[] yCoords = new Float[table.size()];
                    Overlay overlay = new Overlay();

                    // Loop through the table and extract points for this marker type
                    for (int i = 0; i < table.size(); i++) {
                            float xCsvValue = (float) table.getValueAsDouble(table.getColumnIndex("X"), i);
                            float yCsvValue = (float) table.getValueAsDouble(table.getColumnIndex("Y"), i);
                            
                            userPickedXCoordsGlobal[i] = xCsvValue;
                            userPickedYCoordsGlobal[i] = yCsvValue;
                            String markerType = table.getStringValue(table.getColumnIndex("MarkerType"), i);
                            MarkerTypeGlobal[i] = markerType;
                            PointRoi pointRoi = new PointRoi(xCsvValue, yCsvValue);
                            pointRoi.setStrokeColor(Color.GREEN);
                            System.out.println(pointRoi.getStrokeColor());
                            overlay.add(pointRoi);
                            WindowManager.getCurrentImage().setOverlay(overlay);
                    }
                    
                    // Store the coordinates and overlay for this marker type
                    xCoordsPerMarker.add(xCoords);
                    yCoordsPerMarker.add(yCoords);
                    overlaysPerMarker.add(overlay);

                    // Display points for this marker type on the image
                    WindowManager.getCurrentImage().setOverlay(overlay);
                    new WaitForUserDialog("Please click OK to confirm the points for marker type: " + markerName).show();
                }
         
                new WaitForUserDialog("Please click OK to confirm your points").show();
                RoiManager roiManager = RoiManager.getInstance();
                roiManager.setVisible(true);
                roiManager.close();

            }  else {
                // Code for manual point input
                List<Float[]> xCoordsPerMarker = new ArrayList<>();
                List<Float[]> yCoordsPerMarker = new ArrayList<>();
                List<Overlay> overlaysPerMarker = new ArrayList<>();

                // Prompt the user to input the number of markers
                String markerNumber = JOptionPane.showInputDialog("How many markers will be used on this image?");
                Integer chosenMarkerNumber = Integer.parseInt(markerNumber);

                // Initialize global arrays with enough space
                userPickedXCoordsGlobal = new Float[0];
                userPickedYCoordsGlobal = new Float[0];
                MarkerTypeGlobal = new String[0];

                for (int markerIndex = 0; markerIndex < chosenMarkerNumber; markerIndex++) {
                    String markerName = JOptionPane.showInputDialog("Enter the name for marker type " + (markerIndex + 1));

                    boolean pointsGood = false;
                    while (!pointsGood) {
                        // Wait for the user to select points
                        IJ.setTool("multipoint");
                        new WaitForUserDialog("Select points for marker type: " + markerName, "Select points using the multipoint tool").show();
                        Roi thisRoi = WindowManager.getCurrentImage().getRoi();
                        if (thisRoi != null && thisRoi.getType() == Roi.POINT) {
                            pointsGood = true;
                        } else {
                            IJ.getImage().deleteRoi();
                            IJ.showMessage("Error!", "Retry: Wrong type of selection! Select Points");
                        }
                    }

                    PointRoi pointRoi = (PointRoi) WindowManager.getCurrentImage().getRoi();
                    // Get the selected points as a FloatPolygon
                    FloatPolygon floatPolygon = pointRoi.getFloatPolygon();

                    // Create arrays to store the coordinates
                    Float[] xCoords = new Float[floatPolygon.npoints];
                    Float[] yCoords = new Float[floatPolygon.npoints];
                    String[] markerTypes = new String[floatPolygon.npoints];

                    // Populate the coordinate arrays and marker types array
                    for (int i = 0; i < floatPolygon.npoints; i++) {
                        xCoords[i] = floatPolygon.xpoints[i];
                        yCoords[i] = floatPolygon.ypoints[i];
                        markerTypes[i] = markerName;
                    }

                    // Create an overlay for the selected points
                    Overlay overlay = new Overlay();
                    for (int i = 0; i < floatPolygon.npoints; i++) {
                        PointRoi point = new PointRoi(floatPolygon.xpoints[i], floatPolygon.ypoints[i]);
                        point.setStrokeColor(Color.GREEN);
                        overlay.add(point);
                    }

                    // Store the coordinates and overlay for this marker type
                    xCoordsPerMarker.add(xCoords);
                    yCoordsPerMarker.add(yCoords);
                    overlaysPerMarker.add(overlay);

                    // Update the global arrays
                    int existingLength = userPickedXCoordsGlobal.length;
                    int newLength = existingLength + xCoords.length;

                    Float[] newXCoordsGlobal = new Float[newLength];
                    Float[] newYCoordsGlobal = new Float[newLength];
                    String[] newMarkerTypeGlobal = new String[newLength];

                    System.arraycopy(userPickedXCoordsGlobal, 0, newXCoordsGlobal, 0, existingLength);
                    System.arraycopy(userPickedYCoordsGlobal, 0, newYCoordsGlobal, 0, existingLength);
                    System.arraycopy(MarkerTypeGlobal, 0, newMarkerTypeGlobal, 0, existingLength);

                    System.arraycopy(xCoords, 0, newXCoordsGlobal, existingLength, xCoords.length);
                    System.arraycopy(yCoords, 0, newYCoordsGlobal, existingLength, yCoords.length);
                    System.arraycopy(markerTypes, 0, newMarkerTypeGlobal, existingLength, markerTypes.length);

                    userPickedXCoordsGlobal = newXCoordsGlobal;
                    userPickedYCoordsGlobal = newYCoordsGlobal;
                    MarkerTypeGlobal = newMarkerTypeGlobal;

                    // Display points for this marker type on the image
                    WindowManager.getCurrentImage().setOverlay(overlay);
                    new WaitForUserDialog("Please click OK to confirm the points for marker type: " + markerName).show();
                }

                new WaitForUserDialog("Please click OK to confirm your points").show();
                RoiManager roiManager = RoiManager.getInstance();
                roiManager.setVisible(true);
                roiManager.close();
            }
            
            DirectoryChooser directoryChooserImage = new DirectoryChooser("Select Folder to save your drawn Image");
            String savedImageFilePath = directoryChooserImage.getDirectory();
            // SimpleDateFormat is not thread safe, careful when using multithreading
            // (future prob)
            // String timeStamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss").format(new
            // java.util.Date());
            String filePath = savedImageFilePath + imageFileName + "_" + outputFileNameTS + ".tiff";
            FileSaver fileSaver = new FileSaver(imp0);

            fileSaver.saveAsTiff(filePath);

            // LINE 1 DISTANCE CALC///////////////////////////////////////////////////
            // roiManager.close();
            bottomLineArrayDistancesGlobal = new float[userPickedXCoordsGlobal.length];
            for (int i = 0; i < userPickedXCoordsGlobal.length; i++) {
                // This first point is for the edge case where the closest point is the first
                // one chosen
                bottomLineArrayDistancesGlobal[i] = (float) sqrt((xPointsGlobal[0] - userPickedXCoordsGlobal[i])
                        * (xPointsGlobal[0] - userPickedXCoordsGlobal[i])
                        + (yPointsGlobal[0] - userPickedYCoordsGlobal[i])
                                * (yPointsGlobal[0] - userPickedYCoordsGlobal[i]));
                for (int k = 1; k < xPointsGlobal.length - 1; k++) {
                    float firstCompVal = (float) sqrt((xPointsGlobal[k - 1] - userPickedXCoordsGlobal[i])
                            * (xPointsGlobal[k - 1] - userPickedXCoordsGlobal[i])
                            + (yPointsGlobal[k - 1] - userPickedYCoordsGlobal[i])
                                    * (yPointsGlobal[k - 1] - userPickedYCoordsGlobal[i]));
                    float secondCompVal = (float) sqrt((xPointsGlobal[k] - userPickedXCoordsGlobal[i])
                            * (xPointsGlobal[k] - userPickedXCoordsGlobal[i])
                            + (yPointsGlobal[k] - userPickedYCoordsGlobal[i])
                                    * (yPointsGlobal[k] - userPickedYCoordsGlobal[i]));
                    if (firstCompVal > secondCompVal) {
                        bottomLineArrayDistancesGlobal[i] = secondCompVal;
                        
                    }
                }
            }

            // LINE 2 DISTANCE CALC///////////////////////////////////////////////////
            topLineArrayDistancesGlobal = new float[userPickedXCoordsGlobal.length];
            for (int i = 0; i < userPickedXCoordsGlobal.length; i++) {
                // This first point is for the edge case where the closest point is the first
                // one chosen
                topLineArrayDistancesGlobal[i] = (float) sqrt((xPoints2Global[0] - userPickedXCoordsGlobal[i])
                        * (xPoints2Global[0] - userPickedXCoordsGlobal[i])
                        + (yPoints2Global[0] - userPickedYCoordsGlobal[i])
                                * (yPoints2Global[0] - userPickedYCoordsGlobal[i]));
                for (int k = 1; k < xPoints2Global.length - 1; k++) {
                    float firstCompVal = (float) sqrt((xPoints2Global[k - 1] - userPickedXCoordsGlobal[i])
                            * (xPoints2Global[k - 1] - userPickedXCoordsGlobal[i])
                            + (yPoints2Global[k - 1] - userPickedYCoordsGlobal[i])
                                    * (yPoints2Global[k - 1] - userPickedYCoordsGlobal[i]));
                    float secondCompVal = (float) sqrt((xPoints2Global[k] - userPickedXCoordsGlobal[i])
                            * (xPoints2Global[k] - userPickedXCoordsGlobal[i])
                            + (yPoints2Global[k] - userPickedYCoordsGlobal[i])
                                    * (yPoints2Global[k] - userPickedYCoordsGlobal[i]));
                    if (firstCompVal > secondCompVal) {
                        topLineArrayDistancesGlobal[i] = secondCompVal;
                        
                    }
                }
            }
            //Average Length Distance Calc
            distanceToDivide = 0;
            distanceToDivideArray = new float[xPoints2Global.length + 1];
            for (int i = 0; i < xPoints2Global.length; i++) {
                // This first point is for the edge case where the closest point is the first one chosen
                distanceToDivideArray[i] = (float) Math.sqrt((xPointsGlobal[0] - xPoints2Global[i])
                        * (xPointsGlobal[0] - xPoints2Global[i])
                        + (yPointsGlobal[0] - yPoints2Global[i])
                        * (yPointsGlobal[0] - yPoints2Global[i]));
                for (int k = 1; k < xPointsGlobal.length - 1; k++) { // Use xPointsGlobal.length for the outer loop
                    float firstCompVal = (float) Math.sqrt((xPointsGlobal[k - 1] - xPoints2Global[i])
                            * (xPointsGlobal[k - 1] - xPoints2Global[i])
                            + (yPointsGlobal[k - 1] - yPoints2Global[i])
                            * (yPointsGlobal[k - 1] - yPoints2Global[i]));
                    float secondCompVal = (float) Math.sqrt((xPointsGlobal[k] - xPoints2Global[i])
                            * (xPointsGlobal[k] - xPoints2Global[i])
                            + (yPointsGlobal[k] - yPoints2Global[i])
                            * (yPointsGlobal[k] - yPoints2Global[i]));
                    if (firstCompVal > secondCompVal) {
                        distanceToDivideArray[i] = secondCompVal;
                    }
                }
            }


            

for (int i = 0; i < distanceToDivideArray.length; i++) {
    distanceToDivide += distanceToDivideArray[i];
}

            try {
                
                String filename = outputFileName;
                HSSFWorkbook workbook = new HSSFWorkbook();
                HSSFSheet sheet = workbook.createSheet(imageFileName);
                HSSFRow rowhead = sheet.createRow((short) 0);
                // creating cell by using the createCell() method and setting the values to the
                // cell by using the setCellValue() method
                rowhead.createCell(0).setCellValue("Centroid Coordinates (x)");
                rowhead.createCell(1).setCellValue("Centroid Coordinates (y)");
                rowhead.createCell(2).setCellValue("Distance from the base to each point in order");
                rowhead.createCell(3).setCellValue("Distance from the top to each point in order");
                rowhead.createCell(4).setCellValue("Normalized Distance");
                rowhead.createCell(5).setCellValue("Length of Base");
                rowhead.createCell(6).setCellValue("Length of Top");
                rowhead.createCell(7).setCellValue("Chosen Bin Interval");
                rowhead.createCell(8).setCellValue("Average Distance between Lines");
                rowhead.createCell(9).setCellValue("Marker Type");
                for(int i = 1; i < userPickedXCoordsGlobal.length + 1; i++) {
                	System.out.println("Marker Type is: " + MarkerTypeGlobal[i-1]);
                }
                HSSFRow[] rowArray = new HSSFRow[999];
                // All of the plus ones are to make space for the title row created above
                for (int i = 1; i < userPickedXCoordsGlobal.length + 1; i++) {
                    rowArray[i] = sheet.createRow((short) i);
                    rowArray[i].createCell(0).setCellValue(userPickedXCoordsGlobal[i - 1]);
                    rowArray[i].createCell(1).setCellValue(userPickedYCoordsGlobal[i - 1]);
                    rowArray[i].createCell(2).setCellValue(bottomLineArrayDistancesGlobal[i - 1]);
                    rowArray[i].createCell(3).setCellValue(topLineArrayDistancesGlobal[i - 1]);
                    rowArray[i].createCell(4).setCellValue(bottomLineArrayDistancesGlobal[i - 1]
                            / (bottomLineArrayDistancesGlobal[i - 1] + topLineArrayDistancesGlobal[i - 1]));
                    rowArray[i].createCell(9).setCellValue(MarkerTypeGlobal[i-1]);

                }
                rowArray[1].createCell(5).setCellValue(lengthLine1);
                rowArray[1].createCell(6).setCellValue(lengthLine2);
                rowArray[1].createCell(7).setCellValue(chosenInterval);
                rowArray[1].createCell(8).setCellValue(distanceToDivide / (distanceToDivideArray.length));
                FileOutputStream fileOut = new FileOutputStream(filename);
                workbook.write(fileOut);
                // closing the Stream
                fileOut.close();
                // closing the workbook
                workbook.close();
                // prints the message on the console
                System.out.println("Excel file has been generated successfully.");
            } catch (Exception e) {
                e.printStackTrace();
            }

            int result2 = JOptionPane.showConfirmDialog((Component) null,
                    "Do you want to repeat this macro for a new image? Clicking No will close the macro and generate the session histogram",
                    "Batch Processing", JOptionPane.YES_NO_OPTION);

            boolean repeatInput = false;
            switch (result2) {
            case 0:
                repeatInput = true;
                break;
            case 1:
            case -1:
                repeatInput = false;
                break;
            }
            System.out.println("User wants to repeat macro : " + repeatInput);


            if (repeatInput) {
                // Yes Repeat
                ApachePoiLineChart(chosenInterval);
                IJ.getImage().close();
                repeatFlag = 1;
            } else {
                // No repeat
                notConclusionRun = false;
                ApachePoiLineChart(chosenInterval);
                repeatFlag = 0;
            }

        }

    }

    public static String makeOutputFileNameTS() {
        Calendar calendar = Calendar.getInstance();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy_MMdd_HHmmss");
        return dateFormat.format(calendar.getTime());
    }


    public void ApachePoiLineChart(double chosenInterval) {
        // Create a DirectoryChooser to prompt the user for a folder
        DirectoryChooser directoryChooser = new DirectoryChooser("Select Folder to save your Histogram and bincount");

        // Display the DirectoryChooser dialog and wait for user input
        selectedFolderPath = directoryChooser.getDirectory();

        // Check if the user canceled the selection
        if (selectedFolderPath == null) {
            IJ.log("Folder selection canceled.");
            return;
        }

        // Display the selected folder path
        IJ.log("Selected folder path: " + selectedFolderPath);

        // Retrieve the selected folder path from the SaveDialog

        XSSFWorkbook wb = new XSSFWorkbook();
        Calendar calendar = Calendar.getInstance();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy_MMdd_HHmmss");
        String dateTime = dateFormat.format(calendar.getTime());
        
        String sheetName = "MacroSessionSummary" + dateTime;
        XSSFSheet sheet = wb.createSheet(sheetName);
    

        try {
            lineChartIterationCount = lineChartIterationCount + 1;
            
            // Bin Ranges (x axis)
            Row row = sheet.createRow((short) 0);
            Cell cell;
            int columnIndex = 0;
            double epsilon = 1e-10; 
               for (double i = 0; i <= 1.0 + epsilon; i += chosenInterval) {
                    cell = row.createCell(columnIndex++);
                    cell.setCellValue(i);
                }
               
               double columnIndexRow2 = 0;
               epsilon = 1e-10; // Choose an appropriate epsilon value based on the required precision
               for (double k = 0; k <= 1.0 + epsilon; k += chosenInterval) {
                   XSSFRow firstDataRow = sheet.getRow(1);
                    if (firstDataRow == null) {
                        firstDataRow = sheet.createRow(1);
                    }
                   BincountHisto = 0;
                   cell = firstDataRow.createCell((int) columnIndexRow2++);
                   for (int i = 1; i < userPickedXCoordsGlobal.length + 1; i++) {
                        
                      
                        if ((bottomLineArrayDistancesGlobal[i - 1]
                                / (bottomLineArrayDistancesGlobal[i - 1] + topLineArrayDistancesGlobal[i - 1])) < (k+chosenInterval)
                                && bottomLineArrayDistancesGlobal[i - 1]
                                        / (bottomLineArrayDistancesGlobal[i - 1] + topLineArrayDistancesGlobal[i - 1]) > k ) {
                
                            globalBinCountArray[(int) columnIndexRow2]++;
                            
                       
                            cell.setCellValue(((double)globalBinCountArray[(int) columnIndexRow2]) / lineChartIterationCount);
                            
                        
                            
                            
                            
                        
                        }

                        

                    }
               }
               

            

            // Write output to an excel file
            Random random = new Random();
            int rand_int1 = random.nextInt(1000);
            String filename = imgFilePath + "_" + outputFileNameTS + + rand_int1 + "_SessionHistogram.xlsx";
            try (FileOutputStream fileOut = new FileOutputStream(filename)) {
                wb.write(fileOut);
                wb.close();
                fileOut.close();
            }
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
        //Code for image specific histogram
        
    }

    private boolean checkIfRowIsEmpty(Row row) {
        if (row == null) {
            return true;
        }
        if (row.getLastCellNum() <= 0) {
            return true;
        }
        for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
            Cell cell = row.getCell(cellNum);
            if (cell != null && cell.getCellType() != CellType.BLANK && StringUtils.isNotBlank(cell.toString())) {
                return false;
            }
        }
        return true;
    }
}
