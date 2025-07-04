// ID Card PSD Automation Script for Photoshop CC 2019
// This script processes ID card PSDs and places student images at specified coordinates

// Configuration - Paths will be selected via dialog boxes

// Image positioning (center point coordinates)
var imageCenterX = 394.0;  // X position for image center
var imageCenterY = 459.0;  // Y position for image center

// Image dimensions
var imageWidth = 195;   // Image width in pixels
var imageHeight = 247;  // Image height in pixels

// Supported image formats
var imageExtensions = [".jpg", ".jpeg", ".png", ".tif", ".tiff"];

function main() {
    // Select folders using dialog boxes
    alert("Please select the folder containing your PSD files");
    var psdFolder = Folder.selectDialog("Select PSD Folder");
    if (!psdFolder) {
        alert("No PSD folder selected. Script cancelled.");
        return;
    }
    
    alert("Please select the folder containing your student images");
    var imagesFolder = Folder.selectDialog("Select Images Folder");
    if (!imagesFolder) {
        alert("No images folder selected. Script cancelled.");
        return;
    }
    
    alert("Please select the output folder where processed JPEGs will be saved");
    var outputFolder = Folder.selectDialog("Select Output Folder");
    if (!outputFolder) {
        alert("No output folder selected. Script cancelled.");
        return;
    }
    
    // Create output folder if it doesn't exist
    if (!outputFolder.exists) {
        outputFolder.create();
    }
    
    // Get all PSD files
    var psdFiles = psdFolder.getFiles("*.psd");
    
    if (psdFiles.length === 0) {
        alert("No PSD files found in the selected folder.");
        return;
    }
    
    var processedCount = 0;
    var skippedCount = 0;
    
    // Process each PSD file
    for (var i = 0; i < psdFiles.length; i++) {
        var psdFile = psdFiles[i];
        var rollNo = getRollNoFromFileName(psdFile.name);
        
        // Find corresponding image file
        var imageFile = findImageFile(imagesFolder, rollNo);
        
        if (imageFile) {
            try {
                processIDCard(psdFile, imageFile, rollNo, outputFolder);
                processedCount++;
                $.writeln("Processed: " + rollNo);
            } catch (e) {
                alert("Error processing " + rollNo + ": " + e.message);
                skippedCount++;
            }
        } else {
            $.writeln("No image found for roll no: " + rollNo);
            skippedCount++;
        }
    }
    
    // Show completion message
    alert("Processing complete!\nProcessed: " + processedCount + "\nSkipped: " + skippedCount);
}

function getRollNoFromFileName(fileName) {
    // Remove .psd extension and return roll number
    return fileName.replace(/\.psd$/i, "");
}

function findImageFile(folder, rollNo) {
    // Try to find image file with matching roll number
    for (var i = 0; i < imageExtensions.length; i++) {
        var imageFile = new File(folder.fsName + "/" + rollNo + imageExtensions[i]);
        if (imageFile.exists) {
            return imageFile;
        }
    }
    return null;
}

function processIDCard(psdFile, imageFile, rollNo, outputFolder) {
    // Open the PSD file
    var doc = app.open(psdFile);
    
    try {
        // Set ruler units to pixels
        var originalUnits = app.preferences.rulerUnits;
        app.preferences.rulerUnits = Units.PIXELS;
        
        // Place the image
        var placedLayer = placeImageFile(doc, imageFile);
        
        if (placedLayer) {
            // Position the image
            positionLayer(placedLayer, imageCenterX, imageCenterY);
            
            // Flatten the image before saving as JPEG
            doc.flatten();
            
            // Save the document as JPEG
            var outputFile = new File(outputFolder.fsName + "/" + rollNo + ".jpg");
            var jpegOptions = new JPEGSaveOptions();
            jpegOptions.quality = 12; // Maximum quality (1-12)
            jpegOptions.embedColorProfile = true;
            jpegOptions.formatOptions = FormatOptions.STANDARDBASELINE;
            doc.saveAs(outputFile, jpegOptions);
        }
        
        // Restore original units
        app.preferences.rulerUnits = originalUnits;
        
    } finally {
        // Close the document
        doc.close(SaveOptions.DONOTSAVECHANGES);
    }
}

function placeImageFile(doc, imageFile) {
    try {
        // Use the place method to add the image as a smart object
        var idPlc = charIDToTypeID("Plc ");
        var desc = new ActionDescriptor();
        desc.putPath(charIDToTypeID("null"), imageFile);
        desc.putBoolean(charIDToTypeID("LnkD"), true); // Link the file
        executeAction(idPlc, desc);
        
        // Return the active layer (newly placed image)
        return doc.activeLayer;
        
    } catch (e) {
        // If place fails, try opening and copying the image
        return placeImageAlternative(doc, imageFile);
    }
}

function placeImageAlternative(doc, imageFile) {
    try {
        // Open the image file
        var imageDoc = app.open(imageFile);
        
        // Select all and copy
        imageDoc.selection.selectAll();
        imageDoc.selection.copy();
        
        // Close the image document
        imageDoc.close(SaveOptions.DONOTSAVECHANGES);
        
        // Switch back to the main document and paste
        app.activeDocument = doc;
        doc.paste();
        
        // Return the active layer (pasted image)
        return doc.activeLayer;
        
    } catch (e) {
        throw new Error("Failed to place image: " + e.message);
    }
}

function positionLayer(layer, centerX, centerY) {
    try {
        // Get layer bounds
        var bounds = layer.bounds;
        
        // Calculate current center position
        var currentCenterX = (bounds[0].value + bounds[2].value) / 2;
        var currentCenterY = (bounds[1].value + bounds[3].value) / 2;
        
        // Calculate offset needed to move center to target position
        var deltaX = centerX - currentCenterX;
        var deltaY = centerY - currentCenterY;
        
        // Translate the layer to position center at target coordinates
        layer.translate(deltaX, deltaY);
        
    } catch (e) {
        throw new Error("Failed to position layer: " + e.message);
    }
}

// Run the script
main();