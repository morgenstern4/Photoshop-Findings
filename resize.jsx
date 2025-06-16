// Batch Image Resize Script for Photoshop CC 2019
// Resizes images to 195x247 pixels

// Target dimensions
var targetWidth = 195;   // Width in pixels
var targetHeight = 247;  // Height in pixels

// Supported image formats
var imageExtensions = [".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif"];

// Resampling method options:
// ResampleMethod.BICUBIC - Good for smooth gradients
// ResampleMethod.BICUBICSHARPER - Good for reducing size (recommended)
// ResampleMethod.BICUBICSMOOTHER - Good for enlarging
// ResampleMethod.BILINEAR - Faster but lower quality
var resampleMethod = ResampleMethod.BICUBICSHARPER;

function main() {
    // Select input folder
    alert("Please select the folder containing images to resize");
    var inputFolder = Folder.selectDialog("Select Input Folder (Images to Resize)");
    if (!inputFolder) {
        alert("No input folder selected. Script cancelled.");
        return;
    }
    
    // Select output folder
    alert("Please select the output folder for resized images");
    var outputFolder = Folder.selectDialog("Select Output Folder (Resized Images)");
    if (!outputFolder) {
        alert("No output folder selected. Script cancelled.");
        return;
    }
    
    // Create output folder if it doesn't exist
    if (!outputFolder.exists) {
        outputFolder.create();
    }
    
    // Get all image files
    var imageFiles = [];
    for (var i = 0; i < imageExtensions.length; i++) {
        var files = inputFolder.getFiles("*" + imageExtensions[i]);
        imageFiles = imageFiles.concat(files);
    }
    
    if (imageFiles.length === 0) {
        alert("No supported image files found in the selected folder.\nSupported formats: " + imageExtensions.join(", "));
        return;
    }
    
    // Ask user about resize method
    var resizeChoice = confirm(
        "Choose resize method:\n\n" +
        "OK (Yes) = Stretch to exact dimensions (195×247)\n" +
        "Cancel (No) = Maintain aspect ratio (fit within 195×247)\n\n" +
        "Note: Your images have 1.5:1.9 ratio which matches 195:247 exactly,\n" +
        "so both methods will give the same result."
    );
    
    var processedCount = 0;
    var skippedCount = 0;
    
    // Process each image
    for (var i = 0; i < imageFiles.length; i++) {
        var imageFile = imageFiles[i];
        
        try {
            if (resizeChoice) {
                resizeImageStretch(imageFile, outputFolder);
            } else {
                resizeImageFit(imageFile, outputFolder);
            }
            processedCount++;
            $.writeln("Processed: " + imageFile.name);
        } catch (e) {
            alert("Error processing " + imageFile.name + ": " + e.message);
            skippedCount++;
        }
    }
    
    // Show completion message
    alert("Resize complete!\n" +
          "Processed: " + processedCount + " images\n" +
          "Skipped: " + skippedCount + " images\n" +
          "Output folder: " + outputFolder.fsName);
}

function resizeImageStretch(imageFile, outputFolder) {
    // Open the image
    var doc = app.open(imageFile);
    
    try {
        // Set ruler units to pixels
        var originalUnits = app.preferences.rulerUnits;
        app.preferences.rulerUnits = Units.PIXELS;
        
        // Resize to exact dimensions (stretch if needed)
        doc.resizeImage(targetWidth, targetHeight, null, resampleMethod);
        
        // Save the resized image with original filename
        var fileName = getFileNameWithoutExtension(imageFile.name);
        var outputFile = new File(outputFolder.fsName + "/" + fileName + ".jpg");
        
        // Save as JPEG with high quality
        var jpegOptions = new JPEGSaveOptions();
        jpegOptions.quality = 12; // Maximum quality
        jpegOptions.embedColorProfile = true;
        jpegOptions.formatOptions = FormatOptions.STANDARDBASELINE;
        jpegOptions.matte = MatteType.NONE;
        
        doc.saveAs(outputFile, jpegOptions);
        
        // Restore original units
        app.preferences.rulerUnits = originalUnits;
        
    } finally {
        // Close the document
        doc.close(SaveOptions.DONOTSAVECHANGES);
    }
}

function resizeImageFit(imageFile, outputFolder) {
    // Open the image
    var doc = app.open(imageFile);
    
    try {
        // Set ruler units to pixels
        var originalUnits = app.preferences.rulerUnits;
        app.preferences.rulerUnits = Units.PIXELS;
        
        // Get current dimensions
        var currentWidth = doc.width.value;
        var currentHeight = doc.height.value;
        
        // Calculate scaling to fit within target dimensions
        var scaleX = targetWidth / currentWidth;
        var scaleY = targetHeight / currentHeight;
        var scale = Math.min(scaleX, scaleY); // Use smaller scale to fit within bounds
        
        // Calculate new dimensions
        var newWidth = Math.round(currentWidth * scale);
        var newHeight = Math.round(currentHeight * scale);
        
        // Resize maintaining aspect ratio
        doc.resizeImage(newWidth, newHeight, null, resampleMethod);
        
        // If image is smaller than target, create a canvas and center the image
        if (newWidth < targetWidth || newHeight < targetHeight) {
            // Resize canvas to target dimensions
            var deltaWidth = targetWidth - newWidth;
            var deltaHeight = targetHeight - newHeight;
            
            doc.resizeCanvas(targetWidth, targetHeight, AnchorPosition.MIDDLECENTER);
        }
        
        // Save the resized image with original filename
        var fileName = getFileNameWithoutExtension(imageFile.name);
        var outputFile = new File(outputFolder.fsName + "/" + fileName + ".jpg");
        
        // Save as JPEG with high quality
        var jpegOptions = new JPEGSaveOptions();
        jpegOptions.quality = 12; // Maximum quality
        jpegOptions.embedColorProfile = true;
        jpegOptions.formatOptions = FormatOptions.STANDARDBASELINE;
        jpegOptions.matte = MatteType.NONE;
        
        doc.saveAs(outputFile, jpegOptions);
        
        // Restore original units
        app.preferences.rulerUnits = originalUnits;
        
    } finally {
        // Close the document
        doc.close(SaveOptions.DONOTSAVECHANGES);
    }
}

function getFileNameWithoutExtension(fileName) {
    var lastDot = fileName.lastIndexOf(".");
    if (lastDot === -1) return fileName;
    return fileName.substring(0, lastDot);
}

// Run the script
main();