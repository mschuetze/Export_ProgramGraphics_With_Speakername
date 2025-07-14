// Version: 0.1.7

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALERTS;

var doc = app.activeDocument;
var myFolder = doc.filePath;
var jpg_name = doc.name.replace(/\.indd$/i, "");

// ---------- Utility-Funktionen ----------
function cleanString(text) {
    text = text.replace(/[\u200B-\u200D\uFEFF]/g, '');
    text = text.toLowerCase();
    text = text.replace(/^\s+|\s+$/g, ''); // statt .trim() – kompatibel mit ExtendScript
    return text;
}

function replaceGermanUmlauts(text) {
    return text
        .replace(/Ä/g, 'Ae').replace(/ä/g, 'ae')
        .replace(/Ö/g, 'Oe').replace(/ö/g, 'oe')
        .replace(/Ü/g, 'Ue').replace(/ü/g, 'ue')
        .replace(/ß/g, 'ss')
        .replace(/Á/g, 'A').replace(/á/g, 'a')
        .replace(/É/g, 'E').replace(/é/g, 'e')
        .replace(/È/g, 'E').replace(/è/g, 'e')
        .replace(/Ç/g, 'C').replace(/ç/g, 'c')
        .replace(/ñ/g, 'n'); 
}

// ---------- Dateiname-Teile ----------
var fileNamePart1 = jpg_name.split(/_GT-\d{1,5}_/g)[0];
var fileNamePart2 = jpg_name.match(/_GT-\d{1,5}_/g) || [""];
var fileNamePart3 = jpg_name.match(/v\d{1,2}([a-z]*)/g) || [""];

// ---------- Dialog ----------
var dialog = new Window("dialog", "Export-Optionen");
dialog.orientation = "column";
dialog.alignChildren = ["fill", "top"];

// Dateiname enthält (über gesamte Breite)
var namingGroup = dialog.add("group");
namingGroup.orientation = "row"; // horizontal, aber nur 1 Element -> füllt Breite
namingGroup.alignChildren = ["fill", "center"];
namingGroup.add("statictext", undefined, "Dateiname enthält:");
var namingDropdown = namingGroup.add("dropdownlist", undefined, ["ohne", "Maße"]);
namingDropdown.selection = 1; // Standard: Maße
namingDropdown.preferredSize.width = 200; // Optional: Breite definieren

// Gruppe für Auflösung + Format nebeneinander
var bottomGroup = dialog.add("group");
bottomGroup.orientation = "row";
bottomGroup.alignChildren = ["fill", "center"];

// Auflösung
var resolutionGroup = bottomGroup.add("group");
resolutionGroup.orientation = "row";
resolutionGroup.alignChildren = ["fill", "center"];
resolutionGroup.add("statictext", undefined, "Auflösung:");
var resolutionDropdown = resolutionGroup.add("dropdownlist", undefined, ["72", "144", "300"]);
resolutionDropdown.selection = 0;
resolutionDropdown.preferredSize.width = 100; // Breite anpassen

// Format
var formatGroup = bottomGroup.add("group");
formatGroup.orientation = "row";
formatGroup.alignChildren = ["fill", "center"];
formatGroup.add("statictext", undefined, "Format:");
var formatDropdown = formatGroup.add("dropdownlist", undefined, ["JPG", "PNG"]);
formatDropdown.selection = 0;
formatDropdown.preferredSize.width = 100;

// Buttons
var buttonGroup = dialog.add("group");
buttonGroup.alignment = "right";
buttonGroup.add("button", undefined, "Abbrechen", { name: "cancel" });
buttonGroup.add("button", undefined, "OK", { name: "ok" });

if (dialog.show() !== 1) exit();

var selectedFormat = formatDropdown.selection.text;
var selectedResolution = parseInt(resolutionDropdown.selection.text, 10);
var namingOption = namingDropdown.selection.text;  // "ohne" oder "Maße"

// ---------- Funktion: Sprecher-Nachnamen extrahieren ----------
function getFormattedLastNamesFromPage(page) {
    var lastNames = [];

    for (var i = 0; i < page.textFrames.length; i++) {
        var tf = page.textFrames[i];
        var paras = tf.paragraphs;

        for (var j = 0; j < paras.length; j++) {
            var para = paras[j];

            if (para.appliedParagraphStyle.name === "speakers") {
                var text = para.contents;

                // Kürzen bei Pipe
                if (text.indexOf("|") !== -1) {
                    text = text.split("|")[0];
                }

                // Split bei Kommas für mehrere Namen
                var speakerChunks = text.split(",");

                for (var k = 0; k < speakerChunks.length; k++) {
                    var person = cleanString(speakerChunks[k]);
                    person = replaceGermanUmlauts(person);
                    person = person.replace(/^\s+|\s+$/g, ''); // trim-Ersatz

                    var nameParts = person.split(/\s+/);
                    if (nameParts.length === 0) continue;

                    var lastName = nameParts[nameParts.length - 1]; // Letztes Wort = Nachname

                    // Doppelnamen erkennen (Bindestrich behalten für CamelCase-Zerlegung)
                    var namePartsSplit = lastName.split("-");

                    var formattedParts = [];
                    for (var p = 0; p < namePartsSplit.length; p++) {
                        var part = namePartsSplit[p].replace(/[^a-zA-Z0-9]/g, '');
                        if (part.length > 0) {
                            part = part.charAt(0).toUpperCase() + part.slice(1).toLowerCase();
                            formattedParts.push(part);
                        }
                    }

                    lastNames.push(formattedParts.join("")); // Ohne Bindestrich, aber CamelCase
                }

                return lastNames.join("_"); // Neu: mit Unterstrich getrennt
            }
        }
    }

    return "Unknown";
}

// ---------- Export ----------
for (var i = 0; i < doc.pages.length; i++) {
    var page = doc.pages[i];

    var speakerLastName = getFormattedLastNamesFromPage(page);

    var fileName = fileNamePart1 + "_" + speakerLastName;

    if (namingOption === "Maße") {
        var bounds = page.bounds;
        var width = Math.round(bounds[3] - bounds[1]);
        var height = Math.round(bounds[2] - bounds[0]);
        fileName += "_" + width + "x" + height;
    }

    if (fileNamePart2[0]) fileName += fileNamePart2[0];
    if (fileNamePart3[0]) {
        if (fileNamePart2[0]) {
            fileName += fileNamePart3[0];
        } else {
            fileName += "_" + fileNamePart3[0];
        }
    }

    fileName += "_P" + (page.documentOffset + 1);

    var fileExtension = (selectedFormat === "JPG") ? ".jpg" : ".png";
    fileName += fileExtension;

    var file = new File(myFolder + "/" + fileName);
    var pageString = page.name;

    if (selectedFormat === "JPG") {
        app.jpegExportPreferences.properties = {
            exportResolution: selectedResolution,
            jpegQuality: JPEGOptionsQuality.HIGH,
            jpegExportRange: ExportRangeOrAllPages.EXPORT_RANGE,
            pageString: pageString
        };
        try {
            doc.exportFile(ExportFormat.JPG, file);
        } catch (e) {
            $.writeln("Fehler beim JPG-Export Seite " + page.name + ": " + e.message);
        }
    } else {
        app.pngExportPreferences.properties = {
            exportResolution: selectedResolution,
            pngQuality: PNGQualityEnum.HIGH,
            pngExportRange: ExportRangeOrAllPages.EXPORT_RANGE,
            pageString: pageString,
            transparentBackground: true
        };
        try {
            doc.exportFile(ExportFormat.PNG_FORMAT, file);
        } catch (e) {
            $.writeln("Fehler beim PNG-Export Seite " + page.name + ": " + e.message);
        }
    }

    $.writeln("Exportiert: " + file.name);
}

alert("Export abgeschlossen!");
