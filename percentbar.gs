function percentbar() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var rich = SpreadsheetApp.newRichTextValue();
    var style = SpreadsheetApp.newTextStyle();

    var accumulative_val = 0.00
    var accumulative_str = ""
    var Allocated_val = ""

    for (var mot = 2; mot++; mot < 170) {
        var addrvar = "H" + mot.toString();

        var rng = SpreadsheetApp.getActiveSheet().getRange("H" + mot.toString());
        var accumrng = SpreadsheetApp.getActiveSheet().getRange("D" + mot.toString());

        var val = rng.getValue().toString();
        var valtemp = val.split("\n")

        accumulative_val = accumulative_val + (accumrng.getValue() * 100);
        accumulative_val = parseFloat(accumulative_val)


        accumulative_str = " " + parseFloat(accumulative_val).toFixed(2).toString() + "%";
        Allocated_val = " " + parseFloat(accumrng.getValue() * 100).toFixed(2).toString() + "%";

        var baseoffset = valtemp[0].length;

        valtemp[0] = valtemp[0] + " " + Allocated_val;
        val = valtemp.join("\n");

        if (insertT(val, "fir", "", "#b5e1f9", 10, addrvar, Allocated_val.length, baseoffset) == 1) {
            val = rng.getValue().toString();
            if (insertT(val, " " + accumulative_str + "", " ", "#be4a83", 8, addrvar) == 1) {
                if (mot == 3) {
                    val = rng.getValue().toString();
                    insertT(val, "    ~ Accumulative allocations", " ", "#be4a83", 8, addrvar)
                }
            }
        }

    }

    return

}


function insertT(textV, textToadd, seperatorText, colorV, fontisizeV, rangeV, lenV, baseoffset) {

    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(colorV).setFontSize(fontisizeV).build(); // Please set this for the additional text.

    var ss = sheet.getSheets()[0];

    var rangeVar = ss.getRange(rangeV);
    var rich = rangeVar.getRichTextValue();
    var existingStyles = rich.getRuns().map(e => ({ start: e.getStartIndex(), end: e.getEndIndex(), style: e.getTextStyle() }));


    var valtemp = rangeVar.getValue().toString();


    value = textV;


    var startOffset = value.length + 1;


    rangeVar.setValue(value + "" + textToadd)
    if (textToadd == "fir") {
        var valtemp = value.split("\n")
        startOffset = baseoffset + 1;
        existingStyles.push({ start: startOffset, end: startOffset + lenV, style: textStyle });
        textToadd = "";

    } else {
        existingStyles.push({ start: startOffset, end: startOffset + textToadd.length, style: textStyle });
    }

    var richTexts = SpreadsheetApp.newRichTextValue().setText(value + " " + textToadd);
    existingStyles.forEach(e => richTexts.setTextStyle(e.start, e.end, e.style));
    rangeVar.setRichTextValue(richTexts.build());
    rangeVar.setNumberFormat('@STRING@');

    return 1;

}