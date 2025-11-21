#target photoshop
app.bringToFront();

if (documents.length == 0) { alert("Vui lòng mở một file trước!"); exit(); }

var doc = activeDocument;
var originalFileName = doc.name.replace(/\.[^\.]+$/, ""); // Lấy tên file không có đuôi

// --- Lấy thông số từ file ---
var inputWidth = doc.width.as("mm");
var inputHeight = doc.height.as("mm");
var overlap = 20;  // mm
var margin = 50;   // mm tổng 2 cạnh
var decalWidths = [914, 1070, 1270, 1520]; // mm

alert("File hiện tại: " + inputWidth.toFixed(1) + " x " + inputHeight.toFixed(1) + " mm");

// --- Hàm tính optimize ---
function makeOptimizePlan(inputWidth, decalWidths, margin, overlap, inputHeight) {
    var plans = [];
    for (var i = 0; i < decalWidths.length; i++) {
        var r = decalWidths[i];
        var maxIn = r - margin;
        var n = 1;
        while (true) {
            var s = Math.ceil((inputWidth + (n-1)*overlap)/n);
            if (s <= maxIn) break;
            n++;
            if (n > 100) { n=null; break; }
        }
        if (n) {
            var totalArea = n * s * inputHeight;
            plans.push({
                mode: "optimize",
                r: r,
                maxIn: maxIn,
                n: n,
                cuts: n-1,
                s: s,
                area: totalArea
            });
        }
    }
    return plans;
}

// --- Hàm tính fill ---
function makeFillPlan(inputWidth, decalWidths, margin, overlap, inputHeight) {
    var plans = [];
    var usableWidth = inputWidth - margin;
    
    for (var i = 0; i < decalWidths.length; i++) {
        var r = decalWidths[i] - margin;
        if (r <= 0) continue;

        var strips = [];
        var remaining = usableWidth;
        
        while (remaining > 0) {
            if (remaining >= r) {
                strips.push(r);
                remaining -= (r - overlap);
            } else {
                strips.push(remaining);
                remaining = 0;
            }
        }
        
        var totalArea = 0;
        for (var j = 0; j < strips.length; j++) {
            totalArea += strips[j] * inputHeight;
        }
        
        plans.push({
            mode: "fill",
            r: decalWidths[i],
            maxIn: r,
            strips: strips,
            cuts: strips.length - 1,
            n: strips.length,
            area: totalArea
        });
    }
    
    return plans;
}

// --- Tính cả 2 phương án ---
var optimizePlans = makeOptimizePlan(inputWidth, decalWidths, margin, overlap, inputHeight);
var fillPlans = makeFillPlan(inputWidth, decalWidths, margin, overlap, inputHeight);

// Gộp 2 mảng
var allPlans = optimizePlans.concat(fillPlans);

// Sắp xếp theo cuts, rồi theo r
allPlans.sort(function(a,b){
    if(a.cuts != b.cuts) return a.cuts - b.cuts;
    return a.r - b.r;
});

// --- Tính diện tích file gốc ---
var originalArea = inputWidth * inputHeight; // mm²

// --- Hàm repeat cho ExtendScript ---
function repeatStr(str, times) {
    var result = "";
    for (var i = 0; i < times; i++) {
        result += str;
    }
    return result;
}

// --- Hiển thị bảng đầy đủ ---
var msg = "BẢNG PHƯƠNG ÁN CẮT\n";
msg += "File gốc: " + originalArea.toFixed(0) + " mm² (" + (originalArea/1000000).toFixed(3) + " m²)\n";
msg += repeatStr("=", 70) + "\n";
msg += "STT | Mode     | Khổ in | Cuts | Tổng diện tích | Tiết kiệm\n";
msg += repeatStr("-", 70) + "\n";

for (var i = 0; i < allPlans.length; i++) {
    var plan = allPlans[i];
    var stt = i + 1; // STT từ 1
    var saved = originalArea - plan.area;
    var savedPercent = (saved / originalArea * 100).toFixed(1);
    
    msg += stt + " | ";
    msg += plan.mode + " | ";
    msg += plan.maxIn + " mm | ";
    msg += plan.cuts + " | ";
    msg += (plan.area/1000000).toFixed(3) + " m² | ";
    msg += savedPercent + "%\n";
}

alert(msg);

// --- Chọn STT (không phải index) ---
var stt = parseInt(prompt("Chọn STT phương án (1 là tốt nhất)", "1"));
if (stt < 1 || stt > allPlans.length) { 
    alert("Lựa chọn không hợp lệ"); 
    exit(); 
}

var plan = allPlans[stt - 1]; // Chuyển STT về index

// --- Slice vertical ---
var docWidthPx = doc.width.as("px");
var docHeightPx = doc.height.as("px");
var dpi = doc.resolution;
var bitDepth = doc.bitsPerChannel;
var colorProfile = doc.colorProfileName;

var overlapPx = overlap / 25.4 * dpi;

if (plan.mode == "fill") {
    var accWidth = 0;
    for (var i = 0; i < plan.strips.length; i++) {
        var w = plan.strips[i];
        var stripWidthPx = w / 25.4 * dpi;
        var sliceNumber = i + 1; // STT từ 1

        var x1 = accWidth;
        var x2 = x1 + stripWidthPx;
        if (x2 > docWidthPx) x2 = docWidthPx;

        doc.selection.select([[x1,0],[x2,0],[x2,docHeightPx],[x1,docHeightPx]]);
        doc.selection.copy();

        var tempDoc = app.documents.add(
            stripWidthPx,
            docHeightPx,
            dpi,
            originalFileName + "_" + sliceNumber,
            getNewDocumentModeFromDocMode(doc.mode),
            DocumentFill.TRANSPARENT
        );

        if (tempDoc.bitsPerChannel != bitDepth) tempDoc.bitsPerChannel = bitDepth;
        tempDoc.paste();
        try { tempDoc.assignProfile(colorProfile, true); } catch(e) {}

        var tiffOptions = new TiffSaveOptions();
        tiffOptions.imageCompression = TIFFEncoding.NONE;
        tiffOptions.embedColorProfile = true;

        var saveFile = new File("~/Desktop/" + originalFileName + "_" + sliceNumber + "_" + w.toFixed(0) + "mm.tif");
        tempDoc.saveAs(saveFile, tiffOptions);
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);

        doc.selection.deselect();

        accWidth += stripWidthPx - overlapPx;
    }
} else { // optimize
    var stripWidthPx = plan.s / 25.4 * dpi;
    for (var i = 0; i < plan.n; i++) {
        var sliceNumber = i + 1; // STT từ 1
        
        var x1 = i * (stripWidthPx - overlapPx);
        var x2 = x1 + stripWidthPx;
        if (x2 > docWidthPx) x2 = docWidthPx;

        doc.selection.select([[x1,0],[x2,0],[x2,docHeightPx],[x1,docHeightPx]]);
        doc.selection.copy();

        var tempDoc = app.documents.add(
            stripWidthPx,
            docHeightPx,
            dpi,
            originalFileName + "_" + sliceNumber,
            getNewDocumentModeFromDocMode(doc.mode),
            DocumentFill.TRANSPARENT
        );

        if (tempDoc.bitsPerChannel != bitDepth) tempDoc.bitsPerChannel = bitDepth;
        tempDoc.paste();
        try { tempDoc.assignProfile(colorProfile, true); } catch(e) {}

        var tiffOptions = new TiffSaveOptions();
        tiffOptions.imageCompression = TIFFEncoding.NONE;
        tiffOptions.embedColorProfile = true;

        var saveFile = new File("~/Desktop/" + originalFileName + "_" + sliceNumber + "_" + plan.s.toFixed(0) + "mm.tif");
        tempDoc.saveAs(saveFile, tiffOptions);
        tempDoc.close(SaveOptions.DONOTSAVECHANGES);

        doc.selection.deselect();
    }
}

function getNewDocumentModeFromDocMode(docMode) {
    switch(docMode) {
        case DocumentMode.RGB: return NewDocumentMode.RGB;
        case DocumentMode.CMYK: return NewDocumentMode.CMYK;
        case DocumentMode.GRAYSCALE: return NewDocumentMode.GRAYSCALE;
        case DocumentMode.LAB: return NewDocumentMode.LAB;
        default: return NewDocumentMode.RGB;
    }
}

alert("Đã slice xong " + plan.n + " dải (STT 1-" + plan.n + "), lưu ra Desktop!\nTên file: " + originalFileName + "_1, " + originalFileName + "_2, ...");