const { app, core } = require("photoshop");
const { executeAsModal } = require("photoshop").core;
const fs = require("uxp").storage.localFileSystem;
const batchPlay = require("photoshop").action.batchPlay;

let currentDoc = null;
let allPlans = [];
let selectedPlan = null;
const decalWidths = [914, 1070, 1270, 1520];

// Grid slicing variables
let gridCalculation = null;

// Khởi tạo
document.addEventListener("DOMContentLoaded", async () => {
    await loadDocumentInfo();

    document.getElementById("calculateBtn").addEventListener("click", calculatePlans);
    document.getElementById("sliceBtn").addEventListener("click", executeSlice);
    document.getElementById("gridSliceBtn").addEventListener("click", executeGridSlice);

    // Grid inputs event listeners - use 'change' for Spectrum components
    document.getElementById("gridCols").addEventListener("change", updateGridCalculation);
    document.getElementById("gridRows").addEventListener("change", updateGridCalculation);
    document.getElementById("overlap").addEventListener("change", updateGridCalculation);

    // Theo dõi document changes
    app.eventNotifier.addNotifier("select", async () => {
        await loadDocumentInfo();
        updateGridCalculation();
    });
});

// Load thông tin document
async function loadDocumentInfo() {
    try {
        if (app.documents.length === 0) {
            document.getElementById("fileInfo").innerHTML =
                '<div class="loading">Vui lòng mở một file trong Photoshop</div>';
            currentDoc = null;
            return;
        }

        currentDoc = app.activeDocument;
        const width = currentDoc.width;
        const height = currentDoc.height;
        const widthMM = convertToMM(width, currentDoc.resolution);
        const heightMM = convertToMM(height, currentDoc.resolution);
        const area = (widthMM * heightMM / 1000000).toFixed(3);

        document.getElementById("fileInfo").innerHTML = `
            <strong>File:</strong> ${currentDoc.name}<br>
            <strong>Kích thước:</strong> ${widthMM.toFixed(1)} × ${heightMM.toFixed(1)} mm<br>
            <strong>Diện tích:</strong> ${area} m² | <strong>DPI:</strong> ${currentDoc.resolution.toFixed(0)}
        `;
    } catch (error) {
        console.error("Error loading document:", error);
    }
}

// Convert px to mm
function convertToMM(px, dpi) {
    return (px / dpi) * 25.4;
}

// Convert mm to px
function convertToPx(mm, dpi) {
    return (mm / 25.4) * dpi;
}

// Tính phương án optimize
function makeOptimizePlan(inputWidth, overlap, margin, inputHeight) {
    const plans = [];
    for (const r of decalWidths) {
        const maxIn = r - margin;
        let n = 1;
        let s = 0;
        while (true) {
            s = Math.ceil((inputWidth + (n - 1) * overlap) / n);
            if (s <= maxIn) break;
            n++;
            if (n > 100) { n = null; break; }
        }
        if (n) {
            const totalArea = n * s * inputHeight;
            plans.push({
                mode: "optimize",
                r: r,
                maxIn: maxIn,
                n: n,
                cuts: n - 1,
                s: s,
                area: totalArea
            });
        }
    }
    return plans;
}

// Tính phương án fill
function makeFillPlan(inputWidth, overlap, margin, inputHeight) {
    const plans = [];
    const usableWidth = inputWidth - margin;

    for (const r of decalWidths) {
        const stripWidth = r - margin;
        if (stripWidth <= 0) continue;

        const strips = [];
        let remaining = usableWidth;

        while (remaining > 0) {
            if (remaining >= stripWidth) {
                strips.push(stripWidth);
                remaining -= (stripWidth - overlap);
            } else {
                strips.push(remaining);
                remaining = 0;
            }
        }

        let totalArea = 0;
        for (const w of strips) {
            totalArea += w * inputHeight;
        }

        plans.push({
            mode: "fill",
            r: r,
            maxIn: stripWidth,
            strips: strips,
            cuts: strips.length - 1,
            n: strips.length,
            area: totalArea
        });
    }

    return plans;
}

// Calculate và hiển thị bảng
async function calculatePlans() {
    try {
        await loadDocumentInfo(); // Reload info to be sure
        if (!currentDoc) {
            showStatus("Vui lòng mở một file trước!", "error");
            return;
        }
        const overlap = parseFloat(document.getElementById("overlap").value);
        const margin = parseFloat(document.getElementById("margin").value);

        if (isNaN(overlap) || isNaN(margin)) {
            showStatus("Vui lòng nhập số hợp lệ cho Overlap và Margin", "error");
            return;
        }

        const widthMM = convertToMM(currentDoc.width, currentDoc.resolution);
        const heightMM = convertToMM(currentDoc.height, currentDoc.resolution);
        const originalArea = widthMM * heightMM;

        const optimizePlans = makeOptimizePlan(widthMM, overlap, margin, heightMM);
        const fillPlans = makeFillPlan(widthMM, overlap, margin, heightMM);

        allPlans = [...optimizePlans, ...fillPlans];
        allPlans.sort((a, b) => {
            if (a.cuts !== b.cuts) return a.cuts - b.cuts;
            return a.r - b.r;
        });

        // Render bảng
        const tbody = document.getElementById("planTable");
        tbody.innerHTML = "";

        if (allPlans.length === 0) {
            showStatus("Không tìm thấy phương án phù hợp", "error");
            return;
        }

        allPlans.forEach((plan, index) => {
            const stt = index + 1;
            const saved = originalArea - plan.area;
            const savedPercent = (saved / originalArea * 100).toFixed(1);

            const row = document.createElement("tr");
            if (index === 0) row.classList.add("selected");

            row.innerHTML = `
                <td><strong>${stt}</strong> ${index === 0 ? '<span class="best-label">TỐT NHẤT</span>' : ''}</td>
                <td><span class="badge badge-${plan.mode}">${plan.mode}</span></td>
                <td>${plan.r} mm</td>
                <td>${plan.cuts + 1}</td>
                <td>${(plan.area / 1000000).toFixed(3)} m²</td>
                <td class="savings">↓ ${savedPercent}%</td>
            `;

            row.addEventListener("click", () => {
                document.querySelectorAll("#planTable tr").forEach(r => r.classList.remove("selected"));
                row.classList.add("selected");
                selectedPlan = index;
            });

            tbody.appendChild(row);
        });

        selectedPlan = 0;
        document.getElementById("tableContainer").style.display = "block";
        document.getElementById("actionButtons").style.display = "flex";
        showStatus(`Đã tính ${allPlans.length} phương án`, "success");

    } catch (e) {
        console.error(e);
        showStatus(`Lỗi: ${e.message}`, "error");
    }
}

// Execute slice
async function executeSlice() {
    if (!currentDoc || selectedPlan === null) {
        showStatus("Vui lòng chọn phương án trước!", "error");
        return;
    }

    const plan = allPlans[selectedPlan];
    const overlap = parseFloat(document.getElementById("overlap").value);

    try {
        showStatus("Đang cắt ảnh...", "");

        await executeAsModal(async () => {
            const folder = await fs.getFolder();
            if (!folder) {
                showStatus("Đã hủy chọn thư mục", "error");
                return;
            }

            const originalName = currentDoc.name.replace(/\.[^.]+$/, "");
            const dpi = currentDoc.resolution;
            const docMode = currentDoc.mode; // Get original color mode
            const widthMM = convertToMM(currentDoc.width, dpi);
            const heightMM = convertToMM(currentDoc.height, dpi);
            const overlapPx = convertToPx(overlap, dpi);
            const docHeightPx = currentDoc.height; // Keep original height in px
            const docWidthPx = currentDoc.width;

            if (plan.mode === "fill") {
                let accWidth = 0;
                for (let i = 0; i < plan.strips.length; i++) {
                    const w = plan.strips[i];
                    const stripWidthPx = convertToPx(w, dpi);
                    const sliceNum = i + 1;

                    // Calculate coordinates
                    let x1 = accWidth;
                    let x2 = x1 + stripWidthPx;
                    if (x2 > docWidthPx) x2 = docWidthPx;

                    await sliceAndSave(
                        x1, x2, docHeightPx,
                        stripWidthPx, // Target width for new doc
                        dpi,
                        docMode, // Pass color mode
                        folder,
                        `${originalName}_${sliceNum}_${w.toFixed(0)}mm`
                    );

                    accWidth += stripWidthPx - overlapPx;
                }
            } else {
                const stripWidthPx = convertToPx(plan.s, dpi);
                for (let i = 0; i < plan.n; i++) {
                    const sliceNum = i + 1;
                    const x1 = i * (stripWidthPx - overlapPx);
                    let x2 = x1 + stripWidthPx;
                    if (x2 > docWidthPx) x2 = docWidthPx;

                    await sliceAndSave(
                        x1, x2, docHeightPx,
                        stripWidthPx,
                        dpi,
                        docMode, // Pass color mode
                        folder,
                        `${originalName}_${sliceNum}_${plan.s.toFixed(0)}mm`
                    );
                }
            }

            showStatus(`✅ Đã cắt xong ${plan.n} dải!`, "success");
        }, { commandName: "Slice Tool Pro" });

    } catch (error) {
        showStatus(`Lỗi: ${error.message}`, "error");
        console.error(error);
    }
}

// Slice và save một strip
async function sliceAndSave(x1, x2, heightPx, targetWidthPx, dpi, docMode, folder, fileName) {
    console.log(`[DEBUG] Slicing: x1=${x1}, x2=${x2}, height=${heightPx}, targetWidth=${targetWidthPx}`);

    // NEW APPROACH: Duplicate document then crop
    // This is more reliable than copy/paste for TIFF files

    // 1. Duplicate the original document
    await batchPlay([{
        _obj: "duplicate",
        _target: [{ _ref: "document", _enum: "ordinal", _value: "targetEnum" }],
        name: fileName
    }], {});
    console.log("[DEBUG] Document duplicated");

    const newDoc = app.activeDocument;
    console.log(`[DEBUG] Working on duplicated document: ${newDoc.name}`);

    // 2. Flatten the duplicate (in case it has layers)
    try {
        await batchPlay([{ _obj: "flattenImage" }], {});
        console.log("[DEBUG] Duplicate flattened");
    } catch (e) {
        console.log("[DEBUG] Already flat or flatten failed (OK)");
    }

    // 3. Crop to the desired area
    await batchPlay([{
        _obj: "crop",
        to: {
            _obj: "rectangle",
            top: { _unit: "pixelsUnit", _value: 0 },
            left: { _unit: "pixelsUnit", _value: x1 },
            bottom: { _unit: "pixelsUnit", _value: heightPx },
            right: { _unit: "pixelsUnit", _value: x2 }
        },
        angle: { _unit: "angleUnit", _value: 0 },
        delete: true,
        cropAspectRatioModeKey: { _enum: "cropAspectRatioModeClass", _value: "pureAspectRatio" }
    }], {});
    console.log(`[DEBUG] Cropped to x1=${x1}, x2=${x2}`);

    // 4. Resize canvas if needed (to ensure exact target width)
    const currentWidth = newDoc.width;
    console.log(`[DEBUG] Current width after crop: ${currentWidth}, target: ${targetWidthPx}`);

    if (Math.abs(currentWidth - targetWidthPx) > 1) {
        await batchPlay([{
            _obj: "canvasSize",
            width: { _unit: "pixelsUnit", _value: targetWidthPx },
            height: { _unit: "pixelsUnit", _value: heightPx },
            horizontal: { _enum: "horizontalLocation", _value: "left" },
            vertical: { _enum: "verticalLocation", _value: "top" }
        }], {});
        console.log(`[DEBUG] Canvas resized to ${targetWidthPx}x${heightPx}`);
    }

    // 5. Save as TIFF
    const file = await folder.createFile(`${fileName}.tif`, { overwrite: true });
    const token = await fs.createSessionToken(file);

    await batchPlay([{
        _obj: "save",
        as: {
            _obj: "TIFF",
            byteOrder: { _enum: "platform", _value: "IBMPC" },
            imageCompression: { _enum: "TIFFEncoding", _value: "none" }
        },
        in: { _path: token, _kind: "local" },
        copy: true
    }], {});
    console.log(`[DEBUG] File saved: ${fileName}.tif`);

    // 6. Close the duplicated doc
    await newDoc.closeWithoutSaving();
    console.log("[DEBUG] Duplicate closed");

    // 7. Switch back to original doc
    if (currentDoc) {
        await batchPlay([{
            _obj: "select",
            _target: [{ _ref: "document", _name: currentDoc.name }]
        }], {});
        console.log("[DEBUG] Switched back to original document");
    }
}

// Show status message
function showStatus(message, type = "") {
    const status = document.getElementById("status");
    status.textContent = message;
    status.className = `status ${type}`;
}

// Grid slicing functions
function updateGridCalculation() {
    if (!currentDoc) {
        document.getElementById("gridResult").innerHTML = '<div class="result-item">Vui lòng mở một file trong Photoshop</div>';
        document.getElementById("gridSliceBtn").style.display = "none";
        return;
    }

    const cols = parseInt(document.getElementById("gridCols").value);
    const rows = parseInt(document.getElementById("gridRows").value);
    const overlap = parseFloat(document.getElementById("overlap").value);

    if (isNaN(cols) || isNaN(rows) || cols < 1 || rows < 1) {
        document.getElementById("gridResult").innerHTML = '<div class="result-item">Nhập số cột và dòng hợp lệ</div>';
        document.getElementById("gridSliceBtn").style.display = "none";
        return;
    }

    const widthMM = convertToMM(currentDoc.width, currentDoc.resolution);
    const heightMM = convertToMM(currentDoc.height, currentDoc.resolution);

    // Calculate piece size with overlap
    const pieceWidth = (widthMM + (cols - 1) * overlap) / cols;
    const pieceHeight = (heightMM + (rows - 1) * overlap) / rows;
    const totalPieces = cols * rows;

    gridCalculation = {
        cols: cols,
        rows: rows,
        pieceWidth: pieceWidth,
        pieceHeight: pieceHeight,
        overlap: overlap
    };

    document.getElementById("gridResult").innerHTML = `
        <div class="result-item"><strong>Kích thước mỗi phần:</strong> ${pieceWidth.toFixed(1)} × ${pieceHeight.toFixed(1)} mm</div>
        <div class="result-item"><strong>Số cột:</strong> ${cols} | <strong>Số dòng:</strong> ${rows}</div>
        <div class="result-item"><strong>Tổng số phần:</strong> ${totalPieces}</div>
        <div class="result-item"><strong>Overlap:</strong> ${overlap} mm</div>
    `;
    document.getElementById("gridSliceBtn").style.display = "block";
}

async function executeGridSlice() {
    if (!currentDoc || !gridCalculation) {
        showStatus("Vui lòng nhập số cột và dòng trước!", "error");
        return;
    }

    try {
        showStatus("Đang cắt theo lưới...", "");

        await executeAsModal(async () => {
            const folder = await fs.getFolder();
            if (!folder) {
                showStatus("Đã hủy chọn thư mục", "error");
                return;
            }

            const originalName = currentDoc.name.replace(/\.[^.]+$/, "");
            const dpi = currentDoc.resolution;
            const docMode = currentDoc.mode;
            const docWidthPx = currentDoc.width;
            const docHeightPx = currentDoc.height;

            const { cols, rows, pieceWidth, pieceHeight, overlap } = gridCalculation;
            const overlapPx = convertToPx(overlap, dpi);
            const pieceWidthPx = convertToPx(pieceWidth, dpi);
            const pieceHeightPx = convertToPx(pieceHeight, dpi);

            let pieceCount = 0;

            for (let row = 0; row < rows; row++) {
                for (let col = 0; col < cols; col++) {
                    pieceCount++;

                    // Calculate coordinates
                    const x1 = col * (pieceWidthPx - overlapPx);
                    let x2 = x1 + pieceWidthPx;
                    if (x2 > docWidthPx) x2 = docWidthPx;

                    const y1 = row * (pieceHeightPx - overlapPx);
                    let y2 = y1 + pieceHeightPx;
                    if (y2 > docHeightPx) y2 = docHeightPx;

                    await sliceAndSaveGrid(
                        x1, y1, x2, y2,
                        pieceWidthPx, pieceHeightPx,
                        dpi,
                        docMode,
                        folder,
                        `${originalName}_R${row + 1}C${col + 1}_${pieceWidth.toFixed(0)}x${pieceHeight.toFixed(0)}mm`
                    );
                }
            }

            showStatus(`✅ Đã cắt xong ${pieceCount} phần!`, "success");
        }, { commandName: "Grid Slice Tool Pro" });

    } catch (error) {
        showStatus(`Lỗi: ${error.message}`, "error");
        console.error(error);
    }
}

// Slice và save một phần grid
async function sliceAndSaveGrid(x1, y1, x2, y2, targetWidthPx, targetHeightPx, dpi, docMode, folder, fileName) {
    console.log(`[DEBUG] Grid slicing: x1=${x1}, y1=${y1}, x2=${x2}, y2=${y2}, targetW=${targetWidthPx}, targetH=${targetHeightPx}`);

    // NEW APPROACH: Duplicate document then crop (same as vertical slicing)

    // 1. Duplicate the original document
    await batchPlay([{
        _obj: "duplicate",
        _target: [{ _ref: "document", _enum: "ordinal", _value: "targetEnum" }],
        name: fileName
    }], {});
    console.log("[DEBUG] Grid document duplicated");

    const newDoc = app.activeDocument;
    console.log(`[DEBUG] Grid working on duplicated document: ${newDoc.name}`);

    // 2. Flatten the duplicate
    try {
        await batchPlay([{ _obj: "flattenImage" }], {});
        console.log("[DEBUG] Grid duplicate flattened");
    } catch (e) {
        console.log("[DEBUG] Grid already flat or flatten failed (OK)");
    }

    // 3. Crop to the desired area
    await batchPlay([{
        _obj: "crop",
        to: {
            _obj: "rectangle",
            top: { _unit: "pixelsUnit", _value: y1 },
            left: { _unit: "pixelsUnit", _value: x1 },
            bottom: { _unit: "pixelsUnit", _value: y2 },
            right: { _unit: "pixelsUnit", _value: x2 }
        },
        angle: { _unit: "angleUnit", _value: 0 },
        delete: true,
        cropAspectRatioModeKey: { _enum: "cropAspectRatioModeClass", _value: "pureAspectRatio" }
    }], {});
    console.log(`[DEBUG] Grid cropped to x1=${x1}, y1=${y1}, x2=${x2}, y2=${y2}`);

    // 4. Resize canvas if needed
    const currentWidth = newDoc.width;
    const currentHeight = newDoc.height;
    console.log(`[DEBUG] Grid current size after crop: ${currentWidth}x${currentHeight}, target: ${targetWidthPx}x${targetHeightPx}`);

    if (Math.abs(currentWidth - targetWidthPx) > 1 || Math.abs(currentHeight - targetHeightPx) > 1) {
        await batchPlay([{
            _obj: "canvasSize",
            width: { _unit: "pixelsUnit", _value: targetWidthPx },
            height: { _unit: "pixelsUnit", _value: targetHeightPx },
            horizontal: { _enum: "horizontalLocation", _value: "left" },
            vertical: { _enum: "verticalLocation", _value: "top" }
        }], {});
        console.log(`[DEBUG] Grid canvas resized to ${targetWidthPx}x${targetHeightPx}`);
    }

    // 5. Save as TIFF
    const file = await folder.createFile(`${fileName}.tif`, { overwrite: true });
    const token = await fs.createSessionToken(file);

    await batchPlay([{
        _obj: "save",
        as: {
            _obj: "TIFF",
            byteOrder: { _enum: "platform", _value: "IBMPC" },
            imageCompression: { _enum: "TIFFEncoding", _value: "none" }
        },
        in: { _path: token, _kind: "local" },
        copy: true
    }], {});
    console.log(`[DEBUG] Grid file saved: ${fileName}.tif`);

    // 6. Close the duplicated doc
    await newDoc.closeWithoutSaving();
    console.log("[DEBUG] Grid duplicate closed");

    // 7. Switch back to original doc
    if (currentDoc) {
        await batchPlay([{
            _obj: "select",
            _target: [{ _ref: "document", _name: currentDoc.name }]
        }], {});
        console.log("[DEBUG] Grid switched back to original document");
    }
}
