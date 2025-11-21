const { app, core } = require("photoshop");
const { executeAsModal } = require("photoshop").core;
const fs = require("uxp").storage.localFileSystem;
const batchPlay = require("photoshop").action.batchPlay;

let currentDoc = null;
let allPlans = [];
let selectedPlan = null;
const decalWidths = [914, 1070, 1270, 1520];

// Khởi tạo
document.addEventListener("DOMContentLoaded", async () => {
    await loadDocumentInfo();

    document.getElementById("calculateBtn").addEventListener("click", calculatePlans);
    document.getElementById("sliceBtn").addEventListener("click", executeSlice);

    // Theo dõi document changes
    app.eventNotifier.addNotifier("select", async () => {
        await loadDocumentInfo();
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
                <td>${plan.cuts}</td>
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
async function sliceAndSave(x1, x2, heightPx, targetWidthPx, dpi, folder, fileName) {
    // 1. Select area
    await batchPlay([{
        _obj: "set",
        _target: [{ _ref: "channel", _enum: "channel", _value: "selection" }],
        to: {
            _obj: "rectangle",
            top: { _unit: "pixelsUnit", _value: 0 },
            left: { _unit: "pixelsUnit", _value: x1 },
            bottom: { _unit: "pixelsUnit", _value: heightPx },
            right: { _unit: "pixelsUnit", _value: x2 }
        }
    }], {});

    // 2. Copy Merged (to get all layers)
    await batchPlay([{
        _obj: "copyEvent",
        merged: true
    }], {});

    // 3. Create New Document
    const { constants } = require("photoshop");
    await app.documents.add({
        width: targetWidthPx,
        height: heightPx,
        resolution: dpi,
        mode: constants.NewDocumentMode.RGB,
        fill: constants.DocumentFill.TRANSPARENT
    });

    const newDoc = app.activeDocument;

    // 4. Paste
    await batchPlay([{
        _obj: "paste",
        antiAlias: { _enum: "antiAliasType", _value: "antiAliasNone" }
    }], {});

    // 5. Flatten (optional, but good for TIFF)
    await batchPlay([{ _obj: "flattenImage" }], {});

    // 6. Save as TIFF
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

    // 7. Close new doc
    await newDoc.closeWithoutSaving();

    // 8. Switch back to original doc and deselect
    if (currentDoc) {
        await batchPlay([{
            _obj: "select",
            _target: [{ _ref: "document", _name: currentDoc.name }]
        }], {});

        await batchPlay([{
            _obj: "set",
            _target: [{ _ref: "channel", _enum: "channel", _value: "selection" }],
            to: { _enum: "ordinal", _value: "none" }
        }], {});
    }
}

// Show status message
function showStatus(message, type = "") {
    const status = document.getElementById("status");
    status.textContent = message;
    status.className = `status ${type}`;
}