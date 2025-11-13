// main.js - TẠO GUIDES GRID
const photoshop = require("photoshop");
const { app, core } = photoshop;
const { batchPlay } = photoshop.action;
const { entrypoints } = require("uxp");

function showAlert(message) {
  const dialog = document.createElement("dialog");
  dialog.innerHTML = `
    <form method="dialog" style="padding: 20px; min-width: 300px;">
      <p style="margin: 0 0 20px 0; white-space: pre-line;">${message}</p>
      <button type="submit" style="padding: 8px 20px; width: 100%;">OK</button>
    </form>
  `;
  document.body.appendChild(dialog);
  dialog.showModal();
  dialog.addEventListener("close", () => dialog.remove());
}

entrypoints.setup({
  panels: {
    vanilla: {
      show() {
        if (document.readyState === "loading") {
          document.addEventListener("DOMContentLoaded", init);
        } else {
          init();
        }
      }
    }
  }
});

function init() {
  document.getElementById("btnGrid")?.addEventListener("click", async () => {
    const rows = parseInt(document.getElementById("rows").value);
    const cols = parseInt(document.getElementById("cols").value);
    await createGuidesGrid(rows, cols);
  });
  
  document.getElementById("btnClearGuides")?.addEventListener("click", clearGuides);
  document.getElementById("btnExportRegions")?.addEventListener("click", exportRegions);
}

// ✅ TẠO GUIDES GRID (Thay thế Slices)
async function createGuidesGrid(rows, cols) {
  try {
    await core.executeAsModal(async () => {
      const doc = app.activeDocument;
      
      if (!doc) {
        showAlert("⚠️ Mở file ảnh trước!");
        return;
      }
      
      const guides = doc.guides;
      guides.removeAll();
      
      const cellWidth = doc.width / cols;
      const cellHeight = doc.height / rows;
      
      console.log(`Creating ${rows}x${cols} grid`);
      console.log(`Cell size: ${cellWidth} x ${cellHeight}`);
      
      // Tạo guides dọc (vertical)
      for (let i = 1; i < cols; i++) {
        const position = i * cellWidth;
        guides.add("vertical", position);
        console.log(`Vertical guide at ${position}px`);
      }
      
      // Tạo guides ngang (horizontal)
      for (let i = 1; i < rows; i++) {
        const position = i * cellHeight;
        guides.add("horizontal", position);
        console.log(`Horizontal guide at ${position}px`);
      }
      
      showAlert(`✅ Đã tạo lưới ${rows}x${cols}!\n\nGuides đã được tạo để chia ảnh.\nDùng: View → Show → Guides để xem.`);
      
    }, { commandName: "Create Guides Grid" });
  } catch (err) {
    console.error("❌ Error:", err);
    showAlert("Lỗi: " + err.message);
  }
}

// ✅ XÓA TẤT CẢ GUIDES
async function clearGuides() {
  try {
    await core.executeAsModal(async () => {
      const doc = app.activeDocument;
      if (!doc) {
        showAlert("⚠️ Mở file ảnh trước!");
        return;
      }
      
      doc.guides.removeAll();
      showAlert("✅ Đã xóa tất cả guides!");
      
    }, { commandName: "Clear Guides" });
  } catch (err) {
    console.error("❌ Error:", err);
    showAlert("Lỗi: " + err.message);
  }
}

// ✅ EXPORT TỪNG VÙNG (Dựa trên guides)
async function exportRegions() {
  try {
    await core.executeAsModal(async () => {
      const doc = app.activeDocument;
      if (!doc) {
        showAlert("⚠️ Mở file ảnh trước!");
        return;
      }
      
      const guides = doc.guides;
      const allGuides = guides.getAll();
      
      if (allGuides.length === 0) {
        showAlert("⚠️ Tạo guides grid trước!");
        return;
      }
      
      // Lấy tọa độ guides
      const vGuides = allGuides.filter(g => g.direction === "vertical")
        .map(g => g.coordinate)
        .sort((a, b) => a - b);
      
      const hGuides = allGuides.filter(g => g.direction === "horizontal")
        .map(g => g.coordinate)
        .sort((a, b) => a - b);
      
      // Thêm biên
      const xPositions = [0, ...vGuides, doc.width];
      const yPositions = [0, ...hGuides, doc.height];
      
      console.log("X positions:", xPositions);
      console.log("Y positions:", yPositions);
      
      showAlert(`📊 Phát hiện lưới:\n${yPositions.length - 1} hàng x ${xPositions.length - 1} cột\n\n⚠️ Export thủ công:\n1. Dùng Crop Tool (C)\n2. Crop từng vùng theo guides\n3. File → Export → Export As...\n4. Undo để quay lại`);
      
    }, { commandName: "Export Regions" });
  } catch (err) {
    console.error("❌ Error:", err);
    showAlert("Lỗi: " + err.message);
  }
}