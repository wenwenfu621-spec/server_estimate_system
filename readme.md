# CAD Bounding Box Parser (.STEP / .IGS)

本專案用於自動讀取並解析 `.step` / `.stp` 與 `.igs` / `.iges` 3D CAD 檔案的邊界外框尺寸（Bounding Box），並依「長 >= 寬 >= 高」規則輸出胚料估算尺寸。

---

## 🛠️ 環境需求與安裝步驟

### 1. 克隆 GitHub 專案
```bash
git clone [https://github.com/](https://github.com/)<your-username>/cad-boundingbox-parser.git
cd cad-boundingbox-parser