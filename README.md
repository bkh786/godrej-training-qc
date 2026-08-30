# Godrej Training & QC Performance Dashboard

Live interactive intelligence suite for Godrej Training & QC Operations (Visual Merchandising Program).

## 🚀 Live Features

1. **Multi-Page Tabbed Architecture**:
   - **QC Performance Dashboard (`index.html`)**: Complete QC audit analytics, compliance scoring (Display Categories, Campaign/POSM, AOM-wise scores, trends, recurring mistakes, management insights, and audit records).
   - **Training Performance Dashboard (`training.html`)**: Field workforce deployment (VMs), regional coverage (North, East, West, South 1, South 2), branch distribution, pitfall root-cause analysis, and trainer action guidelines.
   - **Merchandising Score Dashboard (`m_score.html`)**: Execution & compliance intelligence extracted directly from the `Product-VM-Score` sheet. Features Overall Merchandising Score, Products Captured, Target Deployment, POSM Executed, POSM Not Executed, trend/category/AOM charts, and detailed `<100%` score audit records table with instant search and pagination.
   - **Program Performance Dashboard (`program_performance.html`)**: Comprehensive executive KPI intelligence extracted from the `Program Performance` sheet. Features Overall KPI, Detailed Channel Segment Coverage (EBO, Chain Store, Cool Studio, Others), Manpower Productivity & Market Time, interactive click-to-filter stacked column charts, Top/Bottom 5 performer comparison, AOM leadership ranking, and paginated master roster.

2. **Automated SharePoint Live Sync**:
   - Directly syncs live Excel data from SharePoint with automatic multi-tier fallback (Direct Fetch → CORS Proxies → Local/Embedded Cache).
   - Real-time indicator (`Live data connected`) with last-synced timestamp.
   - Dedicated `Download Report` button to download the latest report file directly.

3. **Modern Design System (`style.css`)**:
   - Browser-style top navigation tabs with curved borders and glowing active indicators.
   - Deep forest teal gradient main header matching corporate specifications.
   - Responsive KPI cards, Chart.js column & doughnut charts, quadrant insights, and sortable tables.
   - Built-in Dark / Light mode toggle.

## 📊 Live Deployment
- Repository: [https://github.com/bkh786/godrej-training-qc](https://github.com/bkh786/godrej-training-qc)
