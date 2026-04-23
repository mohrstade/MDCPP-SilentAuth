import { getGraphToken, graphGet, graphPost, initializeAuth, signIn, signOut } from "./aadauth";

type GraphCollection<T> = {
    value?: T[];
};

type GraphUser = {
    displayName?: string;
    userPrincipalName?: string;
};

type GraphDriveItem = {
    id?: string;
    name?: string;
    webUrl?: string;
    size?: number;
    lastModifiedDateTime?: string;
    createdBy?: {
        user?: {
            displayName?: string;
        };
    };
    lastModifiedBy?: {
        user?: {
            displayName?: string;
        };
    };
    file?: Record<string, unknown>;
    folder?: Record<string, unknown>;
    parentReference?: {
        path?: string;
    };
    remoteItem?: GraphDriveItem;
};

type PreviewInfo = {
    getUrl?: string;
    postUrl?: string;
    postParameters?: string;
};

type ViewName = "home" | "myfiles" | "folders" | "shared" | "favorites" | "recycle";
type FilterName = "all" | "word" | "excel" | "powerpoint" | "pdf";
type SortName = "recent" | "alphabetical" | "size" | "modified";

type StartPageItem = {
    id: string;
    name: string;
    type: "Word" | "Excel" | "Powerpoint" | "PDF" | "Folder" | "File";
    webUrl: string;
    source: "Recent" | "My files" | "Shared" | "Favorites" | "Recycle bin";
    modifiedRaw?: string;
    modified: string;
    size?: number;
    owner: string;
    location: string;
    folder: boolean;
    raw: GraphDriveItem;
};

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const GRAPH_BETA_BASE = "https://graph.microsoft.com/beta";
const FAVORITES_KEY = "team-startpage:favorites";

const appStyles = `
    :root {
        color-scheme: dark;
        --bg: #07111d;
        --panel: rgba(9, 18, 33, 0.78);
        --panel-strong: rgba(12, 24, 43, 0.92);
        --panel-soft: rgba(255, 255, 255, 0.05);
        --border: rgba(255, 255, 255, 0.09);
        --border-strong: rgba(255, 255, 255, 0.16);
        --text: #ecf4ff;
        --muted: #9fb1c7;
        --muted-strong: #c7d4e7;
        --accent: #66d9ef;
        --accent-2: #4f8cff;
        --accent-warm: #ffb454;
        --success: #29d391;
        --danger: #ff7b8f;
        --shadow-xl: 0 28px 90px rgba(0, 0, 0, 0.35);
        --shadow-md: 0 16px 36px rgba(0, 0, 0, 0.22);
    }

    * {
        box-sizing: border-box;
    }

    html {
        min-height: 100%;
        background:
            radial-gradient(circle at top left, rgba(79, 140, 255, 0.24), transparent 28%),
            radial-gradient(circle at 85% 12%, rgba(41, 211, 145, 0.12), transparent 22%),
            radial-gradient(circle at 50% 100%, rgba(255, 180, 84, 0.1), transparent 30%),
            linear-gradient(180deg, #08111f 0%, #030812 100%);
    }

    html, body {
        margin: 0;
        min-height: 100%;
        font-family: "Aptos", "Segoe UI Variable Text", "Segoe UI", sans-serif;
        color: var(--text);
        background: transparent;
    }

    body::before,
    body::after {
        content: "";
        position: fixed;
        pointer-events: none;
        z-index: 0;
        filter: blur(28px);
        border-radius: 50%;
    }

    body::before {
        width: 220px;
        height: 220px;
        top: 96px;
        right: 6%;
        background: rgba(79, 140, 255, 0.12);
    }

    body::after {
        width: 260px;
        height: 260px;
        bottom: 8%;
        left: 18%;
        background: rgba(41, 211, 145, 0.08);
    }

    button, input {
        font: inherit;
    }

    .hidden {
        display: none !important;
    }

    .app-shell {
        position: relative;
        z-index: 1;
        display: flex;
        min-height: 100vh;
    }

    .sidebar {
        width: 320px;
        padding: 28px 22px;
        display: flex;
        flex-direction: column;
        gap: 24px;
        background: linear-gradient(180deg, rgba(10, 20, 36, 0.98) 0%, rgba(5, 12, 23, 0.98) 100%);
        border-right: 1px solid rgba(255, 255, 255, 0.07);
        box-shadow: inset -1px 0 0 rgba(255, 255, 255, 0.03);
    }

    .brand-block {
        display: flex;
        align-items: center;
        gap: 14px;
        padding: 16px 18px;
        border-radius: 24px;
        background: linear-gradient(135deg, rgba(79, 140, 255, 0.2), rgba(14, 30, 55, 0.45));
        border: 1px solid rgba(128, 187, 255, 0.16);
        box-shadow: var(--shadow-md);
    }

    .brand-mark {
        width: 44px;
        height: 44px;
        border-radius: 14px;
        position: relative;
        background: linear-gradient(135deg, #5ea8ff 0%, #4f8cff 100%);
        box-shadow: 0 14px 28px rgba(79, 140, 255, 0.32);
    }

    .brand-mark::before,
    .brand-mark::after {
        content: "";
        position: absolute;
        border-radius: 10px;
        background: rgba(255, 255, 255, 0.92);
    }

    .brand-mark::before {
        inset: 10px 16px 10px 10px;
    }

    .brand-mark::after {
        inset: 16px 10px 16px 16px;
    }

    .brand {
        font-family: "Bahnschrift", "Segoe UI Variable Display", sans-serif;
        font-size: 29px;
        font-weight: 700;
        letter-spacing: 0.02em;
    }

    .brand-subtitle {
        margin: 4px 0 0;
        color: var(--muted);
        font-size: 13px;
        letter-spacing: 0.03em;
    }

    .create-panel,
    .nav {
        position: relative;
    }

    .primary-btn,
    .secondary-btn,
    .create-action,
    .nav-item,
    .account-dropdown-action,
    .file-menu-btn,
    .favorite-toggle-btn,
    .table-open-btn,
    .topbar-actions button,
    .picker-chip {
        transition:
            transform 180ms ease,
            background 180ms ease,
            border-color 180ms ease,
            box-shadow 180ms ease,
            color 180ms ease,
            opacity 180ms ease;
    }

    .primary-btn,
    .secondary-btn,
    .create-action,
    .account-dropdown-action,
    .table-open-btn,
    .topbar-actions button,
    .picker-chip {
        border-radius: 16px;
        font-size: 14px;
    }

    .primary-btn {
        border: 1px solid rgba(141, 204, 255, 0.22);
        background: linear-gradient(135deg, #59d7f5 0%, #4f8cff 100%);
        color: #03111f;
        padding: 12px 18px;
        font-weight: 800;
        cursor: pointer;
        text-decoration: none;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 18px 36px rgba(79, 140, 255, 0.24);
    }

    .secondary-btn,
    .topbar-actions button,
    .account-dropdown-action,
    .table-open-btn,
    .picker-chip {
        border: 1px solid var(--border);
        background: rgba(255, 255, 255, 0.04);
        color: var(--text);
        padding: 12px 16px;
        cursor: pointer;
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.03);
        text-decoration: none;
    }

    .primary-btn:hover,
    .secondary-btn:hover,
    .create-action:hover,
    .nav-item:hover,
    .file-menu-btn:hover,
    .favorite-toggle-btn:hover,
    .table-open-btn:hover,
    .picker-chip:hover,
    .topbar-actions button:hover {
        transform: translateY(-1px);
    }

    .plus-icon {
        width: 30px;
        height: 30px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border-radius: 12px;
        background: linear-gradient(135deg, rgba(102, 217, 239, 0.24), rgba(79, 140, 255, 0.22));
        color: #8ce3ff;
        font-weight: 800;
        box-shadow: inset 0 0 0 1px rgba(140, 227, 255, 0.1);
    }

    .create-upload-btn {
        width: 100%;
        padding: 14px 16px;
        display: flex;
        align-items: center;
        gap: 12px;
        border-radius: 20px;
        background: linear-gradient(135deg, rgba(79, 140, 255, 0.16), rgba(255, 255, 255, 0.04));
        border: 1px solid rgba(132, 192, 255, 0.14);
        box-shadow: 0 14px 28px rgba(0, 0, 0, 0.16), inset 0 1px 0 rgba(255, 255, 255, 0.04);
        color: var(--text);
        cursor: pointer;
        text-align: left;
    }

    .create-upload-copy {
        min-width: 0;
        display: flex;
        flex-direction: column;
        align-items: flex-start;
        gap: 2px;
    }

    .create-upload-title {
        color: #f7fbff;
        font-size: 15px;
        font-weight: 800;
        line-height: 1.2;
    }

    .create-upload-subtitle {
        color: var(--muted);
        font-size: 12px;
        line-height: 1.25;
    }

    .create-menu {
        margin-top: 14px;
        padding: 18px;
        display: grid;
        gap: 16px;
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 24px;
        box-shadow: var(--shadow-md);
        backdrop-filter: blur(18px);
    }

    .create-section {
        display: grid;
        gap: 10px;
    }

    .create-section-title {
        color: var(--muted);
        font-size: 12px;
        font-weight: 800;
        letter-spacing: 0.12em;
        text-transform: uppercase;
    }

    .create-action {
        width: 100%;
        padding: 14px;
        text-align: left;
        border: 1px solid rgba(255, 255, 255, 0.08);
        background: rgba(255, 255, 255, 0.04);
        color: var(--text);
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 12px;
        position: relative;
        overflow: hidden;
    }

    .create-action::after {
        content: "›";
        margin-left: auto;
        color: rgba(236, 244, 255, 0.42);
        font-size: 20px;
        line-height: 1;
    }

    .create-action-icon {
        width: 42px;
        height: 42px;
        border-radius: 14px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
        font-size: 14px;
        font-weight: 900;
        letter-spacing: 0.04em;
        color: #ffffff;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.14);
    }

    .upload-files-icon { background: linear-gradient(135deg, #4f8cff 0%, #66d9ef 100%); }
    .upload-folder-icon { background: linear-gradient(135deg, #e1a23f 0%, #bb7b18 100%); }
    .word-action-icon { background: linear-gradient(135deg, #2859d9 0%, #1b3fa8 100%); }
    .excel-action-icon { background: linear-gradient(135deg, #0c8d58 0%, #0b6c43 100%); }
    .powerpoint-action-icon { background: linear-gradient(135deg, #f06b2b 0%, #d94713 100%); }
    .forms-action-icon { background: linear-gradient(135deg, #f76794 0%, #d63c76 100%); }
    .teams-action-icon { background: linear-gradient(135deg, #7367f0 0%, #4857d6 100%); }

    .upload-files-icon::before { content: "↑"; font-size: 18px; }
    .upload-folder-icon::before { content: "▣"; font-size: 15px; }

    .create-action-copy {
        min-width: 0;
        display: flex;
        flex-direction: column;
        gap: 3px;
    }

    .create-action-label {
        color: #f5f9ff;
        font-size: 14px;
        font-weight: 800;
    }

    .create-action-meta {
        color: var(--muted);
        font-size: 12px;
        line-height: 1.3;
    }

    .nav {
        display: flex;
        flex-direction: column;
        gap: 10px;
    }

    .nav-item {
        width: 100%;
        border: 1px solid transparent;
        background: transparent;
        color: var(--muted-strong);
        border-radius: 18px;
        padding: 13px 14px;
        display: flex;
        align-items: center;
        gap: 12px;
        cursor: pointer;
        font-size: 15px;
        font-weight: 700;
        text-align: left;
    }

    .nav-item:hover,
    .nav-item.active {
        background: linear-gradient(135deg, rgba(79, 140, 255, 0.18), rgba(102, 217, 239, 0.08));
        border-color: rgba(132, 192, 255, 0.18);
        color: #ffffff;
        transform: translateX(2px);
    }

    .nav-icon,
    .filter-icon {
        flex-shrink: 0;
        position: relative;
        display: inline-flex;
        align-items: center;
        justify-content: center;
    }

    .nav-icon {
        width: 34px;
        height: 34px;
        border-radius: 14px;
        background: rgba(255, 255, 255, 0.07);
        border: 1px solid rgba(255, 255, 255, 0.05);
    }

    .nav-icon::before {
        color: #ebf3ff;
        font-size: 16px;
        line-height: 1;
    }

    .icon-home::before { content: "⌂"; }
    .icon-files::before { content: "▣"; }
    .icon-shared::before { content: "⇄"; }
    .icon-favorites::before { content: "★"; }
    .icon-trash::before { content: "⌫"; }

    .main {
        flex: 1;
        padding: 28px;
        display: grid;
        gap: 18px;
        align-content: start;
    }

    .topbar {
        padding: 24px 26px;
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 20px;
        background: linear-gradient(135deg, rgba(11, 24, 42, 0.86), rgba(7, 17, 31, 0.72));
        border: 1px solid var(--border);
        border-radius: 28px;
        box-shadow: var(--shadow-xl);
        backdrop-filter: blur(18px);
    }

    .headline-block {
        min-width: 0;
    }

    .page-kicker {
        margin: 0 0 8px;
        color: #88d8ff;
        font-size: 12px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.18em;
    }

    .topbar h1 {
        margin: 0;
        font-family: "Bahnschrift", "Segoe UI Variable Display", sans-serif;
        font-size: clamp(34px, 5vw, 46px);
        line-height: 1.02;
        letter-spacing: 0.01em;
    }

    .status-stack {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        margin-top: 14px;
    }

    .status-pill {
        margin: 0;
        padding: 8px 12px;
        border-radius: 999px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        background: rgba(255, 255, 255, 0.04);
        color: var(--muted-strong);
        font-size: 13px;
    }

    .topbar-actions {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        justify-content: flex-end;
    }

    .picker-strip,
    .filters {
        padding: 16px 18px;
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
        align-items: center;
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 24px;
        box-shadow: var(--shadow-md);
        backdrop-filter: blur(18px);
    }

    .picker-strip-label,
    .sort-label {
        color: var(--muted);
        font-size: 13px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.12em;
    }

    .picker-chip {
        display: inline-flex;
        align-items: center;
        gap: 10px;
        font-weight: 800;
    }

    .picker-chip.active,
    .filter-chip.active,
    .sort-chip.active {
        background: linear-gradient(135deg, rgba(79, 140, 255, 0.2), rgba(102, 217, 239, 0.12));
        border-color: rgba(132, 192, 255, 0.18);
        box-shadow: 0 12px 28px rgba(0, 0, 0, 0.2);
    }

    .filter-chip,
    .sort-chip {
        border: 1px solid rgba(255, 255, 255, 0.08);
        background: rgba(255, 255, 255, 0.04);
        color: var(--text);
        padding: 11px 16px 11px 12px;
        border-radius: 999px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 800;
        display: inline-flex;
        align-items: center;
        gap: 10px;
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.03);
    }

    .filter-chip:hover,
    .sort-chip:hover,
    .picker-chip:hover {
        background: rgba(255, 255, 255, 0.08);
        border-color: rgba(255, 255, 255, 0.12);
    }

    .filter-icon {
        width: 34px;
        height: 34px;
        border-radius: 14px;
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.08);
    }

    .filter-icon::before {
        color: #ffffff;
        font-size: 13px;
        font-weight: 900;
        letter-spacing: 0.03em;
    }

    .icon-all::before { content: "•"; font-size: 20px; }
    .icon-word { background: linear-gradient(135deg, #2859d9 0%, #1b3fa8 100%); }
    .icon-word::before { content: "W"; }
    .icon-excel { background: linear-gradient(135deg, #0c8d58 0%, #0b6c43 100%); }
    .icon-excel::before { content: "X"; }
    .icon-powerpoint { background: linear-gradient(135deg, #f06b2b 0%, #d94713 100%); }
    .icon-powerpoint::before { content: "P"; }
    .icon-pdf { background: linear-gradient(135deg, #ff5d6c 0%, #cc2b44 100%); }
    .icon-pdf::before { content: "PDF"; font-size: 10px; }

    .workspace-grid {
        display: grid;
        grid-template-columns: minmax(0, 1fr);
        gap: 18px;
    }

    .table-panel,
    .preview-panel,
    .signin-panel {
        padding: 18px;
        border-radius: 28px;
        background: linear-gradient(180deg, rgba(11, 24, 42, 0.88), rgba(8, 16, 31, 0.92));
        border: 1px solid var(--border);
        box-shadow: var(--shadow-xl);
        backdrop-filter: blur(18px);
    }

    .panel-title {
        margin: 0 0 6px;
        font-size: 22px;
    }

    .panel-subtitle {
        margin: 0;
        color: var(--muted);
        font-size: 13px;
    }

    .panel-head {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 14px;
        margin-bottom: 16px;
    }

    .search-input {
        width: 100%;
        margin-top: 16px;
        padding: 13px 15px;
        border-radius: 16px;
        border: 1px solid var(--border);
        background: rgba(255, 255, 255, 0.04);
        color: var(--text);
        outline: none;
    }

    .search-input:focus {
        border-color: rgba(98, 219, 239, 0.4);
    }

    .file-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        overflow: visible;
        border-radius: 24px;
        background: rgba(255, 255, 255, 0.02);
        border: 1px solid rgba(255, 255, 255, 0.04);
    }

    .file-table thead th {
        padding: 14px 16px;
        text-align: left;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: 0.16em;
        text-transform: uppercase;
        color: #8aa4c6;
        border-bottom: 1px solid rgba(255, 255, 255, 0.08);
    }

    .file-table td {
        padding: 16px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.06);
        vertical-align: middle;
    }

    .file-table tbody tr:hover,
    .file-table tbody tr.selected-row {
        background: rgba(255, 255, 255, 0.04);
    }

    .file-table tbody tr:last-child td {
        border-bottom: none;
    }

    .file-name-cell {
        min-width: 280px;
    }

    .file-name-wrapper {
        display: flex;
        align-items: center;
        gap: 12px;
    }

    .file-visual {
        width: 42px;
        height: 42px;
        border-radius: 15px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 11px;
        font-weight: 900;
        letter-spacing: 0.08em;
        color: #ffffff;
        flex-shrink: 0;
    }

    .file-visual-word,
    .file-type-word {
        background: linear-gradient(135deg, #2859d9 0%, #1b3fa8 100%);
    }

    .file-visual-excel,
    .file-type-excel {
        background: linear-gradient(135deg, #0c8d58 0%, #0b6c43 100%);
    }

    .file-visual-powerpoint,
    .file-type-powerpoint {
        background: linear-gradient(135deg, #f06b2b 0%, #d94713 100%);
    }

    .file-visual-pdf,
    .file-type-pdf {
        background: linear-gradient(135deg, #ff5d6c 0%, #cc2b44 100%);
    }

    .file-visual-folder,
    .file-type-folder {
        background: linear-gradient(135deg, #e1a23f 0%, #bb7b18 100%);
    }

    .file-visual-file,
    .file-type-file {
        background: linear-gradient(135deg, #7284a3 0%, #53627b 100%);
    }

    .file-name-link {
        display: inline-block;
        max-width: 100%;
        color: #f7fbff;
        text-decoration: none;
        font-size: 15px;
        font-weight: 800;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .file-row-meta {
        margin: 6px 0 0;
        color: var(--muted);
        font-size: 12px;
        line-height: 1.35;
    }

    .file-type-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-width: 92px;
        padding: 8px 12px;
        border-radius: 999px;
        color: #ffffff;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: 0.03em;
    }

    .table-actions {
        display: flex;
        align-items: center;
        gap: 8px;
        justify-content: flex-end;
    }

    .favorite-toggle-btn,
    .file-menu-btn {
        width: 36px;
        height: 36px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        background: rgba(255, 255, 255, 0.04);
        color: var(--muted-strong);
        cursor: pointer;
    }

    .favorite-toggle-btn.active {
        color: #ffd66b;
        background: rgba(255, 214, 107, 0.12);
        border-color: rgba(255, 214, 107, 0.22);
    }

    .file-menu {
        position: relative;
    }

    .file-menu-dropdown {
        position: absolute;
        top: calc(100% + 10px);
        right: 0;
        min-width: 180px;
        padding: 10px;
        display: grid;
        gap: 8px;
        background: var(--panel-strong);
        border: 1px solid var(--border);
        border-radius: 18px;
        box-shadow: var(--shadow-md);
        z-index: 14;
    }

    .file-menu-dropdown button {
        width: 100%;
        border: 1px solid rgba(255, 255, 255, 0.06);
        background: rgba(255, 255, 255, 0.04);
        color: var(--text);
        border-radius: 12px;
        padding: 10px 12px;
        text-align: left;
        cursor: pointer;
        font-weight: 700;
    }

    .file-menu-dropdown button:hover {
        background: rgba(255, 255, 255, 0.08);
    }

    .preview-panel {
        display: grid;
        grid-template-rows: auto 1fr;
        gap: 16px;
        min-height: 420px;
    }

    .focused-viewer-shell {
        display: flex;
        justify-content: center;
    }

    .focused-viewer-panel {
        width: min(1440px, 100%);
        padding: 18px;
        border-radius: 28px;
        background: linear-gradient(180deg, rgba(11, 24, 42, 0.92), rgba(8, 16, 31, 0.96));
        border: 1px solid var(--border);
        box-shadow: var(--shadow-xl);
        backdrop-filter: blur(18px);
        display: grid;
        grid-template-rows: auto 1fr;
        gap: 16px;
        min-height: calc(100vh - 56px);
    }

    .preview-header {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 14px;
    }

    .preview-actions {
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
        justify-content: flex-end;
    }

    .preview-surface {
        min-height: 620px;
        border-radius: 24px;
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, 0.06);
        background: rgba(255, 255, 255, 0.03);
        position: relative;
    }

    .preview-frame {
        width: 100%;
        height: 100%;
        min-height: 620px;
        border: none;
        background: white;
    }

    .focused-viewer-panel .preview-surface {
        min-height: calc(100vh - 210px);
    }

    .focused-viewer-panel .preview-frame {
        min-height: calc(100vh - 210px);
    }

    .empty-state,
    .signin-panel {
        display: grid;
        place-items: center;
        min-height: 320px;
        padding: 24px;
        text-align: center;
        color: var(--muted);
    }

    .signin-panel {
        width: min(720px, calc(100% - 28px));
        margin: 10vh auto 0;
    }

    .error-banner {
        padding: 12px 14px;
        border-radius: 16px;
        background: rgba(255, 98, 112, 0.12);
        border: 1px solid rgba(255, 98, 112, 0.16);
        color: #ffcad1;
        font-size: 13px;
    }

    .empty-state strong {
        color: #f7fbff;
    }

    @media (max-width: 1180px) {
        .workspace-grid {
            grid-template-columns: 1fr;
        }
    }

    @media (max-width: 980px) {
        .app-shell {
            flex-direction: column;
        }

        .sidebar {
            width: 100%;
            border-right: none;
            border-bottom: 1px solid rgba(255, 255, 255, 0.07);
        }
    }

    @media (max-width: 720px) {
        .main {
            padding: 18px;
        }

        .topbar,
        .preview-header {
            flex-direction: column;
        }

        .file-table thead {
            display: none;
        }

        .file-table,
        .file-table tbody,
        .file-table tr,
        .file-table td {
            display: block;
            width: 100%;
        }

        .file-table tr {
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        }

        .file-table td::before {
            content: attr(data-label);
            display: block;
            margin-bottom: 6px;
            color: var(--muted);
            font-size: 11px;
            font-weight: 800;
            letter-spacing: 0.14em;
            text-transform: uppercase;
        }

        .file-name-cell {
            min-width: 0;
        }

        .table-actions {
            justify-content: flex-start;
        }
    }
`;

let currentView: ViewName = "home";
let currentFilter: FilterName = "all";
let currentSort: SortName = "recent";
let currentSortDirection: "asc" | "desc" = "desc";
let currentSearch = "";
let currentUserName = "";
let currentUserEmail = "";
let lastError = "";
let createMenuOpen = false;
let openFileMenuId = "";

let homeFiles: StartPageItem[] = [];
let myFiles: StartPageItem[] = [];
let sharedFiles: StartPageItem[] = [];
let recycleItems: StartPageItem[] = [];
let favorites: StartPageItem[] = [];
let folderItems: StartPageItem[] = [];
let currentFolderChildren: StartPageItem[] = [];
let currentFolderTrail: Array<{ id: string; name: string }> = [];
let selectedFileId = "";
let focusedViewerMode = false;

let previewEmbedUrl = "";
let previewPostUrl = "";
let previewPostParameters = "";
let previewState: "idle" | "loading" | "ready" | "error" = "idle";
let previewStatusMessage = "";
let appLoadingMessage = "";
let previewMode: "view" | "edit" = "view";

function ensureStyles() {
    const existing = document.getElementById("team-startpage-styles");
    if (existing) return;
    const style = document.createElement("style");
    style.id = "team-startpage-styles";
    style.textContent = appStyles;
    document.head.appendChild(style);
}

function escapeHtml(value: string) {
    return value
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function formatDate(value?: string) {
    if (!value) return "-";
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? value : date.toLocaleDateString();
}

function formatSize(value?: number) {
    if (value === undefined || value === null) return "-";
    if (value < 1024) return `${value} B`;
    const units = ["KB", "MB", "GB", "TB"];
    let size = value;
    let index = 0;
    while (size >= 1024 && index < units.length - 1) {
        size /= 1024;
        index += 1;
    }
    return `${size.toFixed(1)} ${units[index]}`;
}

function fileTypeOf(item: GraphDriveItem): StartPageItem["type"] {
    const effective = item.remoteItem || item;
    if (effective.folder) return "Folder";
    const name = (effective.name || "").toLowerCase();
    if (name.endsWith(".doc") || name.endsWith(".docx")) return "Word";
    if (name.endsWith(".xls") || name.endsWith(".xlsx")) return "Excel";
    if (name.endsWith(".ppt") || name.endsWith(".pptx")) return "Powerpoint";
    if (name.endsWith(".pdf")) return "PDF";
    return "File";
}

function fileFilterKey(type: StartPageItem["type"]): FilterName | "other" {
    if (type === "Word") return "word";
    if (type === "Excel") return "excel";
    if (type === "Powerpoint") return "powerpoint";
    if (type === "PDF") return "pdf";
    return "other";
}

function getOwnerName(item: GraphDriveItem) {
    const effective = item.remoteItem || item;
    return (
        effective.createdBy?.user?.displayName ||
        effective.lastModifiedBy?.user?.displayName ||
        item.createdBy?.user?.displayName ||
        item.lastModifiedBy?.user?.displayName ||
        "Unknown"
    );
}

function parentLabelOf(item: GraphDriveItem) {
    const effective = item.remoteItem || item;
    const rawPath = effective.parentReference?.path || "";
    if (!rawPath) return "Workspace";
    const normalizedPath = (rawPath.includes(":") ? rawPath.split(":").pop() : rawPath) || "";
    const parts = normalizedPath.split("/").filter(Boolean);
    return parts[parts.length - 1] || "Workspace";
}

function normalizeItem(item: GraphDriveItem, source: StartPageItem["source"]): StartPageItem | null {
    const effective = item.remoteItem || item;
    const id = item.id || effective.id;
    const webUrl = item.webUrl || effective.webUrl;
    const name = item.name || effective.name;
    if (!id || !name || !webUrl) return null;

    return {
        id,
        name,
        type: fileTypeOf(item),
        webUrl,
        source,
        modifiedRaw: item.lastModifiedDateTime || effective.lastModifiedDateTime,
        modified: formatDate(item.lastModifiedDateTime || effective.lastModifiedDateTime),
        size: item.size || effective.size,
        owner: getOwnerName(item),
        location: parentLabelOf(item),
        folder: Boolean(effective.folder),
        raw: item,
    };
}

function isFolderItem(item: StartPageItem) {
    return item.folder || item.type === "Folder";
}

function isFileItem(item: StartPageItem) {
    return !isFolderItem(item);
}

function isEditableOfficeFile(item: StartPageItem | null) {
    if (!item) return false;
    return item.type === "Word" || item.type === "Excel" || item.type === "Powerpoint";
}

function typeClass(type: StartPageItem["type"]) {
    if (type === "Word") return "file-type-word file-visual-word";
    if (type === "Excel") return "file-type-excel file-visual-excel";
    if (type === "Powerpoint") return "file-type-powerpoint file-visual-powerpoint";
    if (type === "PDF") return "file-type-pdf file-visual-pdf";
    if (type === "Folder") return "file-type-folder file-visual-folder";
    return "file-type-file file-visual-file";
}

function typeLabel(type: StartPageItem["type"]) {
    if (type === "Word") return "W";
    if (type === "Excel") return "X";
    if (type === "Powerpoint") return "P";
    if (type === "PDF") return "PDF";
    if (type === "Folder") return "DIR";
    return "DOC";
}

function loadFavorites() {
    try {
        favorites = JSON.parse(window.localStorage.getItem(FAVORITES_KEY) || "[]");
    } catch {
        favorites = [];
    }
}

function saveFavorites() {
    window.localStorage.setItem(FAVORITES_KEY, JSON.stringify(favorites));
}

function isFavorited(item: StartPageItem) {
    return favorites.some(favorite => favorite.id === item.id);
}

function toggleFavorite(item: StartPageItem) {
    if (isFavorited(item)) {
        favorites = favorites.filter(favorite => favorite.id !== item.id);
    } else {
        favorites = [...favorites.filter(favorite => favorite.id !== item.id), item];
    }
    saveFavorites();
}

async function loadFiles() {
    const [me, recent, shared, myFilesResult] = await Promise.allSettled([
        graphGet<GraphUser>(`${GRAPH_BASE}/me?$select=displayName,userPrincipalName`),
        graphGet<GraphCollection<GraphDriveItem>>(`${GRAPH_BASE}/me/drive/recent?$top=30`),
        graphGet<GraphCollection<GraphDriveItem>>(`${GRAPH_BASE}/me/drive/sharedWithMe?$top=30`),
        graphGet<GraphCollection<GraphDriveItem>>(`${GRAPH_BASE}/me/drive/root/children?$top=50&$select=id,name,webUrl,size,lastModifiedDateTime,file,folder,parentReference,createdBy,lastModifiedBy`),
    ]);

    if (me.status === "fulfilled") {
        currentUserName = me.value.displayName || "";
        currentUserEmail = me.value.userPrincipalName || "";
    }

    const recentMap = new Map<string, StartPageItem>();
    if (recent.status === "fulfilled") {
        for (const item of recent.value.value || []) {
            const normalized = normalizeItem(item, "Recent");
            if (normalized) recentMap.set(normalized.id, normalized);
        }
    }

    const myFilesList: StartPageItem[] = [];
    if (myFilesResult.status === "fulfilled") {
        for (const item of myFilesResult.value.value || []) {
            const normalized = normalizeItem(item, "My files");
            if (normalized) myFilesList.push(normalized);
        }
    }

    const sharedList: StartPageItem[] = [];
    if (shared.status === "fulfilled") {
        for (const item of shared.value.value || []) {
            const normalized = normalizeItem(item, "Shared");
            if (normalized) sharedList.push(normalized);
        }
    }

    if (currentView === "recycle") {
        const recycleResult = await Promise.allSettled([
            graphGet<GraphCollection<GraphDriveItem>>(`${GRAPH_BASE}/me/drive/recycleBin?$top=50&$select=id,name,webUrl,size,lastModifiedDateTime,file,folder,parentReference,createdBy,lastModifiedBy`),
        ]);
        const recycleList: StartPageItem[] = [];
        if (recycleResult[0]?.status === "fulfilled") {
            for (const item of recycleResult[0].value.value || []) {
                const normalized = normalizeItem(item, "Recycle bin");
                if (normalized) recycleList.push(normalized);
            }
        }
        recycleItems = recycleList;
    }

    myFiles = myFilesList;
    sharedFiles = sharedList;
    folderItems = myFilesList.filter(isFolderItem);

    const homeMap = new Map<string, StartPageItem>();
    [...Array.from(recentMap.values()), ...myFilesList, ...sharedList].filter(isFileItem).forEach(item => {
        if (!homeMap.has(item.id)) {
            homeMap.set(item.id, item.source === "Recent" ? item : { ...item, source: "Recent" });
        }
    });
    homeFiles = Array.from(homeMap.values());

    loadFavorites();
    favorites = favorites
        .map(favorite => {
            return homeFiles.find(item => item.id === favorite.id)
                || myFiles.find(item => item.id === favorite.id)
                || sharedFiles.find(item => item.id === favorite.id)
                || favorite;
        });
    saveFavorites();

    if (!currentFolderTrail.length) {
        currentFolderChildren = folderItems;
    }

    const available = getFilteredAndSortedItems();
    if (!available.some(item => item.id === selectedFileId)) {
        selectedFileId = available[0]?.id || "";
    }
}

function renderLoadingApp(message = "Loading your Microsoft 365 workspace...") {
    const root = document.getElementById("app");
    if (!root) return;
    root.innerHTML = `
        <div class="signin-panel">
            <div>
                <p class="page-kicker">MDCPP Team Startpage</p>
                <h1>Preparing your workspace</h1>
                <p>${escapeHtml(message)}</p>
            </div>
        </div>
    `;
}

async function browseFolder(item: StartPageItem) {
    const response = await graphGet<GraphCollection<GraphDriveItem>>(
        `${GRAPH_BASE}/me/drive/items/${item.id}/children?$top=100&$select=id,name,webUrl,size,lastModifiedDateTime,file,folder,parentReference,createdBy,lastModifiedBy`,
    );

    currentFolderTrail = [...currentFolderTrail, { id: item.id, name: item.name }];
    currentFolderChildren = (response.value || [])
        .map(child => normalizeItem(child, "My files"))
        .filter((child): child is StartPageItem => Boolean(child));
    currentView = "folders";
    currentSearch = "";
    selectedFileId = "";
    previewState = "idle";
    previewStatusMessage = "";
    previewEmbedUrl = "";
    previewPostUrl = "";
    previewPostParameters = "";
}

async function goBackFolderLevel() {
    if (!currentFolderTrail.length) {
        currentFolderChildren = folderItems;
        return;
    }

    const nextTrail = currentFolderTrail.slice(0, -1);
    currentFolderTrail = [];
    currentFolderChildren = folderItems;

    for (const entry of nextTrail) {
        const folder = [...folderItems, ...currentFolderChildren].find(item => item.id === entry.id);
        if (!folder) break;
        await browseFolder(folder);
    }
}

function getViewTitle(view: ViewName) {
    if (view === "home") return "Home";
    if (view === "myfiles") return "My files";
    if (view === "folders") return currentFolderTrail.length ? currentFolderTrail[currentFolderTrail.length - 1].name : "Folders";
    if (view === "shared") return "Shared";
    if (view === "favorites") return "Favorites";
    return "Recycle bin";
}

function getViewItems(view: ViewName) {
    if (view === "home") return homeFiles.filter(isFileItem);
    if (view === "myfiles") return myFiles.filter(isFileItem);
    if (view === "folders") return currentFolderTrail.length ? currentFolderChildren : folderItems;
    if (view === "shared") return sharedFiles.filter(isFileItem);
    if (view === "favorites") return favorites.filter(isFileItem);
    return recycleItems;
}

function sortItems(items: StartPageItem[]) {
    return [...items].sort((left, right) => {
        let comparison = 0;
        if (currentSort === "alphabetical") {
            comparison = left.name.localeCompare(right.name);
        } else if (currentSort === "size") {
            comparison = (left.size || 0) - (right.size || 0);
        } else if (currentSort === "modified") {
            comparison = new Date(left.modifiedRaw || 0).getTime() - new Date(right.modifiedRaw || 0).getTime();
        } else {
            comparison = new Date(right.modifiedRaw || 0).getTime() - new Date(left.modifiedRaw || 0).getTime();
        }
        return currentSortDirection === "desc" ? -comparison : comparison;
    });
}

function getFilteredAndSortedItems() {
    const query = currentSearch.trim().toLowerCase();
    const items = getViewItems(currentView)
        .filter(item => currentFilter === "all" || fileFilterKey(item.type) === currentFilter)
        .filter(item => !query || item.name.toLowerCase().includes(query) || item.location.toLowerCase().includes(query) || item.owner.toLowerCase().includes(query));
    return sortItems(items);
}

function getSelectedItem() {
    return getFilteredAndSortedItems().find(item => item.id === selectedFileId)
        || getViewItems(currentView).find(item => item.id === selectedFileId)
        || homeFiles.find(item => item.id === selectedFileId)
        || null;
}

async function loadPreviewForSelection(mode: "view" | "edit" = "view") {
    const selected = getSelectedItem();
    previewEmbedUrl = "";
    previewPostUrl = "";
    previewPostParameters = "";
    previewMode = mode;

    if (!selected) {
        previewState = "idle";
        previewStatusMessage = "";
        return;
    }

    if (selected.folder || selected.type === "Folder") {
        previewState = "error";
        previewStatusMessage = "Folders cannot be previewed inside the window.";
        return;
    }

    if (currentView === "recycle") {
        previewState = "error";
        previewStatusMessage = "Recycle bin items can be restored, but not previewed inline from this view.";
        return;
    }

    previewState = "loading";
    previewStatusMessage = mode === "edit"
        ? "Loading editable Microsoft 365 view..."
        : "Loading embeddable preview from Microsoft 365...";

    try {
        let preview: PreviewInfo;

        if (mode === "edit" && isEditableOfficeFile(selected)) {
            preview = await graphPost<PreviewInfo>(
                `${GRAPH_BETA_BASE}/me/drive/items/${selected.id}/preview`,
                {
                    allowEdit: true,
                    chromeless: false,
                    viewer: "office",
                },
                ["User.Read", "Files.ReadWrite"],
            );
        } else {
            preview = await graphPost<PreviewInfo>(`${GRAPH_BASE}/me/drive/items/${selected.id}/preview`, {});
        }

        if (preview.getUrl) {
            previewEmbedUrl = preview.getUrl;
            previewState = "ready";
            previewStatusMessage = mode === "edit"
                ? "Editable mode requested. Microsoft 365 may still ask you to authenticate inside the embedded frame."
                : "";
            return;
        }

        if (preview.postUrl && preview.postParameters) {
            previewPostUrl = preview.postUrl;
            previewPostParameters = preview.postParameters;
            previewState = "ready";
            previewStatusMessage = mode === "edit"
                ? "Editable mode requested. Microsoft 365 may still ask you to authenticate inside the embedded frame."
                : "";
            return;
        }

        previewState = "error";
        previewStatusMessage = mode === "edit"
            ? "Microsoft Graph did not return an editable embedded Office URL for this file."
            : "Microsoft Graph returned no embeddable preview URL for this file.";
    } catch (error: any) {
        previewState = "error";
        previewStatusMessage = error?.message || (mode === "edit"
            ? "Could not load editable mode for this file."
            : "Could not load a preview for this file.");
    }
}

async function putDriveItem(path: string, body: Blob | File, contentType?: string) {
    const token = await getGraphToken(["User.Read", "Files.ReadWrite"]);
    const encodedPath = path.split("/").map(segment => encodeURIComponent(segment)).join("/");
    const response = await fetch(`${GRAPH_BASE}/me/drive/root:/${encodedPath}:/content`, {
        method: "PUT",
        headers: {
            Authorization: `Bearer ${token}`,
            ...(contentType ? { "Content-Type": contentType } : {}),
        },
        body,
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
        throw new Error((data as any)?.error?.message || "Upload failed");
    }

    return data as GraphDriveItem;
}

async function uploadFiles(filesToUpload: File[]) {
    if (!filesToUpload.length) return;
    for (const file of filesToUpload) {
        await putDriveItem(file.name, file, file.type);
    }
}

async function uploadFolder(filesToUpload: File[]) {
    if (!filesToUpload.length) return;
    for (const file of filesToUpload) {
        const targetPath = (file as any).webkitRelativePath || file.name;
        await putDriveItem(targetPath, file, file.type);
    }
}

async function createNewFile(kind: "word" | "excel" | "powerpoint" | "forms") {
    const templates = {
        word: { name: "New Word Document.docx", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
        excel: { name: "New Excel Workbook.xlsx", type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        powerpoint: { name: "New Powerpoint Presentation.pptx", type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
        forms: { name: "New Forms Survey.docx", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
    };
    const template = templates[kind];
    await putDriveItem(template.name, new Blob([""], { type: template.type }), template.type);
}

async function deleteItem(item: StartPageItem) {
    const token = await getGraphToken(["User.Read", "Files.ReadWrite"]);
    const response = await fetch(`${GRAPH_BASE}/me/drive/items/${item.id}`, {
        method: "DELETE",
        headers: {
            Authorization: `Bearer ${token}`,
        },
    });

    if (!response.ok && response.status !== 204) {
        const data = await response.json().catch(() => ({}));
        throw new Error((data as any)?.error?.message || "Delete failed");
    }
}

async function downloadItem(item: StartPageItem) {
    const token = await getGraphToken(["User.Read", "Files.Read"]);
    const response = await fetch(`${GRAPH_BASE}/me/drive/items/${item.id}/content`, {
        headers: {
            Authorization: `Bearer ${token}`,
        },
    });
    if (!response.ok) {
        const data = await response.text().catch(() => "");
        throw new Error(data || "Download failed");
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = item.name;
    anchor.click();
    window.URL.revokeObjectURL(url);
}

async function restoreRecycleItem(item: StartPageItem) {
    await graphPost(`${GRAPH_BASE}/me/drive/items/${item.id}/restore`, {});
}

async function refreshData() {
    try {
        appLoadingMessage = "Loading your Microsoft 365 workspace...";
        renderLoadingApp(appLoadingMessage);
        lastError = "";
        await loadFiles();
        if (selectedFileId) {
            await loadPreviewForSelection();
        } else {
            previewState = "idle";
            previewStatusMessage = "";
            previewEmbedUrl = "";
            previewPostUrl = "";
            previewPostParameters = "";
        }
    } catch (error: any) {
        lastError = error?.message || "Could not load Microsoft 365 files.";
    }
    appLoadingMessage = "";
    renderApp();
}

function renderSignIn() {
    const root = document.getElementById("app");
    if (!root) return;
    root.innerHTML = `
        <div class="signin-panel">
            <div>
                <p class="page-kicker">MDCPP Team Startpage</p>
                <h1>Connect Microsoft 365 to open a Teams-style file workspace.</h1>
                <p>This version merges your custom picker UI with the MDCPP-backed start page and in-window previews.</p>
                ${lastError ? `<div class="error-banner">${escapeHtml(lastError)}</div>` : ""}
                <p><button id="signinBtn" class="primary-btn">Connect Microsoft account</button></p>
            </div>
        </div>
    `;
    document.getElementById("signinBtn")?.addEventListener("click", async () => {
        try {
            await signIn();
            await bootSignedInApp();
        } catch (error: any) {
            lastError = error?.message || "Sign-in failed.";
            renderSignIn();
        }
    });
}

function renderPickerStrip() {
    const homeActive = currentView === "home" ? " active" : "";
    const myFilesActive = currentView === "myfiles" ? " active" : "";
    const sharedActive = currentView === "shared" ? " active" : "";
    return `
        <section class="picker-strip">
            <span class="picker-strip-label">Teams-inspired picker actions</span>
            <button class="picker-chip${homeActive}" data-picker-action="recent">Recent</button>
            <button class="picker-chip${myFilesActive}" data-picker-action="cloud">Attach cloud file</button>
            <button class="picker-chip${sharedActive}" data-picker-action="shared">Shared spaces</button>
            <button class="picker-chip" data-picker-action="upload-device">Upload from device</button>
            <button class="picker-chip" data-picker-action="browse-spaces">Browse Teams & channels</button>
        </section>
    `;
}

function renderCreateMenu() {
    if (!createMenuOpen) return "";
    return `
        <div id="createMenu" class="create-menu">
            <div class="create-section">
                <div class="create-section-title">Attach</div>
                <button class="create-action" data-create-action="recent">
                    <span class="create-action-icon teams-action-icon">R</span>
                    <span class="create-action-copy">
                        <span class="create-action-label">Recent files</span>
                        <span class="create-action-meta">Jump to the latest files like the Teams attach picker.</span>
                    </span>
                </button>
                <button class="create-action" data-create-action="shared">
                    <span class="create-action-icon teams-action-icon">S</span>
                    <span class="create-action-copy">
                        <span class="create-action-label">Shared spaces</span>
                        <span class="create-action-meta">Browse shared SharePoint and collaboration files.</span>
                    </span>
                </button>
                <button class="create-action" data-create-action="browse-spaces">
                    <span class="create-action-icon teams-action-icon">T</span>
                    <span class="create-action-copy">
                        <span class="create-action-label">Browse Teams & channels</span>
                        <span class="create-action-meta">Placeholder for the later Zoom + Teams-style space browser.</span>
                    </span>
                </button>
            </div>
            <div class="create-section">
                <div class="create-section-title">Upload</div>
                <button class="create-action" data-create-action="upload-files">
                    <span class="create-action-icon upload-files-icon"></span>
                    <span class="create-action-copy">
                        <span class="create-action-label">Upload files</span>
                        <span class="create-action-meta">Add one or more documents from this device.</span>
                    </span>
                </button>
                <button class="create-action" data-create-action="upload-folder">
                    <span class="create-action-icon upload-folder-icon"></span>
                    <span class="create-action-copy">
                        <span class="create-action-label">Upload folders</span>
                        <span class="create-action-meta">Bring in a complete folder tree.</span>
                    </span>
                </button>
            </div>
            <div class="create-section">
                <div class="create-section-title">Create</div>
                <button class="create-action" data-create-action="new-word">
                    <span class="create-action-icon word-action-icon">W</span>
                    <span class="create-action-copy">
                        <span class="create-action-label">Word document</span>
                        <span class="create-action-meta">Start a formatted draft quickly.</span>
                    </span>
                </button>
                <button class="create-action" data-create-action="new-excel">
                    <span class="create-action-icon excel-action-icon">X</span>
                    <span class="create-action-copy">
                        <span class="create-action-label">Excel workbook</span>
                        <span class="create-action-meta">Create sheets for numbers and planning.</span>
                    </span>
                </button>
                <button class="create-action" data-create-action="new-powerpoint">
                    <span class="create-action-icon powerpoint-action-icon">P</span>
                    <span class="create-action-copy">
                        <span class="create-action-label">Powerpoint</span>
                        <span class="create-action-meta">Build a deck for updates or demos.</span>
                    </span>
                </button>
                <button class="create-action" data-create-action="new-forms">
                    <span class="create-action-icon forms-action-icon">F</span>
                    <span class="create-action-copy">
                        <span class="create-action-label">Forms survey</span>
                        <span class="create-action-meta">Collect responses from your team.</span>
                    </span>
                </button>
            </div>
        </div>
    `;
}

function renderFileTable(items: StartPageItem[]) {
    if (!items.length) {
        return `<div class="empty-state"><div><strong>No files found.</strong><br>Try another filter, search query, or file source.</div></div>`;
    }

    return `
        <table class="file-table">
            <thead>
                <tr>
                    <th>File</th>
                    <th>Type</th>
                    <th>Source</th>
                    <th>Size</th>
                    <th>Last modified</th>
                    <th>Owner</th>
                    <th></th>
                </tr>
            </thead>
            <tbody>
                ${items.map(item => `
                    <tr class="${item.id === selectedFileId ? "selected-row" : ""}">
                        <td class="file-name-cell" data-label="File">
                            <div class="file-name-wrapper">
                                <span class="file-visual ${typeClass(item.type)}">${typeLabel(item.type)}</span>
                                <div>
                                    <a href="${escapeHtml(item.webUrl)}" class="file-name-link" data-select-file="${escapeHtml(item.id)}">${escapeHtml(item.name)}</a>
                                    <p class="file-row-meta">${escapeHtml(item.type)} • ${escapeHtml(item.location)}</p>
                                </div>
                            </div>
                        </td>
                        <td data-label="Type"><span class="file-type-badge ${typeClass(item.type)}">${escapeHtml(item.type)}</span></td>
                        <td data-label="Source">${escapeHtml(item.source)}</td>
                        <td data-label="Size">${escapeHtml(formatSize(item.size))}</td>
                        <td data-label="Last modified">${escapeHtml(item.modified)}</td>
                        <td data-label="Owner">${escapeHtml(item.owner)}</td>
                        <td data-label="Actions">
                            <div class="table-actions">
                                <button class="favorite-toggle-btn${isFavorited(item) ? " active" : ""}" data-favorite-id="${escapeHtml(item.id)}" title="Toggle favorite">${isFavorited(item) ? "★" : "☆"}</button>
                                <button class="table-open-btn" data-open-id="${escapeHtml(item.id)}">${isFolderItem(item) ? "Browse" : "Open"}</button>
                                <div class="file-menu">
                                    <button class="file-menu-btn" data-menu-id="${escapeHtml(item.id)}">⋯</button>
                                    ${openFileMenuId === item.id ? renderMenu(item) : ""}
                                </div>
                            </div>
                        </td>
                    </tr>
                `).join("")}
            </tbody>
        </table>
    `;
}

function renderMenu(item: StartPageItem) {
    if (isFolderItem(item)) {
        return `
            <div class="file-menu-dropdown">
                <button data-menu-action="browse-folder" data-menu-item="${escapeHtml(item.id)}">Open folder</button>
            </div>
        `;
    }

    if (currentView === "recycle") {
        return `
            <div class="file-menu-dropdown">
                <button data-menu-action="restore" data-menu-item="${escapeHtml(item.id)}">Restore</button>
            </div>
        `;
    }

    return `
            <div class="file-menu-dropdown">
                <button data-menu-action="preview" data-menu-item="${escapeHtml(item.id)}">Preview in pane</button>
                <button data-menu-action="download" data-menu-item="${escapeHtml(item.id)}">Download</button>
            ${isEditableOfficeFile(item) ? `<button data-menu-action="edit-inline" data-menu-item="${escapeHtml(item.id)}">Edit in app</button>` : ""}
            <button data-menu-action="open" data-menu-item="${escapeHtml(item.id)}">Edit in Microsoft 365</button>
                <button data-menu-action="delete" data-menu-item="${escapeHtml(item.id)}">Move to recycle bin</button>
            </div>
        `;
}

function renderPreviewSurface(selected: StartPageItem | null) {
    if (!selected) {
        return `
            <div class="empty-state">
                <div>
                    <strong>Select a file to open it here.</strong><br>
                    This pane uses Microsoft Graph preview so files open inside the same window when Microsoft 365 supports embedding.
                </div>
            </div>
        `;
    }

    if (previewState === "loading") {
        return `
            <div class="empty-state">
                <div><strong>Loading preview...</strong><br>${escapeHtml(previewStatusMessage)}</div>
            </div>
        `;
    }

    if (previewState === "error") {
        return `
            <div class="empty-state">
                <div><strong>Preview unavailable</strong><br>${escapeHtml(previewStatusMessage)}</div>
            </div>
        `;
    }

    if (previewEmbedUrl) {
        return `<iframe class="preview-frame" src="${escapeHtml(previewEmbedUrl)}" title="${escapeHtml(selected.name)}"></iframe>`;
    }

    if (previewPostUrl && previewPostParameters) {
        const formId = "preview-post-form";
        window.setTimeout(() => {
            const frame = document.getElementById("preview-post-frame") as HTMLIFrameElement | null;
            const form = document.getElementById(formId) as HTMLFormElement | null;
            if (!frame || !form) return;
            form.innerHTML = "";
            previewPostParameters.split("&").forEach(entry => {
                const [rawKey, rawValue = ""] = entry.split("=");
                const input = document.createElement("input");
                input.type = "hidden";
                input.name = decodeURIComponent(rawKey);
                input.value = decodeURIComponent(rawValue);
                form.appendChild(input);
            });
            form.submit();
        }, 0);

        return `
            <iframe id="preview-post-frame" class="preview-frame" name="preview-post-frame" title="${escapeHtml(selected.name)}"></iframe>
            <form id="${formId}" method="POST" action="${escapeHtml(previewPostUrl)}" target="preview-post-frame"></form>
        `;
    }

    return `
        <div class="empty-state">
            <div><strong>Select a file to open it here.</strong></div>
        </div>
    `;
}

function renderSidebar() {
    return `
        <aside class="sidebar">
            <div class="brand-block">
                <div class="brand-mark" aria-hidden="true"></div>
                <div>
                    <div class="brand">OneDrive</div>
                    <p class="brand-subtitle">Teams-style start page workspace</p>
                </div>
            </div>

            <div class="create-panel">
                <button id="createUploadToggle" class="create-upload-btn">
                    <span class="plus-icon">+</span>
                    <span class="create-upload-copy">
                        <span class="create-upload-title">Create or upload</span>
                        <span class="create-upload-subtitle">Teams-inspired attach options</span>
                    </span>
                </button>
                ${renderCreateMenu()}
            </div>

            <nav class="nav">
                <button class="nav-item${currentView === "home" ? " active" : ""}" data-view="home"><span class="nav-icon icon-home"></span>Home</button>
                <button class="nav-item${currentView === "myfiles" ? " active" : ""}" data-view="myfiles"><span class="nav-icon icon-files"></span>My files</button>
                <button class="nav-item${currentView === "folders" ? " active" : ""}" data-view="folders"><span class="nav-icon icon-files"></span>Folders</button>
                <button class="nav-item${currentView === "shared" ? " active" : ""}" data-view="shared"><span class="nav-icon icon-shared"></span>Shared</button>
                <button class="nav-item${currentView === "favorites" ? " active" : ""}" data-view="favorites"><span class="nav-icon icon-favorites"></span>Favorites</button>
                <button class="nav-item${currentView === "recycle" ? " active" : ""}" data-view="recycle"><span class="nav-icon icon-trash"></span>Recycle bin</button>
            </nav>

            <input id="fileUploadInput" type="file" multiple class="hidden" />
            <input id="folderUploadInput" type="file" webkitdirectory directory multiple class="hidden" />
        </aside>
    `;
}

function renderFocusedViewer(selected: StartPageItem | null) {
    return `
        <div class="app-shell">
            ${renderSidebar()}

            <main class="main">
                <div class="focused-viewer-shell">
                    <section class="focused-viewer-panel">
                        <div class="preview-header">
                            <div>
                                <p class="page-kicker">In-window file view</p>
                                <h2 class="panel-title">${escapeHtml(selected?.name || "No file selected")}</h2>
                                <p class="panel-subtitle">${selected ? `${escapeHtml(selected.type)} • ${escapeHtml(selected.location)} • ${escapeHtml(selected.source)} • ${escapeHtml(selected.modified)}` : "Select a file from the left to preview it here."}</p>
                            </div>
                            <div class="preview-actions">
                                <button id="backToHomeBtn" class="secondary-btn">Back</button>
                                <button id="previewRefreshBtn" class="secondary-btn">Refresh ${previewMode === "edit" ? "editor" : "preview"}</button>
                                ${selected && isEditableOfficeFile(selected) ? `<button id="editInlineBtn" class="secondary-btn">Edit in app</button>` : ""}
                                ${selected ? `<a class="primary-btn" href="${escapeHtml(selected.webUrl)}" target="_blank" rel="noopener noreferrer">Edit in Microsoft 365</a>` : ""}
                            </div>
                        </div>
                        <div class="preview-surface">
                            ${renderPreviewSurface(selected)}
                        </div>
                    </section>
                </div>
            </main>
        </div>
    `;
}

function renderApp() {
    const root = document.getElementById("app");
    if (!root) return;

    const items = getFilteredAndSortedItems();
    const selected = getSelectedItem();
    if (focusedViewerMode) {
        root.innerHTML = renderFocusedViewer(selected);
        attachEventHandlers();
        return;
    }
    root.innerHTML = `
        <div class="app-shell">
            ${renderSidebar()}

            <main class="main">
                <header class="topbar">
                    <div class="headline-block">
                        <p class="page-kicker">Microsoft 365 Team Startpage</p>
                        <h1>${escapeHtml(getViewTitle(currentView))}</h1>
                        <div class="status-stack">
                            <p class="status-pill">${escapeHtml(currentUserName || "Signed in user")} ${currentUserEmail ? `• ${escapeHtml(currentUserEmail)}` : ""}</p>
                            <p class="status-pill">${escapeHtml(items.length.toString())} file(s)</p>
                            ${selectedFileId && selected ? `<p class="status-pill">Selected: ${escapeHtml(selected.name)}</p>` : `<p class="status-pill">No file selected</p>`}
                        </div>
                        ${lastError ? `<div class="error-banner" style="margin-top:12px;">${escapeHtml(lastError)}</div>` : ""}
                    </div>
                    <div class="topbar-actions">
                        <button id="openSelectedBtn" ${selected ? "" : "disabled"}>Open selected</button>
                        <button id="refreshBtn">Refresh</button>
                        <button id="signoutBtn">Sign out</button>
                    </div>
                </header>

                ${renderPickerStrip()}

                <section class="filters">
                    <button class="filter-chip${currentFilter === "all" ? " active" : ""}" data-filter="all"><span class="filter-icon icon-all"></span>All</button>
                    <button class="filter-chip${currentFilter === "word" ? " active" : ""}" data-filter="word"><span class="filter-icon icon-word"></span>Word</button>
                    <button class="filter-chip${currentFilter === "excel" ? " active" : ""}" data-filter="excel"><span class="filter-icon icon-excel"></span>Excel</button>
                    <button class="filter-chip${currentFilter === "powerpoint" ? " active" : ""}" data-filter="powerpoint"><span class="filter-icon icon-powerpoint"></span>Powerpoint</button>
                    <button class="filter-chip${currentFilter === "pdf" ? " active" : ""}" data-filter="pdf"><span class="filter-icon icon-pdf"></span>PDF</button>
                    <span class="sort-label">Sort by:</span>
                    <button class="sort-chip${currentSort === "recent" ? " active" : ""}" data-sort="recent">Recent</button>
                    <button class="sort-chip${currentSort === "alphabetical" ? " active" : ""}" data-sort="alphabetical">Alphabetical</button>
                    <button class="sort-chip${currentSort === "size" ? " active" : ""}" data-sort="size">File size</button>
                    <button class="sort-chip${currentSort === "modified" ? " active" : ""}" data-sort="modified">Last modified</button>
                </section>

                <section class="workspace-grid">
                    <section class="table-panel">
                        <div class="panel-head">
                            <div>
                                <h2 class="panel-title">${escapeHtml(getViewTitle(currentView))}</h2>
                                <p class="panel-subtitle">${currentView === "folders" ? `Browse folders${currentFolderTrail.length ? ` • ${escapeHtml(currentFolderTrail.map(item => item.name).join(" / "))}` : " from your root drive."}` : "Merged custom picker + MDCPP start page with Teams-inspired file sources."}</p>
                            </div>
                            ${currentView === "folders" && currentFolderTrail.length ? `<button id="folderBackBtn" class="secondary-btn">Back one level</button>` : ""}
                        </div>
                        <input id="searchInput" class="search-input" placeholder="Search files, owners, or locations..." value="${escapeHtml(currentSearch)}" />
                        <div style="margin-top:16px;">${renderFileTable(items)}</div>
                    </section>
                </section>
            </main>
        </div>
    `;

    attachEventHandlers();
}

function attachEventHandlers() {
    document.querySelectorAll<HTMLElement>("[data-view]").forEach(button => {
        button.addEventListener("click", async () => {
            currentView = button.dataset.view as ViewName;
            if (currentView === "folders" && !currentFolderTrail.length) {
                currentFolderChildren = folderItems;
            }
            openFileMenuId = "";
            const items = getFilteredAndSortedItems();
            if (selectedFileId && !items.some(item => item.id === selectedFileId)) {
                selectedFileId = "";
                previewState = "idle";
                previewStatusMessage = "";
                previewEmbedUrl = "";
                previewPostUrl = "";
                previewPostParameters = "";
            }
            if (selectedFileId) {
                await loadPreviewForSelection();
            }
            renderApp();
        });
    });

    document.querySelectorAll<HTMLElement>("[data-filter]").forEach(button => {
        button.addEventListener("click", async () => {
            currentFilter = button.dataset.filter as FilterName;
            const items = getFilteredAndSortedItems();
            if (selectedFileId && !items.some(item => item.id === selectedFileId)) {
                selectedFileId = "";
                previewState = "idle";
                previewStatusMessage = "";
                previewEmbedUrl = "";
                previewPostUrl = "";
                previewPostParameters = "";
            }
            if (selectedFileId) {
                await loadPreviewForSelection();
            }
            renderApp();
        });
    });

    document.querySelectorAll<HTMLElement>("[data-sort]").forEach(button => {
        button.addEventListener("click", async () => {
            const nextSort = button.dataset.sort as SortName;
            if (nextSort === currentSort) {
                currentSortDirection = currentSortDirection === "asc" ? "desc" : "asc";
            } else {
                currentSort = nextSort;
                currentSortDirection = nextSort === "recent" ? "desc" : "asc";
            }
            const items = getFilteredAndSortedItems();
            if (selectedFileId && !items.some(item => item.id === selectedFileId)) {
                selectedFileId = "";
                previewState = "idle";
                previewStatusMessage = "";
                previewEmbedUrl = "";
                previewPostUrl = "";
                previewPostParameters = "";
            }
            if (selectedFileId) {
                await loadPreviewForSelection();
            }
            renderApp();
        });
    });

    document.querySelectorAll<HTMLElement>("[data-picker-action]").forEach(button => {
        button.addEventListener("click", async () => {
            const action = button.dataset.pickerAction;
            if (action === "recent") {
                currentView = "home";
            } else if (action === "cloud") {
                currentView = "myfiles";
            } else if (action === "shared" || action === "browse-spaces") {
                currentView = "shared";
                if (action === "browse-spaces") {
                    lastError = "Teams & channels browsing is mapped to shared SharePoint spaces for this prototype.";
                } else {
                    lastError = "";
                }
            } else if (action === "upload-device") {
                (document.getElementById("fileUploadInput") as HTMLInputElement | null)?.click();
                return;
            }

            const items = getFilteredAndSortedItems();
            if (selectedFileId && !items.some(item => item.id === selectedFileId)) {
                selectedFileId = "";
                previewState = "idle";
                previewStatusMessage = "";
                previewEmbedUrl = "";
                previewPostUrl = "";
                previewPostParameters = "";
            }
            if (selectedFileId) {
                await loadPreviewForSelection();
            }
            renderApp();
        });
    });

    document.getElementById("createUploadToggle")?.addEventListener("click", () => {
        createMenuOpen = !createMenuOpen;
        renderApp();
    });

    document.querySelectorAll<HTMLElement>("[data-create-action]").forEach(button => {
        button.addEventListener("click", async () => {
            const action = button.dataset.createAction;
            createMenuOpen = false;
            try {
                lastError = "";
                if (action === "recent") {
                    currentView = "home";
                } else if (action === "shared" || action === "browse-spaces") {
                    currentView = "shared";
                    if (action === "browse-spaces") {
                        lastError = "Teams & channels browsing is represented by shared SharePoint spaces in this prototype.";
                    }
                } else if (action === "upload-files") {
                    (document.getElementById("fileUploadInput") as HTMLInputElement | null)?.click();
                    renderApp();
                    return;
                } else if (action === "upload-folder") {
                    (document.getElementById("folderUploadInput") as HTMLInputElement | null)?.click();
                    renderApp();
                    return;
                } else if (action === "new-word") {
                    await createNewFile("word");
                } else if (action === "new-excel") {
                    await createNewFile("excel");
                } else if (action === "new-powerpoint") {
                    await createNewFile("powerpoint");
                } else if (action === "new-forms") {
                    await createNewFile("forms");
                }
                await refreshData();
            } catch (error: any) {
                lastError = error?.message || "Action failed.";
                renderApp();
            }
        });
    });

    (document.getElementById("fileUploadInput") as HTMLInputElement | null)?.addEventListener("change", async event => {
        const input = event.target as HTMLInputElement;
        const filesToUpload = Array.from(input.files || []);
        input.value = "";
        if (!filesToUpload.length) return;
        try {
            lastError = "";
            await uploadFiles(filesToUpload);
            await refreshData();
        } catch (error: any) {
            lastError = error?.message || "Upload failed.";
            renderApp();
        }
    });

    (document.getElementById("folderUploadInput") as HTMLInputElement | null)?.addEventListener("change", async event => {
        const input = event.target as HTMLInputElement;
        const filesToUpload = Array.from(input.files || []);
        input.value = "";
        if (!filesToUpload.length) return;
        try {
            lastError = "";
            await uploadFolder(filesToUpload);
            await refreshData();
        } catch (error: any) {
            lastError = error?.message || "Folder upload failed.";
            renderApp();
        }
    });

    document.getElementById("refreshBtn")?.addEventListener("click", async () => {
        await refreshData();
    });

    document.getElementById("previewRefreshBtn")?.addEventListener("click", async () => {
        await loadPreviewForSelection(previewMode);
        renderApp();
    });

    document.getElementById("editInlineBtn")?.addEventListener("click", async () => {
        const selected = getSelectedItem();
        if (!selected || !isEditableOfficeFile(selected)) return;
        await loadPreviewForSelection("edit");
        focusedViewerMode = true;
        renderApp();
    });

    document.getElementById("folderBackBtn")?.addEventListener("click", async () => {
        await goBackFolderLevel();
        renderApp();
    });

    document.getElementById("openSelectedBtn")?.addEventListener("click", async () => {
        const selected = getSelectedItem();
        if (selected) {
            focusedViewerMode = true;
            await loadPreviewForSelection();
            renderApp();
        }
    });

    document.getElementById("backToHomeBtn")?.addEventListener("click", () => {
        focusedViewerMode = false;
        renderApp();
    });

    document.getElementById("signoutBtn")?.addEventListener("click", async () => {
        await signOut();
        homeFiles = [];
        myFiles = [];
        sharedFiles = [];
        recycleItems = [];
        folderItems = [];
        currentFolderChildren = [];
        currentFolderTrail = [];
        favorites = [];
        selectedFileId = "";
        currentUserName = "";
        currentUserEmail = "";
        currentSearch = "";
        renderSignIn();
    });

    (document.getElementById("searchInput") as HTMLInputElement | null)?.addEventListener("input", async event => {
        currentSearch = (event.target as HTMLInputElement).value;
        const items = getFilteredAndSortedItems();
        if (selectedFileId && !items.some(item => item.id === selectedFileId)) {
            selectedFileId = "";
            previewState = "idle";
            previewStatusMessage = "";
            previewEmbedUrl = "";
            previewPostUrl = "";
            previewPostParameters = "";
        }
        renderApp();
    });

    document.querySelectorAll<HTMLElement>("[data-select-file],[data-open-id]").forEach(element => {
        element.addEventListener("click", async event => {
            event.preventDefault();
            const fileId = (event.currentTarget as HTMLElement).dataset.selectFile || (event.currentTarget as HTMLElement).dataset.openId || "";
            const item = [...homeFiles, ...myFiles, ...sharedFiles, ...favorites, ...recycleItems, ...folderItems, ...currentFolderChildren].find(candidate => candidate.id === fileId);
            openFileMenuId = "";
            if (!item) return;
            if (isFolderItem(item)) {
                await browseFolder(item);
            } else {
                selectedFileId = fileId;
                await loadPreviewForSelection("view");
                focusedViewerMode = true;
            }
            renderApp();
        });
    });

    document.querySelectorAll<HTMLElement>("[data-favorite-id]").forEach(button => {
        button.addEventListener("click", event => {
            event.preventDefault();
            const fileId = (event.currentTarget as HTMLElement).dataset.favoriteId || "";
            const item = [...homeFiles, ...myFiles, ...sharedFiles, ...favorites].find(candidate => candidate.id === fileId);
            if (!item) return;
            toggleFavorite(item);
            renderApp();
        });
    });

    document.querySelectorAll<HTMLElement>("[data-menu-id]").forEach(button => {
        button.addEventListener("click", event => {
            event.preventDefault();
            const fileId = (event.currentTarget as HTMLElement).dataset.menuId || "";
            openFileMenuId = openFileMenuId === fileId ? "" : fileId;
            renderApp();
        });
    });

    document.querySelectorAll<HTMLElement>("[data-menu-action]").forEach(button => {
        button.addEventListener("click", async event => {
            event.preventDefault();
            const action = (event.currentTarget as HTMLElement).dataset.menuAction || "";
            const itemId = (event.currentTarget as HTMLElement).dataset.menuItem || "";
            const item = [...homeFiles, ...myFiles, ...sharedFiles, ...favorites, ...recycleItems].find(candidate => candidate.id === itemId);
            openFileMenuId = "";
            if (!item) {
                renderApp();
                return;
            }

            try {
                if (action === "preview") {
                    selectedFileId = item.id;
                    await loadPreviewForSelection("view");
                    focusedViewerMode = true;
                } else if (action === "edit-inline") {
                    selectedFileId = item.id;
                    await loadPreviewForSelection("edit");
                    focusedViewerMode = true;
                } else if (action === "browse-folder") {
                    await browseFolder(item);
                } else if (action === "download") {
                    await downloadItem(item);
                } else if (action === "open") {
                    selectedFileId = item.id;
                    await loadPreviewForSelection("view");
                    focusedViewerMode = true;
                } else if (action === "delete") {
                    await deleteItem(item);
                    favorites = favorites.filter(favorite => favorite.id !== item.id);
                    saveFavorites();
                    await refreshData();
                    return;
                } else if (action === "restore") {
                    await restoreRecycleItem(item);
                    await refreshData();
                    return;
                }
            } catch (error: any) {
                lastError = error?.message || "Action failed.";
            }
            renderApp();
        });
    });

    document.addEventListener("click", event => {
        const target = event.target as HTMLElement;
        if (!target.closest(".file-menu")) {
            if (openFileMenuId) {
                openFileMenuId = "";
                renderApp();
            }
        }
        if (!target.closest(".create-panel")) {
            if (createMenuOpen) {
                createMenuOpen = false;
                renderApp();
            }
        }
    }, { once: true });
}

async function bootSignedInApp() {
    await refreshData();
}

export async function launch() {
    ensureStyles();
    document.body.innerHTML = `<div id="app"></div>`;
    await initializeAuth();
    renderSignIn();
}
