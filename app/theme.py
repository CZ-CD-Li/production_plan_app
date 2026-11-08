"""全局 QSS 主题：黑、白、灰三段式配色，轻量 Bootstrap 风格。"""

from __future__ import annotations


def build_stylesheet() -> str:
    return """
    * {
        font-family: "Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", sans-serif;
        outline: none;
    }
    QWidget {
        color: #1A1A1A;
        font-size: 13px;
    }
    QMainWindow, #windowFrame, #appBody, #contentArea, #loginFrame {
        background-color: #F4F4F4;
    }
    QToolTip {
        background-color: #1A1A1A;
        color: #FFFFFF;
        border: none;
        padding: 6px 9px;
        border-radius: 4px;
        font-size: 12px;
    }

    /* ---------- 自绘标题栏 ---------- */
    #titleBar {
        background: #FFFFFF;
        border-bottom: 1px solid #E6E6E6;
    }
    #titleBar[theme="dark"] {
        background: #0D0D0D;
        border-bottom: 1px solid #242424;
    }
    #titleBarTitle {
        color: #1A1A1A;
        font-size: 12px;
        font-weight: 600;
    }
    #titleBar[theme="dark"] #titleBarTitle {
        color: #F2F2F2;
    }
    #titleBarSub {
        color: #9B9B9B;
        font-size: 11px;
    }
    #titleBar[theme="dark"] #titleBarSub {
        color: #7A7A7A;
    }

    /* ---------- 登录 ---------- */
    #loginFrame {
        background: #0B0B0B;
        border: 1px solid #242424;
    }
    #loginCard {
        background: #FFFFFF;
        border-radius: 18px;
    }
    #brandPanel {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
            stop:0 #1C1C1C, stop:0.55 #121212, stop:1 #070707);
        border-top-left-radius: 18px;
        border-bottom-left-radius: 18px;
    }
    #brandTitle {
        color: #FFFFFF;
        font-size: 27px;
        font-weight: 600;
    }
    #brandSubtitle {
        color: #A8A8A8;
        font-size: 13px;
    }
    #brandFeature {
        color: #CFCFCF;
        font-size: 12px;
    }
    #brandFooter {
        color: #6E6E6E;
        font-size: 11px;
    }
    #brandBadge {
        background: rgba(255, 255, 255, 0.07);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 9px;
        color: #D6D6D6;
        font-size: 11px;
        padding: 4px 10px;
    }
    #loginTitle {
        color: #111111;
        font-size: 25px;
        font-weight: 600;
    }
    #loginSubtitle {
        color: #808080;
        font-size: 12px;
    }
    #fieldLabel {
        color: #4A4A4A;
        font-size: 12px;
        font-weight: 600;
    }
    QLineEdit#loginInput {
        background: #FFFFFF;
        border: 1px solid #DCDCDC;
        border-radius: 9px;
        padding: 11px 12px;
        font-size: 13px;
        min-height: 22px;
        selection-background-color: #D9D9D9;
        selection-color: #111111;
    }
    QLineEdit#loginInput:hover {
        border-color: #C4C4C4;
    }
    QLineEdit#loginInput:focus {
        border-color: #1A1A1A;
        background: #FFFFFF;
    }
    QPushButton#loginButton {
        background: #1A1A1A;
        color: #FFFFFF;
        border: none;
        border-radius: 10px;
        font-size: 14px;
        font-weight: 600;
        min-height: 44px;
    }
    QPushButton#loginButton:hover {
        background: #333333;
    }
    QPushButton#loginButton:pressed {
        background: #000000;
    }
    #loginError {
        color: #1A1A1A;
        font-size: 12px;
        font-weight: 600;
    }
    #loginHintBox {
        background: #F7F7F7;
        border: 1px solid #EAEAEA;
        border-radius: 9px;
    }
    #loginHint {
        color: #808080;
        font-size: 11px;
    }
    #loginHintStrong {
        color: #333333;
        font-size: 11px;
        font-weight: 600;
    }

    /* ---------- 侧边栏 ---------- */
    #sidebar {
        background-color: #141414;
    }
    #avatarCard {
        background: #1C1C1C;
        border: 1px solid #2A2A2A;
        border-radius: 12px;
    }
    #avatarCard:hover {
        background: #242424;
        border-color: #3A3A3A;
    }
    #userName {
        color: #FFFFFF;
        font-size: 14px;
        font-weight: 600;
    }
    #userMeta {
        color: #8E8E8E;
        font-size: 11px;
    }
    #navSection {
        color: #6A6A6A;
        font-size: 11px;
        font-weight: 600;
    }
    #navButton {
        background: transparent;
        border: none;
        border-radius: 9px;
        color: #A3A3A3;
        font-size: 13px;
        text-align: left;
        padding-left: 17px;
        min-height: 44px;
    }
    #navButton:hover {
        color: #F5F5F5;
    }
    #navButton:checked {
        color: #FFFFFF;
        font-weight: 600;
    }
    #navSubButton {
        background: transparent;
        border: none;
        border-radius: 7px;
        color: #8B8B8B;
        font-size: 12px;
        text-align: left;
        padding-left: 34px;
        min-height: 32px;
    }
    #navSubButton:hover {
        color: #EFEFEF;
    }
    #navSubButton:checked {
        color: #FFFFFF;
        font-weight: 600;
    }
    #sidebarFooter {
        color: #5E5E5E;
        font-size: 11px;
    }
    #sidebarVersion {
        color: #4B4B4B;
        font-size: 10px;
    }

    /* ---------- 顶栏 ---------- */
    #topbar {
        background: #FFFFFF;
        border-bottom: 1px solid #E6E6E6;
    }
    #pageTitle {
        color: #111111;
        font-size: 20px;
        font-weight: 600;
    }
    #pageSubtitle {
        color: #8A8A8A;
        font-size: 12px;
    }
    #dateChip {
        background: #F5F5F5;
        border-radius: 8px;
        color: #5C5C5C;
        font-size: 12px;
        padding: 6px 12px;
    }
    #statusText {
        color: #7A7A7A;
        font-size: 12px;
    }
    #statusText[state="error"] {
        color: #111111;
    }
    #statusText[state="ok"] {
        color: #5C5C5C;
    }

    /* ---------- 通用控件 ---------- */
    QPushButton#primaryButton {
        background: #1A1A1A;
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        padding: 9px 20px;
        font-weight: 600;
    }
    QPushButton#primaryButton:hover {
        background: #333333;
    }
    QPushButton#primaryButton:pressed {
        background: #000000;
    }
    QPushButton#primaryButton:disabled {
        background: #D6D6D6;
        color: #F7F7F7;
    }
    QPushButton#secondaryButton {
        background: #FFFFFF;
        border: 1px solid #DCDCDC;
        border-radius: 8px;
        color: #333333;
        padding: 8px 18px;
    }
    QPushButton#secondaryButton:hover {
        background: #F5F5F5;
        border-color: #B8B8B8;
        color: #111111;
    }
    QPushButton#secondaryButton:pressed {
        background: #EDEDED;
    }
    QPushButton#ghostButton {
        background: transparent;
        border: none;
        color: #6B6B6B;
        padding: 6px 10px;
        border-radius: 6px;
    }
    QPushButton#ghostButton:hover {
        background: #F2F2F2;
        color: #111111;
    }
    QPushButton#linkButton {
        background: transparent;
        border: none;
        color: #333333;
        font-size: 12px;
        padding: 4px 8px;
        border-radius: 5px;
    }
    QPushButton#linkButton:hover {
        background: #F0F0F0;
        color: #000000;
    }
    QPushButton#dangerLinkButton {
        background: transparent;
        border: none;
        color: #4A4A4A;
        font-size: 12px;
        padding: 4px 8px;
        border-radius: 5px;
    }
    QPushButton#dangerLinkButton:hover {
        background: #EFEFEF;
        color: #000000;
    }
    QFrame#dropArea {
        background: #FAFAFA;
        border: 1px dashed #C9C9C9;
        border-radius: 10px;
    }
    QFrame#dropArea:hover {
        background: #F5F5F5;
        border-color: #9E9E9E;
    }
    QFrame#dropArea[dragging="true"] {
        background: #EFEFEF;
        border: 1px dashed #1A1A1A;
    }
    QListWidget {
        background: #FBFBFB;
        border: 1px solid #ECECEC;
        border-radius: 8px;
        padding: 4px;
        outline: none;
    }
    QListWidget::item {
        color: #333333;
        padding: 7px 9px;
        border-radius: 6px;
    }
    QListWidget::item:hover {
        background: #F2F2F2;
    }
    QListWidget::item:selected {
        background: #EAEAEA;
        color: #111111;
    }
    QLineEdit, QComboBox {
        background: #FFFFFF;
        border: 1px solid #DCDCDC;
        border-radius: 8px;
        padding: 8px 11px;
        min-height: 20px;
        selection-background-color: #D9D9D9;
        selection-color: #111111;
    }
    QLineEdit:hover, QComboBox:hover {
        border-color: #C4C4C4;
    }
    QLineEdit:focus, QComboBox:focus {
        border-color: #1A1A1A;
        background: #FFFFFF;
    }
    QLineEdit[readOnly="true"] {
        background: #F7F7F7;
        color: #5C5C5C;
    }
    QComboBox::drop-down {
        border: none;
        width: 26px;
    }
    QComboBox QAbstractItemView {
        background: #FFFFFF;
        border: 1px solid #DCDCDC;
        border-radius: 6px;
        padding: 4px;
        selection-background-color: #EAEAEA;
        selection-color: #111111;
        outline: none;
    }
    QFrame#card {
        background: #FFFFFF;
        border: 1px solid #E6E6E6;
        border-radius: 12px;
    }
    #sectionTitle {
        color: #111111;
        font-size: 15px;
        font-weight: 600;
    }
    #sectionHint {
        color: #8A8A8A;
        font-size: 12px;
    }
    #fieldHint {
        color: #9B9B9B;
        font-size: 11px;
    }
    #pill {
        background: #EDEDED;
        border-radius: 10px;
        color: #333333;
        font-size: 11px;
        padding: 4px 9px;
    }
    #pillMuted {
        background: #F3F3F3;
        border-radius: 10px;
        color: #7A7A7A;
        font-size: 11px;
        padding: 4px 9px;
    }
    #divider {
        background: #ECECEC;
        border: none;
    }
    #inlineSuccess {
        color: #333333;
        font-size: 12px;
    }
    #inlineError {
        color: #111111;
        font-size: 12px;
        font-weight: 600;
    }

    /* ---------- 计划页 ---------- */
    #noticeBanner {
        background: #F5F5F5;
        border: 1px solid #E3E3E3;
        border-radius: 10px;
    }
    #noticeBannerError {
        background: #FFFFFF;
        border: 1px solid #C9C9C9;
        border-radius: 10px;
    }
    #noticeText {
        color: #444444;
        font-size: 12px;
    }
    #noticeText[level="error"] {
        color: #111111;
    }
    QFrame#segmented {
        background: #F1F1F1;
        border: 1px solid #E4E4E4;
        border-radius: 10px;
    }
    #slideSegmentButton, QPushButton#segmentButton {
        background: transparent;
        border: none;
        border-radius: 8px;
        color: #6B6B6B;
        font-weight: 600;
        font-size: 12px;
        padding: 7px 24px;
        min-width: 58px;
    }
    #slideSegmentButton:hover, QPushButton#segmentButton:hover {
        color: #1A1A1A;
    }
    #slideSegmentButton:checked, QPushButton#segmentButton:checked {
        color: #111111;
    }
    #tableMeta {
        color: #9B9B9B;
        font-size: 11px;
    }
    #countPill {
        background: #F2F2F2;
        border-radius: 10px;
        color: #5C5C5C;
        font-size: 11px;
        padding: 4px 10px;
    }
    QTableWidget {
        background: #FFFFFF;
        alternate-background-color: #FBFBFB;
        border: 1px solid #ECECEC;
        border-radius: 10px;
        gridline-color: transparent;
        selection-background-color: #EAEAEA;
        selection-color: #111111;
    }
    QTableWidget::item {
        padding: 0 10px;
        border-bottom: 1px solid #F2F2F2;
    }
    QTableWidget::item:selected {
        background: #EAEAEA;
        color: #111111;
    }
    QTableWidget::item:hover {
        background: #F7F7F7;
    }
    QHeaderView::section {
        background: #F7F7F7;
        border: none;
        border-bottom: 1px solid #E8E8E8;
        color: #5C5C5C;
        font-size: 12px;
        font-weight: 600;
        padding: 10px 12px;
    }
    QTableCornerButton::section {
        background: #F7F7F7;
        border: none;
    }

    /* ---------- 设置 ---------- */
    #settingsFooter {
        background: #FFFFFF;
        border-top: 1px solid #E6E6E6;
    }
    #avatarLarge {
        background: transparent;
    }
    #accountName {
        color: #111111;
        font-size: 15px;
        font-weight: 600;
    }
    #accountMeta {
        color: #8A8A8A;
        font-size: 12px;
    }

    /* ---------- 模块占位 ---------- */
    #placeholderTitle {
        color: #111111;
        font-size: 19px;
        font-weight: 600;
    }
    #placeholderText {
        color: #8A8A8A;
        font-size: 13px;
    }
    #placeholderFeature {
        color: #5C5C5C;
        font-size: 12px;
    }
    #placeholderTag {
        background: #F0F0F0;
        border-radius: 10px;
        color: #444444;
        font-size: 11px;
        padding: 4px 10px;
    }

    /* ---------- 滚动条 ---------- */
    QScrollArea {
        background: transparent;
        border: none;
    }
    QScrollArea > QWidget > QWidget {
        background: transparent;
    }
    QScrollBar:vertical {
        background: transparent;
        width: 8px;
        margin: 2px;
    }
    QScrollBar::handle:vertical {
        background: #D4D4D4;
        border-radius: 4px;
        min-height: 28px;
    }
    QScrollBar::handle:vertical:hover {
        background: #B8B8B8;
    }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
        height: 0;
    }
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
        background: transparent;
    }
    QScrollBar:horizontal {
        background: transparent;
        height: 8px;
        margin: 2px;
    }
    QScrollBar::handle:horizontal {
        background: #D4D4D4;
        border-radius: 4px;
        min-width: 28px;
    }
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
        width: 0;
    }
    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
        background: transparent;
    }

    QMessageBox {
        background: #FFFFFF;
    }
    QMessageBox QLabel {
        color: #1A1A1A;
    }
    """
