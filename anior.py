"""
AniOr - 动漫视频手动整理工具
通过 GUI 拖放操作，将视频文件硬链接/剪切/复制到目标目录
"""
import os
import sys
import json
import shutil
import re
import requests
from pathlib import Path
from typing import Optional, Dict, List, Any, Tuple

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QTreeWidget, QTreeWidgetItem, QGroupBox, QLabel,
    QPushButton, QLineEdit, QFileDialog, QMessageBox, QFrame,
    QScrollArea, QHeaderView, QStatusBar, QCheckBox, QComboBox,
    QDialog, QDialogButtonBox, QFormLayout, QTabWidget, QSizePolicy,
    QAbstractItemView, QTableWidget, QTableWidgetItem
)
from PyQt5.QtCore import Qt, QMimeData, QThread, pyqtSignal, QSize, QTimer, QUrl, QSettings, QByteArray
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QDrag, QPixmap, QColor, QBrush, QDesktopServices, QFont, QIcon

# 配置文件路径
if getattr(sys, 'frozen', False):
    CONFIG_DIR = Path(sys.executable).parent / 'config'
else:
    CONFIG_DIR = Path(__file__).parent / 'config'

CONFIG_PATH = CONFIG_DIR / 'config.json'


# ==================== 用户配置管理 ====================
class Config:
    # 默认视频格式（覆盖主流 + 参考项目格式）
    DEFAULT_VIDEO_EXTENSIONS = [
        '.mp4', '.mkv', '.avi', '.wmv', '.flv',  # 主流格式
        '.webm', '.m4v', '.mov', '.ts',          # 现代格式
        '.mpg', '.mpeg',                         # MPEG 格式
        '.rm', '.rmvb',                          # RealMedia（参考项目支持）
    ]
    
    # 字幕文件格式
    SUBTITLE_EXTENSIONS = ['.srt', '.ass', '.ssa', '.sub', '.idx', '.vtt']

    DEFAULT = {
        'source_dir': '',           # 源目录
        'target_dir': '',           # 目标目录
        'tmdb_api_key': '',         # TMDB API Key
        'move_mode': 'link',        # 整理模式：link=硬链接，cut=剪切，copy=复制
        'video_extensions': DEFAULT_VIDEO_EXTENSIONS.copy(),  # 支持的视频格式列表
        'auto_extras': True,        # 自动将未匹配视频整理到 extras 文件夹
        'embyignore_extras': True,  # 在 extras 文件夹生成.embyignore 文件
    }

    def __init__(self):
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        self.config = self._load()
        self._pending = False

    def _load(self) -> dict:
        data = dict(self.DEFAULT)
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                    data.update(json.load(f))
            except: pass
        return data

    def save(self):
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=2, ensure_ascii=False)
        self._pending = False

    def get(self, key: str, default=None):
        return self.config.get(key, default)

    def set(self, key: str, value, save_later=True):
        self.config[key] = value
        if save_later:
            self._pending = True
        else:
            self.save()

    @staticmethod
    def check_duplicate_files(paths: List[Path], matched_files: set, parent=None) -> bool:
        """
        检查文件是否已匹配（标绿）
        
        Args:
            paths: 要检查的文件路径列表
            matched_files: 已匹配的文件集合
            parent: 父窗口（用于弹窗）
            
        Returns:
            True 如果有重复，False 如果没有重复
        """
        duplicates = [p for p in paths if p in matched_files]
        if duplicates:
            dup_names = '\n'.join(p.name for p in duplicates[:5])
            if len(duplicates) > 5:
                dup_names += f'\n... 还有 {len(duplicates) - 5} 个文件'
            QMessageBox.warning(parent, "文件已匹配", f"以下文件已经在其他位置匹配（视频列表中标绿），不能重复添加：\n\n{dup_names}")
            return True
        return False

    def save_if_needed(self):
        if self._pending:
            self.save()

    def get_video_extensions(self) -> set:
        """获取视频格式集合（小写）"""
        exts = self.config.get('video_extensions', self.DEFAULT_VIDEO_EXTENSIONS)
        return {ext.lower() for ext in exts}


# ==================== TMDB API ====================
class TMDBClient:
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.base_url = 'https://api.themoviedb.org/3'

    def search_tv(self, query: str) -> List[dict]:
        url = f'{self.base_url}/search/tv'
        params = {'api_key': self.api_key, 'query': query, 'language': 'zh-CN', 'page': 1}
        try:
            resp = requests.get(url, params=params, timeout=15)
            return resp.json().get('results', [])
        except:
            return []

    def get_tv_details(self, tv_id: int) -> Optional[dict]:
        url = f'{self.base_url}/tv/{tv_id}'
        params = {'api_key': self.api_key, 'language': 'zh-CN'}
        try:
            resp = requests.get(url, params=params, timeout=15)
            return resp.json()
        except:
            return None

    def get_season_details(self, tv_id: int, season_num: int) -> Optional[dict]:
        url = f'{self.base_url}/tv/{tv_id}/season/{season_num}'
        params = {'api_key': self.api_key, 'language': 'zh-CN'}
        try:
            resp = requests.get(url, params=params, timeout=15)
            return resp.json()
        except:
            return None


# ==================== 文件操作 ====================
class FileOperator:
    @staticmethod
    def operate(src: Path, dst: Path, mode: str) -> Tuple[bool, str]:
        """
        执行文件操作
        
        Returns:
            (success, error_message) 元组
            - 成功：(True, "")
            - 失败：(False, "错误原因")
        """
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            
            # 如果目标文件已存在，报错（防止误覆盖）
            if dst.exists():
                return False, f"目标文件已存在：{dst.name}"
            
            if mode == 'link':
                os.link(src, dst)
            elif mode == 'cut':
                shutil.move(src, dst)
            elif mode == 'copy':
                shutil.copy2(src, dst)
            else:
                return False, "无效的整理模式"
            
            return True, ""
            
        except Exception as e:
            return False, f"{type(e).__name__}: {str(e)}"


# ==================== 整理完成弹窗 ====================
class OrganizeResultDialog(QDialog):
    """整理完成结果弹窗 - 显示统计、未整理文件（含重名高亮）"""
    
    def __init__(self, success_count, fail_count, unorganized_files, mode_name, parent=None):
        super().__init__(parent)
        self.success_count = success_count
        self.fail_count = fail_count
        self.unorganized_files = unorganized_files
        self.mode_name = mode_name
        self.selected_files = []
        
        self.setWindowTitle("整理完成")
        self.setMinimumSize(800, 600)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        
        # 统计信息
        stats_group = QGroupBox("整理统计")
        stats_layout = QVBoxLayout()
        stats_text = f"✅ 成功：{self.success_count} 个文件\n"
        if self.fail_count > 0:
            stats_text += f"❌ 失败：{self.fail_count} 个文件"
        stats_label = QLabel(stats_text)
        stats_label.setStyleSheet("font-size: 14px; font-weight: bold; padding: 10px;")
        stats_layout.addWidget(stats_label)
        stats_group.setLayout(stats_layout)
        layout.addWidget(stats_group)
        
        # 说明文字和图例
        info_layout = QHBoxLayout()
        info_label = QLabel("以下是源目录中未被整理的文件（包含所有格式），勾选确认后将整理到 extras 文件夹：")
        info_label.setStyleSheet("font-weight: bold; padding: 5px;")
        info_layout.addWidget(info_label)
        legend_label = QLabel("🔴 红色 = 因重名未整理（置顶显示）")
        legend_label.setStyleSheet("color: #c00; font-weight: bold; padding: 5px;")
        info_layout.addWidget(legend_label)
        info_layout.addStretch()
        layout.addLayout(info_layout)
        
        # 文件列表
        self.file_table = QTableWidget()
        self.file_table.setColumnCount(2)
        self.file_table.setHorizontalHeaderLabels(["✓ 全选", "文件路径"])
        self.file_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.file_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.file_table.setEditTriggers(QAbstractItemView.NoEditTriggers)  # 禁用编辑
        self.file_row_map = {}
        
        # 点击表头时切换全选状态
        self.file_table.horizontalHeader().sectionClicked.connect(self.on_header_clicked)

        for file_path, is_duplicate in self.unorganized_files:
            row = self.file_table.rowCount()
            self.file_table.insertRow(row)
            
            # 复选框单元格
            check_item = QTableWidgetItem("")
            check_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            check_item.setCheckState(Qt.Unchecked)
            self.file_table.setItem(row, 0, check_item)
            
            # 文件路径单元格
            path_item = QTableWidgetItem(str(file_path))
            path_item.setFlags(Qt.ItemIsEnabled)
            path_item.setToolTip(str(file_path))
            if is_duplicate:
                path_item.setBackground(QColor("#ffc0c0"))
                path_item.setForeground(QColor("#c00"))
                path_item.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
            self.file_table.setItem(row, 1, path_item)
            self.file_row_map[row] = file_path
        
        layout.addWidget(self.file_table)

        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.on_accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def on_header_clicked(self, logical_index):
        """点击表头时切换全选状态"""
        if logical_index == 0:
            all_checked = True
            for row in range(self.file_table.rowCount()):
                check_item = self.file_table.item(row, 0)
                if not check_item or check_item.checkState() != Qt.Checked:
                    all_checked = False
                    break
            state = Qt.Unchecked if all_checked else Qt.Checked
            for row in range(self.file_table.rowCount()):
                check_item = self.file_table.item(row, 0)
                if check_item:
                    check_item.setCheckState(state)

    def on_accept(self):
        self.selected_files = []
        for row in range(self.file_table.rowCount()):
            check_item = self.file_table.item(row, 0)
            if check_item and check_item.checkState() == Qt.Checked:
                file_path = self.file_row_map.get(row)
                if file_path:
                    self.selected_files.append(file_path)
        self.accept()


# ==================== 工作线程 ====================
class SeasonWorker(QThread):
    finished = pyqtSignal(int, dict)
    def __init__(self, tmdb: TMDBClient, tv_id: int, season_num: int):
        super().__init__()
        self.tmdb = tmdb
        self.tv_id = tv_id
        self.season_num = season_num
    def run(self):
        details = self.tmdb.get_season_details(self.tv_id, self.season_num)
        if details:
            self.finished.emit(self.season_num, details)


# ==================== 搜索选择对话框 ====================
class SearchSelectDialog(QDialog):
    """搜索并选择番剧的对话框"""
    def __init__(self, tmdb: TMDBClient, query: str, parent=None):
        super().__init__(parent)
        self.tmdb = tmdb
        self.selected_tv = None
        self.setWindowTitle(f"搜索结果：{query}")
        self.setMinimumSize(800, 500)
        self.resize(800, 500)

        layout = QVBoxLayout(self)
        layout.setSpacing(6)

        # 结果列表
        self.result_list = QTreeWidget()
        self.result_list.setHeaderLabels(["名称", "年份", "评分"])
        self.result_list.header().setSectionResizeMode(0, QHeaderView.Interactive)  # 名称列可调整
        self.result_list.header().setSectionResizeMode(1, QHeaderView.Interactive)  # 年份列可调整
        self.result_list.header().setSectionResizeMode(2, QHeaderView.Interactive)  # 评分列可调整
        self.result_list.setColumnWidth(0, 600)  # 名称列初始宽度（适配 800 宽弹窗）
        self.result_list.setColumnWidth(1, 100)  # 年份列初始宽度
        self.result_list.setColumnWidth(2, 60)   # 评分列初始宽度
        self.result_list.header().setSectionsMovable(True)  # 允许移动列
        self.result_list.header().setSectionsClickable(True)  # 允许点击排序
        self.result_list.setSortingEnabled(True)  # 启用排序
        self.result_list.setSelectionMode(QTreeWidget.SingleSelection)
        self.result_list.itemDoubleClicked.connect(self.select_item)
        layout.addWidget(self.result_list)

        # 禁用自动排序，保持 API 返回的原始顺序（相关性排序）
        self.result_list.header().setSortIndicator(-1, Qt.AscendingOrder)

        # 按钮
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.select_item)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        # 立即搜索
        self.search(query)

    def search(self, query: str):
        self.result_list.clear()
        results = self.tmdb.search_tv(query)
        for tv in results:
            item = QTreeWidgetItem()
            item.setText(0, tv.get('name', '未知'))
            item.setText(1, tv.get('first_air_date', '')[:4] if tv.get('first_air_date') else '')
            item.setText(2, f"{tv.get('vote_average', 0):.1f}")
            item.setData(0, Qt.UserRole, tv)
            self.result_list.addTopLevelItem(item)

    def select_item(self):
        item = self.result_list.currentItem()
        if item:
            self.selected_tv = item.data(0, Qt.UserRole)
            self.accept()


# ==================== 剧集拖放行 ====================
class MatchItem(QFrame):
    """批量匹配列表中的单项"""
    def __init__(self, ep_num: int, path: Path, index: int, parent_tab=None):
        super().__init__()
        self.ep_num = ep_num  # 动态集号
        self.path = path
        self.index = index
        self.parent_tab = parent_tab
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.Box | QFrame.Plain)
        self.setStyleSheet("""
            QFrame {
                border: 1px solid #ddd;
                border-radius: 3px;
                background-color: #f9f9f9;
            }
            QFrame:hover {
                border-color: #2196F3;
                background-color: #e3f2fd;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(6)

        # 拖放手柄
        handle = QLabel("☰")
        handle.setStyleSheet("color: #999; font-size: 14px; min-width: 20px;")
        handle.setAlignment(Qt.AlignCenter)
        layout.addWidget(handle)

        # 集号（动态更新）
        self.ep_label = QLabel(f"E{ep_num:02d}")
        self.ep_label.setStyleSheet("font-weight: bold; color: #2196F3; min-width: 45px;")
        layout.addWidget(self.ep_label)

        # 文件名
        name_label = QLabel(path.name)
        name_label.setStyleSheet("color: #333;")
        name_label.setWordWrap(True)
        layout.addWidget(name_label, 1)

        # 删除按钮
        del_btn = QPushButton("✕")
        del_btn.setFixedSize(24, 24)
        del_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border-radius: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)
        del_btn.clicked.connect(self.remove_self)
        layout.addWidget(del_btn)

    def update_ep_num(self, new_ep_num: int):
        """更新集号显示"""
        self.ep_num = new_ep_num
        self.ep_label.setText(f"E{new_ep_num:02d}")

    def remove_self(self):
        if self.parent_tab:
            self.parent_tab.remove_match_item(self)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_pos = event.pos()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            if (event.pos() - self.drag_start_pos).manhattanLength() > QApplication.startDragDistance():
                drag = QDrag(self)
                mime_data = QMimeData()
                mime_data.setData('application/x-match-item', str(self.index).encode('utf-8'))
                drag.setMimeData(mime_data)

                # 创建拖放预览
                pixmap = QPixmap(self.size())
                self.render(pixmap)
                drag.setPixmap(pixmap)
                drag.setHotSpot(event.pos())

                if drag.exec_(Qt.MoveAction) == Qt.MoveAction:
                    pass

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasFormat('application/x-match-item'):
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasFormat('application/x-match-item'):
            data = event.mimeData().data('application/x-match-item').data()
            source_index = int(data.decode('utf-8'))
            if source_index != self.index and self.parent_tab:
                self.parent_tab.reorder_match_item(source_index, self.index)


class BatchDropArea(QFrame):
    """批量拖放区域"""
    def __init__(self, parent_tab=None, drop_type="add"):
        """
        drop_type: "add" - 新增文件（追加到列表）
                   "sort" - 覆盖/排序（替换列表或拖放排序）
        """
        super().__init__()
        self.parent_tab = parent_tab
        self.drop_type = drop_type
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.Box | QFrame.Plain)
        self.setMinimumHeight(50)

        if drop_type == "add":
            self.setStyleSheet("""
                QFrame {
                    border: 2px dashed #4CAF50;
                    border-radius: 5px;
                    background-color: #f1f8e9;
                }
                QFrame:hover {
                    border-color: #2E7D32;
                    background-color: #dcedc8;
                }
            """)
        else:  # sort
            self.setStyleSheet("""
                QFrame {
                    border: 2px dashed #2196F3;
                    border-radius: 5px;
                    background-color: #e3f2fd;
                }
                QFrame:hover {
                    border-color: #1565C0;
                    background-color: #bbdefb;
                }
            """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasFormat('application/x-video-files'):
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if self.parent_tab:
            self.parent_tab.handle_batch_drop(event, self.drop_type)


# ==================== 单集拖放行 ====================
class EpisodeRow(QFrame):
    dropped = pyqtSignal(int, int, list, list)  # season_num, episode_num, new_paths, old_paths
    cancel_match = pyqtSignal(int, int)  # season_num, episode_num

    def __init__(self, season_num: int, episode_num: int, episode_name: str, air_date: str = '', runtime: int = None, parent_window=None):
        super().__init__()
        self.season_num = season_num
        self.episode_num = episode_num
        self.parent_window = parent_window
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.Box | QFrame.Plain)
        self.setStyleSheet("""
            QFrame {
                border: 1px solid #ddd;
                border-radius: 3px;
                background-color: white;
            }
            QFrame:hover {
                border-color: #2196F3;
                background-color: #e3f2fd;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(4)

        # 第一行：集号 + 剧集名 + 日期 + 时长 + 取消按钮
        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        ep_label = QLabel(f"E{episode_num:02d}")
        ep_label.setStyleSheet("font-weight: bold; color: #2196F3; min-width: 45px;")
        top_row.addWidget(ep_label)

        name_label = QLabel(episode_name)
        name_label.setStyleSheet("font-weight: bold;")
        name_label.setWordWrap(True)
        top_row.addWidget(name_label, 1)

        # 时长
        if runtime and runtime > 0:
            runtime_label = QLabel(f"⏱ {runtime} 分钟")
            runtime_label.setStyleSheet("color: #555; font-size: 13px; font-weight: bold;")
            top_row.addWidget(runtime_label)

        # 日期
        if air_date:
            date_label = QLabel(f"📅 {air_date}")
            date_label.setStyleSheet("color: #555; font-size: 13px; font-weight: bold;")
            top_row.addWidget(date_label)

        # 取消匹配按钮（默认隐藏）
        self.cancel_btn = QPushButton("✕ 取消")
        self.cancel_btn.setFixedSize(60, 24)
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #9e9e9e;
                color: white;
                border-radius: 12px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #757575;
            }
        """)
        self.cancel_btn.setVisible(False)
        self.cancel_btn.clicked.connect(self.on_cancel_clicked)
        top_row.addWidget(self.cancel_btn)

        top_row.addStretch()
        layout.addLayout(top_row)

        # 分割线
        line = QFrame()
        line.setFrameStyle(QFrame.HLine | QFrame.Sunken)
        line.setStyleSheet("color: #eee;")
        layout.addWidget(line)

        # 第三行：拖放状态
        self.status_label = QLabel("← 拖放视频文件到此处")
        self.status_label.setStyleSheet("color: #888; font-size: 12px; padding: 4px; background-color: #f9f9f9; border-radius: 3px;")
        self.status_label.setWordWrap(True)
        self.status_label.setMinimumHeight(30)
        layout.addWidget(self.status_label)

        self.dropped_files = []

    def on_cancel_clicked(self):
        """取消匹配"""
        self.cancel_match.emit(self.season_num, self.episode_num)

    def set_matched(self, files: List[Path]):
        """设置为已匹配状态"""
        self.dropped_files = list(files)
        if len(files) == 1:
            self.status_label.setText(f"✓ 已匹配：{files[0].name}")
        else:
            self.status_label.setText(f"✓ 已匹配 {len(files)} 个文件:\n" + "\n".join(p.name for p in files))

        self.status_label.setStyleSheet("color: #2E7D32; font-weight: bold; font-size: 13px; padding: 4px; background-color: #d4edda; border-radius: 3px;")
        self.setStyleSheet("""
            QFrame {
                border: 1px solid #4CAF50;
                border-radius: 3px;
                background-color: #d4edda;
            }
        """)
        self.cancel_btn.setVisible(True)

    def reset(self):
        """重置为未匹配状态"""
        self.dropped_files = []
        self.status_label.setText("← 拖放视频文件到此处")
        self.status_label.setStyleSheet("color: #888; font-size: 12px; padding: 4px; background-color: #f9f9f9; border-radius: 3px;")
        self.setStyleSheet("""
            QFrame {
                border: 1px solid #ddd;
                border-radius: 3px;
                background-color: white;
            }
            QFrame:hover {
                border-color: #2196F3;
                background-color: #e3f2fd;
            }
        """)
        self.cancel_btn.setVisible(False)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasFormat('application/x-video-files'):
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        data = event.mimeData().data('application/x-video-files').data()
        paths = [Path(p.decode('utf-8')) for p in data.split(b'\n') if p]

        if not paths:
            return

        # 单集模式只匹配 1 个文件
        if len(paths) > 1:
            QMessageBox.warning(None, "警告", "单集模式只能匹配 1 个文件，请逐个拖放")
            return

        # 先检测文件是否已匹配（标绿）- 防止重复添加
        if self.parent_window:
            matched_files = self.parent_window.get_matched_files()
            # 排除当前行已匹配的文件（允许重新拖放同一文件）
            if paths[0] in matched_files and paths[0] not in self.dropped_files:
                Config.check_duplicate_files([paths[0]], matched_files, self)
                return

        # 保存旧文件路径（用于移除）
        old_files = list(self.dropped_files)

        # 更新 UI
        self.dropped_files = [paths[0]]
        self.set_matched(self.dropped_files)

        # 发射信号（传递新旧文件路径）
        self.dropped.emit(self.season_num, self.episode_num, [paths[0]], old_files)


class VideoTreeItem(QTreeWidgetItem):
    """支持数值排序的视频列表项"""
    def __lt__(self, other):
        """自定义比较：按 UserRole 数值排序"""
        column = self.treeWidget().header().sortIndicatorSection()
        
        # 获取 UserRole 数据
        self_data = self.data(column, Qt.UserRole)
        other_data = other.data(column, Qt.UserRole)
        
        # 如果都有数值数据，按数值比较
        if self_data is not None and other_data is not None:
            if isinstance(self_data, (int, float)) and isinstance(other_data, (int, float)):
                return self_data < other_data
        
        # 否则按文本比较
        return self.text(column).lower() < other.text(column).lower()


class FolderTreeItem(QTreeWidgetItem):
    """支持数值排序的文件夹列表项"""
    def __lt__(self, other):
        """自定义比较：按 UserRole 数值排序"""
        column = self.treeWidget().header().sortIndicatorSection()
        
        # 获取 UserRole 数据
        self_data = self.data(column, Qt.UserRole)
        other_data = other.data(column, Qt.UserRole)
        
        # 如果都有数值数据，按数值比较
        if self_data is not None and other_data is not None:
            if isinstance(self_data, (int, float)) and isinstance(other_data, (int, float)):
                return self_data < other_data
        
        # 否则按文本比较
        return self.text(column).lower() < other.text(column).lower()


class FolderTreeWidget(QTreeWidget):
    """支持排序的文件夹列表"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSortingEnabled(True)
        self.header().setSectionsClickable(True)


class VideoTreeWidget(QTreeWidget):
    """支持拖放的视频列表"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setDragDropMode(QTreeWidget.DragOnly)
        # 启用排序
        self.setSortingEnabled(True)
        # 允许点击表头排序
        self.header().setSectionsClickable(True)

    def startDrag(self, supportedActions):
        # 如果正在调整列宽或移动列，不启动拖放
        if self.header().sectionsMovable():
            pass  # 允许列移动
        selected = self.selectedItems()
        if not selected:
            return

        paths = []
        for item in selected:
            path = item.data(0, Qt.UserRole)
            if path:
                paths.append(str(path))

        if paths:
            mime_data = QMimeData()
            mime_data.setData('application/x-video-files', b'\n'.join(p.encode('utf-8') for p in paths))

            drag = QDrag(self)
            drag.setMimeData(mime_data)
            pixmap = QPixmap(200, 50)
            pixmap.fill(Qt.lightGray)
            drag.setPixmap(pixmap)
            drag.exec_(Qt.CopyAction)


# ==================== extras 标签页 ====================
class ExtrasTab(QWidget):
    """extras 标签页 - 用于放置未匹配/特典文件"""
    def __init__(self, parent=None):
        super().__init__()
        self.parent_window = parent
        self.file_mappings = {}  # {path: "extras"}
        self.setAcceptDrops(True)  # 启用拖放
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)
        
        # 提示说明
        tip = QLabel("💡 拖放视频文件到此处，整理时会移动到 extras 文件夹（保留原文件名）")
        tip.setStyleSheet("color: #666; font-size: 12px; padding: 4px; background-color: #f0f8ff; border-radius: 4px;")
        tip.setWordWrap(True)
        layout.addWidget(tip)
        
        # 文件列表
        self.file_list = QTreeWidget()
        self.file_list.setHeaderLabels(["文件名", "大小", "操作"])
        self.file_list.header().setStretchLastSection(False)
        self.file_list.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.file_list.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.file_list.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.file_list.setMinimumHeight(150)
        self.file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)  # 支持多选（Ctrl/Shift）
        self.file_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_context_menu)
        layout.addWidget(self.file_list, 1)  # stretch=1 占满剩余空间
        
        # 底部按钮行
        btn_layout = QHBoxLayout()
        
        clear_btn = QPushButton("🗑️ 清空所有")
        clear_btn.setFixedHeight(30)
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)
        clear_btn.clicked.connect(self.clear_all)
        btn_layout.addWidget(clear_btn)
        
        remove_btn = QPushButton("🗑️ 移除选中")
        remove_btn.setFixedHeight(30)
        remove_btn.setStyleSheet("""
            QPushButton {
                background-color: #ff9800;
                color: white;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #f57c00;
            }
        """)
        remove_btn.clicked.connect(self.remove_selected)
        btn_layout.addWidget(remove_btn)
        
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
    
    def show_context_menu(self, pos):
        """显示右键菜单"""
        item = self.file_list.itemAt(pos)
        if item:
            menu = self.file_list.createStandardContextMenu()
            
            # 如果选中了多个，添加批量移除选项
            selected_count = len(self.file_list.selectedItems())
            if selected_count > 1:
                remove_action = menu.addAction(f"🗑️ 移除选中的 {selected_count} 个文件")
            else:
                remove_action = menu.addAction("🗑️ 移除")
            
            action = menu.exec_(self.file_list.mapToGlobal(pos))
            if action == remove_action:
                self.remove_selected()
    
    def add_files(self, paths: List[Path]):
        """添加文件到 extras - 检测标绿文件（已匹配）并报错"""
        # 检测是否有文件已匹配（标绿）
        if self.parent_window:
            matched_files = self.parent_window.get_matched_files()
            if Config.check_duplicate_files(paths, matched_files, self):
                return
        
        # 添加文件
        added = 0
        for path in paths:
            # 验证文件存在性
            if not path.exists():
                continue
            if path not in self.file_mappings:
                self.file_mappings[path] = "extras"
                item = QTreeWidgetItem()
                item.setText(0, path.name)
                try:
                    size_mb = path.stat().st_size / 1024 / 1024
                    item.setText(1, f"{size_mb:.1f} MB")
                except:
                    item.setText(1, "未知")
                item.setData(0, Qt.UserRole, path)
                self.file_list.addTopLevelItem(item)

                # 添加移除按钮
                btn = QPushButton("❌")
                btn.setFixedSize(24, 24)
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #f44336;
                        color: white;
                        border-radius: 12px;
                        font-weight: bold;
                        font-size: 14px;
                    }
                    QPushButton:hover {
                        background-color: #d32f2f;
                    }
                """)
                btn.clicked.connect(lambda checked, p=path: self.remove_file(p))
                self.file_list.setItemWidget(item, 2, btn)
                added += 1

        if added > 0 and self.parent_window:
            self.parent_window._update_status()

    def remove_file(self, path: Path):
        """移除指定文件"""
        if path in self.file_mappings:
            del self.file_mappings[path]
        # 查找并移除 item
        for i in range(self.file_list.topLevelItemCount()):
            item = self.file_list.topLevelItem(i)
            if item.data(0, Qt.UserRole) == path:
                self.file_list.takeTopLevelItem(i)
                break
        if self.parent_window:
            self.parent_window._update_status()

    def remove_item(self, item):
        """移除单个文件"""
        path = item.data(0, Qt.UserRole)
        if path and path in self.file_mappings:
            del self.file_mappings[path]
        self.file_list.takeTopLevelItem(self.file_list.indexOfTopLevelItem(item))
        if self.parent_window:
            self.parent_window._update_status()

    def remove_selected(self):
        """移除选中的文件"""
        for item in self.file_list.selectedItems():
            self.remove_item(item)

    def clear_all(self):
        """清空所有 extras 文件"""
        self.file_mappings.clear()
        self.file_list.clear()
        if self.parent_window:
            self.parent_window._update_status()
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖放进入"""
        if event.mimeData().hasUrls() or event.mimeData().hasFormat('application/x-video-files'):
            event.acceptProposedAction()
            # 视觉反馈：改变背景色
            self.file_list.setStyleSheet("QTreeWidget { border: 2px solid #2196F3; background-color: #e3f2fd; }")

    def dragLeaveEvent(self, event):
        """拖放离开"""
        # 恢复原始样式
        self.file_list.setStyleSheet("")

    def dropEvent(self, event: QDropEvent):
        """拖放完成"""
        # 恢复原始样式
        self.file_list.setStyleSheet("")
        
        try:
            paths = []
            if event.mimeData().hasUrls():
                for url in event.mimeData().urls():
                    path = Path(url.toLocalFile())
                    if self.parent_window:
                        if path.suffix.lower() in self.parent_window._video_extensions:
                            paths.append(path)
            elif event.mimeData().hasFormat('application/x-video-files'):
                data = event.mimeData().data('application/x-video-files')
                for line in bytes(data).split(b'\n'):
                    if line:
                        paths.append(Path(line.decode('utf-8')))

            if paths:
                self.add_files(paths)
        except Exception as e:
            import traceback
            print(f"extras drop error: {e}")
            print(traceback.format_exc())


# ==================== 季度页面 ====================
class SeasonTab(QWidget):
    def __init__(self, season_num: int, season_name: str, episode_count: int, tmdb: TMDBClient, tv_id: int, parent_window=None):
        super().__init__()
        self.season_num = season_num
        self.tmdb = tmdb
        self.tv_id = tv_id
        self.parent_window = parent_window
        self.batch_paths = []  # 批量匹配的文件路径列表（按顺序）
        self.file_mappings = {}  # 最终输出：{path: key}
        self._workers = []  # 保持线程引用
        self.match_mode = "batch"  # "batch" 或 "single"

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # ========== 顶部：模式切换 ==========
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("匹配模式:"))
        
        self.mode_batch_btn = QPushButton("📥 批量匹配")
        self.mode_batch_btn.setCheckable(True)
        self.mode_batch_btn.setChecked(True)
        self.mode_batch_btn.setFixedHeight(30)
        self.mode_batch_btn.setStyleSheet("""
            QPushButton:checked {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:unchecked {
                background-color: #e0e0e0;
                color: #333;
                border-radius: 4px;
            }
        """)
        self.mode_batch_btn.clicked.connect(lambda: self.switch_mode("batch"))
        mode_layout.addWidget(self.mode_batch_btn)
        
        self.mode_single_btn = QPushButton("🎬 单集匹配")
        self.mode_single_btn.setCheckable(True)
        self.mode_single_btn.setFixedHeight(30)
        self.mode_single_btn.setStyleSheet("""
            QPushButton:checked {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:unchecked {
                background-color: #e0e0e0;
                color: #333;
                border-radius: 4px;
            }
        """)
        self.mode_single_btn.clicked.connect(lambda: self.switch_mode("single"))
        mode_layout.addWidget(self.mode_single_btn)
        
        mode_layout.addStretch()
        
        self.clear_all_btn = QPushButton("🗑️ 清空所有")
        self.clear_all_btn.setFixedHeight(30)
        self.clear_all_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-weight: bold;
                border-radius: 4px;
                padding: 4px 12px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)
        self.clear_all_btn.setVisible(False)
        self.clear_all_btn.clicked.connect(self.clear_all_matches)
        mode_layout.addWidget(self.clear_all_btn)
        
        layout.addLayout(mode_layout)

        # ========== 批量匹配区 ==========
        self.batch_widget = QWidget()
        batch_layout = QVBoxLayout(self.batch_widget)
        batch_layout.setContentsMargins(0, 0, 0, 0)
        batch_layout.setSpacing(4)

        # 左右分栏：左边新增，右边覆盖
        drop_area_layout = QHBoxLayout()
        drop_area_layout.setSpacing(8)

        # 左侧：新增文件
        add_group = QGroupBox("➕ 新增")
        add_layout = QVBoxLayout(add_group)
        
        self.batch_drop_add = BatchDropArea(self, drop_type="add")
        self.batch_drop_add.setMinimumHeight(80)
        add_layout.addWidget(self.batch_drop_add)
        
        add_hint = QLabel("拖放到此处\n按文件名排序追加到列表末尾")
        add_hint.setStyleSheet("color: #888; font-size: 11px;")
        add_hint.setAlignment(Qt.AlignCenter)
        add_layout.addWidget(add_hint)
        
        drop_area_layout.addWidget(add_group, 1)

        # 右侧：覆盖/排序
        sort_group = QGroupBox("📋 覆盖/排序")
        sort_layout = QVBoxLayout(sort_group)
        
        self.batch_drop_sort = BatchDropArea(self, drop_type="sort")
        self.batch_drop_sort.setMinimumHeight(80)
        sort_layout.addWidget(self.batch_drop_sort)
        
        sort_hint = QLabel("拖放到此处\n按文件名排序覆盖当前列表")
        sort_hint.setStyleSheet("color: #888; font-size: 11px;")
        sort_hint.setAlignment(Qt.AlignCenter)
        sort_layout.addWidget(sort_hint)
        
        drop_area_layout.addWidget(sort_group, 1)
        batch_layout.addLayout(drop_area_layout)

        # 匹配列表（带滚动 - 占据剩余空间）
        match_scroll = QScrollArea()
        match_scroll.setWidgetResizable(True)
        match_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        match_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self.match_list_widget = QWidget()
        self.match_list_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.match_list_layout = QVBoxLayout(self.match_list_widget)
        self.match_list_layout.setContentsMargins(4, 4, 4, 4)
        self.match_list_layout.setSpacing(2)
        self.match_list_widget.setVisible(False)
        match_scroll.setWidget(self.match_list_widget)
        batch_layout.addWidget(match_scroll, 1)  # 拉伸因子为 1，占据剩余空间

        # 批量操作按钮
        batch_btn_layout = QHBoxLayout()
        self.batch_status = QLabel("")
        self.batch_status.setStyleSheet("color: #4CAF50; font-weight: bold; font-size: 12px;")
        batch_btn_layout.addWidget(self.batch_status)
        batch_btn_layout.addStretch()
        batch_layout.addLayout(batch_btn_layout)
        layout.addWidget(self.batch_widget, 1)  # 拉伸因子为 1

        # ========== 单集匹配区 ==========
        self.episode_widget = QWidget()
        episode_layout = QVBoxLayout(self.episode_widget)
        episode_layout.setContentsMargins(0, 0, 0, 0)
        
        # 剧集滚动区
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        self.episode_container = QWidget()
        self.episode_layout = QVBoxLayout(self.episode_container)
        self.episode_layout.setAlignment(Qt.AlignTop)
        self.episode_layout.setSpacing(4)
        scroll.setWidget(self.episode_container)
        
        episode_layout.addWidget(scroll)
        layout.addWidget(self.episode_widget)

        # 初始显示：Season 0 默认单集模式，其他季度默认批量模式
        if season_num == 0:
            self.match_mode = "single"
            self.batch_widget.setVisible(False)
            self.episode_widget.setVisible(True)
            self.mode_batch_btn.setChecked(False)
            self.mode_single_btn.setChecked(True)
        else:
            self.match_mode = "batch"
            self.batch_widget.setVisible(True)
            self.episode_widget.setVisible(False)
            self.mode_batch_btn.setChecked(True)
            self.mode_single_btn.setChecked(False)

        # 加载剧集（预加载，但不显示）
        self._load_episodes()

    def switch_mode(self, mode: str):
        """切换匹配模式"""
        try:
            self.match_mode = mode

            if mode == "batch":
                # 切换到批量模式：清空单集匹配数据
                for i in range(self.episode_layout.count()):
                    item = self.episode_layout.itemAt(i)
                    if item and isinstance(item.widget(), EpisodeRow):
                        row = item.widget()
                        # 从 file_mappings 移除单集匹配的文件
                        for path in list(row.dropped_files):
                            if path in self.file_mappings:
                                del self.file_mappings[path]
                        row.reset()
                
                self.batch_widget.setVisible(True)
                self.episode_widget.setVisible(False)
                self.mode_batch_btn.setChecked(True)
                self.mode_single_btn.setChecked(False)
                # 批量模式：启用批量拖放，禁用单集行拖放
                self.batch_drop_add.setEnabled(True)
                self.batch_drop_sort.setEnabled(True)
                for i in range(self.episode_layout.count()):
                    item = self.episode_layout.itemAt(i)
                    if item:
                        row = item.widget()
                        if isinstance(row, EpisodeRow):
                            row.setAcceptDrops(False)
            else:
                # 切换到单集模式：清空批量匹配数据
                # 从 file_mappings 移除批量匹配的文件
                for path in list(self.batch_paths):
                    if path in self.file_mappings:
                        del self.file_mappings[path]
                
                self.batch_paths = []
                while self.match_list_layout.count():
                    item = self.match_list_layout.takeAt(0)
                    if item.widget():
                        item.widget().deleteLater()
                self.match_list_widget.setVisible(False)
                self.batch_status.setText("")
                
                self.batch_widget.setVisible(False)
                self.episode_widget.setVisible(True)
                self.mode_batch_btn.setChecked(False)
                self.mode_single_btn.setChecked(True)
                # 单集模式：禁用批量拖放，启用单集行拖放
                self.batch_drop_add.setEnabled(False)
                self.batch_drop_sort.setEnabled(False)
                for i in range(self.episode_layout.count()):
                    item = self.episode_layout.itemAt(i)
                    if item:
                        row = item.widget()
                        if isinstance(row, EpisodeRow):
                            row.setAcceptDrops(True)

            self._update_buttons()
            # 通知主窗口刷新高亮
            if self.parent_window:
                self.parent_window._update_status()
        except Exception as e:
            import traceback
            traceback.print_exc()

    def _load_episodes(self):
        worker = SeasonWorker(self.tmdb, self.tv_id, self.season_num)
        self._workers.append(worker)  # 保持线程引用防止被 GC 回收
        worker.finished.connect(self._on_episodes_loaded)
        worker.finished.connect(lambda: self._workers.remove(worker))  # 完成后移除引用，防止内存泄漏
        worker.start()

    def _on_episodes_loaded(self, season_num: int, details: dict):
        episodes = details.get('episodes', [])
        for ep in episodes:
            ep_num = ep.get('episode_number', 0)
            ep_name = ep.get('name', f'第{ep_num}集')
            air_date = ep.get('air_date', '')
            runtime = ep.get('runtime')  # 获取单集时长（分钟）
            row = EpisodeRow(season_num, ep_num, ep_name, air_date, runtime, self.parent_window)
            row.dropped.connect(self._on_episode_dropped)
            row.cancel_match.connect(self._on_cancel_match)
            self.episode_layout.addWidget(row)

    def _on_episode_dropped(self, season_num: int, episode_num: int, new_paths: List[Path], old_paths: List[Path]):
        """单集拖放处理 - 只在单集模式下生效"""
        if self.match_mode != "single":
            return  # 非单集模式忽略

        # 移除旧文件
        for old_path in old_paths:
            if old_path in self.file_mappings:
                del self.file_mappings[old_path]

        # 添加新文件
        for path in new_paths:
            key = f"S{season_num:02d}E{episode_num:02d}"
            self.file_mappings[path] = key

        self._update_buttons()
        if self.parent_window:
            self.parent_window._update_status(new_paths)

    def _on_cancel_match(self, season_num: int, episode_num: int):
        """取消单集匹配"""
        # 找到该集的行，获取文件路径
        for i in range(self.episode_layout.count()):
            item = self.episode_layout.itemAt(i)
            if item and isinstance(item.widget(), EpisodeRow):
                row = item.widget()
                if row.season_num == season_num and row.episode_num == episode_num:
                    # 从 file_mappings 移除
                    old_paths = list(row.dropped_files)
                    for path in old_paths:
                        if path in self.file_mappings:
                            del self.file_mappings[path]
                    row.reset()
                    break

        # 如果单集取消了，批量列表里有这个文件，要重新计算集号
        self._sync_batch_from_mappings()
        self._update_buttons()
        if self.parent_window:
            self.parent_window._update_status(old_paths)

    def handle_batch_drop(self, event, drop_type: str = "add"):
        """处理批量拖放 - 检测标绿文件（已匹配）并报错"""
        if self.match_mode != "batch":
            return  # 非批量模式忽略

        data = event.mimeData().data('application/x-video-files').data()
        paths = [Path(p.decode('utf-8')) for p in data.split(b'\n') if p]
        
        # 检测是否有文件已匹配（标绿）
        if self.parent_window:
            matched_files = self.parent_window.get_matched_files()
            if Config.check_duplicate_files(paths, matched_files, self):
                return
        
        sorted_paths = sorted(paths, key=lambda p: p.name)

        if drop_type == "add":
            # 新增模式：添加新文件，然后整体按文件名排序
            existing_paths = list(self.batch_paths)  # 复制现有列表
            for path in sorted_paths:
                if path not in existing_paths:  # 去重
                    existing_paths.append(path)
            # 整体按文件名排序
            self.batch_paths = sorted(existing_paths, key=lambda p: p.name)
        else:  # sort
            # 覆盖模式：替换当前列表，按文件名排序
            # 先清除旧的 file_mappings
            for old_path in self.batch_paths:
                if old_path in self.file_mappings:
                    del self.file_mappings[old_path]
            self.batch_paths = sorted_paths

        # 更新 file_mappings
        for idx, path in enumerate(self.batch_paths):
            ep_num = idx + 1
            key = f"S{self.season_num:02d}E{ep_num:02d}"
            self.file_mappings[path] = key

        # 更新批量列表显示
        self._refresh_match_list()
        self._update_buttons()

        # 通知主窗口刷新高亮（实时高亮）
        if self.parent_window:
            self.parent_window._update_status(sorted_paths)

    def _sync_batch_from_mappings(self):
        """根据 file_mappings 同步批量列表（处理单集取消后的情况）"""
        # 只在批量模式下才显示批量列表
        if self.match_mode != "batch":
            return

        # 获取所有已匹配的文件
        matched_paths = set(self.file_mappings.keys())

        # 过滤批量列表，只保留还在 file_mappings 中的文件
        self.batch_paths = [p for p in self.batch_paths if p in matched_paths]

        # 如果批量列表为空，隐藏
        if not self.batch_paths:
            self.match_list_widget.setVisible(False)
            self.batch_status.setText("")
        else:
            self._refresh_match_list()

    def _clear_episode_status(self):
        """清空所有单集行的状态显示"""
        for i in range(self.episode_layout.count()):
            item = self.episode_layout.itemAt(i)
            if item and item.widget():
                row = item.widget()
                if isinstance(row, EpisodeRow):
                    row.reset()

    def _update_buttons(self):
        """更新按钮显示状态"""
        has_any = bool(self.file_mappings)
        self.clear_all_btn.setVisible(has_any)

    def clear_all_matches(self):
        """清空所有匹配"""
        # 清空 file_mappings
        self.file_mappings.clear()

        # 清空批量列表
        self.batch_paths = []
        while self.match_list_layout.count():
            item = self.match_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.match_list_widget.setVisible(False)
        self.batch_status.setText("")

        # 清空单集行
        self._clear_episode_status()
        self.clear_all_btn.setVisible(False)

        if self.parent_window:
            self.parent_window._update_status()

    def _refresh_match_list(self):
        """刷新匹配列表显示"""
        # 清空
        while self.match_list_layout.count():
            item = self.match_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.batch_paths:
            self.match_list_widget.setVisible(False)
            self._update_buttons()
            return

        # 添加匹配项（集号根据索引动态计算）
        for idx, path in enumerate(self.batch_paths):
            ep_num = idx + 1
            item = MatchItem(ep_num, path, idx, self)
            self.match_list_layout.addWidget(item)

        self.match_list_widget.setVisible(True)
        self.batch_status.setText(f"✓ 已匹配 {len(self.batch_paths)} 个文件 - 可拖动右侧调整顺序")
        self._update_buttons()

    def remove_match_item(self, item):
        """删除匹配项 - 从 file_mappings 移除"""
        idx = item.index
        if 0 <= idx < len(self.batch_paths):
            path = self.batch_paths[idx]
            if path in self.file_mappings:
                del self.file_mappings[path]
            self.batch_paths.pop(idx)
            self._refresh_match_list()
            self._update_buttons()
            if self.parent_window:
                self.parent_window._update_status()

    def reorder_match_item(self, from_index: int, to_index: int):
        """重新排序匹配项 - 拖动后重新分配集号并更新 file_mappings"""
        if 0 <= from_index < len(self.batch_paths) and 0 <= to_index < len(self.batch_paths):
            # 移动项目
            moved_path = self.batch_paths.pop(from_index)
            self.batch_paths.insert(to_index, moved_path)
            
            # 重新分配集号并更新 file_mappings
            for idx, path in enumerate(self.batch_paths):
                ep_num = idx + 1
                key = f"S{self.season_num:02d}E{ep_num:02d}"
                self.file_mappings[path] = key
            
            self._refresh_match_list()
            self._update_buttons()
            if self.parent_window:
                self.parent_window._update_status()


# ==================== 工具函数 ====================
def apply_highlight(item: QTreeWidgetItem, is_matched: bool):
    """统一高亮样式"""
    if is_matched:
        item.setBackground(0, QColor(144, 238, 144))  # 浅绿色背景
        item.setForeground(0, QColor(0, 100, 0))  # 深绿色文字
        item.setForeground(1, QColor(0, 100, 0))
        item.setForeground(2, QColor(0, 100, 0))
    else:
        item.setBackground(0, QBrush())
        item.setForeground(0, QBrush())
        item.setForeground(1, QBrush())
        item.setForeground(2, QBrush())


# ==================== 主窗口 ====================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = Config()
        self.tmdb: Optional[TMDBClient] = None
        self.tv_info: Optional[dict] = None
        self.current_folder: Optional[Path] = None  # 当前加载的文件夹路径
        self.file_mappings: Dict[Path, str] = {}
        self.season_tabs: Dict[int, SeasonTab] = {}
        self._folder_video_cache: Dict[Path, List[Path]] = {}  # 文件夹视频缓存
        self._first_show = True  # 首次启动标记

        self.setWindowTitle("AniOr - 动漫视频整理工具")
        self.setMinimumSize(1200, 800)
        self.resize(1400, 1000)  # 默认窗口大小

        # 保持线程引用
        self._workers = []

        # 视频格式缓存
        self._video_extensions = self.config.get_video_extensions()

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self._create_toolbar(main_layout)
        self._create_main_splitter(main_layout)

        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("就绪 - 请先配置源目录、目标目录和 TMDB API Key")

    def _create_toolbar(self, layout):
        toolbar = QFrame()
        toolbar.setFrameStyle(QFrame.StyledPanel)
        toolbar.setMaximumHeight(44)
        toolbar.setStyleSheet("QFrame { background-color: #f8f8f8; border-bottom: 1px solid #ddd; }")
        tl = QHBoxLayout(toolbar)
        tl.setContentsMargins(8, 4, 8, 4)

        refresh_btn = QPushButton("🔄 刷新文件夹")
        refresh_btn.clicked.connect(self.load_anime_folders)
        tl.addWidget(refresh_btn)
        
        self.refresh_video_btn = QPushButton("🔄 刷新该文件夹视频列表")
        self.refresh_video_btn.clicked.connect(self.on_refresh_video_clicked)
        self.refresh_video_btn.setEnabled(False)
        tl.addWidget(self.refresh_video_btn)
        
        tl.addStretch()

        config_btn = QPushButton("⚙️ 设置")
        config_btn.clicked.connect(self.open_config)
        tl.addWidget(config_btn)

        layout.addWidget(toolbar)

    def _create_main_splitter(self, layout):
        splitter = QSplitter(Qt.Horizontal)

        # === 左侧：文件夹 + 视频 ===
        left_splitter = QSplitter(Qt.Vertical)

        # 文件夹
        folder_group = QGroupBox("📁 动漫文件夹")
        folder_layout = QVBoxLayout(folder_group)

        self.folder_tree = FolderTreeWidget()
        self.folder_tree.setHeaderLabels(["文件夹", "文件数", "日期"])
        self.folder_tree.header().setStretchLastSection(False)
        self.folder_tree.header().setSectionResizeMode(0, QHeaderView.Interactive)
        self.folder_tree.header().setSectionResizeMode(1, QHeaderView.Interactive)
        self.folder_tree.header().setSectionResizeMode(2, QHeaderView.Interactive)
        self.folder_tree.setColumnWidth(0, 600)  # 文件夹
        self.folder_tree.setColumnWidth(1, 90)   # 文件数
        self.folder_tree.setColumnWidth(2, 150)  # 日期
        self.folder_tree.header().setMinimumSectionSize(50)
        self.folder_tree.header().setSectionsMovable(True)  # 允许移动列
        self.folder_tree.header().setSectionsClickable(True)  # 允许点击排序
        self.folder_tree.setSortingEnabled(True)  # 允许点击排序
        # 默认按文件夹名称升序排序
        self.folder_tree.header().setSortIndicator(0, Qt.AscendingOrder)
        self.folder_tree.setStyleSheet("""
            QTreeWidget {
                border: 1px solid #ccc;
                border-radius: 3px;
            }
            QTreeWidget::item {
                padding: 4px 2px;
            }
            QTreeWidget::item:hover { background-color: #e8f4fc; }
            QTreeWidget::item:selected { background-color: #2196F3; color: white; }
        """)
        self.folder_tree.itemClicked.connect(self.on_folder_selected)
        folder_layout.addWidget(self.folder_tree)
        left_splitter.addWidget(folder_group)

        # 视频
        video_group = QGroupBox("🎬 视频文件")
        video_layout = QVBoxLayout(video_group)
        
        self.video_list = VideoTreeWidget()
        self.video_list.setHeaderLabels(["文件名", "大小", "日期"])
        self.video_list.header().setStretchLastSection(False)
        self.video_list.header().setSectionResizeMode(0, QHeaderView.Interactive)
        self.video_list.header().setSectionResizeMode(1, QHeaderView.Interactive)
        self.video_list.header().setSectionResizeMode(2, QHeaderView.Interactive)
        self.video_list.setColumnWidth(0, 560)  # 文件名
        self.video_list.setColumnWidth(1, 100)  # 大小
        self.video_list.setColumnWidth(2, 180)  # 日期
        self.video_list.header().setMinimumSectionSize(50)
        self.video_list.header().setSectionsMovable(True)  # 允许移动列
        self.video_list.header().setSectionsClickable(True)  # 允许点击排序
        self.video_list.setSortingEnabled(True)  # 启用排序
        # 默认按文件名升序排序
        self.video_list.header().setSortIndicator(0, Qt.AscendingOrder)
        self.video_list.setSelectionMode(QTreeWidget.ExtendedSelection)
        self.video_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.video_list.customContextMenuRequested.connect(self._on_video_context_menu)
        self.video_list.setStyleSheet("""
            QTreeWidget {
                border: 1px solid #ccc;
                border-radius: 3px;
            }
            QTreeWidget::item {
                padding: 4px 2px;
            }
            QTreeWidget::item:hover { background-color: #e8f4fc; }
            QTreeWidget::item:selected { background-color: #2196F3; color: white; }
        """)
        self.video_list.itemSelectionChanged.connect(self._on_video_selection_changed)
        video_layout.addWidget(self.video_list)

        btn_row = QHBoxLayout()
        btn_all = QPushButton("全选")
        btn_all.clicked.connect(lambda: [self.video_list.topLevelItem(i).setSelected(True) for i in range(self.video_list.topLevelItemCount())])
        btn_row.addWidget(btn_all)
        btn_none = QPushButton("不选")
        btn_none.clicked.connect(lambda: [self.video_list.topLevelItem(i).setSelected(False) for i in range(self.video_list.topLevelItemCount())])
        btn_row.addWidget(btn_none)
        btn_row.addStretch()
        video_layout.addLayout(btn_row)
        left_splitter.addWidget(video_group)
        left_splitter.setStretchFactor(0, 1)
        left_splitter.setStretchFactor(1, 1)
        splitter.addWidget(left_splitter)

        # === 右侧：搜索 + 季度标签 ===
        season_group = QGroupBox("📺 季度列表")
        season_main_layout = QVBoxLayout(season_group)
        season_main_layout.setSpacing(8)

        # 搜索行 - 始终在最上面
        search_row = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("输入番剧名称...")
        self.search_edit.setFixedHeight(28)
        self.search_edit.returnPressed.connect(self.search_and_select)
        search_row.addWidget(self.search_edit)

        search_btn = QPushButton("🔍 搜索")
        search_btn.setFixedHeight(28)
        search_btn.clicked.connect(self.search_and_select)
        search_row.addWidget(search_btn)
        season_main_layout.addLayout(search_row)

        # 已选信息 - 始终显示
        self.selected_info = QLabel("未选择番剧")
        self.selected_info.setStyleSheet("font-weight: bold; padding: 8px; background-color: #f0f8ff; border-radius: 4px; color: #888;")
        self.selected_info.setOpenExternalLinks(True)  # 允许点击链接
        season_main_layout.addWidget(self.selected_info)

        # 季度标签页 - 放在可滚动区域中间
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        # 内部容器
        scroll_content = QWidget()
        self.season_tabs_widget = QTabWidget()
        self.season_tabs_widget.setVisible(False)
        self.season_tabs = {}  # 季度标签页字典
        self.extras_tab = None  # extras 标签页

        content_layout = QVBoxLayout(scroll_content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.addWidget(self.season_tabs_widget)

        scroll_area.setWidget(scroll_content)
        season_main_layout.addWidget(scroll_area, 1)  # stretch=1 占满剩余空间

        # 整理按钮 - 始终在最下面
        self.link_btn = QPushButton("✅ 开始整理")
        self.link_btn.setFixedHeight(40)
        self.link_btn.setStyleSheet("""
            QPushButton { background-color: #4CAF50; color: white; font-size: 15px; font-weight: bold; border-radius: 5px; }
            QPushButton:hover { background-color: #45a049; }
            QPushButton:disabled { background-color: #ccc; }
        """)
        self.link_btn.clicked.connect(self.start_link)
        self.link_btn.setEnabled(False)
        season_main_layout.addWidget(self.link_btn)

        splitter.addWidget(season_group)
        splitter.setStretchFactor(0, 3)  # 左侧 60%
        splitter.setStretchFactor(1, 2)  # 右侧 40%

        layout.addWidget(splitter, 1)

    def load_anime_folders(self):
        """加载动漫文件夹（只加载一级文件夹名称）"""
        source = self.config.get('source_dir')
        if not source:
            QMessageBox.warning(self, "警告", "请先在设置中配置源目录")
            return
        source_path = Path(source)
        if not source_path.exists():
            QMessageBox.warning(self, "警告", "源目录不存在")
            return

        # 清空缓存
        self._folder_video_cache.clear()

        self.folder_tree.clear()
        folders = [d for d in source_path.iterdir() if d.is_dir()]

        # 批量 UI 更新
        self.folder_tree.setUpdatesEnabled(False)

        from datetime import datetime
        for folder in folders:
            item = FolderTreeItem()
            item.setText(0, folder.name)
            item.setText(1, "-")  # 占位，点击时才加载
            item.setText(2, datetime.fromtimestamp(folder.stat().st_mtime).strftime('%Y-%m-%d %H:%M'))
            item.setData(0, Qt.UserRole, folder)
            item.setExpanded(False)

            self.folder_tree.addTopLevelItem(item)
        
        self.folder_tree.setUpdatesEnabled(True)

    def on_folder_selected(self, item):
        """点击文件夹加载视频和子文件夹信息"""
        folder = item.data(0, Qt.UserRole)
        if not folder:
            return

        # 保存当前加载的文件夹路径
        self.current_folder = folder

        # 检查是否是子文件夹（通过文本判断）
        parent_text = item.text(0)
        is_subfolder = parent_text.startswith("📁 ")

        if is_subfolder:
            # 子文件夹：只加载该子文件夹根目录的视频（不包含更深层子文件夹）
            self._load_videos_to_list(folder, item, root_only=True)
        else:
            # 主文件夹：加载该文件夹的完整信息（统计、高亮、子文件夹、视频列表）
            self._load_folder_full_info(folder, item)

    def _load_folder_full_info(self, folder, item):
        """加载一级文件夹的完整信息（首次点击时加载子文件夹）"""
        # 检查是否已加载过子文件夹（通过子项数量判断）
        has_subfolders = item.childCount() > 0

        if not has_subfolders:
            # 首次加载：需要收集已匹配文件用于添加子文件夹
            matched_files = set()
            for tab in self.season_tabs.values():
                matched_files.update(tab.file_mappings.keys())
            if self.extras_tab:
                matched_files.update(self.extras_tab.file_mappings.keys())

            # 扫描该文件夹的所有视频（包含子文件夹），用于统计和高亮
            all_video_files = self._get_folder_videos(folder)
            matched_count = sum(1 for f in all_video_files if f in matched_files)
            total_count = len(all_video_files)

            # 更新统计和高亮
            item.setText(1, f"{matched_count}/{total_count}")
            is_all_matched = total_count > 0 and matched_count == total_count
            apply_highlight(item, is_all_matched)

            # 递归添加子文件夹（支持多级）
            self._add_subfolders_recursive(folder, item, matched_files)

        # 展开父文件夹
        item.setExpanded(True)

        # 启用刷新视频列表按钮
        self.refresh_video_btn.setEnabled(True)

        # 加载根目录视频到左下（不包含子文件夹）
        self._load_videos_to_list(folder, item, root_only=True)

    def _add_subfolders_recursive(self, parent_folder, parent_item, matched_files):
        """递归添加子文件夹（支持多级，每级独立统计和高亮）"""
        subfolders = [d for d in parent_folder.iterdir() if d.is_dir()]

        for sub in subfolders:
            # 获取该子文件夹的所有视频（包含其子文件夹）
            sub_video_files = self._get_folder_videos(sub)
            if not sub_video_files:
                continue

            child = FolderTreeItem()
            child.setText(0, "📁 " + sub.name)

            # 计算已匹配数量
            sub_matched = sum(1 for f in sub_video_files if f in matched_files)
            child.setText(1, f"{sub_matched}/{len(sub_video_files)}")
            child.setData(0, Qt.UserRole, sub)

            # 高亮判断：所有视频（包含子文件夹）都匹配才高亮
            is_sub_matched = len(sub_video_files) > 0 and sub_matched == len(sub_video_files)
            apply_highlight(child, is_sub_matched)
            parent_item.addChild(child)

            # 递归添加下一级子文件夹
            self._add_subfolders_recursive(sub, child, matched_files)

    def _load_videos_to_list(self, folder, item=None, root_only=False):
        """加载视频到左下列表"""
        # 从所有季度和 extras 收集已匹配的文件
        matched_files = set()
        for tab in self.season_tabs.values():
            matched_files.update(tab.file_mappings.keys())
        if self.extras_tab:
            matched_files.update(self.extras_tab.file_mappings.keys())

        # 加载视频
        if root_only:
            # 只加载根目录视频
            videos = sorted([f for f in folder.glob('*') if f.is_file() and f.suffix.lower() in self._video_extensions], key=lambda x: x.name)
        else:
            # 加载该文件夹的所有视频（包含子文件夹）
            videos = sorted(self._get_folder_videos(folder), key=lambda x: x.name)

        # 批量 UI 更新
        self.video_list.setUpdatesEnabled(False)
        self.video_list.setSortingEnabled(False)
        self.video_list.clear()

        from datetime import datetime
        items = []
        for v in videos:
            item = VideoTreeItem()
            is_matched = v in matched_files
            item.setText(0, "✓ " + v.name if is_matched else v.name)
            size_mb = v.stat().st_size / 1024 / 1024
            item.setText(1, f"{size_mb:.1f} MB")
            item.setData(1, Qt.UserRole, v.stat().st_size)
            date_str = datetime.fromtimestamp(v.stat().st_mtime).strftime('%Y-%m-%d %H:%M')
            item.setText(2, date_str)
            item.setData(2, Qt.UserRole, v.stat().st_mtime)
            item.setData(0, Qt.UserRole, v)
            apply_highlight(item, is_matched)
            items.append(item)

        self.video_list.addTopLevelItems(items)
        self.video_list.setSortingEnabled(True)
        self.video_list.setUpdatesEnabled(True)
        self.statusBar.showMessage(f"已加载 {len(videos)} 个视频文件 - {folder.name}")

    def on_refresh_video_clicked(self):
        """点击刷新视频列表按钮"""
        if self.current_folder:
            # 只清空当前文件夹的缓存
            if self.current_folder in self._folder_video_cache:
                del self._folder_video_cache[self.current_folder]
            
            # 重新扫描当前文件夹
            item = self.folder_tree.currentItem()
            if item:
                # 检查是否是子文件夹
                parent_text = item.text(0)
                is_subfolder = parent_text.startswith("📁 ")
                
                if is_subfolder:
                    # 子文件夹：重新加载视频
                    self._load_videos_to_list(self.current_folder, item)
                else:
                    # 主文件夹：重新加载完整信息
                    self._load_folder_full_info(self.current_folder, item)
            
            self.statusBar.showMessage(f"已刷新视频列表 - {self.current_folder.name}")

    def _on_video_selection_changed(self):
        """视频列表选择变化时显示选中数量"""
        selected_count = len(self.video_list.selectedItems())
        if selected_count > 0:
            self.statusBar.showMessage(f"已选中 {selected_count} 个文件")

    def _on_video_context_menu(self, pos):
        """视频列表右键菜单"""
        item = self.video_list.itemAt(pos)
        if not item:
            return

        path = item.data(0, Qt.UserRole)
        if not path:
            return

        from PyQt5.QtWidgets import QMenu
        from PyQt5.QtGui import QClipboard
        menu = QMenu(self)

        # 播放选项
        play_action = menu.addAction("▶️ 播放")
        play_action.triggered.connect(lambda: self._play_video(path))

        menu.addSeparator()

        # 打开文件夹选项
        open_folder_action = menu.addAction("📂 打开文件所在文件夹")
        open_folder_action.triggered.connect(lambda: self._open_file_location(path))

        menu.addSeparator()

        # 复制文件名选项
        copy_name_action = menu.addAction("📋 复制文件名")
        copy_name_action.triggered.connect(lambda: QApplication.clipboard().setText(path.name))

        menu.exec_(self.video_list.mapToGlobal(pos))

    def _play_video(self, path: Path):
        """使用默认播放器打开视频"""
        try:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        except Exception as e:
            QMessageBox.warning(self, "播放失败", f"无法打开视频文件：{e}")

    def _open_file_location(self, path: Path):
        """打开文件所在文件夹并选中文件"""
        try:
            # Windows 使用 explorer /select 命令
            if sys.platform == "win32":
                import subprocess
                # explorer 是异步启动的，不检查返回码
                subprocess.Popen(['explorer', '/select,', str(path)])
            else:
                # 其他系统只打开文件夹
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.parent)))
        except Exception as e:
            QMessageBox.warning(self, "打开失败", f"无法打开文件夹：{e}")

    def search_and_select(self):
        """搜索并选择番剧"""
        api_key = self.config.get('tmdb_api_key')
        if not api_key:
            QMessageBox.warning(self, "警告", "请先在设置中配置 TMDB API Key")
            return

        query = self.search_edit.text().strip()
        if not query:
            QMessageBox.warning(self, "警告", "请输入搜索关键词")
            return

        self.tmdb = TMDBClient(api_key)

        # 弹出搜索结果对话框
        dialog = SearchSelectDialog(self.tmdb, query, self)
        result = dialog.exec_()
        if result == QDialog.Accepted and dialog.selected_tv:
            # 获取详细信息（包含季度列表）
            tv_id = dialog.selected_tv.get('id')
            self.tv_info = self.tmdb.get_tv_details(tv_id)
            if not self.tv_info:
                QMessageBox.critical(self, "错误", "无法获取番剧详细信息")
                return
            tv_name = self.tv_info.get('name', '未知')
            # 显示番剧名称和 TMDB 链接
            self.selected_info.setText(f"📺 {tv_name} &nbsp;&nbsp;<a href='https://www.themoviedb.org/tv/{tv_id}' style='color: #03254C; text-decoration: none;'>🔗 TMDB</a>")
            self._load_season_tabs()
            self.statusBar.showMessage(f"已选择：{tv_name}")

    def _load_season_tabs(self):
        try:
            self.season_tabs_widget.clear()
            self.season_tabs.clear()
            self.file_mappings.clear()
            
            # 先添加 extras 标签页（始终在第一个）
            self.extras_tab = ExtrasTab(self)
            self.season_tabs_widget.addTab(self.extras_tab, "📦 extras")

            tv_id = self.tv_info.get('id')
            seasons = self.tv_info.get('seasons', [])

            for season in seasons:
                num = season.get('season_number', 0)
                name = season.get('name', f'Season {num}')
                count = season.get('episode_count', 0)

                tab = SeasonTab(num, name, count, self.tmdb, tv_id, self)  # 传入主窗口引用
                self.season_tabs[num] = tab
                self.season_tabs_widget.addTab(tab, f"S{num} ({count}集)")

            self.season_tabs_widget.setVisible(True)
            self.link_btn.setEnabled(True)
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"加载季度失败：{e}")

    def _get_folder_videos(self, folder: Path) -> List[Path]:
        """获取文件夹下的所有视频文件（带缓存）"""
        if folder in self._folder_video_cache:
            return self._folder_video_cache[folder]

        videos = [f for f in folder.rglob('*') if f.is_file() and f.suffix.lower() in self._video_extensions]
        self._folder_video_cache[folder] = videos
        return videos

    def _update_status(self, specific_paths=None):
        """
        更新状态（拖动视频后调用）

        Args:
            specific_paths: 指定要刷新的视频路径列表，如果为 None 则刷新所有已展开的文件夹
        """
        total = sum(len(tab.file_mappings) for tab in self.season_tabs.values())
        if self.extras_tab:
            total += len(self.extras_tab.file_mappings)
        self.statusBar.showMessage(f"已匹配 {total} 个文件")

        # 刷新视频列表高亮
        self._refresh_video_highlight()

        # 刷新已展开的文件夹高亮（轻量级操作）
        self._refresh_expanded_folders(specific_paths)

    def _refresh_expanded_folders(self, specific_paths=None):
        """
        刷新文件夹高亮（轻量级）

        Args:
            specific_paths: 指定要刷新的视频路径列表，如果为 None 则刷新所有已展开的文件夹
        """
        matched_files = self.get_matched_files()

        if specific_paths:
            # 指定了视频路径：只刷新这些视频所属的文件夹
            folders_to_refresh = set()
            for path in specific_paths:
                # 找到该视频所属的主文件夹
                folder = self._find_parent_folder(path)
                if folder and self._is_folder_expanded(folder):
                    folders_to_refresh.add(folder)

            for folder in folders_to_refresh:
                self._refresh_single_folder(folder, matched_files)
        else:
            # 未指定路径：刷新所有已展开的文件夹
            for i in range(self.folder_tree.topLevelItemCount()):
                parent = self.folder_tree.topLevelItem(i)
                folder = parent.data(0, Qt.UserRole)
                if not folder:
                    continue

                # 只刷新已展开的文件夹
                if not parent.isExpanded():
                    continue

                self._refresh_single_folder(folder, matched_files)

    def _find_parent_folder(self, video_path: Path) -> Optional[Path]:
        """查找视频所属的主文件夹（一级子目录）"""
        source = self.config.get('source_dir')
        if not source:
            return None

        source_path = Path(source)
        try:
            # 获取相对于源目录的路径
            rel_path = video_path.relative_to(source_path)
            # 第一级目录就是主文件夹
            return source_path / rel_path.parts[0]
        except ValueError:
            # 视频不在源目录下
            return None

    def _is_folder_expanded(self, folder: Path) -> bool:
        """检查文件夹是否已展开"""
        for i in range(self.folder_tree.topLevelItemCount()):
            item = self.folder_tree.topLevelItem(i)
            if item.data(0, Qt.UserRole) == folder:
                return item.isExpanded()
        return False

    def _refresh_single_folder(self, folder, matched_files):
        """刷新单个文件夹及其子文件夹的高亮"""
        for i in range(self.folder_tree.topLevelItemCount()):
            parent = self.folder_tree.topLevelItem(i)
            if parent.data(0, Qt.UserRole) == folder:
                # 使用缓存统计视频文件
                all_video_files = self._folder_video_cache.get(folder, [])

                matched_count = sum(1 for f in all_video_files if f in matched_files)
                total_count = len(all_video_files)

                # 更新计数和高亮
                parent.setText(1, f"{matched_count}/{total_count}")
                is_all_matched = total_count > 0 and matched_count == total_count
                apply_highlight(parent, is_all_matched)

                # 递归更新子文件夹
                self._refresh_expanded_subfolders(parent, matched_files)
                break

    def _refresh_expanded_subfolders(self, item, matched_files):
        """递归刷新已展开的子文件夹高亮"""
        for i in range(item.childCount()):
            child = item.child(i)
            folder = child.data(0, Qt.UserRole)
            if not folder:
                continue

            # 使用缓存统计
            all_video_files = self._folder_video_cache.get(folder, [])
            
            sub_matched = sum(1 for f in all_video_files if f in matched_files)
            child.setText(1, f"{sub_matched}/{len(all_video_files)}")
            is_sub_matched = len(all_video_files) > 0 and sub_matched == len(all_video_files)
            apply_highlight(child, is_sub_matched)

            # 递归更新下一级（只递归已展开的子文件夹）
            if child.isExpanded():
                self._refresh_expanded_subfolders(child, matched_files)

    def get_matched_files(self) -> set:
        """获取所有已匹配的文件路径（用于拖放检查）"""
        matched = set()
        for tab in self.season_tabs.values():
            matched.update(tab.file_mappings.keys())
        if self.extras_tab:
            matched.update(self.extras_tab.file_mappings.keys())
        return matched

    def _refresh_video_highlight(self):
        """刷新视频列表的高亮显示（不重新加载）"""
        matched_files = set()
        for tab in self.season_tabs.values():
            matched_files.update(tab.file_mappings.keys())
        # 加上 extras 标签页
        if self.extras_tab:
            matched_files.update(self.extras_tab.file_mappings.keys())

        # 禁用排序和更新，防止修改文本时 Qt 重新排序导致高亮错位
        self.video_list.setUpdatesEnabled(False)
        self.video_list.setSortingEnabled(False)

        # 遍历视频列表，更新高亮
        for i in range(self.video_list.topLevelItemCount()):
            item = self.video_list.topLevelItem(i)
            path = item.data(0, Qt.UserRole)
            if not path:
                continue

            is_matched = path in matched_files
            apply_highlight(item, is_matched)

            # 更新 ✓ 标记：始终使用原始文件名
            current_text = item.text(0)
            if is_matched:
                if not current_text.startswith("✓ "):
                    item.setText(0, "✓ " + path.name)
            else:
                if current_text.startswith("✓ "):
                    item.setText(0, path.name)

        # 恢复排序和更新
        self.video_list.setSortingEnabled(True)
        self.video_list.setUpdatesEnabled(True)

    def _move_subtitles(self, video_src: Path, target_folder: Path, ep_key: Optional[str], processed_files: set, mode: str) -> Tuple[int, List[Tuple[Path, Path, str]]]:
        """
        处理视频文件关联的字幕文件

        Args:
            video_src: 视频文件路径
            target_folder: 目标文件夹
            ep_key: 集数标识（如 "S01E01"），如果为 None 则不重命名
            processed_files: 已处理文件集合
            mode: 整理模式

        Returns:
            (success_count, fail_details) 元组
            - success_count: 成功处理的字幕数量
            - fail_details: 失败详情列表 [(src, dst, error), ...]
        """
        success = 0
        fail_details = []

        video_filename = video_src.stem
        video_parent = video_src.parent

        # 收集关联字幕文件
        sub_files_to_move = []
        for f in video_parent.iterdir():
            if f.is_file() and f.name.startswith(f"{video_filename}.") and f != video_src:
                if f.suffix.lower() in Config.SUBTITLE_EXTENSIONS:
                    if f not in processed_files:
                        sub_files_to_move.append(f)

        # 处理字幕文件
        for sub_src in sub_files_to_move:
            if ep_key:
                sub_dst = target_folder / f"{ep_key} - {sub_src.name}"
            else:
                sub_dst = target_folder / sub_src.name

            ok, error = FileOperator.operate(sub_src, sub_dst, mode)
            if ok:
                success += 1
                processed_files.add(sub_src)
            else:
                fail_details.append((sub_src, sub_dst, error))

        return success, fail_details

    def start_link(self):
        try:
            # 收集所有映射
            self.file_mappings.clear()
            for tab in self.season_tabs.values():
                self.file_mappings.update(tab.file_mappings)

            # 添加 extras 标签页的文件
            if self.extras_tab:
                self.file_mappings.update(self.extras_tab.file_mappings)

            if not self.file_mappings:
                QMessageBox.warning(self, "警告", "请先拖放文件到季度区域或 extras")
                return

            target = self.config.get('target_dir')
            if not target:
                QMessageBox.warning(self, "警告", "请先在设置中配置目标目录")
                return

            mode = self.config.get('move_mode', 'link')
            mode_names = {'link': '硬链接', 'cut': '剪切', 'copy': '复制'}

            if QMessageBox.question(self, "确认", f"使用 {mode_names.get(mode, '硬链接')} 模式整理 {len(self.file_mappings)} 个文件？") != QMessageBox.Yes:
                return

            target_path = Path(target)
            tv_name = self.tv_info.get('name', 'Unknown')
            year = (self.tv_info.get('first_air_date', '') or '')[:4]

            success, fail = 0, 0
            fail_details = []  # 记录失败详情
            extras_files = []  # 收集 extras 文件（包括 extras 标签页和 auto_extras 的）
            processed_files = set(self.file_mappings.keys())  # 记录所有已处理的文件（初始为视频文件，后续会添加字幕）

            # 1. 收集 extras 标签页的文件
            for src, ep_key in self.file_mappings.items():
                if not src.exists():
                    fail += 1
                    continue

                if ep_key == "extras":
                    extras_files.append(src)

            # 2. auto_extras: 扫描未匹配文件并添加到 extras_files
            if self.config.get('auto_extras', True):
                source_dir = Path(self.config.get('source_dir', ''))
                if source_dir.exists():
                    # 找到每个已匹配文件所属的动漫文件夹
                    anime_folders = set()
                    for f in self.file_mappings.keys():
                        try:
                            relative = f.relative_to(source_dir)
                            anime_folder = source_dir / relative.parts[0]
                            anime_folders.add(anime_folder)
                        except ValueError:
                            continue

                    # 扫描每个动漫文件夹的未匹配文件
                    for anime_folder in anime_folders:
                        all_videos = self._get_folder_videos(anime_folder)
                        matched_paths = set(self.file_mappings.keys())
                        unmatched_videos = [v for v in all_videos if v not in matched_paths]
                        extras_files.extend(unmatched_videos)

            # 3. 处理季度文件（重命名 + 字幕）
            for src, ep_key in self.file_mappings.items():
                if not src.exists():
                    fail += 1
                    continue

                if ep_key == "extras":
                    continue  # extras 文件稍后处理

                s_num = int(ep_key[1:3])
                folder = target_path / f"{tv_name} ({year})" / (f"Season0" if s_num == 0 else f"Season{s_num}")
                dst = folder / f"{ep_key} - {src.name}"

                # 处理视频文件
                ok, error = FileOperator.operate(src, dst, mode)
                if ok:
                    success += 1
                else:
                    fail += 1
                    fail_details.append((src, dst, error))
                    continue

                # 处理关联字幕文件
                sub_success, sub_fail = self._move_subtitles(src, folder, ep_key, processed_files, mode)
                success += sub_success
                for sub_src, sub_dst, error in sub_fail:
                    fail += 1
                    fail_details.append((sub_src, sub_dst, error))

            # 4. 处理所有 extras 文件（不重命名）
            extras_folder = target_path / f"{tv_name} ({year})" / "extras"
            extras_folder.mkdir(parents=True, exist_ok=True)

            for src in extras_files:
                if src.exists():
                    dst = extras_folder / src.name
                    ok, error = FileOperator.operate(src, dst, mode)
                    if ok:
                        success += 1
                        processed_files.add(src)
                        # 处理关联字幕文件
                        sub_success, sub_fail = self._move_subtitles(src, extras_folder, None, processed_files, mode)
                        success += sub_success
                        for sub_src, sub_dst, error in sub_fail:
                            fail += 1
                            fail_details.append((sub_src, sub_dst, error))
                    else:
                        fail += 1
                        fail_details.append((src, dst, error))

            # 生成.embyignore 文件（如果开启）
            if self.config.get('embyignore_extras', True) and extras_files:
                embyignore_file = extras_folder / ".embyignore"
                if not embyignore_file.exists():
                    embyignore_file.touch()

            # 收集未整理的文件
            unorganized_files = []
            
            # 从 fail_details 中提取因重名未整理的文件（完整路径）
            duplicate_files = set()
            for src, dst, error in fail_details:
                if "目标文件已存在" in error:
                    duplicate_files.add(src)  # 添加源文件完整路径
            
            # 扫描源目录中的所有文件，排除已处理的
            source_dir = Path(self.config.get('source_dir', ''))
            
            if source_dir.exists():
                # 找到每个已匹配文件所属的动漫文件夹
                anime_folders = set()
                for f in self.file_mappings.keys():
                    try:
                        relative = f.relative_to(source_dir)
                        anime_folder = source_dir / relative.parts[0]
                        anime_folders.add(anime_folder)
                    except ValueError:
                        continue
                
                # 扫描每个动漫文件夹中的所有文件
                for anime_folder in anime_folders:
                    # 先添加重名失败的文件（标红置顶）
                    for dup_file in duplicate_files:
                        if dup_file.is_relative_to(anime_folder):
                            unorganized_files.append((dup_file, True))
                    
                    # 再添加普通未整理文件
                    for f in anime_folder.rglob('*'):
                        if f.is_file():
                            if f not in processed_files:
                                unorganized_files.append((f, False))

            # 显示完成弹窗
            mode_names = {'link': '硬链接', 'cut': '剪切', 'copy': '复制'}
            mode_name = mode_names.get(mode, '硬链接')
            
            dialog = OrganizeResultDialog(success, fail, unorganized_files, mode_name, self)
            if dialog.exec_() == QDialog.Accepted:
                selected_files = dialog.selected_files
                if selected_files:
                    reply = QMessageBox.question(
                        self, "确认",
                        f"确定要将 {len(selected_files)} 个文件以 {mode_name} 方式整理到 extras 文件夹吗？",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    if reply == QMessageBox.Yes:
                        extras_move_count = 0
                        extras_folder = target_path / f"{tv_name} ({year})" / "extras"
                        extras_folder.mkdir(parents=True, exist_ok=True)
                        processed_subs = set()

                        for src in selected_files:
                            if src.exists():
                                dst = extras_folder / src.name
                                ok, error = FileOperator.operate(src, dst, mode)
                                if ok:
                                    extras_move_count += 1
                                    # 处理关联字幕文件
                                    sub_success, _ = self._move_subtitles(src, extras_folder, None, processed_subs, mode)
                                    extras_move_count += sub_success

                        QMessageBox.information(self, "完成", f"已{mode_name} {extras_move_count} 个文件到 extras 文件夹")

            # 不清空映射，保持按钮可用（用户可以修改映射后重新整理，或整理下一个动漫）
            # 搜索新番剧时 _load_season_tabs() 会自动清空映射
            
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(f"[错误] 整理失败：{error_msg}")
            QMessageBox.critical(self, "错误", f"整理失败：{e}\n\n{error_msg}")

    def open_config(self):
        dialog = ConfigDialog(self.config, self)
        if dialog.exec_() == QDialog.Accepted:
            dialog.save()
            if self.config.get('tmdb_api_key'):
                self.tmdb = TMDBClient(self.config.get('tmdb_api_key'))

    def closeEvent(self, event):
        # 使用 Qt 标准方式保存窗口状态
        settings = QSettings("AniOr", "AniOr")
        settings.setValue("geometry", self.saveGeometry())
        settings.setValue("windowState", self.saveState())
        settings.setValue("splitterState", self.centralWidget().findChild(QSplitter).saveState())
        settings.setValue("folderTreeHeaderState", self.folder_tree.header().saveState())
        settings.setValue("videoTreeHeaderState", self.video_list.header().saveState())
        
        # 保存用户配置
        self.config.save_if_needed()
        event.accept()

    def showEvent(self, event):
        super().showEvent(event)
        
        # 只在首次启动时加载文件夹
        if self._first_show:
            self._first_show = False
            
            # 恢复窗口状态
            settings = QSettings("AniOr", "AniOr")
            geometry = settings.value("geometry")
            if geometry:
                self.restoreGeometry(geometry)

            # 恢复 splitter 状态
            splitter = self.centralWidget().findChild(QSplitter)
            if splitter:
                splitterState = settings.value("splitterState")
                if splitterState:
                    splitter.restoreState(splitterState)

            # 恢复表头状态（列宽、顺序、排序）
            folderTreeHeaderState = settings.value("folderTreeHeaderState")
            if folderTreeHeaderState:
                self.folder_tree.header().restoreState(folderTreeHeaderState)

            videoTreeHeaderState = settings.value("videoTreeHeaderState")
            if videoTreeHeaderState:
                self.video_list.header().restoreState(videoTreeHeaderState)

            # 延迟加载文件夹
            QTimer.singleShot(500, lambda: self.load_anime_folders())


# ==================== 配置对话框 ====================
class ConfigDialog(QDialog):
    def __init__(self, config: Config, parent=None):
        super().__init__(parent)
        self.config = config
        self.setWindowTitle("设置")
        self.setMinimumWidth(500)

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.source_edit = QLineEdit(config.get('source_dir'))
        self.source_edit.setPlaceholderText("选择源目录...")
        source_btn = QPushButton("浏览...")
        source_btn.clicked.connect(lambda: self._browse(self.source_edit))
        sw = QWidget()
        sl = QHBoxLayout(sw)
        sl.setContentsMargins(0,0,0,0)
        sl.addWidget(self.source_edit)
        sl.addWidget(source_btn)
        form.addRow("源目录:", sw)

        self.target_edit = QLineEdit(config.get('target_dir'))
        target_btn = QPushButton("浏览...")
        target_btn.clicked.connect(lambda: self._browse(self.target_edit))
        tw = QWidget()
        tl = QHBoxLayout(tw)
        tl.setContentsMargins(0,0,0,0)
        tl.addWidget(self.target_edit)
        tl.addWidget(target_btn)
        form.addRow("目标目录:", tw)

        self.api_edit = QLineEdit(config.get('tmdb_api_key'))
        form.addRow("TMDB API Key:", self.api_edit)

        self.mode_combo = QComboBox()
        self.mode_combo.addItem("硬链接 (推荐)", "link")
        self.mode_combo.addItem("剪切", "cut")
        self.mode_combo.addItem("复制", "copy")
        idx = self.mode_combo.findData(config.get('move_mode', 'link'))
        self.mode_combo.setCurrentIndex(idx if idx >= 0 else 0)
        form.addRow("整理模式:", self.mode_combo)

        # 视频格式配置
        self.ext_edit = QLineEdit()
        self.ext_edit.setPlaceholderText(".mp4,.mkv,.avi,...")
        self.ext_edit.setText(",".join(config.get('video_extensions', config.DEFAULT_VIDEO_EXTENSIONS)))
        form.addRow("视频格式:", self.ext_edit)
        
        # 视频格式说明
        ext_tip = QLabel("用逗号或空格分隔，如：.mp4,.mkv,.avi")
        ext_tip.setStyleSheet("color: #666; font-size: 12px;")
        form.addRow("", ext_tip)

        # 未匹配文件配置
        self.extras_check = QCheckBox("整理到 extras 文件夹")
        self.extras_check.setChecked(config.get('auto_extras', True))
        form.addRow("未匹配文件:", self.extras_check)

        # 未匹配文件说明
        extras_tip = QLabel("开启后，未匹配的视频文件会自动移动到 extras 文件夹")
        extras_tip.setStyleSheet("color: #666; font-size: 12px;")
        form.addRow("", extras_tip)

        # embyignore 配置
        self.embyignore_check = QCheckBox("生成.embyignore 文件")
        self.embyignore_check.setChecked(config.get('embyignore_extras', True))
        form.addRow("extras 忽略:", self.embyignore_check)

        # embyignore 说明
        embyignore_tip = QLabel("在 extras 文件夹生成.embyignore 文件，让 Emby 忽略该文件夹")
        embyignore_tip.setStyleSheet("color: #666; font-size: 12px;")
        form.addRow("", embyignore_tip)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _browse(self, edit: QLineEdit):
        path = QFileDialog.getExistingDirectory(self, "选择目录")
        if path:
            edit.setText(path)

    def save(self):
        self.config.set('source_dir', self.source_edit.text())
        self.config.set('target_dir', self.target_edit.text())
        self.config.set('tmdb_api_key', self.api_edit.text())
        self.config.set('move_mode', self.mode_combo.currentData())
        # 解析视频格式配置
        ext_text = self.ext_edit.text().strip()
        # 支持逗号和空格分隔
        exts = [ext.strip().lower() for ext in re.split(r'[,\s]+', ext_text) if ext.strip()]
        if exts:
            self.config.set('video_extensions', exts)
        # 保存 auto_extras 配置
        self.config.set('auto_extras', self.extras_check.isChecked())
        # 保存 embyignore_extras 配置
        self.config.set('embyignore_extras', self.embyignore_check.isChecked())
        self.config.save()


# ==================== 主函数 ====================
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet("""
        QTreeWidget::item { padding: 4px; }
        QLineEdit { padding: 4px 8px; border: 1px solid #ccc; border-radius: 4px; }
        QStatusBar { background-color: #f5f5f5; }
    """)

    # 设置应用程序图标
    icon_path = Path(__file__).parent / "docs" / "icon.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    try:
        window = MainWindow()
        window.show()
    except Exception as e:
        import traceback
        print(f"启动错误：{e}")
        print(traceback.format_exc())
        QMessageBox.critical(None, "启动错误", f"程序启动失败：{e}")
        sys.exit(1)

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
