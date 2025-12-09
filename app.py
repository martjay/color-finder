"""
颜色块识别程序 - PyQt6版本
支持：上传图片、粘贴图片、全屏截图框选、窗口置顶、DPI校准、暗色主题适配
"""
import sys
import os
import ctypes
import numpy as np
from io import BytesIO
from collections import Counter
from PIL import Image, ImageGrab

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QFrame, QCheckBox, QDialog,
    QSpinBox, QDialogButtonBox, QGroupBox, QFormLayout, QSizePolicy, QSlider
)
from PyQt6.QtCore import Qt, QRect, QBuffer, QIODevice, QTimer, QSettings
from PyQt6.QtGui import (
    QPixmap, QPainter, QColor, QPen, QGuiApplication, QPalette,
    QShortcut, QKeySequence, QDragEnterEvent, QDropEvent, QIcon
)

# Windows DPI感知
if sys.platform == 'win32':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except:
            pass


def is_dark_theme():
    """检测是否为暗色主题"""
    palette = QApplication.palette()
    bg_color = palette.color(QPalette.ColorRole.Window)
    # 亮度小于128认为是暗色主题
    return bg_color.lightness() < 128


def get_style(dark=False):
    """获取样式表"""
    if dark:
        return '''
            QMainWindow, QDialog {
                background-color: #1e1e1e;
            }
            QLabel {
                color: #e0e0e0;
            }
            QCheckBox {
                color: #e0e0e0;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QGroupBox {
                color: #e0e0e0;
                border: 1px solid #444;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QSpinBox {
                background: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #444;
                border-radius: 4px;
                padding: 5px 10px;
                min-width: 100px;
                min-height: 28px;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                width: 20px;
            }
            QFrame#previewFrame {
                background: #2d2d2d;
                border: 2px dashed #555;
                border-radius: 12px;
            }
            QPushButton#calibrateBtn {
                background: #3d3d3d;
                color: #e0e0e0;
                border: 1px solid #555;
                padding: 6px 14px;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton#calibrateBtn:hover {
                background: #4d4d4d;
            }
            QDialogButtonBox QPushButton {
                background: #3d3d3d;
                color: #e0e0e0;
                border: 1px solid #555;
                padding: 8px 20px;
                border-radius: 4px;
                min-width: 80px;
            }
            QDialogButtonBox QPushButton:hover {
                background: #4d4d4d;
            }
        '''
    else:
        return '''
            QFrame#previewFrame {
                background: #f8f9fa;
                border: 2px dashed #ddd;
                border-radius: 12px;
            }
            QPushButton#calibrateBtn {
                background: #f0f0f0;
                color: #333;
                border: 1px solid #ddd;
                padding: 6px 14px;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton#calibrateBtn:hover {
                background: #e0e0e0;
            }
            QSpinBox {
                padding: 5px 10px;
                min-width: 100px;
                min-height: 28px;
            }
            QGroupBox {
                border: 1px solid #ccc;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        '''


class CalibrationDialog(QDialog):
    """DPI校准对话框"""
    
    def __init__(self, parent=None, dark=False):
        super().__init__(parent)
        self.dark = dark
        self.setWindowTitle('DPI校准设置')
        self.setMinimumWidth(400)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 说明
        info = QLabel(
            '如果截图位置不准确，请调整缩放比例：\n\n'
            '• 值 > 100%：截图区域会放大\n'
            '• 值 < 100%：截图区域会缩小\n\n'
            '一般设置为系统显示缩放比例即可'
        )
        info.setWordWrap(True)
        info.setStyleSheet(f'color: {"#aaa" if dark else "#666"}; line-height: 1.5;')
        layout.addWidget(info)
        
        # 缩放设置
        group = QGroupBox('缩放比例')
        form = QFormLayout(group)
        form.setSpacing(12)
        form.setContentsMargins(15, 20, 15, 15)
        
        self.scale_x = QSpinBox()
        self.scale_x.setRange(50, 300)
        self.scale_x.setValue(100)
        self.scale_x.setSuffix(' %')
        
        self.scale_y = QSpinBox()
        self.scale_y.setRange(50, 300)
        self.scale_y.setValue(100)
        self.scale_y.setSuffix(' %')
        
        form.addRow('水平缩放:', self.scale_x)
        form.addRow('垂直缩放:', self.scale_y)
        layout.addWidget(group)
        
        # 自动检测按钮
        auto_btn = QPushButton('🔍 自动检测系统DPI')
        auto_btn.setStyleSheet(f'''
            QPushButton {{
                background: {"#4a6cf7" if dark else "#667eea"};
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-size: 14px;
            }}
            QPushButton:hover {{
                background: {"#3a5ce7" if dark else "#5a6fd6"};
            }}
        ''')
        auto_btn.clicked.connect(self.auto_detect)
        layout.addWidget(auto_btn)
        
        layout.addSpacing(10)
        
        # 确定取消
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.adjustSize()
    
    def auto_detect(self):
        screen = QGuiApplication.primaryScreen()
        if screen:
            ratio = screen.devicePixelRatio()
            self.scale_x.setValue(int(ratio * 100))
            self.scale_y.setValue(int(ratio * 100))
    
    def get_scale(self):
        return self.scale_x.value() / 100.0, self.scale_y.value() / 100.0
    
    def set_scale(self, sx, sy):
        self.scale_x.setValue(int(sx * 100))
        self.scale_y.setValue(int(sy * 100))


class ScreenshotOverlay(QWidget):
    """全屏截图选区覆盖层"""
    
    def __init__(self, screenshot, callback, scale_x=1.0, scale_y=1.0):
        super().__init__()
        self.screenshot = screenshot
        self.callback = callback
        self.scale_x = scale_x
        self.scale_y = scale_y
        self.start_pos = None
        self.end_pos = None
        self.is_selecting = False
        
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint | 
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        
        screen = QGuiApplication.primaryScreen()
        self.setGeometry(screen.geometry())
        self.setCursor(Qt.CursorShape.CrossCursor)
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self.width(), self.height(), self.screenshot)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 80))
        
        if self.start_pos and self.end_pos:
            rect = QRect(self.start_pos, self.end_pos).normalized()
            src_rect = QRect(
                int(rect.x() * self.scale_x), int(rect.y() * self.scale_y),
                int(rect.width() * self.scale_x), int(rect.height() * self.scale_y)
            )
            painter.drawPixmap(rect, self.screenshot, src_rect)
            painter.setPen(QPen(QColor(102, 126, 234), 2))
            painter.drawRect(rect)
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(rect.x() + 5, rect.y() - 8, 
                           f'{int(rect.width() * self.scale_x)} x {int(rect.height() * self.scale_y)}')
        
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(20, 30, "拖动鼠标选择区域，松开完成截图，按 ESC 取消")
    
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.start_pos = self.end_pos = event.pos()
            self.is_selecting = True
    
    def mouseMoveEvent(self, event):
        if self.is_selecting:
            self.end_pos = event.pos()
            self.update()
    
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.is_selecting:
            self.is_selecting = False
            if self.start_pos and self.end_pos:
                rect = QRect(self.start_pos, self.end_pos).normalized()
                if rect.width() > 10 and rect.height() > 10:
                    src_rect = QRect(
                        int(rect.x() * self.scale_x), int(rect.y() * self.scale_y),
                        int(rect.width() * self.scale_x), int(rect.height() * self.scale_y)
                    )
                    self.callback(self.screenshot.copy(src_rect))
                else:
                    self.callback(None)
            self.close()
    
    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.callback(None)
            self.close()


class ColorAnalyzer:
    """颜色分析器 - 简化版本"""
    
    @staticmethod
    def analyze(pixmap, sensitivity=1.5, debug=False):
        """分析图片
        
        sensitivity: 敏感度 0.5-3.0
        debug: 是否返回调试信息
        """
        buffer = QBuffer()
        buffer.open(QIODevice.OpenModeFlag.WriteOnly)
        pixmap.save(buffer, "PNG")
        pil_image = Image.open(BytesIO(buffer.data().data())).convert('RGB')
        img_array = np.array(pil_image)
        
        # 使用最简单的方法：尝试所有可能的网格
        grid_info = ColorAnalyzer._detect_grid_simple(img_array)
        if not grid_info:
            return None, "无法检测到有效的颜色网格"
        
        rows, cols, cells = grid_info
        
        # 提取每个单元格的颜色
        cell_colors = []
        for x, y, w, h in cells:
            # 采样中心区域
            mx, my = int(w * 0.2), int(h * 0.2)
            cx, cy = x + w//2, y + h//2
            sample_w, sample_h = max(5, w - 2*mx), max(5, h - 2*my)
            
            region = img_array[
                max(0, cy-sample_h//2):min(img_array.shape[0], cy+sample_h//2),
                max(0, cx-sample_w//2):min(img_array.shape[1], cx+sample_w//2)
            ]
            
            if region.size > 0:
                cell_colors.append(np.mean(region.reshape(-1, 3), axis=0))
            else:
                cell_colors.append(np.array([128, 128, 128]))
        
        # 找出不同颜色
        different_cells = ColorAnalyzer._find_different(cell_colors, sensitivity)
        results = []
        for idx in different_cells:
            c = cell_colors[idx]
            row = idx // cols + 1
            col = idx % cols + 1
            results.append({
                'row': row,
                'col': col,
                'color': f'RGB({int(c[0])}, {int(c[1])}, {int(c[2])})'
            })
        
        result = {'grid_size': f'{rows}x{cols}', 'results': results}
        
        # 调试信息
        if debug:
            result['debug_colors'] = []
            for idx, c in enumerate(cell_colors):
                row = idx // cols + 1
                col = idx % cols + 1
                result['debug_colors'].append({
                    'row': row,
                    'col': col,
                    'color': f'RGB({int(c[0])}, {int(c[1])}, {int(c[2])})'
                })
        
        return result, None
    
    @staticmethod
    def _detect_grid_simple(img_array):
        """检测网格 - 基于连通区域"""
        height, width = img_array.shape[:2]
        
        # 转换为灰度
        gray = np.mean(img_array, axis=2)
        
        # 检测背景色（通常是白色或边缘颜色）
        # 采样四个角
        corners = [
            gray[0:5, 0:5],
            gray[0:5, -5:],
            gray[-5:, 0:5],
            gray[-5:, -5:]
        ]
        bg_values = [np.mean(c) for c in corners]
        bg_gray = np.median(bg_values)
        
        # 二值化：背景vs前景
        threshold = bg_gray - 20
        foreground = gray < threshold
        
        # 找到前景区域的边界
        rows_with_fg = np.any(foreground, axis=1)
        cols_with_fg = np.any(foreground, axis=0)
        
        if not np.any(rows_with_fg) or not np.any(cols_with_fg):
            # 没有明显前景，尝试简单分割
            return ColorAnalyzer._fallback_grid(img_array)
        
        # 找到前景的起止位置
        row_start = np.argmax(rows_with_fg)
        row_end = len(rows_with_fg) - np.argmax(rows_with_fg[::-1])
        col_start = np.argmax(cols_with_fg)
        col_end = len(cols_with_fg) - np.argmax(cols_with_fg[::-1])
        
        # 在前景区域内检测行列分割
        fg_region = foreground[row_start:row_end, col_start:col_end]
        
        # 检测行分割（水平方向的空白）
        row_density = np.mean(fg_region, axis=1)
        row_gaps = row_density < 0.3  # 空白行
        
        # 检测列分割（垂直方向的空白）
        col_density = np.mean(fg_region, axis=0)
        col_gaps = col_density < 0.3  # 空白列
        
        # 找到分割线
        h_splits = ColorAnalyzer._find_gaps(row_gaps, row_start, row_end)
        v_splits = ColorAnalyzer._find_gaps(col_gaps, col_start, col_end)
        
        if len(h_splits) < 2 or len(v_splits) < 2:
            return ColorAnalyzer._fallback_grid(img_array)
        
        rows = len(h_splits) - 1
        cols = len(v_splits) - 1
        
        # 生成单元格
        cells = []
        for i in range(rows):
            for j in range(cols):
                x = v_splits[j]
                y = h_splits[i]
                w = v_splits[j+1] - v_splits[j]
                h = h_splits[i+1] - h_splits[i]
                cells.append((x, y, w, h))
        
        return rows, cols, cells
    
    @staticmethod
    def _find_gaps(gaps, start, end):
        """找到间隙位置"""
        splits = [start]
        in_gap = False
        gap_start = 0
        
        for i, is_gap in enumerate(gaps):
            if is_gap and not in_gap:
                in_gap = True
                gap_start = i
            elif not is_gap and in_gap:
                in_gap = False
                # 间隙中点
                gap_mid = start + (gap_start + i) // 2
                if gap_mid - splits[-1] > (end - start) * 0.05:  # 至少5%间隔
                    splits.append(gap_mid)
        
        splits.append(end)
        return splits
    
    @staticmethod
    def _fallback_grid(img_array):
        """备用方案 - 简单均匀分割"""
        height, width = img_array.shape[:2]
        best = None
        best_score = 0
        
        for rows in range(2, 8):
            for cols in range(2, 8):
                if rows * cols > 50:
                    continue
                
                ch, cw = height // rows, width // cols
                if ch < 20 or cw < 20:
                    continue
                
                cells = [(j*cw, i*ch, cw, ch) for i in range(rows) for j in range(cols)]
                score = ColorAnalyzer._score_grid(img_array, cells, rows, cols)
                
                if score > best_score:
                    best_score = score
                    best = (rows, cols, cells)
        
        return best
    
    @staticmethod
    def _score_grid(img_array, cells, rows, cols):
        """评估网格质量"""
        colors = []
        
        for x, y, w, h in cells:
            # 采样中心区域
            margin_x = int(w * 0.15)
            margin_y = int(h * 0.15)
            
            y1 = max(0, y + margin_y)
            y2 = min(img_array.shape[0], y + h - margin_y)
            x1 = max(0, x + margin_x)
            x2 = min(img_array.shape[1], x + w - margin_x)
            
            region = img_array[y1:y2, x1:x2]
            
            if region.size > 0:
                # 检查是否是背景（白色或灰色）
                avg_color = np.mean(region.reshape(-1, 3), axis=0)
                if np.mean(avg_color) > 200:  # 太亮，可能是背景
                    return 0
                colors.append(avg_color)
            else:
                return 0
        
        if len(colors) != rows * cols:
            return 0
        
        colors = np.array(colors)
        
        # 检查颜色方差
        color_std = np.std(colors, axis=0)
        avg_std = np.mean(color_std)
        
        # 颜色应该有一定的一致性
        median_color = np.median(colors, axis=0)
        distances = np.sqrt(np.sum((colors - median_color) ** 2, axis=1))
        
        # 至少70%的颜色相似
        similar_count = np.sum(distances < 80)
        similarity = similar_count / len(colors)
        
        # 评分：相似度高，但不是完全相同
        if similarity < 0.7:
            return 0
        
        if avg_std < 5:  # 颜色太相似
            return similarity * 0.6
        
        return similarity
    
    @staticmethod
    def _find_different(cell_colors, sensitivity=1.0):
        """找出颜色不同的单元格 - 多策略组合
        
        sensitivity: 敏感度 0.5-3.0
        """
        if len(cell_colors) < 2:
            return []
        
        colors = np.array(cell_colors)
        n = len(colors)
        
        # 策略1: 计算每个颜色到所有其他颜色的平均距离
        avg_distances = []
        for i in range(n):
            dists = [np.sqrt(np.sum((colors[i] - colors[j]) ** 2)) for j in range(n) if i != j]
            avg_distances.append(np.mean(dists))
        
        avg_distances = np.array(avg_distances)
        
        # 找出平均距离最大的（最孤立的）
        max_avg_dist = avg_distances.max()
        min_threshold = 10 / sensitivity
        
        if max_avg_dist < min_threshold:
            return []
        
        # 策略2: 使用Z-score找异常值
        mean_avg_dist = np.mean(avg_distances)
        std_avg_dist = np.std(avg_distances)
        
        if std_avg_dist > 0.1:
            z_scores = (avg_distances - mean_avg_dist) / std_avg_dist
            # Z-score > 1.5 认为是异常
            outliers_z = np.where(z_scores > 1.5)[0]
            if len(outliers_z) > 0:
                return outliers_z.tolist()
        
        # 策略3: 使用中位数绝对偏差(MAD)
        median_color = np.median(colors, axis=0)
        distances_to_median = np.sqrt(np.sum((colors - median_color) ** 2, axis=1))
        
        mad = np.median(np.abs(distances_to_median - np.median(distances_to_median)))
        if mad > 0.1:
            # 修正的Z-score
            modified_z = 0.6745 * (distances_to_median - np.median(distances_to_median)) / mad
            outliers_mad = np.where(modified_z > 2.5)[0]
            if len(outliers_mad) > 0:
                return outliers_mad.tolist()
        
        # 策略4: K-means聚类（备用）
        max_pair_dist = 0
        center1_idx = 0
        center2_idx = 1
        
        for i in range(n):
            for j in range(i+1, n):
                dist = np.sqrt(np.sum((colors[i] - colors[j]) ** 2))
                if dist > max_pair_dist:
                    max_pair_dist = dist
                    center1_idx = i
                    center2_idx = j
        
        if max_pair_dist < min_threshold:
            return []
        
        center1 = colors[center1_idx].copy()
        center2 = colors[center2_idx].copy()
        
        # 迭代优化
        for _ in range(10):
            dist_to_c1 = np.sqrt(np.sum((colors - center1) ** 2, axis=1))
            dist_to_c2 = np.sqrt(np.sum((colors - center2) ** 2, axis=1))
            
            cluster1 = np.where(dist_to_c1 <= dist_to_c2)[0]
            cluster2 = np.where(dist_to_c1 > dist_to_c2)[0]
            
            if len(cluster1) == 0 or len(cluster2) == 0:
                break
            
            new_center1 = np.mean(colors[cluster1], axis=0)
            new_center2 = np.mean(colors[cluster2], axis=0)
            
            if np.allclose(center1, new_center1) and np.allclose(center2, new_center2):
                break
            
            center1 = new_center1
            center2 = new_center2
        
        # 返回少数组
        if len(cluster1) > 0 and len(cluster2) > 0:
            if len(cluster1) < len(cluster2):
                return cluster1.tolist()
            elif len(cluster2) < len(cluster1):
                return cluster2.tolist()
        
        # 策略5: 如果都失败，找距离中位数最远的那个
        max_idx = int(np.argmax(distances_to_median))
        if distances_to_median[max_idx] > min_threshold:
            return [max_idx]
        
        return []


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.current_pixmap = None
        self.screenshot_overlay = None
        self.settings = QSettings('ColorFinder', 'ColorFinder')
        self.dark = is_dark_theme()
        
        self.dpi_scale_x = self.settings.value('dpi_scale_x', 1.0, type=float)
        self.dpi_scale_y = self.settings.value('dpi_scale_y', 1.0, type=float)
        self.sensitivity = self.settings.value('sensitivity', 1.5, type=float)
        
        if self.settings.value('first_run', True, type=bool):
            self._auto_detect_dpi()
            self.settings.setValue('first_run', False)
        
        self.init_ui()
        self.setStyleSheet(get_style(self.dark))
    
    def _auto_detect_dpi(self):
        screen = QGuiApplication.primaryScreen()
        if screen:
            ratio = screen.devicePixelRatio()
            self.dpi_scale_x = self.dpi_scale_y = ratio
            self.settings.setValue('dpi_scale_x', ratio)
            self.settings.setValue('dpi_scale_y', ratio)
    
    def init_ui(self):
        self.setWindowTitle('颜色块识别器')
        
        # 设置窗口图标 - 支持打包后的路径
        if getattr(sys, 'frozen', False):
            # 打包后的路径
            base_path = sys._MEIPASS
        else:
            # 开发环境路径
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        icon_path = os.path.join(base_path, 'icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        self.setMinimumSize(700, 500)
        
        # 恢复窗口大小和位置
        geometry = self.settings.value('window_geometry')
        if geometry:
            self.restoreGeometry(geometry)
        else:
            self.resize(850, 650)
        
        self.setAcceptDrops(True)
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 控制栏
        title_layout = QHBoxLayout()
        
        # 敏感度滑块
        sens_label = QLabel('敏感度:')
        sens_label.setStyleSheet('font-size: 12px;')
        title_layout.addWidget(sens_label)
        
        self.sensitivity_slider = QSlider(Qt.Orientation.Horizontal)
        self.sensitivity_slider.setMinimum(5)
        self.sensitivity_slider.setMaximum(30)  # 提高到3.0
        self.sensitivity_slider.setValue(int(self.sensitivity * 10))
        self.sensitivity_slider.setFixedWidth(100)
        self.sensitivity_slider.setToolTip('调节颜色识别敏感度\n左：只识别明显差异\n右：识别极细微差异')
        self.sensitivity_slider.valueChanged.connect(self.on_sensitivity_changed)
        title_layout.addWidget(self.sensitivity_slider)
        
        self.sens_value_label = QLabel(f'{self.sensitivity:.1f}')
        self.sens_value_label.setStyleSheet('font-size: 12px; min-width: 30px;')
        title_layout.addWidget(self.sens_value_label)
        
        # 调试模式开关
        self.debug_cb = QCheckBox('调试')
        self.debug_cb.setStyleSheet('font-size: 11px;')
        self.debug_cb.setToolTip('显示所有方块的颜色值')
        self.debug_cb.stateChanged.connect(lambda: self.analyze() if self.current_pixmap else None)
        title_layout.addWidget(self.debug_cb)
        
        title_layout.addSpacing(10)
        
        self.btn_calibrate = QPushButton('⚙️ DPI校准')
        self.btn_calibrate.setObjectName('calibrateBtn')
        self.btn_calibrate.clicked.connect(self.show_calibration)
        title_layout.addWidget(self.btn_calibrate)
        
        self.topmost_cb = QCheckBox('窗口置顶')
        self.topmost_cb.setStyleSheet('font-size: 14px;')
        self.topmost_cb.stateChanged.connect(self.toggle_topmost)
        title_layout.addWidget(self.topmost_cb)
        layout.addLayout(title_layout)
        
        # 按钮栏
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(15)
        
        btn_style = '''
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #667eea, stop:1 #764ba2);
                color: white; border: none; padding: 12px 25px;
                font-size: 14px; border-radius: 8px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #5a6fd6, stop:1 #6a4190);
            }
        '''
        
        for text, handler in [('📁 选择图片', self.open_file), 
                               ('📋 粘贴图片', self.paste_image),
                               ('✂️ 截图选区', self.start_screenshot)]:
            btn = QPushButton(text)
            btn.setStyleSheet(btn_style)
            btn.clicked.connect(handler)
            btn_layout.addWidget(btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        # 图片预览区
        preview_frame = QFrame()
        preview_frame.setObjectName('previewFrame')
        preview_layout = QVBoxLayout(preview_frame)
        
        self.preview_label = QLabel('拖拽图片到这里，或点击上方按钮\n支持 Ctrl+V 粘贴')
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setStyleSheet(f'color: {"#888" if not self.dark else "#aaa"}; font-size: 16px; padding: 30px;')
        self.preview_label.setMinimumHeight(200)
        self.preview_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        preview_layout.addWidget(self.preview_label)
        layout.addWidget(preview_frame, stretch=1)
        
        # 结果区
        self.result_label = QLabel('')
        self.result_label.setStyleSheet('''
            QLabel {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #667eea, stop:1 #764ba2);
                color: white; padding: 20px; border-radius: 12px; font-size: 18px;
            }
        ''')
        self.result_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.result_label.setWordWrap(True)
        self.result_label.hide()
        layout.addWidget(self.result_label)
        
        # DPI提示
        self.dpi_label = QLabel(f'当前DPI缩放: {self.dpi_scale_x*100:.0f}% x {self.dpi_scale_y*100:.0f}%')
        self.dpi_label.setStyleSheet(f'color: {"#666" if not self.dark else "#888"}; font-size: 11px;')
        self.dpi_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.dpi_label)
        
        QShortcut(QKeySequence('Ctrl+V'), self, self.paste_image)
        QShortcut(QKeySequence('Ctrl+O'), self, self.open_file)

    def show_calibration(self):
        dialog = CalibrationDialog(self, self.dark)
        dialog.set_scale(self.dpi_scale_x, self.dpi_scale_y)
        dialog.setStyleSheet(get_style(self.dark))
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.dpi_scale_x, self.dpi_scale_y = dialog.get_scale()
            self.settings.setValue('dpi_scale_x', self.dpi_scale_x)
            self.settings.setValue('dpi_scale_y', self.dpi_scale_y)
            self.dpi_label.setText(f'当前DPI缩放: {self.dpi_scale_x*100:.0f}% x {self.dpi_scale_y*100:.0f}%')
    
    def toggle_topmost(self, state):
        if state == Qt.CheckState.Checked.value:
            self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowStaysOnTopHint)
        self.show()
    
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择图片', '', 'Images (*.png *.jpg *.jpeg *.bmp *.gif)')
        if path:
            self.load_image(QPixmap(path))
    
    def paste_image(self):
        clipboard = QApplication.clipboard()
        pixmap = clipboard.pixmap()
        if not pixmap.isNull():
            self.load_image(pixmap)
        else:
            try:
                img = ImageGrab.grabclipboard()
                if img and not isinstance(img, list):
                    buffer = BytesIO()
                    img.convert('RGB').save(buffer, format='PNG')
                    pixmap = QPixmap()
                    pixmap.loadFromData(buffer.getvalue())
                    if not pixmap.isNull():
                        self.load_image(pixmap)
                        return
            except:
                pass
            self.result_label.setText('剪贴板中没有图片')
            self.result_label.show()
    
    def start_screenshot(self):
        self.hide()
        QApplication.processEvents()
        QTimer.singleShot(250, self._do_screenshot)
    
    def _do_screenshot(self):
        screen = QGuiApplication.primaryScreen()
        screenshot = screen.grabWindow(0)
        self.screenshot_overlay = ScreenshotOverlay(
            screenshot, self._on_screenshot_done,
            self.dpi_scale_x, self.dpi_scale_y
        )
        self.screenshot_overlay.show()
    
    def _on_screenshot_done(self, pixmap):
        self.show()
        self.activateWindow()
        if pixmap and not pixmap.isNull():
            self.load_image(pixmap)
    
    def load_image(self, pixmap):
        """加载图片并自动识别"""
        self.current_pixmap = pixmap
        
        # 根据预览区域大小自适应缩放
        label_size = self.preview_label.size()
        max_w = max(label_size.width() - 40, 400)
        max_h = max(label_size.height() - 40, 300)
        
        scaled = pixmap.scaled(max_w, max_h, Qt.AspectRatioMode.KeepAspectRatio,
                               Qt.TransformationMode.SmoothTransformation)
        self.preview_label.setPixmap(scaled)
        
        # 自动识别
        QTimer.singleShot(100, self.analyze)
    
    def on_sensitivity_changed(self, value):
        """敏感度改变"""
        self.sensitivity = value / 10.0
        self.sens_value_label.setText(f'{self.sensitivity:.1f}')
        self.settings.setValue('sensitivity', self.sensitivity)
        
        # 如果有图片，重新分析
        if self.current_pixmap:
            QTimer.singleShot(100, self.analyze)
    
    def analyze(self):
        """分析图片"""
        if not self.current_pixmap:
            return
        
        self.result_label.setText('🔍 识别中...')
        self.result_label.show()
        QApplication.processEvents()
        
        debug_mode = self.debug_cb.isChecked()
        result, error = ColorAnalyzer.analyze(self.current_pixmap, self.sensitivity, debug_mode)
        
        if error:
            self.result_label.setText(f'❌ {error}')
        elif debug_mode and 'debug_colors' in result:
            # 调试模式：显示所有方块颜色
            lines = [f"网格: {result['grid_size']} (调试模式)"]
            for info in result['debug_colors'][:20]:  # 最多显示20个
                marker = '🎯' if any(r['row'] == info['row'] and r['col'] == info['col'] 
                                    for r in result['results']) else '  '
                lines.append(f"{marker} [{info['row']},{info['col']}] {info['color']}")
            if len(result['debug_colors']) > 20:
                lines.append(f"... 还有 {len(result['debug_colors']) - 20} 个方块")
            self.result_label.setText('\n'.join(lines))
        elif result['results']:
            lines = [f"网格: {result['grid_size']}"]
            for r in result['results']:
                lines.append(f"🎯 第 {r['row']} 行，第 {r['col']} 列  ({r['color']})")
            self.result_label.setText('\n'.join(lines))
        else:
            self.result_label.setText(f"网格: {result['grid_size']}\n✅ 所有方块颜色相同")
        
        self.result_label.show()
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                self.load_image(QPixmap(path))
    
    def closeEvent(self, event):
        """关闭时保存窗口状态"""
        self.settings.setValue('window_geometry', self.saveGeometry())
        if self.screenshot_overlay:
            self.screenshot_overlay.close()
        event.accept()
    
    def resizeEvent(self, event):
        """窗口大小改变时重新缩放图片"""
        super().resizeEvent(event)
        if self.current_pixmap and not self.current_pixmap.isNull():
            label_size = self.preview_label.size()
            max_w = max(label_size.width() - 40, 400)
            max_h = max(label_size.height() - 40, 300)
            scaled = self.current_pixmap.scaled(max_w, max_h, 
                                                Qt.AspectRatioMode.KeepAspectRatio,
                                                Qt.TransformationMode.SmoothTransformation)
            self.preview_label.setPixmap(scaled)


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
