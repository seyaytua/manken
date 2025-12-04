import sys
import os
from pathlib import Path
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QListWidget, QLabel, 
                               QFileDialog, QComboBox, QSpinBox, QGroupBox,
                               QMessageBox, QProgressBar, QTabWidget, QLineEdit,
                               QCheckBox, QTextEdit, QSplitter, QDialog, QDialogButtonBox,
                               QFormLayout, QRadioButton, QButtonGroup, QInputDialog,
                               QScrollArea, QGridLayout)
from PySide6.QtCore import Qt, QThread, Signal, QSize
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QPixmap, QImage
import PyPDF2
from pdf2image import convert_from_path
from PIL import Image
import io


class PDFProcessThread(QThread):
    """PDFの処理を別スレッドで実行"""
    progress = Signal(int)
    finished = Signal(bool, str)
    
    def __init__(self, mode, files, output_path, **kwargs):
        super().__init__()
        self.mode = mode
        self.files = files
        self.output_path = output_path
        self.kwargs = kwargs
    
    def run(self):
        try:
            if self.mode == "merge":
                self.merge_pdfs()
            elif self.mode == "convert":
                self.convert_to_images()
            elif self.mode == "split":
                self.split_pdf()
            elif self.mode == "compress":
                self.compress_pdf()
            elif self.mode == "rotate":
                self.rotate_pdf()
            elif self.mode == "extract_pages":
                self.extract_pages()
            self.finished.emit(True, "処理が完了しました！")
        except Exception as e:
            self.finished.emit(False, f"エラーが発生しました: {str(e)}")
    
    def merge_pdfs(self):
        """複数のPDFを1つにまとめる（パスワード付き）"""
        pdf_writer = PyPDF2.PdfWriter()
        total_files = len(self.files)
        
        for idx, file_path in enumerate(self.files):
            try:
                with open(file_path, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)
            except Exception as e:
                raise Exception(f"ファイル '{Path(file_path).name}' の処理中にエラー: {str(e)}")
            
            progress = int((idx + 1) / total_files * 100)
            self.progress.emit(progress)
        
        # パスワード設定
        password = self.kwargs.get('password')
        if password:
            pdf_writer.encrypt(password)
        
        with open(self.output_path, 'wb') as output_file:
            pdf_writer.write(output_file)
    
    def convert_to_images(self):
        """PDFを画像に変換"""
        image_format = self.kwargs.get('image_format', 'PNG')
        dpi = self.kwargs.get('dpi', 200)
        total_files = len(self.files)
        
        for idx, file_path in enumerate(self.files):
            # PDFを画像に変換
            images = convert_from_path(file_path, dpi=dpi)
            
            # ファイル名を生成
            base_name = Path(file_path).stem
            
            for page_num, image in enumerate(images, start=1):
                if len(images) > 1:
                    output_file = f"{self.output_path}/{base_name}_page_{page_num}.{image_format.lower()}"
                else:
                    output_file = f"{self.output_path}/{base_name}.{image_format.lower()}"
                
                # 画像を保存
                image.save(output_file, image_format)
            
            progress = int((idx + 1) / total_files * 100)
            self.progress.emit(progress)
    
    def split_pdf(self):
        """PDFを1ページずつ分割"""
        file_path = self.files[0]
        
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            total_pages = len(pdf_reader.pages)
            
            base_name = Path(file_path).stem
            
            for page_num in range(total_pages):
                pdf_writer = PyPDF2.PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[page_num])
                
                output_file = f"{self.output_path}/{base_name}_page_{page_num + 1}.pdf"
                
                with open(output_file, 'wb') as output:
                    pdf_writer.write(output)
                
                progress = int((page_num + 1) / total_pages * 100)
                self.progress.emit(progress)
    
    def compress_pdf(self):
        """PDFを圧縮"""
        file_path = self.files[0]
        
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            pdf_writer = PyPDF2.PdfWriter()
            
            total_pages = len(pdf_reader.pages)
            
            for page_num, page in enumerate(pdf_reader.pages):
                page.compress_content_streams()
                pdf_writer.add_page(page)
                
                progress = int((page_num + 1) / total_pages * 100)
                self.progress.emit(progress)
            
            with open(self.output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
    
    def rotate_pdf(self):
        """PDFを回転"""
        file_path = self.files[0]
        pages_to_rotate = self.kwargs.get('pages_to_rotate', [])
        angle = self.kwargs.get('angle', 90)
        
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            pdf_writer = PyPDF2.PdfWriter()
            
            total_pages = len(pdf_reader.pages)
            
            for page_num, page in enumerate(pdf_reader.pages):
                if page_num in pages_to_rotate:
                    page.rotate(angle)
                pdf_writer.add_page(page)
                
                progress = int((page_num + 1) / total_pages * 100)
                self.progress.emit(progress)
            
            # パスワード設定
            password = self.kwargs.get('password')
            if password:
                pdf_writer.encrypt(password)
            
            with open(self.output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
    
    def extract_pages(self):
        """特定のページを抽出"""
        file_path = self.files[0]
        pages = self.kwargs.get('pages', [])
        
        with open(file_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            pdf_writer = PyPDF2.PdfWriter()
            
            for idx, page_num in enumerate(pages):
                if 0 <= page_num < len(pdf_reader.pages):
                    pdf_writer.add_page(pdf_reader.pages[page_num])
                
                progress = int((idx + 1) / len(pages) * 100)
                self.progress.emit(progress)
            
            # パスワード設定
            password = self.kwargs.get('password')
            if password:
                pdf_writer.encrypt(password)
            
            with open(self.output_path, 'wb') as output_file:
                pdf_writer.write(output_file)


class PDFPreviewWidget(QWidget):
    """PDFプレビューウィジェット"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.pdf_path = None
        self.pages = []
        self.page_labels = []
        self.selected_pages = set()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 情報ラベル
        self.info_label = QLabel("👆 上のボタンからPDFファイルを読み込んでください")
        self.info_label.setStyleSheet("""
            QLabel {
                background-color: #fff3cd;
                color: #856404;
                padding: 10px;
                border-radius: 5px;
                border: 1px solid #ffeaa7;
            }
        """)
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.info_label)
        
        # スクロールエリア
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # グリッドウィジェット
        self.grid_widget = QWidget()
        self.grid_layout = QGridLayout(self.grid_widget)
        self.grid_layout.setSpacing(10)
        
        scroll.setWidget(self.grid_widget)
        layout.addWidget(scroll)
    
    def load_pdf(self, file_path=None):
        """PDFを読み込んでプレビュー表示"""
        if file_path is None:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "PDFファイルを選択",
                "",
                "PDF Files (*.pdf)"
            )
        
        if not file_path:
            return False
        
        try:
            self.pdf_path = file_path
            self.clear_preview()
            
            # 情報ラベルを更新
            self.info_label.setText(f"📄 読み込み中: {Path(file_path).name}")
            self.info_label.setStyleSheet("""
                QLabel {
                    background-color: #cce5ff;
                    color: #004085;
                    padding: 10px;
                    border-radius: 5px;
                    border: 1px solid #b8daff;
                }
            """)
            QApplication.processEvents()  # UIを更新
            
            # PDFを画像に変換（低解像度でプレビュー）
            images = convert_from_path(file_path, dpi=100)
            self.pages = images
            
            # グリッドに配置（1行に3列）
            for idx, image in enumerate(images):
                # PIL ImageをQPixmapに変換
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                
                qimage = QImage.fromData(img_byte_arr.read())
                pixmap = QPixmap.fromImage(qimage)
                
                # サムネイルサイズに縮小
                pixmap = pixmap.scaled(250, 350, Qt.AspectRatioMode.KeepAspectRatio, 
                                      Qt.TransformationMode.SmoothTransformation)
                
                # ページウィジェット作成
                page_widget = QWidget()
                page_layout = QVBoxLayout(page_widget)
                page_layout.setContentsMargins(5, 5, 5, 5)
                
                # チェックボックス
                checkbox = QCheckBox(f"ページ {idx + 1}")
                checkbox.setStyleSheet("font-weight: bold; font-size: 13px;")
                checkbox.stateChanged.connect(lambda state, p=idx: self.on_page_selected(p, state))
                page_layout.addWidget(checkbox)
                
                # 画像ラベル
                label = QLabel()
                label.setPixmap(pixmap)
                label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                label.setStyleSheet("""
                    QLabel {
                        border: 2px solid #bdc3c7;
                        border-radius: 5px;
                        padding: 5px;
                        background-color: white;
                    }
                """)
                page_layout.addWidget(label)
                
                # グリッドに追加（3列）
                row = idx // 3
                col = idx % 3
                self.grid_layout.addWidget(page_widget, row, col)
                
                self.page_labels.append({'checkbox': checkbox, 'label': label, 'widget': page_widget})
            
            # 情報ラベルを更新
            self.info_label.setText(f"✅ 読み込み完了: {Path(file_path).name} ({len(images)}ページ)")
            self.info_label.setStyleSheet("""
                QLabel {
                    background-color: #d4edda;
                    color: #155724;
                    padding: 10px;
                    border-radius: 5px;
                    border: 1px solid #c3e6cb;
                }
            """)
            
            return True
        except Exception as e:
            self.info_label.setText(f"❌ エラー: {str(e)}")
            self.info_label.setStyleSheet("""
                QLabel {
                    background-color: #f8d7da;
                    color: #721c24;
                    padding: 10px;
                    border-radius: 5px;
                    border: 1px solid #f5c6cb;
                }
            """)
            QMessageBox.critical(self, "エラー", f"PDFの読み込みに失敗しました: {str(e)}")
            return False
    
    def clear_preview(self):
        """プレビューをクリア"""
        self.selected_pages.clear()
        self.page_labels.clear()
        
        # 既存のウィジェットを削除
        for i in reversed(range(self.grid_layout.count())): 
            widget = self.grid_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
    
    def on_page_selected(self, page_num, state):
        """ページが選択された時の処理"""
        if state == Qt.CheckState.Checked.value:
            self.selected_pages.add(page_num)
            # ボーダーを青に
            self.page_labels[page_num]['label'].setStyleSheet("""
                QLabel {
                    border: 3px solid #3498db;
                    border-radius: 5px;
                    padding: 5px;
                    background-color: #e3f2fd;
                }
            """)
        else:
            self.selected_pages.discard(page_num)
            # ボーダーを元に戻す
            self.page_labels[page_num]['label'].setStyleSheet("""
                QLabel {
                    border: 2px solid #bdc3c7;
                    border-radius: 5px;
                    padding: 5px;
                    background-color: white;
                }
            """)
    
    def select_all(self):
        """すべてのページを選択"""
        for item in self.page_labels:
            item['checkbox'].setChecked(True)
    
    def deselect_all(self):
        """すべての選択を解除"""
        for item in self.page_labels:
            item['checkbox'].setChecked(False)
    
    def get_selected_pages(self):
        """選択されたページのリストを取得"""
        return sorted(list(self.selected_pages))
    
    def get_pdf_path(self):
        """現在読み込まれているPDFのパスを取得"""
        return self.pdf_path
    
    def get_total_pages(self):
        """総ページ数を取得"""
        return len(self.pages)


class PDFInfoDialog(QDialog):
    """PDFの情報を表示するダイアログ"""
    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PDF情報")
        self.setGeometry(200, 200, 500, 400)
        
        layout = QVBoxLayout(self)
        
        info_text = QTextEdit()
        info_text.setReadOnly(True)
        
        try:
            with open(file_path, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                
                info = f"ファイル名: {Path(file_path).name}\n"
                info += f"ファイルパス: {file_path}\n"
                info += f"ファイルサイズ: {os.path.getsize(file_path) / 1024:.2f} KB\n"
                info += f"ページ数: {len(pdf_reader.pages)}\n"
                info += f"暗号化: {'はい' if pdf_reader.is_encrypted else 'いいえ'}\n\n"
                
                if pdf_reader.metadata:
                    info += "メタデータ:\n"
                    for key, value in pdf_reader.metadata.items():
                        info += f"  {key}: {value}\n"
                
                info_text.setText(info)
        except Exception as e:
            info_text.setText(f"エラー: {str(e)}")
        
        layout.addWidget(info_text)
        
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(self.accept)
        layout.addWidget(button_box)


class PasswordDialog(QDialog):
    """パスワード設定ダイアログ"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("パスワード設定")
        self.setGeometry(300, 300, 400, 150)
        
        layout = QFormLayout(self)
        
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addRow("パスワード:", self.password_input)
        
        self.confirm_input = QLineEdit()
        self.confirm_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addRow("パスワード確認:", self.confirm_input)
        
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.validate_and_accept)
        button_box.rejected.connect(self.reject)
        layout.addRow(button_box)
    
    def validate_and_accept(self):
        if self.password_input.text() != self.confirm_input.text():
            QMessageBox.warning(self, "エラー", "パスワードが一致しません")
            return
        if self.password_input.text():
            self.accept()
        else:
            QMessageBox.warning(self, "エラー", "パスワードを入力してください")
    
    def get_password(self):
        return self.password_input.text()


class DragDropListWidget(QListWidget):
    """ドラッグ&ドロップ対応のリストウィジェット"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.parent_widget = parent
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls() 
                 if url.toLocalFile().lower().endswith('.pdf')]
        if self.parent_widget and hasattr(self.parent_widget, 'add_files_to_current_tab'):
            self.parent_widget.add_files_to_current_tab(files)


class PDFConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_files = []
        self.process_thread = None
        self.init_ui()
    
    def init_ui(self):
        """UIの初期化"""
        self.setWindowTitle("PDF統合・変換ツール")
        self.setGeometry(100, 100, 1200, 800)
        
        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # タイトル
        title_label = QLabel("📄 PDF統合・変換ツール")
        title_label.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #2c3e50;
            padding: 10px;
            background-color: #ecf0f1;
            border-radius: 5px;
        """)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # タブウィジェット
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # タブ1: PDF統合
        merge_tab = self.create_merge_tab()
        self.tab_widget.addTab(merge_tab, "📚 PDF統合")
        
        # タブ2: PDF→画像変換
        convert_tab = self.create_convert_tab()
        self.tab_widget.addTab(convert_tab, "🖼️ 画像変換")
        
        # タブ3: PDF分割
        split_tab = self.create_split_tab()
        self.tab_widget.addTab(split_tab, "✂️ PDF分割")
        
        # タブ4: PDF圧縮
        compress_tab = self.create_compress_tab()
        self.tab_widget.addTab(compress_tab, "📦 PDF圧縮")
        
        # タブ5: PDF回転（プレビュー付き）
        rotate_tab = self.create_rotate_tab_with_preview()
        self.tab_widget.addTab(rotate_tab, "🔄 PDF回転")
        
        # タブ6: ページ抽出（プレビュー付き）
        extract_tab = self.create_extract_tab_with_preview()
        self.tab_widget.addTab(extract_tab, "📑 ページ抽出")
        
        # プログレスバー
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                text-align: center;
                height: 25px;
            }
            QProgressBar::chunk {
                background-color: #3498db;
            }
        """)
        main_layout.addWidget(self.progress_bar)
        
        # ステータスラベル
        self.status_label = QLabel("準備完了")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("font-size: 14px; padding: 5px;")
        main_layout.addWidget(self.status_label)
    
    def create_merge_tab(self):
        """PDF統合タブの作成"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 説明ラベル
        info_label = QLabel("複数のPDFファイルを1つのPDFファイルに統合します。")
        info_label.setStyleSheet("font-weight: bold; color: #2c3e50; padding: 10px;")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # ファイルリストグループ
        file_group = QGroupBox("PDFファイル一覧")
        file_layout = QVBoxLayout()
        
        # リストウィジェット
        self.merge_file_list = DragDropListWidget(self)
        self.merge_file_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        file_layout.addWidget(self.merge_file_list)
        
        # ボタンレイアウト
        button_layout = QHBoxLayout()
        
        add_button = QPushButton("📁 ファイルを追加")
        add_button.clicked.connect(self.add_files_dialog)
        button_layout.addWidget(add_button)
        
        remove_button = QPushButton("🗑️ 選択を削除")
        remove_button.clicked.connect(self.remove_selected_files)
        button_layout.addWidget(remove_button)
        
        clear_button = QPushButton("🧹 すべてクリア")
        clear_button.clicked.connect(self.clear_all_files)
        button_layout.addWidget(clear_button)
        
        up_button = QPushButton("⬆️ 上へ")
        up_button.clicked.connect(self.move_up)
        button_layout.addWidget(up_button)
        
        down_button = QPushButton("⬇️ 下へ")
        down_button.clicked.connect(self.move_down)
        button_layout.addWidget(down_button)
        
        info_button = QPushButton("ℹ️ 情報")
        info_button.clicked.connect(self.show_pdf_info)
        button_layout.addWidget(info_button)
        
        file_layout.addLayout(button_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # パスワード設定
        password_group = QGroupBox("パスワード設定（オプション）")
        password_layout = QVBoxLayout()
        
        self.merge_password_check = QCheckBox("🔒 作成するPDFにパスワードをかける")
        self.merge_password_check.setToolTip("チェックすると、出力されるPDFファイルにパスワードが設定されます")
        password_layout.addWidget(self.merge_password_check)
        
        password_group.setLayout(password_layout)
        layout.addWidget(password_group)
        
        # 統合実行ボタン
        merge_button = QPushButton("📄 PDFを統合")
        merge_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        merge_button.clicked.connect(self.merge_pdfs)
        layout.addWidget(merge_button)
        
        return tab
    
    def create_convert_tab(self):
        """画像変換タブの作成"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 説明ラベル
        info_label = QLabel("PDFファイルをJPEGまたはPNG画像に変換します")
        info_label.setStyleSheet("font-weight: bold; color: #2c3e50; padding: 10px;")
        layout.addWidget(info_label)
        
        # ファイルリストグループ
        file_group = QGroupBox("PDFファイル一覧")
        file_layout = QVBoxLayout()
        
        # リストウィジェット
        self.convert_file_list = DragDropListWidget(self)
        self.convert_file_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        file_layout.addWidget(self.convert_file_list)
        
        # ボタンレイアウト
        button_layout = QHBoxLayout()
        
        add_button = QPushButton("📁 ファイルを追加")
        add_button.clicked.connect(self.add_files_dialog_convert)
        button_layout.addWidget(add_button)
        
        remove_button = QPushButton("🗑️ 選択を削除")
        remove_button.clicked.connect(self.remove_selected_files_convert)
        button_layout.addWidget(remove_button)
        
        clear_button = QPushButton("🧹 すべてクリア")
        clear_button.clicked.connect(self.clear_all_files_convert)
        button_layout.addWidget(clear_button)
        
        file_layout.addLayout(button_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 設定グループ
        settings_group = QGroupBox("変換設定")
        settings_layout = QHBoxLayout()
        
        # 画像フォーマット選択
        format_label = QLabel("出力形式:")
        settings_layout.addWidget(format_label)
        
        self.format_combo = QComboBox()
        self.format_combo.addItems(["JPEG", "PNG"])
        settings_layout.addWidget(self.format_combo)
        
        # DPI設定
        dpi_label = QLabel("解像度(DPI):")
        settings_layout.addWidget(dpi_label)
        
        self.dpi_spinbox = QSpinBox()
        self.dpi_spinbox.setRange(72, 600)
        self.dpi_spinbox.setValue(200)
        self.dpi_spinbox.setSuffix(" dpi")
        settings_layout.addWidget(self.dpi_spinbox)
        
        settings_layout.addStretch()
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # 変換実行ボタン
        convert_button = QPushButton("🖼️ 画像に変換")
        convert_button.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)
        convert_button.clicked.connect(self.convert_to_images)
        layout.addWidget(convert_button)
        
        return tab
    
    def create_split_tab(self):
        """PDF分割タブの作成"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        info_label = QLabel("PDFファイルを1ページずつ分割します")
        info_label.setStyleSheet("font-weight: bold; color: #2c3e50; padding: 10px;")
        layout.addWidget(info_label)
        
        file_group = QGroupBox("PDFファイル選択")
        file_layout = QVBoxLayout()
        
        self.split_file_list = DragDropListWidget(self)
        file_layout.addWidget(self.split_file_list)
        
        button_layout = QHBoxLayout()
        add_button = QPushButton("📁 ファイルを選択")
        add_button.clicked.connect(lambda: self.add_single_file(self.split_file_list))
        button_layout.addWidget(add_button)
        
        clear_button = QPushButton("🧹 クリア")
        clear_button.clicked.connect(lambda: self.split_file_list.clear())
        button_layout.addWidget(clear_button)
        
        file_layout.addLayout(button_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        split_button = QPushButton("✂️ PDFを分割")
        split_button.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        split_button.clicked.connect(self.split_pdf)
        layout.addWidget(split_button)
        
        return tab
    
    def create_compress_tab(self):
        """PDF圧縮タブの作成"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        info_label = QLabel("PDFファイルを圧縮してファイルサイズを削減します")
        info_label.setStyleSheet("font-weight: bold; color: #2c3e50; padding: 10px;")
        layout.addWidget(info_label)
        
        file_group = QGroupBox("PDFファイル選択")
        file_layout = QVBoxLayout()
        
        self.compress_file_list = DragDropListWidget(self)
        file_layout.addWidget(self.compress_file_list)
        
        button_layout = QHBoxLayout()
        add_button = QPushButton("📁 ファイルを選択")
        add_button.clicked.connect(lambda: self.add_single_file(self.compress_file_list))
        button_layout.addWidget(add_button)
        
        clear_button = QPushButton("🧹 クリア")
        clear_button.clicked.connect(lambda: self.compress_file_list.clear())
        button_layout.addWidget(clear_button)
        
        file_layout.addLayout(button_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        compress_button = QPushButton("📦 PDFを圧縮")
        compress_button.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
        """)
        compress_button.clicked.connect(self.compress_pdf)
        layout.addWidget(compress_button)
        
        return tab
    
    def create_rotate_tab_with_preview(self):
        """PDF回転タブの作成（プレビュー付き）"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        info_label = QLabel("PDFファイルのページを回転します。まずPDFを読み込んで、回転したいページを選択してください。")
        info_label.setStyleSheet("font-weight: bold; color: #2c3e50; padding: 10px;")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # プレビューウィジェットを先に作成
        self.rotate_preview = PDFPreviewWidget()
        
        # ファイル読み込みボタン
        load_button_layout = QHBoxLayout()
        load_pdf_button = QPushButton("📂 PDFファイルを読み込む")
        load_pdf_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        load_pdf_button.clicked.connect(lambda: self.rotate_preview.load_pdf())
        load_button_layout.addWidget(load_pdf_button)
        
        select_all_btn = QPushButton("✅ すべて選択")
        select_all_btn.clicked.connect(self.rotate_preview.select_all)
        load_button_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("❌ すべて解除")
        deselect_all_btn.clicked.connect(self.rotate_preview.deselect_all)
        load_button_layout.addWidget(deselect_all_btn)
        
        load_button_layout.addStretch()
        layout.addLayout(load_button_layout)
        
        # プレビューウィジェットを追加
        layout.addWidget(self.rotate_preview)
        
        # 回転角度選択
        angle_group = QGroupBox("回転角度")
        angle_layout = QHBoxLayout()
        
        self.rotate_angle_group = QButtonGroup()
        
        for angle in [90, 180, 270]:
            radio = QRadioButton(f"{angle}度（時計回り）")
            self.rotate_angle_group.addButton(radio, angle)
            angle_layout.addWidget(radio)
            if angle == 90:
                radio.setChecked(True)
        
        angle_layout.addStretch()
        angle_group.setLayout(angle_layout)
        layout.addWidget(angle_group)
        
        # パスワード設定
        password_group = QGroupBox("パスワード設定（オプション）")
        password_layout = QVBoxLayout()
        
        self.rotate_password_check = QCheckBox("🔒 作成するPDFにパスワードをかける")
        self.rotate_password_check.setToolTip("チェックすると、出力されるPDFファイルにパスワードが設定されます")
        password_layout.addWidget(self.rotate_password_check)
        
        password_group.setLayout(password_layout)
        layout.addWidget(password_group)
        
        rotate_button = QPushButton("🔄 選択したページを回転")
        rotate_button.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        rotate_button.clicked.connect(self.rotate_pdf)
        layout.addWidget(rotate_button)
        
        return tab
    
    def create_extract_tab_with_preview(self):
        """ページ抽出タブの作成（プレビュー付き）"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        info_label = QLabel("PDFから特定のページを抽出して新しいPDFを作成します。まずPDFを読み込んで、抽出したいページを選択してください。")
        info_label.setStyleSheet("font-weight: bold; color: #2c3e50; padding: 10px;")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # プレビューウィジェットを先に作成
        self.extract_preview = PDFPreviewWidget()
        
        # ファイル読み込みボタン
        load_button_layout = QHBoxLayout()
        load_pdf_button = QPushButton("📂 PDFファイルを読み込む")
        load_pdf_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        load_pdf_button.clicked.connect(lambda: self.extract_preview.load_pdf())
        load_button_layout.addWidget(load_pdf_button)
        
        select_all_btn = QPushButton("✅ すべて選択")
        select_all_btn.clicked.connect(self.extract_preview.select_all)
        load_button_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("❌ すべて解除")
        deselect_all_btn.clicked.connect(self.extract_preview.deselect_all)
        load_button_layout.addWidget(deselect_all_btn)
        
        load_button_layout.addStretch()
        layout.addLayout(load_button_layout)
        
        # プレビューウィジェットを追加
        layout.addWidget(self.extract_preview)
        
        # パスワード設定
        password_group = QGroupBox("パスワード設定（オプション）")
        password_layout = QVBoxLayout()
        
        self.extract_password_check = QCheckBox("🔒 作成するPDFにパスワードをかける")
        self.extract_password_check.setToolTip("チェックすると、出力されるPDFファイルにパスワードが設定されます")
        password_layout.addWidget(self.extract_password_check)
        
        password_group.setLayout(password_layout)
        layout.addWidget(password_group)
        
        extract_button = QPushButton("📑 選択したページを抽出")
        extract_button.setStyleSheet("""
            QPushButton {
                background-color: #16a085;
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #138d75;
            }
        """)
        extract_button.clicked.connect(self.extract_pages)
        layout.addWidget(extract_button)
        
        return tab
    
    def add_files_dialog(self):
        """ファイル選択ダイアログを開く（統合用）"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "PDFファイルを選択",
            "",
            "PDF Files (*.pdf)"
        )
        if files:
            self.add_files(files)
    
    def add_files_dialog_convert(self):
        """ファイル選択ダイアログを開く（変換用）"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "PDFファイルを選択",
            "",
            "PDF Files (*.pdf)"
        )
        if files:
            self.add_files_convert(files)
    
    def add_single_file(self, list_widget):
        """単一ファイル選択"""
        file, _ = QFileDialog.getOpenFileName(
            self,
            "PDFファイルを選択",
            "",
            "PDF Files (*.pdf)"
        )
        if file:
            list_widget.clear()
            list_widget.addItem(Path(file).name)
            list_widget.item(0).setData(Qt.ItemDataRole.UserRole, file)
    
    def add_files(self, files):
        """ファイルをリストに追加（統合用）"""
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.merge_file_list.addItem(Path(file).name)
        self.update_status()
    
    def add_files_convert(self, files):
        """ファイルをリストに追加（変換用）"""
        for file in files:
            items = [self.convert_file_list.item(i).data(Qt.ItemDataRole.UserRole) 
                    for i in range(self.convert_file_list.count())]
            if file not in items:
                self.convert_file_list.addItem(Path(file).name)
                self.convert_file_list.item(self.convert_file_list.count() - 1).setData(
                    Qt.ItemDataRole.UserRole, file)
        self.update_status()
    
    def add_files_to_current_tab(self, files):
        """現在のタブに応じてファイルを追加"""
        current_index = self.tab_widget.currentIndex()
        if current_index == 0:  # 統合タブ
            self.add_files(files)
        elif current_index == 1:  # 変換タブ
            self.add_files_convert(files)
        elif current_index == 2 and files:  # 分割タブ
            self.split_file_list.clear()
            self.split_file_list.addItem(Path(files[0]).name)
            self.split_file_list.item(0).setData(Qt.ItemDataRole.UserRole, files[0])
        elif current_index == 3 and files:  # 圧縮タブ
            self.compress_file_list.clear()
            self.compress_file_list.addItem(Path(files[0]).name)
            self.compress_file_list.item(0).setData(Qt.ItemDataRole.UserRole, files[0])
        elif current_index == 4 and files:  # 回転タブ
            self.rotate_preview.load_pdf(files[0])
        elif current_index == 5 and files:  # 抽出タブ
            self.extract_preview.load_pdf(files[0])
    
    def remove_selected_files(self):
        """選択されたファイルを削除（統合用）"""
        for item in self.merge_file_list.selectedItems():
            row = self.merge_file_list.row(item)
            self.merge_file_list.takeItem(row)
            del self.pdf_files[row]
        self.update_status()
    
    def remove_selected_files_convert(self):
        """選択されたファイルを削除（変換用）"""
        for item in self.convert_file_list.selectedItems():
            row = self.convert_file_list.row(item)
            self.convert_file_list.takeItem(row)
        self.update_status()
    
    def clear_all_files(self):
        """すべてのファイルをクリア（統合用）"""
        self.merge_file_list.clear()
        self.pdf_files.clear()
        self.update_status()
    
    def clear_all_files_convert(self):
        """すべてのファイルをクリア（変換用）"""
        self.convert_file_list.clear()
        self.update_status()
    
    def move_up(self):
        """選択されたファイルを上に移動"""
        current_row = self.merge_file_list.currentRow()
        if current_row > 0:
            current_item = self.merge_file_list.takeItem(current_row)
            self.merge_file_list.insertItem(current_row - 1, current_item)
            self.merge_file_list.setCurrentRow(current_row - 1)
            
            self.pdf_files[current_row], self.pdf_files[current_row - 1] = \
                self.pdf_files[current_row - 1], self.pdf_files[current_row]
    
    def move_down(self):
        """選択されたファイルを下に移動"""
        current_row = self.merge_file_list.currentRow()
        if current_row < self.merge_file_list.count() - 1 and current_row >= 0:
            current_item = self.merge_file_list.takeItem(current_row)
            self.merge_file_list.insertItem(current_row + 1, current_item)
            self.merge_file_list.setCurrentRow(current_row + 1)
            
            self.pdf_files[current_row], self.pdf_files[current_row + 1] = \
                self.pdf_files[current_row + 1], self.pdf_files[current_row]
    
    def show_pdf_info(self):
        """PDFの情報を表示"""
        current_item = self.merge_file_list.currentItem()
        if current_item:
            row = self.merge_file_list.row(current_item)
            file_path = self.pdf_files[row]
            dialog = PDFInfoDialog(file_path, self)
            dialog.exec()
        else:
            QMessageBox.warning(self, "警告", "PDFファイルを選択してください")
    
    def merge_pdfs(self):
        """PDFを統合"""
        if not self.pdf_files:
            QMessageBox.warning(self, "警告", "PDFファイルが選択されていません")
            return
        
        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "統合PDFを保存",
            "merged.pdf",
            "PDF Files (*.pdf)"
        )
        
        if output_file:
            kwargs = {}
            
            # 出力PDFのパスワード設定
            if self.merge_password_check.isChecked():
                dialog = PasswordDialog(self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    kwargs['password'] = dialog.get_password()
                else:
                    return
            
            self.start_process("merge", self.pdf_files, output_file, **kwargs)
    
    def convert_to_images(self):
        """PDFを画像に変換"""
        if self.convert_file_list.count() == 0:
            QMessageBox.warning(self, "警告", "PDFファイルが選択されていません")
            return
        
        output_dir = QFileDialog.getExistingDirectory(
            self,
            "画像の保存先フォルダを選択"
        )
        
        if output_dir:
            files = [self.convert_file_list.item(i).data(Qt.ItemDataRole.UserRole) 
                    for i in range(self.convert_file_list.count())]
            
            kwargs = {
                'image_format': self.format_combo.currentText(),
                'dpi': self.dpi_spinbox.value()
            }
            
            self.start_process("convert", files, output_dir, **kwargs)
    
    def split_pdf(self):
        """PDFを分割"""
        if self.split_file_list.count() == 0:
            QMessageBox.warning(self, "警告", "PDFファイルが選択されていません")
            return
        
        output_dir = QFileDialog.getExistingDirectory(
            self,
            "分割PDFの保存先フォルダを選択"
        )
        
        if output_dir:
            file_path = self.split_file_list.item(0).data(Qt.ItemDataRole.UserRole)
            self.start_process("split", [file_path], output_dir)
    
    def compress_pdf(self):
        """PDFを圧縮"""
        if self.compress_file_list.count() == 0:
            QMessageBox.warning(self, "警告", "PDFファイルが選択されていません")
            return
        
        file_path = self.compress_file_list.item(0).data(Qt.ItemDataRole.UserRole)
        base_name = Path(file_path).stem
        
        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "圧縮PDFを保存",
            f"{base_name}_compressed.pdf",
            "PDF Files (*.pdf)"
        )
        
        if output_file:
            self.start_process("compress", [file_path], output_file)
    
    def rotate_pdf(self):
        """PDFを回転（選択されたページのみ）"""
        pdf_path = self.rotate_preview.get_pdf_path()
        if not pdf_path:
            QMessageBox.warning(self, "警告", "PDFファイルを読み込んでください")
            return
        
        selected_pages = self.rotate_preview.get_selected_pages()
        if not selected_pages:
            QMessageBox.warning(self, "警告", "回転するページを選択してください")
            return
        
        base_name = Path(pdf_path).stem
        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "回転PDFを保存",
            f"{base_name}_rotated.pdf",
            "PDF Files (*.pdf)"
        )
        
        if output_file:
            angle = self.rotate_angle_group.checkedId()
            kwargs = {
                'angle': angle,
                'pages_to_rotate': selected_pages
            }
            
            # パスワード設定
            if self.rotate_password_check.isChecked():
                dialog = PasswordDialog(self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    kwargs['password'] = dialog.get_password()
                else:
                    return
            
            self.start_process("rotate", [pdf_path], output_file, **kwargs)
    
    def extract_pages(self):
        """ページを抽出"""
        pdf_path = self.extract_preview.get_pdf_path()
        if not pdf_path:
            QMessageBox.warning(self, "警告", "PDFファイルを読み込んでください")
            return
        
        selected_pages = self.extract_preview.get_selected_pages()
        if not selected_pages:
            QMessageBox.warning(self, "警告", "抽出するページを選択してください")
            return
        
        base_name = Path(pdf_path).stem
        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "抽出PDFを保存",
            f"{base_name}_extracted.pdf",
            "PDF Files (*.pdf)"
        )
        
        if output_file:
            kwargs = {'pages': selected_pages}
            
            # パスワード設定
            if self.extract_password_check.isChecked():
                dialog = PasswordDialog(self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    kwargs['password'] = dialog.get_password()
                else:
                    return
            
            self.start_process("extract_pages", [pdf_path], output_file, **kwargs)
    
    def start_process(self, mode, files, output_path, **kwargs):
        """処理を開始"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("処理中...")
        
        self.process_thread = PDFProcessThread(mode, files, output_path, **kwargs)
        self.process_thread.progress.connect(self.update_progress)
        self.process_thread.finished.connect(self.process_finished)
        self.process_thread.start()
    
    def update_progress(self, value):
        """プログレスバーを更新"""
        self.progress_bar.setValue(value)
    
    def process_finished(self, success, message):
        """処理完了時の処理"""
        self.progress_bar.setVisible(False)
        self.status_label.setText("準備完了")
        
        if success:
            QMessageBox.information(self, "完了", message)
        else:
            QMessageBox.critical(self, "エラー", message)
    
    def update_status(self):
        """ステータスを更新"""
        merge_count = self.merge_file_list.count()
        convert_count = self.convert_file_list.count()
        self.status_label.setText(
            f"📚 統合: {merge_count}件 | 🖼️ 変換: {convert_count}件"
        )


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # アプリケーションのスタイルシート
    app.setStyleSheet("""
        QMainWindow {
            background-color: #ecf0f1;
        }
        QGroupBox {
            font-weight: bold;
            border: 2px solid #bdc3c7;
            border-radius: 5px;
            margin-top: 10px;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px;
        }
        QPushButton {
            padding: 8px;
            border-radius: 4px;
            background-color: #34495e;
            color: white;
        }
        QPushButton:hover {
            background-color: #2c3e50;
        }
        QListWidget {
            border: 2px solid #bdc3c7;
            border-radius: 5px;
            padding: 5px;
        }
        QCheckBox {
            spacing: 5px;
        }
        QCheckBox::indicator {
            width: 18px;
            height: 18px;
        }
    """)
    
    window = PDFConverterApp()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
