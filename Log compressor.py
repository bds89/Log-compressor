import os, time, sys, inspect, re, zipfile, yaml, datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QAction, QLabel,\
                            QLineEdit, QWidget, QTextEdit, QGridLayout, QSlider,\
                            QFileDialog, QComboBox, QCheckBox, QSystemTrayIcon, QMenu, qApp, QVBoxLayout, QHBoxLayout, QStyle
from PyQt5.QtCore import QThread, QObject, Qt, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QFont, QIcon, QIntValidator

class ContinueLoop(Exception): pass
class Window(QMainWindow):
    resized = pyqtSignal() 
    activeWindow = ["self.initTextBoxes()", "self.page.deleteLater()"]
    logText = ""
    lastResize = time.time()
    trayIcon = None
    def __init__(self):
        super().__init__()
        #zip thread
        self.thread = QThread()
        self.zip = Zip()
        self.zip.moveToThread(self.thread)
        self.thread.started.connect(self.zip.run)

        self.w = 600
        self.h = 400
        self.initMainWindow()
        self.initTextBoxes()
        self.showMainWindow()
        #signals from zip
        self.zip.isRun.connect(self.changeStatusBar)
        self.zip.log.connect(self.writeLog)
        if CONFIG["autoStart"]:
            self.btn.click()
        self.resized.connect(self.changeSize)
        self.trayIcon = QSystemTrayIcon(self)
        pixmapi = getattr(QStyle, "SP_TitleBarMenuButton")
        self.trayIcon.setIcon(self.style().standardIcon(pixmapi))
        show_action = QAction("Развернуть", self)
        quit_action = QAction("Выход", self)
        hide_action = QAction("Свернуть", self)
        show_action.triggered.connect(self.show)
        hide_action.triggered.connect(self.hide)
        quit_action.triggered.connect(qApp.quit)
        tray_menu = QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addAction(quit_action)

        self.trayIcon.setContextMenu(tray_menu)
        self.trayIcon.show()


    def closeEvent(self, event):
        if CONFIG["minimizeToTray"]:
            event.ignore()
            self.hide()
            self.trayIcon.showMessage(
                "Log compressor",
                "Приложение все еще работает",
                QSystemTrayIcon.Information,
                2000
            )
    def resizeEvent(self, event):
        self.resized.emit()
        return super(Window, self).resizeEvent(event)
    def changeSize(self):
        if time.time() - self.lastResize > 0.1:
            self.w = self.geometry().width()
            self.h = self.geometry().height()
            eval(self.activeWindow[1])
            eval(self.activeWindow[0])
            self.lastResize = time.time()

    def initMainWindow(self):
        #menu
        helpAction = QAction('Справка', self)
        helpAction.triggered.connect(self.dialogHelp)
        aboutAction = QAction('О программе', self)
        aboutAction.triggered.connect(self.dialogAbout)
        SettingsAction = QAction('Настройки', self)
        SettingsAction.triggered.connect(self.dialogSettings)
        
        menubar = self.menuBar()
        aboutMenu = menubar.addMenu('Меню')
        aboutMenu.addAction(SettingsAction)
        aboutMenu.addAction(aboutAction)
        aboutMenu.addAction(helpAction)
        #statusBar
        self.statusBar().showMessage('Остановлен')
        self.foldersList = CONFIG["folders"]
        pixmapi = getattr(QStyle, "SP_TitleBarMenuButton")
        self.setWindowIcon(self.style().standardIcon(pixmapi))
        
    def dialogAbout(self):
        self.activeWindow = ["self.dialogAbout()", "self.menuPage.deleteLater()"]
        self.page.hide()
        if hasattr(self, 'menuPage'):
            self.menuPage.hide()
        self.menuPage = QWidget(self)
        hbox = QHBoxLayout()
        vbox = QVBoxLayout()
        text1 = QLabel('Log compressor')
        text1.setFont(QFont('Arial', 16))
        text2 = QLabel('Версия: 1.0')
        text3 = QLabel('Автор: bds89')
        text4 = QLabel('Сайт: <a href="https://github.com/bds89/">github.com/bds89</a>')
        text4.setOpenExternalLinks(True)
        btn = QPushButton('', self)
        pixmapi = getattr(QStyle, "SP_ArrowBack")
        btn.setIcon(self.style().standardIcon(pixmapi))
        btn.resize(btn.sizeHint())
        hbox.addStretch(1)
        hbox.addWidget(text1)
        hbox.addStretch(1)
        vbox.addLayout(hbox)
        vbox.addWidget(text2)
        vbox.addWidget(text3)
        vbox.addWidget(text4)
        vbox.addStretch(1)
        vbox.addWidget(btn)
        self.menuPage.setLayout(vbox)
        self.menuPage.move(0,20)
        self.menuPage.resize(self.w, (self.h-20))
        self.menuPage.show()


        btn.clicked.connect(self.btnBack)

    def btnBack(self):
        self.menuPage.hide()
        self.page.show()
        self.activeWindow = ["self.initTextBoxes()", "self.page.deleteLater()"]

    def dialogHelp(self):
        self.activeWindow = ["self.dialogHelp()", "self.menuPage.deleteLater()"]
        self.page.hide()
        if hasattr(self, 'menuPage'):
            self.menuPage.hide()
        self.menuPage = QWidget(self)
        vbox = QVBoxLayout()
        hbox = QHBoxLayout()
        text1 = QLabel('Справка')
        text1.setFont(QFont('Arial', 16))
        text2 = QLabel('1. Приложение сжимает и удаляет все файлы (кроме файла с последней датой изменения) с указанными расширениями из указанных каталогов\n'+\
                        '2. Расширения указываются через запятую (например: png, pdf, zip)\n'+\
                        '3. Если расширения не указаны, приложение будет сжимать и удалять все файлы из указанных директорий'+\
                            ' (кроме файла с последней датой изменения, самой себя, конфигурационного файла и *.zip файлов)\n'+\
                        '4. Если каталоги не указаны, приложение будет работать в директории своего расположения\n'+\
                        '5. Самый сильный метод сжатия: ZIP_LZMA, но он так-же требует большее количество времени')
        text2.setWordWrap(True)
        btn = QPushButton('', self)
        pixmapi = getattr(QStyle, "SP_ArrowBack")
        btn.setIcon(self.style().standardIcon(pixmapi))
        btn.resize(btn.sizeHint())
        btn.clicked.connect(self.btnBack)
        hbox.addStretch(1)
        hbox.addWidget(text1)
        hbox.addStretch(1)
        vbox.addLayout(hbox)
        vbox.addWidget(text2)
        vbox.addStretch(1)
        vbox.addWidget(btn)
        self.menuPage.setLayout(vbox)
        self.menuPage.move(0,20)
        self.menuPage.resize(self.w, (self.h-20))
        self.menuPage.show()


    def dialogSettings(self):
        self.activeWindow = ["self.dialogSettings()", "self.menuPage.deleteLater()"]
        self.page.hide()
        if hasattr(self, 'menuPage'):
            self.menuPage.hide()
        self.menuPage = QWidget(self)
        vbox = QVBoxLayout()
        hbox1 = QHBoxLayout()
        hbox2 = QHBoxLayout()
        text1 = QLabel('Настройки')
        text1.setFont(QFont('Arial', 16))
        #autostart
        cb1 = QCheckBox('Пуск при старте', self)
        if CONFIG["autoStart"]:
            cb1.toggle()
        cb1.stateChanged.connect(self.checkBox1)
        #minimize to tray
        cb2 = QCheckBox('Сворачивать в трей', self)
        if CONFIG["minimizeToTray"]:
            cb2.toggle()
        cb2.stateChanged.connect(self.checkBox2)
        logLabel = QLabel('Размер лога (строк):')
        logInput = QLineEdit(self)
        logInput.setFixedWidth(60)
        logInput.setValidator(QIntValidator())
        logInput.setText(str(CONFIG["logLength"]))
        logInput.editingFinished.connect(self.settingsLog)
        hbox2.addWidget(logLabel)
        hbox2.addWidget(logInput)
        hbox2.addStretch(1)
        hbox1.addStretch(1)
        hbox1.addWidget(text1)
        hbox1.addStretch(1)
        vbox.addLayout(hbox1)
        vbox.addWidget(cb1)
        vbox.addWidget(cb2)
        vbox.addLayout(hbox2)
        vbox.addStretch(1)
        btn = QPushButton('', self)
        pixmapi = getattr(QStyle, "SP_ArrowBack")
        btn.setIcon(self.style().standardIcon(pixmapi))
        btn.resize(btn.sizeHint())
        vbox.addWidget(btn)
        btn.clicked.connect(self.btnBack)
        self.menuPage.setLayout(vbox)
        self.menuPage.move(0,20)
        self.menuPage.resize(self.w, self.h-20)
        self.menuPage.show()


    def initTextBoxes(self):
        self.activeWindow = ["self.initTextBoxes()", "self.page.deleteLater()"]
        self.forDisable = []
        self.page = QWidget(self)
        r=0
        #mainBoxes
        grid = QGridLayout()
        grid.setSpacing(5)
        #Расширения файлов
        extension = QLabel('Расширения файлов:')
        extensionEdit = QLineEdit()
        extensionEdit.setToolTip('Разделяя запятыми, например: exe, pdf, png')
        extensionEdit.setText(CONFIG["extensions"])
        extensionEdit.editingFinished.connect(self.changeExtension)
        self.forDisable.append(extensionEdit)
        r+=1
        grid.addWidget(extension, r, 0)
        grid.addWidget(extensionEdit, r, 1, 1, 5)
        #Каталоги с логами
        catalogs = QLabel('Каталоги с логами:')
        r+=1
        grid.addWidget(catalogs, r, 0)
        if len(self.foldersList) == 0:
            selfDir = QLabel('Из папки размещения скрипта')
            grid.addWidget(selfDir, r, 1, 1, 4)

        for i, cat in enumerate(self.foldersList):
            catalog = QLineEdit()
            catalog.setText(cat)
            catalog.setObjectName(str(i))
            catalog.editingFinished.connect(self.changeFolders)
            self.forDisable.append(catalog)
            grid.addWidget(catalog, r, 1, 1, 4)

            delFolder = QPushButton('', self)
            pixmapi = getattr(QStyle, "SP_DialogCancelButton")
            delFolder.setIcon(self.style().standardIcon(pixmapi))
            delFolder.setObjectName(cat)
            delFolder.clicked.connect(self.delFolder)
            self.forDisable.append(delFolder)
            grid.addWidget(delFolder, r, 5, 1, 1)
            r+=1

        openFolder = QPushButton('', self)
        pixmapi = getattr(QStyle, "SP_FileDialogNewFolder")
        openFolder.setIcon(self.style().standardIcon(pixmapi))
        openFolder.clicked.connect(self.openFolderDialog)
        self.forDisable.append(openFolder)
        grid.addWidget(openFolder, r, 5)
        #метод сжатия
        metod = QLabel('Метод сжатия:')
        comboMetod = QComboBox(self)
        comboMetod.addItems(["ZIP_DEFLATED", "ZIP_BZIP2",
                        "ZIP_LZMA"])
        comboMetod.setCurrentText(CONFIG["compressMethod"])
        comboMetod.activated[str].connect(self.changeCombo)
        self.forDisable.append(comboMetod)
        r+=1
        grid.addWidget(metod, r, 0)
        grid.addWidget(comboMetod, r, 1, 1, 3)
        #степень сжатия
        if (comboMetod.currentText() != "ZIP_LZMA"):
            self.compressionLabel = QLabel('     '+str(CONFIG["compression"]))
            self.compressionLabel.setFont(QFont('Arial', 12))
            compression = QLabel('Степень сжатия:')
            compression_slider = QSlider(Qt.Horizontal, self)
            compression_slider.setRange(1, 9)
            compression_slider.setTracking(True)
            compression_slider.valueChanged.connect(self.changeCompression)
            self.forDisable.append(compression_slider)
            compression_slider.setValue(CONFIG["compression"])

            r+=1
            grid.addWidget(compression, r, 0)
            grid.addWidget(compression_slider, r, 1, 1, 4)
            grid.addWidget(self.compressionLabel, r, 5, 1, 1)
        #время опроса
        self.intervalLabel = QLabel('     1')
        self.intervalLabel.setFont(QFont('Arial', 12))
        interval = QLabel('Интервал проверки (мин):')
        interval_slider = QSlider(Qt.Horizontal, self)
        interval_slider.setRange(1, 60)
        interval_slider.setTracking(True)
        interval_slider.valueChanged.connect(self.changeInterval)
        self.forDisable.append(interval_slider)
        interval_slider.setValue(CONFIG["interval_min"])

        r+=1
        grid.addWidget(interval, r, 0)
        grid.addWidget(interval_slider, r, 1, 1, 4)
        grid.addWidget(self.intervalLabel, r, 5, 1, 1)

        #Лог
        r+=1
        log = QLabel('Лог:')
        self.logBox = QTextEdit()
        self.logBox.setText(self.logText)
        grid.addWidget(log, r, 1)
        #logbtn
        btnLog = QPushButton('', self)
        pixmapi = getattr(QStyle, "SP_DialogResetButton")
        btnLog.setIcon(self.style().standardIcon(pixmapi))
        grid.addWidget(btnLog, r, 5)
        btnLog.clicked.connect(self.btnClearLog)

        r+=1
        grid.addWidget(self.logBox, r, 1, 1, 5)

        #btn
        self.btn = QPushButton('Пуск', self)
        pixmapi = getattr(QStyle, "SP_MediaPlay")
        self.btn.setIcon(self.style().standardIcon(pixmapi))
        grid.addWidget(self.btn, r, 0)
        self.btn.clicked.connect(self.btnClick)

        #if script runing hide gui
        if self.zip.isRunBool:
            for widget in self.forDisable:
                widget.setEnabled(False)
            self.btn.setText('Стоп')
            pixmapi = getattr(QStyle, "SP_MediaStop")
            self.btn.setIcon(self.style().standardIcon(pixmapi))

        self.page.setLayout(grid)
        self.page.move(0,20)
        self.page.resize(self.w, self.h-20)
        self.page.show()


    def showMainWindow(self):
        self.setGeometry(300, 300, self.w, self.h)
        self.resize(self.w, self.h)
        self.setWindowTitle("Log compressor")
        self.setMinimumSize(500, 300)
        self.show()

    def changeExtension(self):
        source = self.sender()
        CONFIG["extensions"] = source.text()
        save_config(CONFIG)
    def settingsLog(self):
        source = self.sender()
        CONFIG["logLength"] = int(source.text())
        save_config(CONFIG)

    def changeCombo(self, value):
        CONFIG["compressMethod"] = value
        self.page.deleteLater()
        self.initTextBoxes()
        save_config(CONFIG)

    def changeCompression(self, value):
        text = "     "+str(value)
        CONFIG["compression"] = value
        self.compressionLabel.setText(text)
        save_config(CONFIG)

    def changeInterval(self, value):
        text = "     "+str(value)
        self.intervalLabel.setText(text)
        CONFIG["interval_min"] = value
        save_config(CONFIG)

    def openFolderDialog(self):
        fname = QFileDialog.getExistingDirectory(self)
        if fname:
            self.foldersList.append(fname)
            self.page.deleteLater()
            self.initTextBoxes()
            CONFIG["folders"] = self.foldersList
            save_config(CONFIG)

    def delFolder(self):
        source = self.sender()
        self.foldersList.remove(source.objectName())
        self.page.deleteLater()
        self.initTextBoxes()
        CONFIG["folders"] = self.foldersList
        save_config(CONFIG)

    def changeFolders(self):
        source = self.sender()
        self.foldersList[int(source.objectName())] = source.text()
        CONFIG["folders"] = self.foldersList
        save_config(CONFIG)

    def btnClick(self):
        source = self.sender()
        if source.text() == 'Пуск':
            self.thread.start()
            self.btn.setText('Стоп')
            pixmapi = getattr(QStyle, "SP_MediaStop")
            self.btn.setIcon(self.style().standardIcon(pixmapi))

        else:
            self.btn.setEnabled(False)
            self.btn.setText('Останавливаем...')
            self.zip.isRunBool = False
            self.thread.exit()
            
    def checkBox1(self, state):
        if state == Qt.Checked:
            CONFIG["autoStart"] = True
        else:
            CONFIG["autoStart"] = False
        save_config(CONFIG)
    def checkBox2(self, state):
        if state == Qt.Checked:
            CONFIG["minimizeToTray"] = True
        else:
            CONFIG["minimizeToTray"] = False
        save_config(CONFIG)

    def changeStatusBar(self, status):
        self.statusBar().showMessage(status)
        if status == "Выполняется":
            for widget in self.forDisable:
                widget.setEnabled(False)
        else:
            for widget in self.forDisable:
                widget.setEnabled(True)
                self.btn.setEnabled(True)
                self.btn.setText('Пуск')
                pixmapi = getattr(QStyle, "SP_MediaPlay")
                self.btn.setIcon(self.style().standardIcon(pixmapi))

    def writeLog(self, string):
        self.logText += string + "\n"
        logList = self.logText.split("\n")
        if len(logList) > CONFIG["logLength"]+1:
            logList = logList[len(logList)-CONFIG["logLength"]-1:]
            self.logText = "\n".join(logList)
        self.logBox.setPlainText(self.logText)
    def btnClearLog(self):
        self.logText = ""
        self.logBox.setPlainText(self.logText)

class Zip(QObject):
    isRun = pyqtSignal(str)
    log = pyqtSignal(str)
    isRunBool = False
    def run(self):
        self.isRunBool = True
        self.isRun.emit("Выполняется")
        while self.isRunBool:
            try:
                lastLoop = datetime.datetime.now()
                observed_dirs = CONFIG["folders"]
                if not observed_dirs: observed_dirs = [work_dir]
                extensions = CONFIG["extensions"]
                if extensions: extensions = extensions.split(",")
                for dir in observed_dirs:
                    if not os.path.exists(dir):
                        self.log.emit(f"{datetime.datetime.now()} Нет такого каталога: {dir}")
                        continue
                    file_list = os.listdir(dir)
                    file_dict = {}
                    #выберем файлы
                    for filename in file_list:
                        #not for folders
                        if not os.path.isfile((os.path.join(dir, filename))): continue
                        #not for self
                        if  SCRIPT_NAME == filename or filename == "config.yaml": continue
                        #not for zip
                        if re.search(f"(\.zip)", filename[-4:]): continue
                        if extensions:
                            for ext in extensions:
                                if re.search(f"(\.{ext.strip()})", filename[-len(ext.strip())-1:]):
                                    file_dict[filename] = os.path.getmtime(os.path.join(dir, filename))
                        else: file_dict[filename] = os.path.getmtime(os.path.join(dir, filename))
                    #отсорируем список файлов по дате создания
                    sorted_tuples = sorted(file_dict.items(), key=lambda item: item[1], reverse=True)
                    sorted_file_dict = {k: v for k, v in sorted_tuples}

                    for name in list(sorted_file_dict.keys())[1:]:
                        if re.findall(r"(.+)\.", name):
                            name_cleared = re.findall(r"(.+)\.", name)[0]
                            add_zip = os.path.join(dir, name_cleared+'.zip')
                        try:
                            with zipfile.ZipFile(add_zip, mode='w', \
                                                compression=eval("zipfile."+CONFIG["compressMethod"]), compresslevel=CONFIG["compression"]) as zf:
                                add_file = os.path.join(dir, name)
                                start_time = datetime.datetime.now()
                                self.log.emit(f"{start_time} Архивирование файла: {add_file}")
                                zf.write(os.path.join(dir, name), arcname=name)
                                zf.close()
                                self.log.emit(f"{datetime.datetime.now()} Архивирование закончено за {self.chop_microseconds(datetime.datetime.now() - start_time)}")
                                os.remove(add_file)
                                self.log.emit(f"{datetime.datetime.now()} Файл: {add_file} удален")
                        except: 
                            window.logBox.setPlainText(f"{datetime.datetime.now()} Ошибка при архивировании файла: {add_file}. Исходный файл остается без изменений")
                            continue

                while(datetime.datetime.now()-lastLoop < datetime.timedelta(minutes=CONFIG["interval_min"])):
                    time.sleep(2)
                    if not self.isRunBool: raise ContinueLoop()
            except ContinueLoop: break
        self.isRun.emit("Остановлен")

    def chop_microseconds(self, delta):
        return str(delta - datetime.timedelta(microseconds=delta.microseconds))

def check_config(config):
    if not "extensions" in config:
        config.update({"extensions":""})
    if not "folders" in config:
        config.update({"folders":[]})
    if not "compression" in config:
        config.update({"compression":6})
    if not "interval_min" in config:
        config.update({"interval_min":1})
    if not "compressMethod" in config:
        config.update({"compressMethod":"ZIP_DEFLATED"})
    if not "autoStart" in config:
        config.update({"autoStart":False})   
    if not "minimizeToTray" in config:
        config.update({"minimizeToTray":False})   
    if not "logLength" in config:
        config.update({"logLength":100})  
        
    return config

def save_config(config):
    with open(CONFIG_PATCH, "w") as f:
        f.write(yaml.dump(CONFIG, sort_keys=False))

def get_script_dir(follow_symlinks=True):
    if getattr(sys, 'frozen', False):
        path = os.path.abspath(sys.executable)
    else:
        path = inspect.getabsfile(get_script_dir)
    if follow_symlinks:
        path = os.path.realpath(path)
    return os.path.dirname(path)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    work_dir = get_script_dir()
    CONFIG_PATCH = os.path.join(work_dir, "config.yaml")
    if not os.path.exists(CONFIG_PATCH):
        with open(CONFIG_PATCH, mode='w', encoding='utf-8') as f:
            CONFIG = {}
            CONFIG = check_config(CONFIG)
    else:
        with open(CONFIG_PATCH, encoding='utf-8') as f:
            CONFIG = yaml.load(f.read(), Loader=yaml.FullLoader)
            if not CONFIG: CONFIG = {}
            CONFIG = check_config(CONFIG)

    SCRIPT_NAME = os.path.basename(__file__)
    window = Window()
    sys.exit(app.exec_())
