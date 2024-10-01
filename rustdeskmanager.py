# Copyright (c) 2024 Roman Boyarintsev
# All rights reserved.
#
# This software is licensed under the Python Software Foundation License (PSF License) and GNU General Public License (GPL v3).
# See LICENSE_PSF.txt and LICENSE_GPLv3.txt for details.
import sys
import os
import toml
import subprocess
import pickle
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QFileDialog, QListWidget, QLabel,\
    QTextEdit, QTreeWidget, QTreeWidgetItem, QInputDialog, QGroupBox, QMessageBox, QCheckBox, QAbstractItemView, QTreeWidgetItemIterator,\
    QFrame
from PyQt6.QtCore import Qt, QProcess
from PyQt6.QtGui import QIcon

class MyTreeWidget(QTreeWidget):
    def __init__(self):
        super().__init__()

class MyPushButton(QPushButton):
    def __init__(self):
        super().__init__()
    
        self.setStyleSheet("QPushButton:hover { background-color: #ccc; }")

class RustDeskManager(QWidget):
    def __init__(self):
        super().__init__()

        self.work_folder_path = ""
        self.rustdesk_path = ""
        self.ids = []
        self.init_ui()

    def showEvent(self, event):
        super().showEvent(event)
        self.item_value_update()


    def init_ui(self):
        self.setWindowTitle("rustdeskmanager")

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # There is group of id_tree, it buttons
        tree_layout = QVBoxLayout()
        tree_button_layout = QHBoxLayout()

        # There is group of id_tree, ids_list
        horizontal_layout = QHBoxLayout()
        horizontal_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # There is local layout for details
        details_layout = QVBoxLayout()
        
        # There is group of IDs_label, IDs_list, run_button
        IDs_layout = QVBoxLayout()

        # Create a settings frame
        settings_frame = QGroupBox("Настройки")
        settings_layout = QVBoxLayout()
        settings_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        # Work folder path
        work_folder_layout = QHBoxLayout()
        self.work_folder_label = QLabel("Путь к конфигурационным файлам Rustdesk:")
        self.work_folder_input = QLineEdit()
        self.work_folder_input.setReadOnly(True)
        self.work_folder_input.setToolTip(f"Укажите папку, где содержатся конфигурационные файлы *.toml rustdesk.\nОбычно находятся в папке пользователя AppData\\Roaming\\RustDesk\\config\\peers")
        self.work_folder_input.setFixedSize(400, 20)
        #self.work_folder_input.setMaximumWidth(400)
        self.work_folder_button = MyPushButton()
        self.work_folder_button.setText("Выбрать")
        self.work_folder_button.setFixedWidth(80)
        self.work_folder_button.clicked.connect(self.browse_work_folder)
        work_folder_layout.addWidget(self.work_folder_input)
        work_folder_layout.addWidget(self.work_folder_button)
        work_folder_layout.addStretch()
        settings_layout.addWidget(self.work_folder_label)
        settings_layout.addLayout(work_folder_layout)

        # Rustdesk path
        rustdesk_layout = QHBoxLayout()
        self.rustdesk_label = QLabel("Путь к исполняемому файлу Rustdesk:")
        self.rustdesk_input = QLineEdit()
        self.rustdesk_input.setReadOnly(True)
        self.rustdesk_input.setToolTip(f"Укажите экземпляр rustdesk.exe для запуска.\nВозможно указать не установленный в систему экземпляр,\nесли он на данный момент времени не является только установщиком.\nДолжен содержать в имени файла rustdesk")
        self.rustdesk_input.setFixedSize(400, 20)
        #self.rustdesk_input.setMaximumWidth(400)
        self.rustdesk_button = MyPushButton()
        self.rustdesk_button.setText("Выбрать")
        self.rustdesk_button.setFixedWidth(80)
        self.rustdesk_button.clicked.connect(self.browse_rustdesk)
        self.rustdesk_run_button = MyPushButton()
        self.rustdesk_run_button.setText("Запустить")
        self.rustdesk_run_button.setFixedWidth(80)
        self.rustdesk_run_button.setToolTip("Запустить основное приложение Rustdesk")
        self.rustdesk_run_button.clicked.connect(self.rustdesk_run_button_clicked)
        rustdesk_layout.addWidget(self.rustdesk_input)
        rustdesk_layout.addWidget(self.rustdesk_button)
        rustdesk_layout.addWidget(self.rustdesk_run_button)
        rustdesk_layout.addStretch()
        settings_layout.addWidget(self.rustdesk_label)
        settings_layout.addLayout(rustdesk_layout)

        # Add settings layout to settings frame
        settings_frame.setLayout(settings_layout)
        layout.addWidget(settings_frame)

        # IDs list
        self.ids_label = QLabel("Rustdesk IDs (*.toml):")
        self.ids_list = QListWidget()
        self.ids_list.setMaximumWidth(150)
        self.ids_list.setDragEnabled(True)
        self.ids_list.setToolTip(f"Можно выделить элемент и перетащить в структуру.\nМожно выделить несколько элементов и перетащить в структуру.\nКнопка Rustdesk сейчас работает только для этого списка!")
        self.ids_list.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
        self.ids_list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.ids_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.ids_list.currentRowChanged.connect(self.ids_list_selection_changed)
        self.ids_list.clicked.connect(self.ids_list_selection_changed)
        self.ids_list.itemDoubleClicked.connect(self.run_rustdesk)

        # Run button
        self.run_button = MyPushButton()
        self.run_button.setText("Rustdesk...")
        self.run_button.setToolTip("Запустить rustdesk.exe для выбранного ID (rustdesk.exe --connect ID)")
        self.run_button.clicked.connect(self.run_rustdesk)
        self.run_button.setDefault(True)

        # Details
        self.details_label = QLabel("Описание:")

        # IDs Details
        self.details_text = QTextEdit(self)
        self.details_text.setReadOnly(True)
        self.details_text.setMaximumHeight(100)
        self.details_text.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.details_text.setStyleSheet("background-color: #f0f0f0;")

        # Group tree
        self.id_tree = MyTreeWidget()
        self.id_tree.setHeaderLabels(["Структура", "Название", "Пользователь", "Компьютер", "Система"])
        self.id_tree.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.id_tree.setDragEnabled(True)
        self.id_tree.setDropIndicatorShown(True)
        self.id_tree.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.id_tree.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.id_tree.setColumnWidth(0, 200)
        self.id_tree.itemDoubleClicked.connect(self.run_rustdesk)
        self.id_tree.currentItemChanged.connect(self.tree_selection_changed)
        self.id_tree.clicked.connect(self.tree_selection_changed)

        # Create checkbox for expanding/collapsing tree
        create_checkbox = QCheckBox("Expand all")
        create_checkbox.setChecked(False)
        create_checkbox.setText("Раскрыть все")
        create_checkbox.stateChanged.connect(self.expand_all_checkbox_changed)

        # Add buttons to create, rename, move, and delete groups
        create_button = MyPushButton()
        create_button.setText("Создать")
        create_button.setToolTip("Создать новую группу или элемент")
        create_button.clicked.connect(self.create_group)
        create_button.setMaximumWidth(100)
        
        rename_button = MyPushButton()
        rename_button.setText("Переименовать")
        rename_button.setToolTip("Переименовать группу или элемент")
        rename_button.clicked.connect(self.rename_group)
        rename_button.setMaximumWidth(100)
        
        delete_button = MyPushButton()
        delete_button.setText("Удалить")
        delete_button.setToolTip("Удалить группу или элемент")
        delete_button.clicked.connect(self.delete_group)
        delete_button.setMaximumWidth(100)

        update_button = MyPushButton()
        update_button.setText("Обновить")
        update_button.setToolTip("Обновить название, пользователь, компьютер, система")
        update_button.clicked.connect(self.item_value_update)
        update_button.setMaximumWidth(100)
        
        # Add buttons to tree layout
        tree_button_layout.addWidget(create_button)
        tree_button_layout.addSpacing(1)
        tree_button_layout.addWidget(rename_button)
        tree_button_layout.addSpacing(1)
        tree_button_layout.addWidget(delete_button)
        tree_button_layout.addSpacing(1)
        tree_button_layout.addWidget(update_button)
        tree_button_layout.addStretch()

        # Add tree layout to tree frame
        tree_layout.addWidget(create_checkbox)
        tree_layout.addWidget(self.id_tree)
        tree_layout.addLayout(tree_button_layout)

        # Add IDs widgets to IDs layout
        IDs_layout.addWidget(self.ids_label)
        IDs_layout.addWidget(self.ids_list)
        IDs_layout.addWidget(self.run_button)

        # Add layouts to horizontal layout
        horizontal_layout.addLayout(IDs_layout)
        horizontal_layout.addLayout(tree_layout)

        horizontal_line = QFrame()
        horizontal_line.setFrameShape(QFrame.Shape.HLine)
        horizontal_line.setFrameShadow(QFrame.Shadow.Sunken)

        layout.addLayout(horizontal_layout)
        layout.addSpacing(15)
        layout.addWidget(horizontal_line)
        layout.addWidget(self.details_label)
        layout.addWidget(self.details_text)

        self.load_config()

        self.setLayout(layout)



    ########For QTreeWidget Start###############
    def expand_all_checkbox_changed(self, state):
        if state == Qt.CheckState.Checked.value:
            self.id_tree.expandAll()
        else:
            self.id_tree.collapseAll()

    def create_group(self):
        selected_item = self.id_tree.currentItem()
        if selected_item:
            group_name, ok = QInputDialog.getText(self, "Создание группы/элемента", "Введите имя группы/элемента:")
            if ok:
                group = QTreeWidgetItem([group_name])
                group.setFlags(group.flags() | Qt.ItemFlag.ItemIsSelectable)
                selected_item.addChild(group)
        else:
            group_name, ok = QInputDialog.getText(self, "Создание группы/элемента", "Введите имя группы/элемента:")
            if ok:
                group = QTreeWidgetItem([group_name])
                group.setFlags(group.flags() | Qt.ItemFlag.ItemIsSelectable)
                self.id_tree.addTopLevelItem(group)

    def rename_group(self):
        selected_item = self.id_tree.currentItem()
        if selected_item:
            new_name, ok = QInputDialog.getText(self, "Переименование группы/элемента", "Введите новое имя группы/элемента:", text=selected_item.text(0))
            if ok:
                selected_item.setText(0, new_name)

    def delete_group(self):
        selected_item = self.id_tree.currentItem()
        if selected_item:
            confirm = QMessageBox.question(self, "Удалить группу", f"Вы уверены, что хотите удалить {selected_item.text(0)}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if confirm == QMessageBox.StandardButton.Yes:
                parent = selected_item.parent()
                parent_count = self.id_tree.topLevelItemCount()
                if parent is None:
                    # remove the item from the tree
                    if parent_count > 1:
                        self.id_tree.currentItem().removeChild(selected_item)
                    else:
                        QMessageBox.warning(self, "Удаление группы", "Нельзя удалить единственную группу верхнего уровня.")
                        pass
                else:
                    for child in range(selected_item.childCount()):
                        child_item = selected_item.child(child)
                        parent.addChild(child_item)
                    parent.removeChild(selected_item)

    def find_item(self, name, parent=None):
        if parent is None:
            parent = self.main_group
        for child in range(parent.childCount()):
            child_item = parent.child(child)
            if child_item.text(0) == name:
                return child_item
            child_item = self.find_item(name, child_item)
            if child_item:
                return child_item
        return None
    
    def item_value_update(self):
        iterator = QTreeWidgetItemIterator(self.id_tree)
        while iterator.value():
            item = iterator.value()
            self.id_tree.setCurrentItem(item)
            iterator += 1

        self.id_tree.clearSelection()

        return True
#########For QTreeWidget End###############

    def rustdesk_run_button_clicked(self):
        try:
            # Run rustdesk
            command = f'{os.path.normpath(self.rustdesk_path)}'
            subprocess.Popen(command, shell=True)
        except FileNotFoundError:
            QMessageBox.information(self, "Информация", "Путь к Rustdesk.exe не задан.")
    
    def load_config(self):
        try:
            with open("config.toml", "r", encoding="utf-8") as f:
                config = toml.load(f)
                if "paths" in config:
                    if "work_folder_path" in config["paths"]:
                        self.work_folder_path = config["paths"]["work_folder_path"]
                        if self.work_folder_path != "":
                            self.work_folder_input.setText(os.path.normpath(self.work_folder_path))
                    if "rustdesk_path" in config["paths"]:
                        self.rustdesk_path = config["paths"]["rustdesk_path"]
                        if self.rustdesk_path != "":
                            self.rustdesk_input.setText(os.path.normpath(self.rustdesk_path))
                if "width" in config["WindowSize"] and "height" in config["WindowSize"]:
                    self.resize(config["WindowSize"]["width"], config["WindowSize"]["height"])
                self.load_ids()
            
        except FileNotFoundError:
            QMessageBox.information(self, "Первый запуск", "При начале работы с программой укажите, пожалуйста, пути Rustdesk в настройках!")
            pass

        try:
            with open("config.dat", "rb") as f:
                config = pickle.load(f)
                if "TreeStructure" in config:
                    self.load_tree_structure(config["TreeStructure"])
        except FileNotFoundError:
            pass

    def load_tree_structure(self, structure):
        def create_item(item, parent=None):
            item_widget = QTreeWidgetItem([item["text"]])
            if parent is None:
                parent = self.id_tree.addTopLevelItem(item_widget)
            else:
                parent.addChild(item_widget)
            for child in item.get("children", []):
                create_item(child, item_widget)
        for item in structure:
            create_item(item)

    def browse_work_folder(self):
        current_path = self.work_folder_input.text()
        folder_path = QFileDialog.getExistingDirectory(self, "Выберите папку RustDesk с toml конфигурационными файлами.",current_path)
        if folder_path:
            self.work_folder_path = folder_path
            self.work_folder_input.setText(os.path.normpath(folder_path))
            self.save_config()
            self.load_ids()

    def browse_rustdesk(self):
        current_path = self.rustdesk_input.text()
        rustdesk_path = QFileDialog.getOpenFileName(self,"Выберите rustdesk.exe",current_path,"Файлы rustdesk (rustdesk*.exe)")[0]
        if rustdesk_path:
            self.rustdesk_path = rustdesk_path
            self.rustdesk_input.setText(os.path.normpath(rustdesk_path))
            self.save_config()

    def load_ids(self):
        self.ids = []
        for filename in os.listdir(self.work_folder_path):
            if filename.endswith(".toml"):
                self.ids.append(os.path.splitext(filename)[0])
        self.ids_list.clear()
        self.ids_list.addItems(self.ids)
        self.ids_list.setCurrentRow(0)
        self.ids_label.setToolTip(f"Всего {len(self.ids_list)} файлов")

    def tree_selection_changed(self):
        selected_id = self.id_tree.currentItem().text(0)
        
        if self.ids_list.findItems(selected_id, Qt.MatchFlag.MatchFixedString):
            details = self.set_details_text(selected_id)
            if details is not None:
                self.id_tree.currentItem().setText(1, details["alias"])
                self.id_tree.currentItem().setText(2, details["username"])
                self.id_tree.currentItem().setText(3, details["hostname"])
                self.id_tree.currentItem().setText(4, details["platform"])

        else:
            pass

    def ids_list_selection_changed(self):
        try:
            selected_id = self.ids_list.currentItem().text()
            self.set_details_text(selected_id)
        except AttributeError:
            pass
    def set_details_text(self,selected_id):
        try:
            filename = os.path.join(self.work_folder_path, selected_id) + ".toml"
            with open(filename, "r", encoding="utf-8") as f:
                config = toml.load(f)
                alias = config.get("options", {}).get("alias")
                username = config.get("info", {}).get("username")
                hostname = config.get("info", {}).get("hostname")
                platform = config.get("info", {}).get("platform")
                details = f"Alias: {alias}\nUsername: {username}\nHostname: {hostname}\nPlatform: {platform}\n"
                self.details_text.setText(details)
                # Внимание! Здесь details перезаписывается
                details = {
                    "alias": config.get("options", {}).get("alias"),
                    "username": config.get("info", {}).get("username"),
                    "hostname": config.get("info", {}).get("hostname"),
                    "platform": config.get("info", {}).get("platform")
                }
            return details
        except FileNotFoundError:
            QMessageBox.warning(self, "Ошибка", f"Файл {filename} не найден.")
            return

    def save_tree_structure(self):
        structure = []
        def traverse(item):
            children = []
            for i in range(item.childCount()):
                child = item.child(i)
                children.append({
                    "text": child.text(0),
                    "children": traverse(child)
                })
            return children
        for i in range(self.id_tree.topLevelItemCount()):
            top_level_item = self.id_tree.topLevelItem(i)
            structure.append({
                "text": top_level_item.text(0),
                "children": traverse(top_level_item)
            })
        return structure
    
    def save_config(self):
        config = {
            "paths": {
                "work_folder_path": self.work_folder_path,
                "rustdesk_path": self.rustdesk_path
            },
            "WindowSize": {
                "width": self.width(),
                "height": self.height()
            }
        }
        with open("config.toml", "w", encoding="utf-8") as f:
            toml.dump(config, f)
        
        config = {
            "TreeStructure": self.save_tree_structure()
        }
        with open("config.dat", "wb") as f:
            pickle.dump(config, f)

    def run_rustdesk(self):
        #if self.ids == []:
        #    QMessageBox.warning(self, "Внимание!", "Список ID пуст.")
        if self.rustdesk_path !="":
            # Get sender. It can be ids_list or id_tree
            sender = self.sender()
            if sender == self.ids_list or sender == self.run_button:
                if self.ids != []:
                    selected_id = self.ids_list.currentItem().text()
                else:
                    #QMessageBox.warning(self, "Внимание!", "Список ID пуст.")
                    return
            elif sender == self.id_tree:
                if self.id_tree.currentItem().childCount() == 0:
                    selected_id = self.id_tree.currentItem().text(0)
                    if not selected_id.isdigit() or len(selected_id) < 7:
                        QMessageBox.warning(self, "Ошибка", "Некорректный ID")
                        return
                else:
                    return
            else:
                return
            
            # Run rustdesk with selected id
            command = f'"{os.path.normpath(self.rustdesk_path)}" --connect {selected_id}'
            subprocess.Popen(command, shell=True)


        elif self.rustdesk_path == "":
            QMessageBox.information(self, "Информация", "Путь к Rustdesk.exe не задан.")

    def closeEvent(self, event):
        self.save_config()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RustDeskManager()
    window.setWindowIcon(QIcon("rustdeskmanager.ico"))
    window.show()
    sys.exit(app.exec())