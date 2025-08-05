import shutil
import configparser
from kivy.app import App
from kivy.config import Config
from kivy.properties import StringProperty
from kivy.lang import Builder
from kivy.uix.popup import Popup
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.label import Label
from docx import Document

import os
from datetime import datetime

import funcs


Builder.load_file("actcreator.kv")

# resolution
Config.set('graphics', 'width', '800')
Config.set('graphics', 'height', '650')
Config.set('graphics', 'resizable', False)


class ActCreatorRoot(BoxLayout):

    config = configparser.ConfigParser()
    config.read('assets\config.ini', encoding='utf-8')
    manager_name = config['data']['manager_name']
    manager_name_gen = config['data']['manager_name_gen']
    power_of_attorney = config['data']['power_of_attorney']
    employee_name = config['data']['employee_name']
    employee_name_gen = config['data']['employee_name_gen']
    inv_num = config['data']['inv_num']
    laptop_condition = config['data']['laptop_condition']

    output_path = StringProperty(os.getcwd())

    date = StringProperty(datetime.today().strftime('%d.%m.%Y'))

    def show_popup(self, title, message):
        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        label = Label(text=message)
        btn_ok = Button(text='OK', size_hint_y=None, height=40)

        layout.add_widget(label)
        layout.add_widget(btn_ok)

        popup = Popup(title=title, content=layout, size_hint=(0.5, 0.3), auto_dismiss=False)
        btn_ok.bind(on_press=popup.dismiss)
        popup.open()

    def open_folder_chooser(self):
        content = BoxLayout(orientation='vertical', spacing=10)
        chooser = FileChooserListView(path=self.output_path, filters=['*/'], dirselect=True)
        btn_select = Button(text="Выбрать", size_hint_y=None, height=40)
        popup = Popup(title="Выбор папки", content=content, size_hint=(0.9, 0.9))

        def select_folder(instance):
            if chooser.selection:
                self.output_path = chooser.selection[0]
                popup.dismiss()

        btn_select.bind(on_press=select_folder)
        content.add_widget(chooser)
        content.add_widget(btn_select)
        popup.open()

    def generate(self):

        sys_info = funcs.SystemInfo.get_all_system_info()
        print(sys_info)

        date = datetime.now().strftime('%d.%m.%Y')
        date_readable = funcs.format_date_readable(date)
        manager = self.manager_name
        manager_gen = self.manager_name_gen
        proxy = self.power_of_attorney
        employee = self.employee_name
        employee_gen = self.employee_name_gen
        condition = self.laptop_condition
        export_path = self.output_path
        inv_num = self.inv_num

        act_name = f"Акт_{date}_{employee.replace(' ', '_')}.docx"
        act_path = os.path.join(export_path, act_name)

        if not os.path.isdir(export_path):
            self.show_popup("Ошибка", "Невозможно сохранить здесь")
            return

        try:
            shutil.copy('assets/act_template.docx', act_path)
        except Exception as e:
            print("Ошибка копирования:", e)
            self.show_popup("Ошибка", "Невозможно сохранить здесь")
            return

        shutil.copy('assets/act_template.docx', act_path)

        doc = Document(act_path)

        os_ws = sys_info.get('OS')
        cpu_ws = sys_info.get('CPU')
        ram_ws = sys_info.get('RAM')
        ram_type_ws = sys_info.get('RAM_TYPE')

        replacements = {
            '{DATE}': date,
            '{DATE_READABLE}': date_readable,
            '{MANAGER}': manager,
            '{MANAGER_GEN}': manager_gen,
            '{PROXY}': proxy,
            '{EMPLOYEE}': employee,
            '{EMPLOYEE_GEN}': employee_gen,
            '{CONDITION}': condition,
            '{LAPTOP_MODEL}': sys_info.get('model', ''),
            '{SERIAL}': sys_info.get('serial', ''),
            '{OS}': os_ws,
            '{CPU}': cpu_ws,
            '{RAM}': ram_ws,
            '{RAM_TYPE}': ram_type_ws,
            '{DRIVES}': sys_info.get('drives'),
            '{INV_NUM}': inv_num
        }

        print(replacements)

        funcs.SystemInfo.replace_placeholders(doc, replacements)

        doc.save(act_path)
        self.show_popup("Успех", "Акт сгенерирован")
        print(f"✅ Акт успешно сгенерирован: {act_path}")


class ActCreatorApp(App):
    def build(self):
        return ActCreatorRoot()

if __name__ == '__main__':
    ActCreatorApp().run()
