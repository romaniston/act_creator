import shutil
from kivy.app import App
from kivy.config import Config
from kivy.properties import StringProperty
from kivy.lang import Builder
from kivy.uix.popup import Popup
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from docx import Document

import os
from datetime import datetime

import funcs


Builder.load_file("actcreator.kv")

# resolution
Config.set('graphics', 'width', '800')
Config.set('graphics', 'height', '495')
Config.set('graphics', 'resizable', False)


class ActCreatorRoot(BoxLayout):

    output_path = StringProperty("")

    date = StringProperty(datetime.today().strftime('%d.%m.%Y'))
    manager_name = StringProperty("")
    manager_name_gen = StringProperty("")
    power_of_attorney = StringProperty("")
    employee_name = StringProperty("")
    employee_name_gen = StringProperty("")
    inv_num = StringProperty("")
    laptop_condition = StringProperty("Вышеуказанное оборудование на момент его передачи находится в надлежащем"
                                      "состоянии, соответствует предъявляемым к нему техническим требованиям.")

    def generate(self):
        variables = {
            "var_date": self.date,
            "var_manager": self.manager_name,
            "var_manager_gen": self.manager_name_gen,
            "var_attorney": self.power_of_attorney,
            "var_employee": self.employee_name,
            "var_employee_gen": self.employee_name_gen,
            "var_condition": self.laptop_condition
        }
        generate_act_with_system_info(
            output_path=self.output_path,
            variables=variables
        )

    def open_folder_chooser(self):
        content = BoxLayout(orientation='vertical', spacing=10)
        chooser = FileChooserListView(path=".", filters=['*/'], dirselect=True)
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

        sys_info = funcs.get_system_info()

        manager = self.manager_name
        manager_gen = self.manager_name_gen
        proxy = self.power_of_attorney
        employee = self.employee_name
        employee_gen = self.employee_name_gen
        condition = self.laptop_condition
        export_path = self.output_path
        inv_num = self.inv_num

        if not os.path.isdir(export_path):
            print("❌ Указанный путь недействителен.")
            return

        today = datetime.now().strftime('%Y-%m-%d')
        act_name = f"Акт_{today}_{employee.replace(' ', '_')}.docx"
        act_path = os.path.join(export_path, act_name)

        shutil.copy('assets/act_template.docx', act_path)

        doc = Document(act_path)

        replacements = {
            '{DATE}': today,
            '{MANAGER}': manager,
            '{MANAGER_GEN}': manager_gen,
            '{PROXY}': proxy,
            '{EMPLOYEE}': employee,
            '{EMPLOYEE_GEN}': employee_gen,
            '{CONDITION}': condition,
            '{LAPTOP_MODEL}': sys_info.get('model', ''),
            '{SERIAL}': sys_info.get('serial', ''),
            '{OS}': sys_info.get('os', ''),
            '{CPU}': sys_info.get('cpu', ''),
            '{RAM}': sys_info.get('ram', ''),
            '{RAM_TYPE}': sys_info.get('ram_type', ''),
            '{DRIVES}': sys_info.get('drives', ''),
            '{INV_NUM}': sys_info.get('inv_num', '')
        }

        for paragraph in doc.paragraphs:
            for key, val in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(val))

        doc.save(act_path)
        print(f"✅ Акт успешно сгенерирован: {act_path}")


class ActCreatorApp(App):
    def build(self):
        return ActCreatorRoot()


if __name__ == '__main__':
    ActCreatorApp().run()
