import subprocess
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox
import getpass
import json


def extract_cpu_info(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    cpu_info = {}
    for subnode in root.iter('SubNode'):
        for property_node in subnode.iter('Property'):
            entry = property_node.find('Entry').text
            description = property_node.find('Description').text
            if entry == 'CPU Brand Name':
                cpu_info['CPU Brand Name'] = description

    return cpu_info


def extract_monitor_info(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    monitor_info = []
    for subnode in root.iter('SubNode'):
        for property_node in subnode.iter('Property'):
            entry = property_node.find('Entry').text
            description = property_node.find('Description').text
            if entry == 'Monitor Name':
                monitor_info.append(description)

    return monitor_info


def extract_gpu_info(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    gpu_info = {}
    for subnode in root.iter('SubNode'):
        for property_node in subnode.iter('Property'):
            entry = property_node.find('Entry').text
            description = property_node.find('Description').text
            if entry == 'Video Chipset':
                gpu_info['Video Chipset'] = description

    return gpu_info


def extract_ram_info(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    ram_info = {}
    for memory_node in root.iter('MEMORY'):
        node_name = memory_node.find('NodeName').text.strip()
        if node_name == 'Memory':
            for property_node in memory_node.iter('Property'):
                entry = property_node.find('Entry').text
                description = property_node.find('Description').text
                if entry == 'Total Memory Size':
                    ram_info['Total Memory Size'] = description

    return ram_info


def extract_motherboard_info(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    motherboard_info = {}
    for motherboard_node in root.iter('MOBO'):
        node_name = motherboard_node.find('NodeName').text.strip()
        if node_name == 'Motherboard':
            for property_node in motherboard_node.iter('Property'):
                entry = property_node.find('Entry').text
                description = property_node.find('Description').text
                if entry == 'Motherboard Model':
                    motherboard_info['Motherboard Model'] = description

    return motherboard_info


def extract_memory_info(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    memory_info = {
        'Drive Capacity': [],
        'Media Rotation Rate': []
    }

    for subnode in root.iter('SubNode'):
        for property_node in subnode.iter('Property'):
            entry = property_node.find('Entry').text
            description = property_node.find('Description').text
            if entry == 'Drive Capacity':
                memory_info['Drive Capacity'].append(description)
            elif entry == 'Media Rotation Rate':
                memory_info['Media Rotation Rate'].append(description)

    return memory_info


def get_domain_info():
    try:
        output = subprocess.check_output(['wmic', 'computersystem', 'get', 'domain'], universal_newlines=True)
        lines = output.splitlines()
        for line in lines:
            if line.strip() and not line.startswith('Domain'):
                domain_info = line.strip()
                return domain_info
    except subprocess.CalledProcessError:
        pass

    return None


def load_template(window, fields):
    try:
        with open('template.json', 'r') as file:
            template = json.load(file)
    except FileNotFoundError:
        return

    for field_name in template:
        field_frame = tk.Frame(window)
        field_frame.pack(anchor='w')

        field_label = tk.Label(field_frame, text=field_name + ':')
        field_label.pack(side='top')

        field_var = tk.StringVar()
        field_entry = tk.Entry(field_frame, textvariable=field_var, state='normal', width=30)
        field_entry.pack(side='left')

        delete_button = tk.Button(field_frame, text='-', command=lambda: delete_custom_field(fields, field_frame))
        delete_button.configure(bg='#a60000', cursor='hand2', fg='#f0f0f0', font=('Arial', 12, 'bold'), relief='flat')
        delete_button.pack(side='left')

        fields.append(field_frame)


def add_custom_field(window, fields):
    field_name = tk.simpledialog.askstring('Custom Field', 'Enter field name:')
    if field_name:
        field_frame = tk.Frame(window)
        field_frame.pack(anchor='w')

        field_label = tk.Label(field_frame, text=field_name + ':')
        field_label.pack(side='top')

        field_var = tk.StringVar()
        field_entry = tk.Entry(field_frame, textvariable=field_var, state='normal', width=30)
        field_entry.pack(side='left')

        delete_button = tk.Button(field_frame, text='-', command=lambda: delete_custom_field(fields, field_frame))
        delete_button.configure(bg='#a60000', cursor='hand2', fg='#f0f0f0', font=('Arial', 12, 'bold'), relief='flat')
        delete_button.pack(side='left')

        fields.append(field_frame)

        save_template(fields)


def save_template(fields):
    template = []
    for field_frame in fields:
        field_name = field_frame.winfo_children()[0]['text'][:-1]
        template.append(field_name)

    with open('template.json', 'w') as file:
        json.dump(template, file)


def delete_custom_field(fields, field_frame):
    result = messagebox.askquestion('Delete Field', 'Are you sure you want to delete this field?', icon='warning')
    if result == 'yes':
        field_label = field_frame.winfo_children()[0]
        field_name = field_label['text'][:-1]
        field_frame.destroy()
        fields.remove(field_frame)
        update_template_file(fields, field_name)

def update_template_file(fields, deleted_field_name):
    template = []
    for field_frame in fields:
        field_label = field_frame.winfo_children()[0]
        field_name = field_label['text'][:-1]
        template.append(field_name)

    with open('template.json', 'r') as file:
        template_data = json.load(file)

    template_data.remove(deleted_field_name)

    with open('template.json', 'w') as file:
        json.dump(template_data, file)


def select_xml_file():
    file_path = filedialog.askopenfilename(filetypes=[('XML files', '*.xml')])
    if file_path:
        cpu_info = extract_cpu_info(file_path)
        gpu_info = extract_gpu_info(file_path)
        monitor_info = extract_monitor_info(file_path)
        ram_info = extract_ram_info(file_path)
        motherboard_info = extract_motherboard_info(file_path)
        memory_info = extract_memory_info(file_path)
        domain_info = get_domain_info()

        create_info_window(cpu_info, gpu_info, monitor_info, ram_info, motherboard_info, domain_info,
                           memory_info)


def save():
    pass


def create_info_window(cpu_info, gpu_info, monitor_info, ram_info, motherboard_info, domain_info, memory_info):
    window = tk.Toplevel()
    window.title('System Information')
    window.geometry('200x900')
    window.resizable(False, True)

    # Set the alignment to the left
    window.grid_propagate(False)

    # Custom fields
    custom_fields = []
    custom_fields_frame = tk.Frame(window)
    custom_fields_frame.pack(anchor='center')

    add_button = tk.Button(custom_fields_frame, text='+', command=lambda: add_custom_field(window, custom_fields))
    add_button.configure(bg='#008a00', cursor='hand2', fg='#f0f0f0', font=('Arial', 12, 'bold'), relief='flat')
    add_button.pack(side='left', padx='10')
    
    add_button = tk.Button(custom_fields_frame, text='Save', command=save())
    add_button.configure(bg='#008a00', cursor='hand2', fg='#f0f0f0', font=('Arial', 12, 'bold'), relief='flat')
    add_button.pack(side='left', padx='10')

    for field_frame in custom_fields:
        field_frame.pack(anchor='w')

    # CPU
    tk.Label(window, text='Processor:').pack(anchor='w')
    processor_name_var = tk.StringVar()
    processor_name_var.set(cpu_info.get('CPU Brand Name'))
    processor_name_entry = tk.Entry(window, textvariable=processor_name_var, state='normal', width=30)
    processor_name_entry.pack(anchor='w')

    # GPU
    tk.Label(window, text='Graphics card:').pack(anchor='w')
    video_controller_var = tk.StringVar()
    video_controller_var.set(gpu_info.get('Video Chipset'))
    video_controller_entry = tk.Entry(window, textvariable=video_controller_var, state='normal', width=30)
    video_controller_entry.pack(anchor='w')

    # Monitor
    for i, monitor_name in enumerate(monitor_info):
        tk.Label(window, text=f'Monitor {i + 1}:').pack(anchor='w')
        monitor_var = tk.StringVar()
        monitor_var.set(monitor_name)
        monitor_entry = tk.Entry(window, textvariable=monitor_var, state='normal', width=30)
        monitor_entry.pack(anchor='w')

    # RAM memory
    tk.Label(window, text='RAM:').pack(anchor='w')
    ram_memory_var = tk.StringVar()
    ram_memory_var.set(ram_info.get('Total Memory Size'))
    ram_memory_entry = tk.Entry(window, textvariable=ram_memory_var, state='normal', width=30)
    ram_memory_entry.pack(anchor='w')

    # Motherboard
    tk.Label(window, text='Motherboard:').pack(anchor='w')
    motherboard_var = tk.StringVar()
    motherboard_var.set(motherboard_info.get('Motherboard Model'))
    motherboard_entry = tk.Entry(window, textvariable=motherboard_var, state='normal', width=30)
    motherboard_entry.pack(anchor='w')

    # Domain
    tk.Label(window, text='Domain:').pack(anchor='w')
    domain_var = tk.StringVar()
    domain_var.set(domain_info)
    domain_entry = tk.Entry(window, textvariable=domain_var, state='normal', width=30)
    domain_entry.pack(anchor='w')

    # Memory size
    memory_sizes = memory_info.get('Drive Capacity', [])[:2]
    memory_types = memory_info.get('Media Rotation Rate', [])[:2]

    for i, (size, mem_type) in enumerate(zip(memory_sizes, memory_types)):
        tk.Label(window, text=f'Memory {i + 1} capacity:').pack(anchor='w')
        capacity_var = tk.StringVar()
        capacity_var.set(size)
        capacity_entry = tk.Entry(window, textvariable=capacity_var, state='normal', width=30)
        capacity_entry.pack(anchor='w')

        tk.Label(window, text=f'Memory {i + 1} type:').pack(anchor='w')
        type_var = tk.StringVar()
        type_var.set(mem_type)
        type_entry = tk.Entry(window, textvariable=type_var, state='normal', width=30)
        type_entry.pack(anchor='w')

    # Username
    tk.Label(window, text='Username:').pack(anchor='w')
    username_var = tk.StringVar()
    username_var.set(getpass.getuser())
    username_entry = tk.Entry(window, textvariable=username_var, state='normal', width=30)
    username_entry.pack(anchor='w')

    load_template(window, custom_fields)

root = tk.Tk()
root.withdraw()

select_xml_file()
root.mainloop()
