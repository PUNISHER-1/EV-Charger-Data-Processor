import pythoncom
import win32com.client
import random
from docxtpl import DocxTemplate
import os
import tempfile
from docx2pdf import convert
from PyQt5.QtCore import QThread, pyqtSignal


class DataThread(QThread):
    progress = pyqtSignal(str)

    def __init__(self, extracted_data):
        super().__init__()
        self.extracted_data = extracted_data

    def run(self):
        process_extracted_data(extracted_data=self.extracted_data, thread=self)

    def execute(self):
        self.start()


def process_extracted_data(extracted_data, thread):
    print(extracted_data)
    # local list
    he531008_list = [] # dynamic 120
    he518741_list = [] # 30 local
    he518847_list = [] # 60 slim 
    he518662_list = [] # 120 1.5g
    # infi list
    he518518_list = [] # 30 dual gun
    he518671_list = [] # 60 infi
    he518675_list = [] # 120 infi
    he518986_list = [] # 180 infi
    he518695_list = [] # 240 infi
    he518114_list = [] # 120 single infi 


    for model_data in extracted_data:
        model_name = model_data["Model"]

        # local

        if model_name == "HE531008":
            for serial_number, details in model_data["Serial Numbers"].items():
                he531008_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he531008(he531008_list, thread)

        if model_name == "HE518662":
            for serial_number, details in model_data["Serial Numbers"].items():
                he518662_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he518662(he518662_list, thread)

        if model_name == "HE518741":
            for serial_number, details in model_data["Serial Numbers"].items():
                he518741_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he518741(he518741_list, thread)

        if model_name == "HE518847":
            for serial_number, details in model_data["Serial Numbers"].items():
                he518847_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he518847(he518847_list, thread)

        # infi

        if model_name == "HE518518":
            for serial_number, details in model_data["Serial Numbers"].items():
                he518518_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he518518(he518518_list, thread)

        if model_name == "HE518671":
            for serial_number, details in model_data["Serial Numbers"].items():
                he518671_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he518671(he518671_list, thread)

        if model_name == "HE518675":
            for serial_number, details in model_data["Serial Numbers"].items():
                he518675_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he518675(he518675_list, thread)

        if model_name == "HE518986":
            for serial_number, details in model_data["Serial Numbers"].items():
                he518986_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he518986(he518986_list, thread)

        if model_name == "HE518695":
            for serial_number, details in model_data["Serial Numbers"].items():
                he518695_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he518695(he518695_list, thread)

        if model_name == "HE518114":
            for serial_number, details in model_data["Serial Numbers"].items():
                he518114_list.append({
                    "Serial No": serial_number,
                    **details
                })
            process_he518114(he518114_list, thread)

# local

def process_he531008(data, thread):
    print("Processing HE531008 data and generating PDF reports...")
    thread.progress.emit("Processing HE531008 data and generating PDF reports...") 
    template_path = 'template docx/dyn 120 local.docx' 
    output_folder = 'DYn 120 kw_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': item.get('Gun B DCEM', ''),
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }
        voltage_1 = (random.randint(4000, 5999))/10
        voltage_2 = (random.randint(5000, 6999))/10
        voltage_3 = (random.randint(5000, 8999))/10
        eff_1 = (random.randint(220, 299))/10
        eff_2 = (random.randint(500, 599))/10
        eff_3 = (random.randint(700, 899))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)
        curr_2 = round(((eff_2*1000) / voltage_2),1)
        curr_3 = round(((eff_3*1000) / voltage_3),1)
        
        voltage_b1 = (random.randint(4000, 5999))/10
        voltage_b2 = (random.randint(5000, 6999))/10
        voltage_b3 = (random.randint(5000, 8999))/10
        eff_b1 = (random.randint(780, 899))/10
        eff_b2 = (random.randint(500, 599))/10
        eff_b3 = (random.randint(220, 299))/10
        curr_b1 = round(((eff_b1*1000) / voltage_b1),1)
        curr_b2 = round(((eff_b2*1000) / voltage_b2),1)
        curr_b3 = round(((eff_b3*1000) / voltage_b3),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(1165, 1199))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['voltage_2'] = voltage_2
        context['voltage_3'] = voltage_3
        context['eff_1'] = eff_1
        context['eff_2'] = eff_2
        context['eff_3'] = eff_3
        context['curr_1'] = curr_1
        context['curr_2'] = curr_2
        context['curr_3'] = curr_3

        context['voltage_b1'] = voltage_b1
        context['voltage_b2'] = voltage_b2
        context['voltage_b3'] = voltage_b3
        context['eff_b1'] = eff_b1
        context['eff_b2'] = eff_b2
        context['eff_b3'] = eff_b3
        context['curr_b1'] = curr_b1
        context['curr_b2'] = curr_b2
        context['curr_b3'] = curr_b3

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)

def process_he518741(data, thread):
    print("Processing HE518741 data and generating PDF reports...")
    thread.progress.emit("Processing HE518741 data and generating PDF reports...") 
    template_path = 'template docx/30 Single local.docx' 
    output_folder = 'output_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': item.get('Gun B DCEM', ''),
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }

        voltage_1 = (random.randint(3000, 4999))/10
        eff_1 = (random.randint(220, 299))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(270, 299))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['eff_1'] = eff_1
        context['curr_1'] = curr_1

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)

def process_he518847(data, thread):
    print("Processing HE518847 data and generating PDF reports...")
    thread.progress.emit("Processing HE518847 data and generating PDF reports...") 
    template_path = 'template docx/local ctrl.docx' 
    output_folder = 'output_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': item.get('Gun B DCEM', ''),
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }

        voltage_1 = (random.randint(4000, 5999))/10
        eff_1 = (random.randint(220, 299))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)

        voltage_b1 = (random.randint(4000, 5999))/10
        eff_b1 = (random.randint(230, 299))/10
        curr_b1 = round(((eff_b1*1000) / voltage_b1),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(570, 599))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['eff_1'] = eff_1
        context['curr_1'] = curr_1

        context['voltage_b1'] = voltage_b1
        context['eff_b1'] = eff_b1
        context['curr_b1'] = curr_b1

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)

def process_he518662(data, thread):
    print("Processing HE518662 data and generating PDF reports...")
    thread.progress.emit("Processing HE518662 data and generating PDF reports...") 
    template_path = 'template docx/local ctrl.docx' 
    output_folder = 'output_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': item.get('Gun B DCEM', ''),
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }

        voltage_1 = (random.randint(3000, 4999))/10
        eff_1 = (random.randint(400, 599))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)

        voltage_b1 = (random.randint(4000, 5999))/10
        eff_b1 = (random.randint(430, 599))/10
        curr_b1 = round(((eff_b1*1000) / voltage_b1),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(1180, 1199))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['eff_1'] = eff_1
        context['curr_1'] = curr_1

        context['voltage_b1'] = voltage_b1
        context['eff_b1'] = eff_b1
        context['curr_b1'] = curr_b1

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)            


# infi

def process_he518671(data, thread):
    print("Processing HE518671 data and generating PDF reports...")
    thread.progress.emit("Processing HE518671 data and generating PDF reports...") 
    template_path = 'template docx/infi controller.docx' 
    output_folder = 'output_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': item.get('Gun B DCEM', ''),
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }

        voltage_1 = (random.randint(4000, 5999))/10
        eff_1 = (random.randint(220, 299))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)

        voltage_b1 = (random.randint(4000, 5999))/10
        eff_b1 = (random.randint(230, 299))/10
        curr_b1 = round(((eff_b1*1000) / voltage_b1),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(570, 599))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['eff_1'] = eff_1
        context['curr_1'] = curr_1

        context['voltage_b1'] = voltage_b1
        context['eff_b1'] = eff_b1
        context['curr_b1'] = curr_b1

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)


def process_he518675(data, thread):
    print("Processing HE518675 data and generating PDF reports...")
    thread.progress.emit("Processing HE518675 data and generating PDF reports...") 
    template_path = 'template docx/infi controller.docx' 
    output_folder = 'output_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': item.get('Gun B DCEM', ''),
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }

        voltage_1 = (random.randint(4000, 5999))/10
        eff_1 = (random.randint(400, 599))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)

        voltage_b1 = (random.randint(4000, 5999))/10
        eff_b1 = (random.randint(450, 599))/10
        curr_b1 = round(((eff_b1*1000) / voltage_b1),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(1180, 1199))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['eff_1'] = eff_1
        context['curr_1'] = curr_1

        context['voltage_b1'] = voltage_b1
        context['eff_b1'] = eff_b1
        context['curr_b1'] = curr_b1

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)


def process_he518986(data, thread):
    print("Processing HE518986 data and generating PDF reports...")
    thread.progress.emit("Processing HE518986 data and generating PDF reports...") 
    template_path = 'template docx/infi controller.docx' 
    output_folder = 'output_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': item.get('Gun B DCEM', ''),
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }

        voltage_1 = (random.randint(4000, 5999))/10
        eff_1 = (random.randint(799, 899))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)

        voltage_b1 = (random.randint(4000, 5999))/10
        eff_b1 = (random.randint(750, 899))/10
        curr_b1 = round(((eff_b1*1000) / voltage_b1),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(1780, 1799))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['eff_1'] = eff_1
        context['curr_1'] = curr_1

        context['voltage_b1'] = voltage_b1
        context['eff_b1'] = eff_b1
        context['curr_b1'] = curr_b1

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)


def process_he518695(data, thread):
    print("Processing HE518695 data and generating PDF reports...")
    thread.progress.emit("Processing HE518695 data and generating PDF reports...") 
    template_path = 'template docx/infi controller.docx' 
    output_folder = 'output_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': item.get('Gun B DCEM', ''),
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }

        voltage_1 = (random.randint(4000, 5999))/10
        eff_1 = (random.randint(999, 1199))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)

        voltage_b1 = (random.randint(4000, 5999))/10
        eff_b1 = (random.randint(1095, 1199))/10
        curr_b1 = round(((eff_b1*1000) / voltage_b1),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(2200, 2399))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['eff_1'] = eff_1
        context['curr_1'] = curr_1

        context['voltage_b1'] = voltage_b1
        context['eff_b1'] = eff_b1
        context['curr_b1'] = curr_b1

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)


        os.remove(temp_docx_path)

def process_he518518(data, thread):
    print("Processing HE518518 data and generating PDF reports...")
    thread.progress.emit("Processing HE518518 data and generating PDF reports...") 
    template_path = 'template docx/infi controller.docx' 
    output_folder = 'output_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': item.get('Gun B DCEM', ''),
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }

        voltage_1 = (random.randint(4000, 4999))/10
        eff_1 = (random.randint(100, 149))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)

        voltage_b1 = (random.randint(4000, 4999))/10
        eff_b1 = (random.randint(100, 149))/10
        curr_b1 = round(((eff_b1*1000) / voltage_b1),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(269, 299))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['eff_1'] = eff_1
        context['curr_1'] = curr_1

        context['voltage_b1'] = voltage_b1
        context['eff_b1'] = eff_b1
        context['curr_b1'] = curr_b1

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)


def process_he518114(data, thread):
    print("Processing HE518114 data and generating PDF reports...")
    thread.progress.emit("Processing HE518114 data and generating PDF reports...") 
    template_path = 'template docx/infi controller.docx' 
    output_folder = 'output_reports'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for idx, item in enumerate(data):
        doc = DocxTemplate(template_path)

        context = {
            'Charger_Serial_No': item.get('Charger Serial No', ''),
            'Gun_A_DCEM': item.get('Gun A DCEM', ''),
            'Gun_B_DCEM': "NA",
            'Gun_C_EM': item.get('Gun C EM', ''),
            'Upper_sw_ver': item.get('Upper sw ver.', ''),
            'Pilot_Cont_sw': item.get('Pilot Cont sw', ''),
            'Tested_Date': item.get('Tested Date', ''),
            'Tested_By': item.get('Tested By', '')
        }

        voltage_1 = (random.randint(4000, 5999))/10
        eff_1 = (random.randint(999, 1199))/10
        curr_1 = round(((eff_1*1000) / voltage_1),1)

        eff_out = (random.randint(957, 967))/10
        out_pwr = (random.randint(1179, 1199))/10
        inp_pwr = round(((out_pwr/ eff_out)* 100), 1)

        context['voltage_1'] = voltage_1
        context['eff_1'] = eff_1
        context['curr_1'] = curr_1

        context['voltage_b1'] = "NA"
        context['eff_b1'] = "NA"
        context['curr_b1'] = "NA"

        context['eff_out'] = eff_out
        context['out_pwr'] = out_pwr
        context['inp_pwr'] = inp_pwr

        doc.render(context)

        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")

            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                temp_docx_path = temp_docx.name
                doc.save(temp_docx_path)

            pdf_file_path = os.path.join(output_folder, f'{item["Charger Serial No"]}.pdf')
            convert(temp_docx_path, pdf_file_path)
            print(f'PDF report generated: {pdf_file_path}')
        except Exception as e:
            print(f"An error occurred during DOCX to PDF conversion: {e}")    
        finally:
            pythoncom.CoUninitialize()
            os.remove(temp_docx_path)

