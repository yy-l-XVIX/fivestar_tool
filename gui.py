import tkinter as tk
import tkinter.ttk as ttk
import validation
import database as db
import qrcode
import platform
import subprocess
import datetime
import os, sys
from tkinter import messagebox
from PIL import ImageTk, Image
from io import BytesIO
from docxtpl import DocxTemplate

class Application:
    def __init__(self, master):

        self.master = master
        self.form_elements = {}
        self.form_data = {}
        self.validated_data = {}
        self.nowtime = datetime.datetime.now()
        # 驗證失敗的樣式
        style = ttk.Style()
        style.configure("Invalid.TEntry", fieldbackground="#ffbfaa")
        style.configure("Invalid.TCombobox", foreground='red', )

        # 創建表單元件
        outer_frame = ttk.Frame(master)
        outer_frame.pack(fill='both', expand=1)
        # canvas
        my_canvas = tk.Canvas(outer_frame)
        my_canvas.pack(side='left', fill='both', expand=1)
        # scrollbar
        my_scrollbar = tk.Scrollbar(outer_frame, orient='vertical', command=my_canvas.yview)
        my_scrollbar.pack(side='right', fill='y')
        # target frame for put something
        container = ttk.Frame(my_canvas)
        container.pack(fill='both')
        # configure the canvas
        my_canvas.configure(yscrollcommand=my_scrollbar.set)
        my_canvas.bind(
            '<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all"))
        )
        my_canvas.create_window((0, 0), window=container, anchor="nw")

        # LabelFrame 表單控件容器

        thm_form_frame = ttk.LabelFrame(container, text="THM表單")
        thm_form_frame.pack(padx=5, pady=5)
        thm_basic_frame = ttk.Frame(thm_form_frame)
        thm_basic_frame.grid(row=0, column=0, padx=5, pady=2)
        thm_param_frame = ttk.Frame(thm_form_frame)
        thm_param_frame.grid(row=0, column=1, padx=5, pady=2)
        thm_process_fram = ttk.Frame(thm_form_frame)
        thm_process_fram.grid(row=0, column=2, padx=5, pady=2)
        thm_qrcode_frame = ttk.Frame(thm_form_frame)
        thm_qrcode_frame.grid(row=0, column=3, padx=5, pady=2)

        # 分割線
        ttk.Separator(container, orient='horizontal').pack(fill='x')

        tvm_form_frame = ttk.LabelFrame(container, text="TVM表單")
        tvm_form_frame.pack(padx=5, pady=5)
        tvm_basic_frame = ttk.Frame(tvm_form_frame)
        tvm_basic_frame.grid(row=0, column=0, padx=5, pady=2)
        tvm_param_frame = ttk.Frame(tvm_form_frame)
        tvm_param_frame.grid(row=0, column=1, padx=5, pady=2)
        tvm_process_fram = ttk.Frame(tvm_form_frame)
        tvm_process_fram.grid(row=0, column=2, padx=5, pady=2)
        tvm_qrcode_frame = ttk.Frame(tvm_form_frame)
        tvm_qrcode_frame.grid(row=0, column=3, padx=5, pady=2)
        action_frame = ttk.Frame(container)
        action_frame.pack(fill='both',side='right')

        # THM 表單控件 ===================================================================
        making_date_label = ttk.Label(thm_basic_frame, text="製單日期")
        making_date_entry = ttk.Entry(thm_basic_frame)
        self.form_elements['making_date'] = making_date_entry
        client_name_label = ttk.Label(thm_basic_frame, text="客戶名稱")
        client_name_entry = ttk.Entry(thm_basic_frame)
        self.form_elements['client_name'] = client_name_entry
        order_number_label = ttk.Label(thm_basic_frame, text="表單號碼")
        order_number_entry = ttk.Entry(thm_basic_frame)
        self.form_elements['order_number'] = order_number_entry
        production_number_label = ttk.Label(thm_basic_frame, text='訂單號碼')
        production_number_entry = ttk.Entry(thm_basic_frame)
        self.form_elements['production_number'] = production_number_entry
        material_name_label = ttk.Label(thm_param_frame, text='鋼材名稱')
        material_name_entry = ttk.Entry(thm_param_frame)
        self.form_elements['material_name'] = material_name_entry
        quantity_label = ttk.Label(thm_param_frame, text='數量')
        quantity_entry = ttk.Entry(thm_param_frame)
        self.form_elements['quantity'] = quantity_entry
        pro_no_label = ttk.Label(thm_param_frame, text='工序')
        pro_no_entry = ttk.Entry(thm_param_frame)
        self.form_elements['pro_no'] = pro_no_entry
        z_heat_label = ttk.Label(thm_param_frame, text='熱處理')
        z_heat_entry = ttk.Entry(thm_param_frame)
        self.form_elements['z_heat'] = z_heat_entry
        grind_label = ttk.Label(thm_param_frame, text='研磨')
        grind_entry = ttk.Entry(thm_param_frame)
        self.form_elements['grind'] = grind_entry
        chamfering_label = ttk.Label(thm_param_frame, text='倒角')
        chamfering_entry = ttk.Entry(thm_param_frame)
        self.form_elements['chamfering'] = chamfering_entry
        cut_parameter_label = ttk.Label(thm_param_frame,text='切削條件')
        cut_parameter_combobox = ttk.Combobox(thm_param_frame,values=['1','2','3','4','5','6','7','8'])
        self.form_elements['cut_parameter'] = cut_parameter_combobox
        machine_number_label = ttk.Label(thm_param_frame, text='機台編號')
        machine_number_entry = ttk.Entry(thm_param_frame)
        self.form_elements['machine_number'] = machine_number_entry
        delivery_date_label = ttk.Label(thm_basic_frame, text='交貨日期')
        delivery_date_entry = ttk.Entry(thm_basic_frame)
        self.form_elements['delivery_date'] = delivery_date_entry
        material_length_label = ttk.Label(thm_process_fram, text='素材長')
        material_length_entry = ttk.Entry(thm_process_fram, width=10)
        self.form_elements['material_length'] = material_length_entry
        material_width_label = ttk.Label(thm_process_fram, text='素材寬')
        material_width_entry = ttk.Entry(thm_process_fram, width=10)
        self.form_elements['material_width'] = material_width_entry
        material_height_label = ttk.Label(thm_process_fram, text='素材高')
        material_height_entry = ttk.Entry(thm_process_fram, width=10)
        self.form_elements['material_height'] = material_height_entry
        finished_length_label = ttk.Label(thm_process_fram, text='完成長')
        finished_length_entry = ttk.Entry(thm_process_fram, width=10)
        self.form_elements['finished_length'] = finished_length_entry
        finished_width_label = ttk.Label(thm_process_fram, text='完成寬')
        finished_width_entry = ttk.Entry(thm_process_fram, width=10)
        self.form_elements['finished_width'] = finished_width_entry
        finished_height_label = ttk.Label(thm_process_fram, text='完成高')
        finished_height_entry = ttk.Entry(thm_process_fram, width=10)
        self.form_elements['finished_height'] = finished_height_entry
        fw_label = ttk.Label(thm_process_fram, text='FW')
        self.fw_label = fw_label
        fl_label = ttk.Label(thm_process_fram, text='FL')
        self.fl_label = fl_label
        processing_technology_label = ttk.Label(thm_process_fram, text='加工工藝')
        processing_technology_entry = ttk.Entry(thm_process_fram, width=10)
        self.form_elements['processing_technology'] = processing_technology_entry
        remark_label = ttk.Label(thm_process_fram, text='備註')
        remark_entry = tk.Text(thm_process_fram, width=50, height=10)
        self.form_elements['remark'] = remark_entry
        pieces_label = ttk.Label(thm_qrcode_frame, text='枚數')
        pieces_entry = ttk.Entry(thm_qrcode_frame, width=13)
        self.form_elements['pieces'] = pieces_entry
        weight_label = ttk.Label(thm_qrcode_frame, text='重量')
        weight_entry = ttk.Entry(thm_qrcode_frame, width=13)
        self.form_elements['weight'] = weight_entry
        processing_date_label = ttk.Label(thm_qrcode_frame, text='加工日期')
        processing_date_entry = ttk.Entry(thm_qrcode_frame, width=13)
        self.form_elements['processing_date'] = processing_date_entry
        processing_staff_label = ttk.Label(thm_qrcode_frame, text='加工人員')
        processing_staff_entry = ttk.Entry(thm_qrcode_frame, width=13)
        self.form_elements['processing_staff'] = processing_staff_entry
        bale_checkbutton = ttk.Checkbutton(thm_qrcode_frame, text='捆包')
        bale_checkbutton.state(['!alternate'])
        self.form_elements['bale_check'] = bale_checkbutton
        packet_checkbutton = ttk.Checkbutton(thm_qrcode_frame, text='封包')
        packet_checkbutton.state(['!alternate'])
        self.form_elements['packet_check'] = packet_checkbutton
        qr_label = ttk.Label(thm_qrcode_frame)
        self.qr_label = qr_label

        # TVM 表單控件 ===================================================================
        tvm_making_date_label = ttk.Label(tvm_basic_frame, text="製單日期")
        tvm_making_date_entry = ttk.Entry(tvm_basic_frame)
        self.form_elements['tvm_making_date'] = tvm_making_date_entry  # 将表单元素添加到数组中
        tvm_client_name_label = ttk.Label(tvm_basic_frame, text="客戶名稱")
        tvm_client_name_entry = ttk.Entry(tvm_basic_frame)
        self.form_elements['tvm_client_name'] = tvm_client_name_entry
        tvm_order_number_label = ttk.Label(tvm_basic_frame, text="工單號碼")
        tvm_order_number_entry = ttk.Entry(tvm_basic_frame)
        self.form_elements['tvm_order_number'] = tvm_order_number_entry
        tvm_production_number_label = ttk.Label(tvm_basic_frame, text='訂單號碼')
        tvm_production_number_entry = ttk.Entry(tvm_basic_frame)
        self.form_elements['tvm_production_number'] = tvm_production_number_entry
        tvm_material_name_label = ttk.Label(tvm_param_frame, text='鋼材名稱')
        tvm_material_name_entry = ttk.Entry(tvm_param_frame)
        self.form_elements['tvm_material_name'] = tvm_material_name_entry
        tvm_quantity_label = ttk.Label(tvm_param_frame, text='數量')
        tvm_quantity_entry = ttk.Entry(tvm_param_frame)
        self.form_elements['tvm_quantity'] = tvm_quantity_entry
        tvm_pro_no_label = ttk.Label(tvm_param_frame, text='工序')
        tvm_pro_no_entry = ttk.Entry(tvm_param_frame)
        self.form_elements['tvm_pro_no'] = tvm_pro_no_entry
        tvm_z_heat_label = ttk.Label(tvm_param_frame, text='熱處理')
        tvm_z_heat_entry = ttk.Entry(tvm_param_frame)
        self.form_elements['tvm_z_heat'] = tvm_z_heat_entry
        tvm_grind_label = ttk.Label(tvm_param_frame, text='研磨')
        tvm_grind_entry = ttk.Entry(tvm_param_frame)
        self.form_elements['tvm_grind'] = tvm_grind_entry
        tvm_rou_magn_label = ttk.Label(tvm_param_frame, text='粗切磁力')
        tvm_rou_magn_entry = ttk.Combobox(tvm_param_frame,values=['1','2','3','4','5','6','7','8'])
        self.form_elements['tvm_rou_magn'] = tvm_rou_magn_entry
        tvm_fin_magn_label = ttk.Label(tvm_param_frame, text='精切磁力')
        tvm_fin_magn_entry = ttk.Combobox(tvm_param_frame,values=['1','2','3','4','5','6','7','8'])
        self.form_elements['tvm_fin_magn'] = tvm_fin_magn_entry
        tvm_cut_parameter_label = ttk.Label(tvm_param_frame,text='切削條件')
        tvm_cut_parameter_combobox = ttk.Combobox(tvm_param_frame,values=['1','2','3','4','5','6','7','8'])
        self.form_elements['tvm_cut_parameter'] = tvm_cut_parameter_combobox
        tvm_machine_number_label = ttk.Label(tvm_param_frame, text='機台編號')
        tvm_machine_number_entry = ttk.Entry(tvm_param_frame)
        self.form_elements['tvm_machine_number'] = tvm_machine_number_entry
        tvm_delivery_date_label = ttk.Label(tvm_basic_frame, text='交貨日期')
        tvm_delivery_date_entry = ttk.Entry(tvm_basic_frame)
        self.form_elements['tvm_delivery_date'] = tvm_delivery_date_entry
        tvm_material_length_label = ttk.Label(tvm_process_fram, text='素材長')
        tvm_material_length_entry = ttk.Entry(tvm_process_fram,width=10)
        self.form_elements['tvm_material_length'] = tvm_material_length_entry
        tvm_material_width_label = ttk.Label(tvm_process_fram, text='素材寬')
        tvm_material_width_entry = ttk.Entry(tvm_process_fram,width=10)
        self.form_elements['tvm_material_width'] = tvm_material_width_entry
        tvm_material_height_label = ttk.Label(tvm_process_fram, text='素材高')
        tvm_material_height_entry = ttk.Entry(tvm_process_fram,width=10)
        self.form_elements['tvm_material_height'] = tvm_material_height_entry
        tvm_single_cut_label = ttk.Label(tvm_process_fram, text='單面切完成厚')
        tvm_single_cut_entry = ttk.Entry(tvm_process_fram,width=10)
        self.form_elements['tvm_single_cut'] = tvm_single_cut_entry
        fh_single_label = ttk.Label(tvm_process_fram, text='FH')
        self.fh_single_label = fh_single_label
        tvm_double_cut_label = ttk.Label(tvm_process_fram, text='兩面切完成厚')
        tvm_double_cut_entry = ttk.Entry(tvm_process_fram,width=10)
        self.form_elements['tvm_double_cut'] = tvm_double_cut_entry
        fh_double_label = ttk.Label(tvm_process_fram, text='FH')
        self.fh_double_label = fh_double_label
        tvm_qr_label = ttk.Label(tvm_qrcode_frame)
        self.tvm_qr_label = tvm_qr_label

        # 指定項目自動填入tvm表單
        copy_entries = [
            (making_date_entry, tvm_making_date_entry),
            (client_name_entry, tvm_client_name_entry),
            (order_number_entry, tvm_order_number_entry),
            (production_number_entry, tvm_production_number_entry),
            (material_name_entry, tvm_material_name_entry),
            (quantity_entry, tvm_quantity_entry),
            (pro_no_entry, tvm_pro_no_entry),
            (z_heat_entry, tvm_z_heat_entry),
            (grind_entry, tvm_grind_entry),
            (cut_parameter_combobox, tvm_cut_parameter_combobox),
            (delivery_date_entry, tvm_delivery_date_entry),
            (material_length_entry, tvm_material_length_entry),
            (material_width_entry, tvm_material_width_entry),
            (material_height_entry, tvm_material_height_entry),
        ]

        for entry, tvm_entry in copy_entries:
            def copy_value(event, entry=entry, tvm_entry=tvm_entry):
                # 清除readonly
                tvm_entry.state(["!readonly"])
                tvm_entry.delete(0, 'end')
                tvm_entry.insert(0, entry.get())
                #設置readonly 不可被手動修改
                tvm_entry.state(["readonly"])

            entry.bind('<FocusOut>', copy_value, add="+")

        # 檢查表單號碼是否有儲存記錄
        order_number_entry.bind("<FocusOut>", lambda event: self.search_data(), add="+")

        # 設定預設值
        making_date_entry.insert(0,self.nowtime.strftime("%Y/%m/%d"))
        tvm_making_date_entry.insert(0,self.nowtime.strftime("%Y/%m/%d"))

        # 創建按鈕
        self.validate_button = ttk.Button(action_frame, text="完成", command=self.validate_form)
        self.clear_button = ttk.Button(action_frame, text="清除", command=self.clear_form)
        self.print_button = ttk.Button(action_frame, text="列印", command=self.print_form)
        self.save_button = ttk.Button(action_frame, text="儲存", command=self.save_data)

        # 佈局表單控件按鈕
        # thm ================================================
        # basic frame
        making_date_label.grid(row=0, column=0,pady=5)
        making_date_entry.grid(row=0, column=1,pady=5)
        client_name_label.grid(row=1, column=0,pady=5)
        client_name_entry.grid(row=1, column=1,pady=5)
        order_number_label.grid(row=2, column=0,pady=5)
        order_number_entry.grid(row=2, column=1,pady=5)
        production_number_label.grid(row=3, column=0,pady=5)
        production_number_entry.grid(row=3, column=1,pady=5)
        delivery_date_label.grid(row=4, column=0,pady=5)
        delivery_date_entry.grid(row=4, column=1,pady=5)
        # parame frame
        material_name_label.grid(row=0, column=0,pady=5)
        material_name_entry.grid(row=0, column=1,pady=5)
        quantity_label.grid(row=1, column=0,pady=5)
        quantity_entry.grid(row=1, column=1,pady=5)
        pro_no_label.grid(row=2, column=0,pady=5)
        pro_no_entry.grid(row=2, column=1,pady=5)
        z_heat_label.grid(row=3, column=0,pady=5)
        z_heat_entry.grid(row=3, column=1,pady=5)
        grind_label.grid(row=4, column=0,pady=5)
        grind_entry.grid(row=4, column=1,pady=5)
        chamfering_label.grid(row=5, column=0,pady=5)
        chamfering_entry.grid(row=5, column=1,pady=5)
        cut_parameter_label.grid(row=6, column=0,pady=5)
        cut_parameter_combobox.grid(row=6, column=1,pady=5)
        machine_number_label.grid(row=7, column=0,pady=5)
        machine_number_entry.grid(row=7, column=1,pady=5)
        # process frame
        material_length_label.grid(row=0, column=0,pady=5)
        material_length_entry.grid(row=0, column=1,pady=5)
        material_width_label.grid(row=0, column=2,pady=5)
        material_width_entry.grid(row=0, column=3,pady=5)
        material_height_label.grid(row=0, column=4,pady=5)
        material_height_entry.grid(row=0, column=5,pady=5)
        finished_length_label.grid(row=1, column=0,pady=5)
        finished_length_entry.grid(row=1, column=1,pady=5)
        finished_width_label.grid(row=1, column=2,pady=5)
        finished_width_entry.grid(row=1, column=3,pady=5)
        finished_height_label.grid(row=1, column=4,pady=5)
        finished_height_entry.grid(row=1, column=5,pady=5)
        fw_label.grid(row=2, column=0,pady=5)
        fl_label.grid(row=2, column=1,pady=5)
        processing_technology_label.grid(row=3, column=0,pady=5)
        processing_technology_entry.grid(row=3, column=1,columnspan=5,pady=5, sticky='w')
        remark_label.grid(row=4, column=0,pady=5)
        remark_entry.grid(row=4, column=1,columnspan=5,pady=5, sticky='w')
        # qrcode frame
        ttk.Label(thm_qrcode_frame,text='THM機台掃描條碼').grid(row=0, columnspan=2,pady=5)
        qr_label.grid(row=1, columnspan=2,pady=5)
        ttk.Label(thm_qrcode_frame,text='Machine scanning barcode').grid(row=2, columnspan=2,pady=5)
        pieces_label.grid(row=3, column=0,pady=5)
        pieces_entry.grid(row=3, column=1,pady=5)
        weight_label.grid(row=4, column=0,pady=5)
        weight_entry.grid(row=4, column=1,pady=5)
        processing_date_label.grid(row=5, column=0,pady=5)
        processing_date_entry.grid(row=5, column=1,pady=5)
        processing_staff_label.grid(row=6, column=0,pady=5)
        processing_staff_entry.grid(row=6, column=1,pady=5)
        bale_checkbutton.grid(row=7, column=0,pady=5)
        packet_checkbutton.grid(row=7, column=1,pady=5)

        # tvm ================================================
        # basic frame
        tvm_making_date_label.grid(row=0, column=0,pady=5)
        tvm_making_date_entry.grid(row=0, column=1,pady=5)
        tvm_client_name_label.grid(row=1, column=0,pady=5)
        tvm_client_name_entry.grid(row=1, column=1,pady=5)
        tvm_order_number_label.grid(row=2, column=0,pady=5)
        tvm_order_number_entry.grid(row=2, column=1,pady=5)
        tvm_production_number_label.grid(row=3, column=0,pady=5)
        tvm_production_number_entry.grid(row=3, column=1,pady=5)
        tvm_delivery_date_label.grid(row=4, column=0,pady=5)
        tvm_delivery_date_entry.grid(row=4, column=1,pady=5)
        # parameter frame
        tvm_material_name_label.grid(row=0, column=0,pady=5)
        tvm_material_name_entry.grid(row=0, column=1,pady=5)
        tvm_quantity_label.grid(row=1, column=0,pady=5)
        tvm_quantity_entry.grid(row=1, column=1,pady=5)
        tvm_pro_no_label.grid(row=2, column=0,pady=5)
        tvm_pro_no_entry.grid(row=2, column=1,pady=5)
        tvm_z_heat_label.grid(row=3, column=0,pady=5)
        tvm_z_heat_entry.grid(row=3, column=1,pady=5)
        tvm_grind_label.grid(row=4, column=0,pady=5)
        tvm_grind_entry.grid(row=4, column=1,pady=5)
        tvm_rou_magn_label.grid(row=5, column=0,pady=5)
        tvm_rou_magn_entry.grid(row=5, column=1,pady=5)
        tvm_fin_magn_label.grid(row=6, column=0,pady=5)
        tvm_fin_magn_entry.grid(row=6, column=1,pady=5)
        tvm_cut_parameter_label.grid(row=7, column=0,pady=5)
        tvm_cut_parameter_combobox.grid(row=7, column=1,pady=5)
        tvm_machine_number_label.grid(row=8, column=0,pady=5)
        tvm_machine_number_entry.grid(row=8, column=1,pady=5)
        # process frame
        tvm_material_length_label.grid(row=0, column=0,pady=5)
        tvm_material_length_entry.grid(row=0, column=1,pady=5)
        tvm_material_width_label.grid(row=0, column=2,pady=5)
        tvm_material_width_entry.grid(row=0, column=3,pady=5)
        tvm_material_height_label.grid(row=0, column=4,pady=5)
        tvm_material_height_entry.grid(row=0, column=5,pady=5)
        tvm_single_cut_label.grid(row=1, column=0,pady=5)
        tvm_single_cut_entry.grid(row=1, column=1,pady=5)
        fh_single_label.grid(row=1, column=2,pady=5)
        tvm_double_cut_label.grid(row=2, column=0,pady=5)
        tvm_double_cut_entry.grid(row=2, column=1,pady=5)
        fh_double_label.grid(row=2, column=2,pady=5)
        # qrcode
        ttk.Label(tvm_qrcode_frame,text='TVM機台掃描條碼').grid(row=0, columnspan=2,pady=5)
        tvm_qr_label.grid(row=1, columnspan=2,pady=5)
        ttk.Label(tvm_qrcode_frame,text='Machine scanning barcode').grid(row=2, columnspan=2,pady=5)

        self.validate_button.grid(row=0, column=0, padx=5, pady=5)
        self.clear_button.grid(row=0, column=1, padx=5, pady=5)
        self.print_button.grid(row=0, column=2, padx=5, pady=5)
        self.save_button.grid(row=0, column=3, padx=5, pady=5)

    def save_data(self):
        thm_data = {}
        tvm_data = {}
        data = {}
        if not self.form_data:
            # 資料未驗證不可以儲存
            messagebox.showinfo('System', '請先填入完整資訊並驗證在執行儲存',)
            return False

        thm_data = {
            'making_date' : self.form_data['making_date'],
            'client_name' : self.form_data['client_name'],
            'order_number' : self.form_data['order_number'],
            'production_number' : self.form_data['production_number'],
            'material_name' : self.form_data['material_name'],
            'quantity' : self.form_data['quantity'],
            'pro_no' : self.form_data['pro_no'],
            'z_heat' : self.form_data['z_heat'],
            'grind' : self.form_data['grind'],
            'chamfering' : self.form_data['chamfering'],
            'cut_parameter' : self.form_data['cut_parameter'],
            'machine_number' : self.form_data['machine_number'],
            'delivery_date' : self.form_data['delivery_date'],
            'material_length' : self.form_data['material_length'],
            'material_width' : self.form_data['material_width'],
            'material_height' : self.form_data['material_height'],
            'finished_length' : self.form_data['finished_length'],
            'finished_width' : self.form_data['finished_width'],
            'finished_height' : self.form_data['finished_height'],
            'fw' : self.form_data['fw'],
            'fl' : self.form_data['fl'],
            'processing_technology' : self.form_data['processing_technology'],
            'remark' : self.form_data['remark'],
            'pieces' : self.form_data['pieces'],
            'weight' : self.form_data['weight'],
            'bale_check' : 1 if self.form_data['bale_check']  else 0,
            'packet_check' : 1 if self.form_data['packet_check']  else 0,
            'updatedate' : self.nowtime.strftime("%Y/%m/%d %H:%M:%S"),
        }
        tvm_data = {
            'making_date' : self.form_data['tvm_making_date'],
            'client_name' : self.form_data['tvm_client_name'],
            'order_number' : self.form_data['tvm_order_number'],
            'production_number' : self.form_data['tvm_production_number'],
            'material_name' : self.form_data['tvm_material_name'],
            'quantity' : self.form_data['tvm_quantity'],
            'pro_no' : self.form_data['tvm_pro_no'],
            'z_heat' : self.form_data['tvm_z_heat'],
            'grind' : self.form_data['tvm_grind'],
            'rou_magn' : self.form_data['tvm_rou_magn'],
            'fin_magn' : self.form_data['tvm_fin_magn'],
            'cut_parameter' : self.form_data['tvm_cut_parameter'],
            'machine_number' : self.form_data['tvm_machine_number'],
            'delivery_date' : self.form_data['tvm_delivery_date'],
            'material_length' : self.form_data['tvm_material_length'],
            'material_width' : self.form_data['tvm_material_width'],
            'material_height' : self.form_data['tvm_material_height'],
            'single_cut' : self.form_data['tvm_single_cut'],
            'double_cut' : self.form_data['tvm_double_cut'],
            'fh_single' : self.form_data['fh_single'],
            'fh_double' : self.form_data['fh_double'],
            'updatedate' : self.nowtime.strftime("%Y/%m/%d %H:%M:%S"),
        }

        # check order_number exist
        sql = f"SELECT order_number FROM thm_form_data WHERE order_number = '{thm_data['order_number']}' LIMIT 1;"
        data = db.query_data(sql)

        if not data:
            # do insert
            thm_data['createdate'] = self.nowtime.strftime("%Y/%m/%d %H:%M:%S")
            tvm_data['createdate'] = self.nowtime.strftime("%Y/%m/%d %H:%M:%S")
            try:
                db.insert_data('thm_form_data',thm_data)
                db.insert_data('tvm_form_data',tvm_data)
                messagebox.showinfo('System', "資料儲存成功",)

            except Exception as e:
                # print("Failed to open the document:", e)
                messagebox.showinfo('System', f"資料儲存失敗 {e}",)

        else:
            try:
                db.update_data('thm_form_data',thm_data,thm_data['order_number'])
                db.update_data('tvm_form_data',tvm_data,tvm_data['order_number'])
                messagebox.showinfo('System', "資料更新成功")

            except Exception as e:
                # print("Failed to open the document:", e)
                messagebox.showinfo('System', f"資料儲存失敗 {e}",)

    def clear_form(self):
        for name, entry in self.form_elements.items():
            if isinstance(entry, tk.Text):
                entry.delete("1.0", tk.END)
                continue
            if isinstance(entry, ttk.Checkbutton):
                entry.state(['!selected'])
                continue
            entry.state(["!readonly"])
            entry.delete(0, tk.END)
            if isinstance(entry, tk.Text) or isinstance(entry, ttk.Checkbutton) or isinstance(entry, ttk.Combobox):
                continue
            entry.configure(style="TEntry")  # 重置樣式

        self.qr_label.configure(image=None)
        self.qr_label.image = None
        self.tvm_qr_label.configure(image=None)
        self.tvm_qr_label.image = None
        self.fw_label['text'] = 'FW'
        self.fl_label['text'] = 'FL'
        self.fh_single_label['text'] = 'FH'
        self.fh_double_label['text'] = 'FH'
        self.form_data = {}

    def print_form(self):

        if not self.form_data:
            messagebox.showinfo('System', '請先填入資訊',)
            return False

        doc = DocxTemplate(self.get_route('form_template.docx'))
        doc.replace_pic('qr_label',self.form_data['qr_label'])
        doc.replace_pic('tvm_qr_label',self.form_data['tvm_qr_label'])
        doc.render(self.form_data)
        doc_path = self.get_route('ready_to_print.docx')
        doc.save(doc_path)

        # 檢測當前系統
        current_os = platform.system()
        try:
            if current_os == "Darwin":
                subprocess.Popen(['open', doc_path])
            elif current_os == 'Windows':
                subprocess.Popen(['start', doc_path], shell=True)

        except Exception as e:
            # print("Failed to open the document:", e)
            messagebox.showinfo('System', f"系統發生錯誤 {e}",)

    def validate_form(self):
        form_data = {}
        # 獲取值
        for name, entry in self.form_elements.items():
            if isinstance(entry, tk.Text):
                form_data[name] = entry.get("1.0", tk.END)
                continue
            if isinstance(entry, ttk.Checkbutton):
                form_data[name] = entry.instate(['selected'])
                continue
            form_data[name] = entry.get()
        form_data['bale_checkbox'] = "☐"
        form_data['packet_checkbox'] = "☐"

        # 表單欄位資料驗證
        validation_result, invalid_fields, form_data = validation.validate_form(form_data)

        #先重置樣式
        for entry in self.form_elements.values():
            if isinstance(entry, tk.Text) or isinstance(entry, ttk.Checkbutton):
                continue
            elif isinstance(entry, ttk.Combobox):
                entry.configure(style="TCombobox")
            else:
                entry.configure(style="TEntry")

        if validation_result:

            self.fw_label['text'] = 'FW ' + form_data['fw']
            self.fl_label['text'] = 'FL ' + form_data['fl']
            self.fh_single_label['text'] = 'FH ' + form_data['fh_single']
            self.fh_double_label['text'] = 'FH ' + form_data['fh_double']

            self.form_data = form_data

            # 驗證成功 產生qrcode
            self.get_qrcode(form_data)
        else:
            # 驗證失敗控件標記失敗
            for field in invalid_fields:
                entry = self.form_elements[field]
                if isinstance(entry, tk.Text) or isinstance(entry, ttk.Checkbutton):
                    continue
                elif isinstance(entry, ttk.Combobox):
                    entry.configure(style="Invalid.TCombobox")
                else:
                    entry.configure(style="Invalid.TEntry")

    def get_qrcode(self, form_data):
        thm_data = [
            form_data['material_height'],
            form_data['material_width'],
            form_data['material_length'],
            form_data['finished_height'],
            form_data['fw'],# form_data['finished_width'],
            form_data['fl'],# form_data['finished_length'],
            form_data['chamfering'],
            form_data['machine_number'],
            form_data['cut_parameter'],
            form_data['order_number'],
            form_data['pro_no'],
        ]
        tvm_data = [
            form_data['tvm_material_height'],
            form_data['tvm_material_width'],
            form_data['tvm_material_length'],
            form_data['fh_single'], # form_data['tvm_single_cut'],
            form_data['fh_double'], # form_data['tvm_double_cut'],
            '0000,000',# 預留用
            form_data['tvm_rou_magn'] + ',' + form_data['tvm_fin_magn'],
            form_data['tvm_machine_number'],
            form_data['tvm_cut_parameter'],
            form_data['tvm_order_number'],
            form_data['tvm_pro_no'],
        ]
        # 資料串接
        thm_string = f"{thm_data[0]}{thm_data[1]}{thm_data[2]}{thm_data[3]}{thm_data[4]}{thm_data[5]}{thm_data[6]}{thm_data[7]}{thm_data[8]}{thm_data[9]}{thm_data[10]}"
        thm_string = thm_string.replace('.', ',') #小數點取代成逗點
        tvm_string = f"{tvm_data[0]}{tvm_data[1]}{tvm_data[2]}{tvm_data[3]}{tvm_data[4]}{tvm_data[5]}{tvm_data[6]}{tvm_data[7]}{tvm_data[8]}{tvm_data[9]}{tvm_data[10]}"
        tvm_string = tvm_string.replace('.', ',') #小數點取代成逗點

        # 生成thm QR Code圖像
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=2, border=2)
        qr.add_data(thm_string)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")

        if img is None:
            # print("Failed to generate QR code image")
            # print("thm_data = ",thm_data)
            # print("data = ",thm_string)
            messagebox.showinfo('System', f"THM圖片生成失敗 {e}",)

        else:
            img_tk = ImageTk.PhotoImage(img)
            self.qr_label.configure(image=img_tk)
            self.qr_label.image = img_tk
            qr_image_path = self.get_route("temp_thm_qrcode.png")
            img.save(qr_image_path)
            self.form_data['qr_label'] = qr_image_path # 顯示qrcode

        # 生成thm QR Code圖像
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=2, border=2)
        qr.add_data(tvm_string)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")

        if img is None:
            # print("Failed to generate QR code image")
            # print("tvm_data = ",tvm_data)
            # print("data = ",tvm_string)
            messagebox.showinfo('System', f"TVM圖片生成失敗 {e}",)

        else:
            img_tk = ImageTk.PhotoImage(img)
            self.tvm_qr_label.configure(image=img_tk)
            self.tvm_qr_label.image = img_tk
            qr_image_path = self.get_route("temp_tvm_qrcode.png")
            img.save(qr_image_path)
            self.form_data['tvm_qr_label'] = qr_image_path

        self.master.geometry("1270x741")

    def search_data(self):
        order_number = self.form_elements['order_number'].get()

        # 查詢order_number是否有紀錄
        sql = f"SELECT * FROM thm_form_data WHERE order_number = '{order_number}' LIMIT 1;"
        thm_data = db.query_data(sql)
        sql = f"SELECT * FROM tvm_form_data WHERE order_number = '{order_number}' LIMIT 1;"
        tvm_data = db.query_data(sql)

        if thm_data:
            msg_box = messagebox.askquestion('System', '系統偵測到一筆紀錄存在，請問是否載入?', icon='question')

            if msg_box == 'yes':
                # 資料記錄回填入控件
                # thm 表單
                for key , value in thm_data.items():
                    if not key in self.form_elements:
                        continue
                    entry = self.form_elements[key]

                    if isinstance(entry, tk.Text):
                        entry.delete("1.0", tk.END)
                        entry.insert("1.0", value)
                        continue
                    if isinstance(entry, ttk.Checkbutton):
                        entry.state(['!selected'])
                        if value == 1:
                            entry.state(['selected'])
                        continue
                    entry.state(["!readonly"])
                    entry.delete(0, tk.END)
                    entry.insert(0, value)

                # tvm 表單
                for key , value in tvm_data.items():
                    if not 'tvm_'+key in self.form_elements:
                        continue
                    entry = self.form_elements['tvm_'+key]
                    if isinstance(entry, tk.Text):
                        entry.delete("1.0", tk.END)
                        entry.insert("1.0", value)
                        continue
                    if isinstance(entry, ttk.Checkbutton):
                        entry.state(['!selected'])
                        if value == 1:
                            entry.state(['selected'])
                        continue
                    entry.state(["!readonly"])
                    entry.delete(0, tk.END)
                    entry.insert(0, value)

    def _get_cwd(self):
        if getattr(sys, "frozen", False):
            path = os.path.dirname(sys.executable)
        elif __file__:
            path = os.path.dirname(__file__)
        else:
            path = os.path.dirname(os.getcwd())
        return path

    def get_route(self,fileName):
        path = os.path.join(self._get_cwd(), fileName)
        path = os.path.abspath(path)

        return path
