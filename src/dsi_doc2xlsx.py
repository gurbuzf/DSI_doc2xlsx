try:
    import xlsxwriter
    import docx2txt
    import re
except ImportError as _err:
    print("ImportError: {0}".format(_err))
    raise _err


class read_DSI_doc_file:
    """ A class handling information in DSI streamflow doc files.  

    ...

    Attributes:
        path: str, 
            directory of the doc file

    Methods:
        doc2text:doc formatli DSI akim yillik sayfasini text ve tex listesine cevirir.

        find_index_YagisAlani: text list'te scan_start ve scan_end arasindaki stringleri tarar ve 
            'YAĞIŞ ALANI' kelimesini arar.  

        parse_files:doc formatli DSI akim yillik sayfasindan tum verileri alir.

    
    """
    def __init__(self, path):
        self.path = path
        text_list = self.doc2text(self.path)
        self.parse_file(text_list)
    
    def parse_file(self, text_list):
        """doc formatli DSI akim yillik sayfasindan tum verileri alir.
        
        Parameters:
            text_list:list, 
                doc2text methodunun ciktisi olan ve doc dosyasindaki satirlari 
                    iceren liste

        Returns:
            None
        
        """
        self.watershed = text_list[0].strip()
        # Istasyon_Kodu Istasyon_Adi
        self.station = text_list[2].strip() 
        # Istasyon_kodu 
        self.sta_number = self.station.split(" ", 1)[0] 
        # Istasyon_Adi
        self. sta_name = self.station.split(" ", 1)[1] 
        # Referans index YAĞIŞ ALANI 
        ind_drArea = self.find_index_YagisAlani(text_list) 
        temp = text_list[ind_drArea]
        _lineDrarea = re.split("km2|:", ''.join(temp.split()))
        #   YAĞIŞ ALANI  
        self.area = _lineDrarea[1]    
        #   YAKLAŞIK KOT  	
        self.altitude = _lineDrarea[-1][:-1] 
        #   KOORDINAT
        self.coords = text_list[ind_drArea - 2].strip() 
        #   YERI
        self.location = '\n'.join([i.strip() for i in \
                                    text_list[4 :ind_drArea - 2]])
        self.location = re.split("YERİ", self.location)[-1].strip()[1:].strip() 
        #   GÖZLEM SÜRESİ
        self.obs_duration = text_list[ind_drArea + 2]\
                    .strip().split(':')[-1].strip() 
        #   ORTALAMA AKIMLAR
        self.mean_flow = ' '.join(text_list[ind_drArea + 4]\
                    .strip().split(':')[-1].strip().split())  
        #    ANLIK EN ÇOK VE EN AZ AKIMLAR - Tablo
        self.stats_InstFlow = [[i.strip() for i in re.split("m3/sn|:", \
                    text_list[ind_drArea + j].strip())] for j in [7, 9, 11, 13]] 
        #   Anahtar Eğrisi
        self.rating_info = text_list[ind_drArea + 15].strip()  
        #   Anahtar Eğrisi Tablo
        self.rating_table =  [text_list[ind_drArea + j]\
                    .strip().split() for j in [17, 19, 20, 21, 22, 23]] 
        self.flow_info = text_list[ind_drArea + 27].strip()
        #   Akim verileri - Tablo
        empty_index = [31, 37, 43, 49, 55, 61, 67, 69]  #TODO: make it not hard coded!
        self.streamflow = [text_list[ind_drArea + j].strip()\
                    .split() for j in range(30, 76) if j not in empty_index] 
        #   [AKIM mm.] ve [MİL. M3] satirlarinda duzeltme
        for i in [-1, -2]:
            self.streamflow[i] = [self.streamflow[i][0] + ' ' + \
                    self.streamflow[i][1]] + self.streamflow[i][2:]

        self.footnote = ' '.join(text_list [-2].strip().split())
        
        return 1

    @staticmethod    
    def doc2text(path):

        """doc formatli DSI akim yillik sayfasini text ve tex listesine cevirir.
        
        Parameters:
            path: doc file directory
        
        Returns:
            list : the list of lines in doc file.
        """
        text = docx2txt.process(path)
        text_list = text.split('\n\n')
        
        return text_list

    @staticmethod 
    def find_index_YagisAlani(text_list, scan_start=4, scan_end=12):
        """ text list'te scan_start ve scan_end arasindaki stringleri tarar ve 
            'YAĞIŞ ALANI' kelimesini arar.
        
        Parameters:
            text_list: list
            scan_start: int, default 4
            scan_start: int, default 12
        Returns:
            int : the index where 'YAĞIŞ ALANI' is found.
        """
        temp_text = text_list[scan_start:scan_end]
        value_founded = False
        
        for i, _char in enumerate(temp_text):
            if "YAĞIŞ ALANI" in _char:
                ind_drArea = scan_start + i 
                value_founded = True
            else:
                pass

        if value_founded is False:
            raise AttributeError("YAĞIŞ ALANI bilgisi bulunamadı!" )
        
        return ind_drArea
        
    def write_xlsx(self, path_xlsx):

        # Create an new Excel file and add a worksheet.
        workbook = xlsxwriter.Workbook(path_xlsx,  {'strings_to_numbers':  True})
        worksheet = workbook.add_worksheet()
        worksheet.set_default_row(14)
        worksheet.set_row(0, 15)

        # Create a format to use in the merged range.
        f_header = workbook.add_format({'font_name':'Courier New','font_size':13, 
                                    'valign':'vcenter', 'align':'center', 'bold': True})
        f_text1 = workbook.add_format({'font_name':'Courier New','font_size':8,
                                    'valign':'vcenter', 'align':'center'})
        f_text2 = workbook.add_format({'font_name':'Courier New','font_size':8,
                                    'valign':'vcenter', 'align':'left'})
        f_text2_w = workbook.add_format({'font_name':'Courier New','font_size':8,
                                    'valign':'top', 'align':'left','text_wrap': True})
        f_text3 = workbook.add_format({'font_name':'Microsoft Sans Serif','font_size':8,
                                    'valign':'left', 'align':'left', 'bold': True, 'text_wrap': True})
        f_text4 = workbook.add_format({'font_name':'Courier New','font_size':8.5, 
                                    'valign':'vcenter', 'align':'center', 'text_wrap': True})
        f_text5 = workbook.add_format({'font_name':'Courier New','font_size':8, 
                                    'valign':'vcenter', 'align':'center', 
                                    'border':7, 'border_color':'black', 
                                    'left':False, 'right':False,'bold': True})
        f_text6 = workbook.add_format({'font_name':'Courier New','font_size':8.5, 
                                    'valign':'vcenter', 'align':'center', 
                                    'border':3, 'border_color':'black', 
                                    'left':False, 'right':False, 'bold': True})
        f_text7 = workbook.add_format({'font_name':'Courier New','font_size':8, 
                                    'valign':'vcenter', 'align':'center', 
                                    'border':3, 'border_color':'black', 
                                    'left':False, 'right':False, 'top':False})
        f_text8 = workbook.add_format({'font_name':'Cambria','font_size':8,
                                    'valign':'vcenter', 'align':'center','bold': True})

        worksheet.write('G1', 'D S İ', f_header)
        worksheet.write('G2', self.watershed, f_text1)
        worksheet.write('G3', self.station, f_text1)
        worksheet.merge_range('A4:C4', "YERİ", f_text2)
        worksheet.merge_range('D4:M5', self.location, f_text3)
        worksheet.merge_range('D6:G6', self.coords, f_text2)
        worksheet.merge_range('A7:C7', "YAĞIŞ ALANI", f_text2)
        worksheet.write('D7', self.area, f_text2)
        worksheet.write('E7', "km2", f_text2)
        worksheet.merge_range('G7:H7', "YAKLAŞIK KOT", f_text2)
        worksheet.write('I7', self.altitude, f_text2)
        worksheet.write('J7', "m", f_text2)
        worksheet.merge_range('A8:C8', "GÖZLEM SÜRESİ", f_text2)
        worksheet.write('D8', self.obs_duration, f_text2)
        worksheet.merge_range('A9:C9', "ORTALAMA AKIMLAR", f_text2)
        worksheet.merge_range('D9:M9', self.mean_flow, f_text2)
        worksheet.merge_range(9,0,12,2, "ANLIK EN ÇOK VE EN AZ AKIMLAR", f_text2_w)

        for row, data in enumerate(self.stats_InstFlow):
            worksheet.write(row+9, 8, 'm3/sn', f_text2)
            for col_index, item in enumerate(data):
                if col_index == 0:
                    worksheet.merge_range(row+9,col_index+3, row+9, col_index+6, item, f_text2)
                elif col_index == 2:
                    worksheet.write(row+9, col_index+7, item, f_text2)
                else:
                    worksheet.write(row+9, col_index+6, item, f_text2)
            

        worksheet.write('G14', self.rating_info, f_text1)
        col = 3
        for row, data in enumerate(self.rating_table):
            if row == 0:
                worksheet.write_row(row+14, col, data, f_text5)
            else:    
                worksheet.write_row(row+14, col, data, f_text1)

        worksheet.set_row(20, 22.5)
        worksheet.merge_range('A21:M21', self.flow_info, f_text4)
        col = 0
        for row, data in enumerate(self.streamflow):
            if row == 0 :
                worksheet.write_row(row+21, col, data, f_text6)
            elif row ==31 :
                worksheet.write_row(row+21, col, data, f_text7)
            elif row in range(32, 38):
                worksheet.write_row(row+21, col, data, f_text8)
            else:
                worksheet.write_row(row+21, col, data, f_text1)

        worksheet.merge_range('A60:M60', self.footnote, f_text6)
        
        workbook.close()