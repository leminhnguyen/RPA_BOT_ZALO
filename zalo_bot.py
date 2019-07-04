from ZaloFunctions import *
from extract_data import ExtractData
from PIL import ImageTk,Image
import codecs

class Zalo(ExtractData):
    ZALO_LOGIN = "https://chat.zalo.me/"
    #Data table parameters
    TBL_FIRST_ROW = 5                       #first row of the table
    TBL_FIRST_COL = 'A'

    #Web parameter
    STEP_WAIT = 1

    #name of report file
    STATUS_FILE = 'status_summary.xlsx'

    #summary field
    SUMMARY_FILEDS =[
        'Ngày giờ gửi tin',
        'Trạng thái gửi tin',
        'Ngày giờ trạng thái (nếu có)',
        'Ghi chú',
    ]
    #GUI
    IMG_BROWSE = CURRENT_DIRECTORY + "/images/browse.png"
    IMG_MESSAGE = CURRENT_DIRECTORY + "/images/message.png"
    IMG_QUIT = CURRENT_DIRECTORY + "/images/quit.png"
    IMG_RESULT = CURRENT_DIRECTORY + "/images/result.png"
    HEIGHT = 30
    ENTRY_WIDTH = 100
    FONT        = ("Arial", 10)

    def __init__(self):
        ExtractData.__init__(self)
        #file containning nick zalo
        self.nicks_path = ""
        self.nicks_ws = None
        #file containning must-send data
        self.data_path = ""
        self.data_wb = None
        #summary file
        self.status_summary = ""
        #Set up chrome profile
        self.browser = None
        self.contact_found = False

    @staticmethod
    def resize_image(image, maxsize):
        r1 = image.size[0]/maxsize[0] # width ratio
        r2 = image.size[1]/maxsize[1] # height ratio
        ratio = max(r1, r2)
        newsize = (int(image.size[0]/ratio), int(image.size[1]/ratio))
        image = image.resize(newsize, Image.ANTIALIAS)
        return image

    def browse_nicks(self):
        #file containning nick zalo
        path = get_excel_file('Chọn file chứa nick Zalo')
        if path != "":
            self.nicks_path = path

    def validate_zalo_file(self):
        if self.nicks_path == "":
            messagebox.showerror("Lỗi", "Bạn chưa chọn file chứa nick Zalo")
        else:
            self.nicks_ws = openpyxl.load_workbook(self.nicks_path).active

    @staticmethod
    def create_driver(chrome_driver_path = 'chromedriver.exe'):
        default_cookie_path = os.path.join(
            os.environ.get("USERPROFILE"),
            r"AppData\Local\Google\Chrome\User Data\Default\Cookies")
        default_cookie_path = default_cookie_path.replace("\\", "/")
        tmp_profile_dir = TemporaryDirectory().name
        tmp_profile_path = os.path.join(tmp_profile_dir, "Default")
        os.makedirs(tmp_profile_path)
        tmp_cookie_path = os.path.join(tmp_profile_path, "Cookies")
        shutil.copy(default_cookie_path, tmp_cookie_path)
        chrome_options = Options()
        chrome_options.add_argument(f"user-data-dir={tmp_profile_dir}")
        chrome_options.add_argument("disable-infobars")
        chrome_options.add_argument("launch-simple-browser")
        chrome_options.add_argument("start-maximized")
        driver = webdriver.Chrome(
            executable_path=chrome_driver_path, options=chrome_options)
        return driver

    def login(self):
        """
        """
        try:
            if self.browser.current_url == self.ZALO_LOGIN:
                return
        except:
            pass
        # profile_path = "C:/Users/Admin/AppData/Roaming/Mozilla/Firefox/Profiles/hy4g9l29.default"
        # self.browser = webdriver.Firefox(FirefoxProfile(profile_path))
        self.browser = self.create_driver()
        self.browser.get(self.ZALO_LOGIN)
        while self.browser.current_url != self.ZALO_LOGIN:
            message('Thông báo', 'Xin vui lòng đăng nhập Zalo và bấm Ctrl+q để tiếp tục')
            keyboard.wait('ctrl+q')

    def _get_range(self, ws):
        """
        return the range storing table in a worksheet
        """
        last_col = pre_char(h_empty_cell(ws, row = self.TBL_FIRST_ROW))
        last_row = v_empty_cell(ws, col = self.TBL_FIRST_COL) - 1
        return 'A5:' + last_col + str(last_row)

    def _find_contact(self, nick):
        """
        find contact and go to message
        contact_found = True if contact is found
        """
        self.contact_found = True
        contact = self.browser.find_element_by_id('contact-search-input')
        contact.clear()
        contact.send_keys(nick)
        contact.click()
        time.sleep(self.STEP_WAIT)
        search_result = self.browser.find_elements_by_class_name("global-search-no-result")
        if len(search_result) != 0:
            #contact not found
            self.contact_found = False
            return
        contact.send_keys(Keys.ENTER)
        time.sleep(self.STEP_WAIT)

    def _paste_to_contact(self, nick):
        """
        paste data in clipboard in a zalo contact
        """
        self._find_contact(nick)
        if not self.contact_found:
            return
        chat = self.browser.find_element_by_id('richInput')
        chat.send_keys(Keys.CONTROL, 'v')
        time.sleep(self.STEP_WAIT)
        self.browser.find_element_by_css_selector(
            '.btn.btn-txt.btn-primary.btn-modal-action').click()
        time.sleep(self.STEP_WAIT)
        clear_clipboard()

    def _copy_and_paste(self, sheet_name, data_range, nick):
        """
        To fix permission error in fee_summary
        """
        xlwb = self.excel.Workbooks.Open(self.data_path)
        try:
            xlwb.Worksheets(sheet_name).Range(data_range).Copy()
            self._paste_to_contact(nick)
        except:
            pass
        xlwb.Close(True)

    def send_data(self):
        if self.ws is None:
            return
        self.create_worksheets_to_send()
        self.data_path = self.fee_summary
        self.data_wb = openpyxl.load_workbook(self.fee_summary)

        #summary file
        status_summary = self.data_path.split('/')[:-1]
        status_summary.append(self.STATUS_FILE)
        self.status_summary = '/'.join(status_summary)

        last_row = v_empty_cell(self.nicks_ws, col = 'B')
        for r in range(2, last_row, 1):
            group_name = self.nicks_ws["A" + str(r)].value
            nick = self.nicks_ws["B" + str(r)].value
            if group_name in [None, ""] or nick in [None, ""]:
                continue

            #check whether a sheet corresponds to a group
            for sheet_name in self.data_wb.sheetnames:
                if normalize(sheet_name) in normalize(group_name):
                    #if sheet name is found
                    data_range = self._get_range(self.data_wb[sheet_name])
                    self._copy_and_paste(sheet_name, data_range, nick)
                    print(group_name + ' : ' + nick, self.contact_found)

    def _get_status(self, nick):
        """
        return receipt acknowledgement of each nick
        """
        summary = dict()
        for field in self.SUMMARY_FILEDS:
            summary.update({field: ""})
        self._find_contact(nick)
        time.sleep(self.STEP_WAIT)
        if not self.contact_found:
            summary.update({"Ghi chú": "Không tìm thấy nick zalo"})
            return summary

        chat_date = self.browser.find_elements_by_class_name('chat-date')
        if len(chat_date) !=0:
            summary.update({"Ngày giờ gửi tin": chat_date[-1].get_attribute('textContent')})

        send_status = self.browser.find_elements_by_class_name('card-send-status')
        if len(send_status) != 0:
            summary.update({"Trạng thái gửi tin": send_status[-1].get_attribute('textContent')})

        receipt_time = self.browser.find_elements_by_class_name('card-send-time__sendTime')
        if len(receipt_time) != 0:
            summary.update({"Ngày giờ trạng thái (nếu có)": receipt_time[-1].get_attribute('textContent')})
        return summary

    def _create_status_summary(self):
        """
        create header for summary file
        """
        wb = openpyxl.load_workbook(self.nicks_path)
        ws = wb.active
        for i, value in enumerate(self.SUMMARY_FILEDS):
            ws.cell(row = 1, column = i+3).value = value
        save_excel(self.status_summary, wb)

    def report_status(self):
        """
        summarize sending status
        """
        #create report file
        self._create_status_summary()
        wb = openpyxl.load_workbook(self.status_summary)
        ws = wb.active
        last_row = v_empty_cell(ws, col = 'B')
        for r in range(2, last_row, 1):
            found = False
            group_name = ws["A" + str(r)].value
            nick = ws["B" + str(r)].value
            if group_name in [None, ""] or nick in [None, ""]:
                continue
            #check whether a sheet corresponds to a group
            for sheet_name in self.data_wb.sheetnames:
                if normalize(sheet_name) in normalize(group_name):
                    found = True
                    summary = self._get_status(nick)
                    ws['C' + str(r)] = summary['Ngày giờ gửi tin']
                    ws['D' + str(r)] = summary['Trạng thái gửi tin']
                    ws['E' + str(r)] = summary['Ngày giờ trạng thái (nếu có)']
                    ws['F' + str(r)] = summary['Ghi chú']
                    break
            if not found:
                ws['F' + str(r)] = 'Không tìm thấy sheet chứa dữ liệu'
        save_excel(self.status_summary, wb)
        message('Thông báo', 'Đã gửi tin nhắn thành công')

    def close(self):
        self.browser.quit()

class Gui(Zalo):
    SAVED_PATHS = CURRENT_DIRECTORY + "/saved_paths.txt"

    def __init__(self):
        print (self.SAVED_PATHS)
        Zalo.__init__(self)
        self.input_ok = True
        self.root = tk.Tk()
        self.root.grid()
        paths = self._get_saved_paths()
        print(paths)
        self.nicks_path, self.file_path = r'D:\Documents\University\PROJECT\Python_RPA\bot_zalo\Bot Zalo_v0\nick.xlsx', paths[1]

        self.btn_open = None

        photo = self.resize_image(Image.open(self.IMG_BROWSE),[self.HEIGHT,self.HEIGHT])
        self.img_browse = ImageTk.PhotoImage(photo)

        photo = self.resize_image(Image.open(self.IMG_MESSAGE),[self.HEIGHT,self.HEIGHT])
        self.img_message = ImageTk.PhotoImage(photo)

        photo = self.resize_image(Image.open(self.IMG_QUIT),[self.HEIGHT,self.HEIGHT])
        self.img_quit = ImageTk.PhotoImage(photo)
        photo = self.resize_image(Image.open(self.IMG_RESULT),[self.HEIGHT,self.HEIGHT])
        self.img_result = ImageTk.PhotoImage(photo)

        self.nick_entry = tk.Entry(self.root, state = tk.NORMAL, width = self.ENTRY_WIDTH)
        self.nick_entry.insert(0, self.nicks_path)
        self.nick_entry.config(state = "readonly")

        self.fee_entry = tk.Entry(self.root, state = tk.NORMAL, width = self.ENTRY_WIDTH)
        self.fee_entry.insert(0, self.file_path)
        self.fee_entry.config(state = "readonly")

    @staticmethod
    def _get_saved_paths(path = SAVED_PATHS):
        f = open(path, encoding='utf-8', mode='r')
        st = f.read()
        print(st)
        f.close()
        return st.split("\n")

    def _save_paths(self, path = SAVED_PATHS):
        f = codecs.open(path, encoding='utf-8', mode='w')
        f.write(self.nicks_path + "\n" + self.file_path)
        f.close()

    def gui_browse_nick(self):
        self.browse_nicks()
        self.nick_entry.config(state = tk.NORMAL)
        self.nick_entry.delete(0, tk.END)
        self.nick_entry.insert(0, self.nicks_path)
        self.nick_entry.config(state = "readonly")
        self._save_paths()

    def gui_browse_fee(self):
        self.browse_file()
        self.fee_entry.config(state = tk.NORMAL)
        self.fee_entry.delete(0, tk.END)
        self.fee_entry.insert(0, self.file_path)
        self.fee_entry.config(state = "readonly")
        self._save_paths()

    def _format_gui(self):
        self.root.geometry("900x150")
        self.root.title("Zalo Automation")

    def open_status(self):
        os.startfile(self.status_summary)

    def check_input(self):
        self.validate_zalo_file()
        self.validate_fee_data()
        if self.nicks_ws is None or self.ws is None:
            self.input_ok = False

    def login_and_send(self):
        self.check_input()
        if not self.input_ok:
            return
        self.login()
        self._hide_excel()
        self.send_data()
        self._show_excel()
        self.report_status()
        if self.status_summary != "":
            self.btn_open.config(state = tk.NORMAL)
        else:
            self.btn_open.config(state = tk.DISABLED)
        self.browser.get(self.ZALO_LOGIN)

    def quit(self):
        self.root.quit()
        self._show_excel()
        exit()

    def main(self):
        self._format_gui()
        self.nick_entry.grid(row = 1, column = 2)
        self.fee_entry.grid(row = 3, column = 2)
        label_nick = Label(
            self.root,
            text = "Chọn file chứa nick Zalo",
            font = self.FONT
        )
        label_nick.grid(row = 1, column = 0, sticky = tk.W)

        btn_get_nick = Button(
            self.root,
            image = self.img_browse,
            command = lambda: self.gui_browse_nick(),
            height  = self.HEIGHT + self.FONT[1],
            font = self.FONT
        )
        btn_get_nick.grid(row = 1, column = 1, sticky = tk.E)

        label_fee = Label(
            self.root,
            text = "Chọn file tổng hợp thu phí",
            font = self.FONT
        )
        label_fee.grid(row = 3, column = 0, sticky = tk.W)

        btn_get_data = Button(
            self.root,
            image = self.img_browse,
            command = lambda: self.gui_browse_fee(),
            font = self.FONT
        )
        btn_get_data.grid(row = 3, column = 1, sticky = tk.E)

        btn_send = Button(
            self.root,
            text = "Gửi tin nhắn",
            image = self.img_message,
            compound = tk.TOP,
            command = lambda: self.login_and_send()
        )
        btn_send.grid(row = 4, column = 1, sticky = tk.S)

        if self.status_summary == "":
            state = tk.DISABLED
        else:
            state = tk.NORMAL

        self.btn_open = Button(
            self.root,
            text = "Mở file tổng hợp tin nhắn",
            image = self.img_result,
            compound = tk.TOP,
            command = lambda: self.open_status(),
            state = state
        )
        self.btn_open.grid(row = 4, column = 2)

        btn_quit = Button(
            self.root,
            text = "Thoát",
            image = self.img_quit,
            compound = tk.TOP,
            command = lambda: self.quit()
        )
        btn_quit.grid(row = 4, column =3)

if __name__ == '__main__':
    bim = Gui()
    bim.main()
    bim.root.mainloop()




