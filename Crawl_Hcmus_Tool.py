from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from time import sleep
import xlsxwriter
import time

while (True) : #Create Excel File
    file_exl = xlsxwriter.Workbook("data_hcmus_k19.xlsx")
    f = file_exl.add_worksheet("data_k19")
    items = ["MSSV", "Tên", "Giới tính", "Năm sinh", "Ngành", "Khoa", 
    "Khối", "Môn1", "Môn2", "Môn3", "Điểm cộng", "Tổng", "SBD", "Năm TN", "Xếp loại", "CMND",
    "Ngày cấp", "SĐT", "SĐT2", "Thường trú", "Tạm trú", "EmailCN", "EmailSV", 
    "Người LL", "ĐC người LL", "SĐT người LL", "Email người LL", "Ngân hàng",
    "Số tài khoản", "Chi nhánh"]
    n = len(items)
    for i in range(0,n) :
        header =(i>25 and "A" + chr(65+i-26) or chr(65+i)) + "1"
        f.write(header,items[i])
    break

while (True) : #Login
    br = webdriver.Chrome(executable_path="chromedriver")
    br.get("https://portal4.hcmus.edu.vn")
    txtUser = br.find_element_by_id("ctl00_ContentPlaceHolder1_txtUsername")
    txtUser.send_keys("20280115")
    txtPass = br.find_element_by_id("ctl00_ContentPlaceHolder1_txtPassword")
    txtPass.send_keys("qfdczs11")
    while (br.current_url != "https://portal4.hcmus.edu.vn/") : sleep(2)
    sleep(0.5)
    br.get("https://portal4.hcmus.edu.vn/Default.aspx?pid=62")
    sleep(0.5)
    br.get("https://portal4.hcmus.edu.vn/Default.aspx?pid=127&t=v&uid=72057594046957354&rel=1")
    h = 1
    print("\nStart Crawling....\n")
    break

start = 56871
end = 60099

for x in range(start, end) : #crawl and save
    a = []
    x = str(x)
    link_1 = "https://portal4.hcmus.edu.vn/default.aspx?pid=126&t=v&uid=720575940468"+x+"&rel=1"
    br.get(link_1)
    if  br.find_elements_by_xpath("//div[@class='ob_iTC']") != [] :
        name = br.find_element_by_xpath("//div[@class='grid_9']/h1")
        mssv = name.get_attribute("innerText")
        name = mssv[22:]
        mssv = mssv[11:19]
        age = br.find_element_by_id("ctl00_MainContent_ctl00_txtNgaySinh").get_attribute("value")
        sex = br.find_element_by_id("ctl00_MainContent_ctl00_radNam").get_attribute("checked") and "Nam" or "Nữ"     
        faculty = br.find_element_by_id("ctl00_MainContent_ctl00_txtKhoa").get_attribute("value")
        id_ = br.find_element_by_id("ctl00_MainContent_ctl00_cboNganhDaoTao_ctl00_MainContent_ctl00_cboNganhDaoTao").get_attribute("value")[-3:]
        for i in br.find_elements_by_xpath("//div[@id='ctl00_MainContent_ctl00_cboNganhDaoTao_ob_CbocboNganhDaoTaoItemsContainer']/div[2]/ul/li"):
            tmp = (i.get_attribute("innerText")) 
            if tmp[-3:] == id_ :
                major = tmp[:-17]
                break
        print(h, name)
        h += 1
        a.extend([mssv, name, age, sex, faculty, major])
        ######################################

        link_2 = "https://portal4.hcmus.edu.vn/default.aspx?pid=155&t=v&uid=720575940469"+x+"&rel=1"
        br.get(link_2)
        sbd = br.find_element_by_id("ctl00_MainContent_ctl00_txtSoBaoDanh").get_attribute("value")
        khoi = br.find_element_by_id("ctl00_MainContent_ctl00_txtKhoi").get_attribute("value")
        mon1 = br.find_element_by_id("ctl00_MainContent_ctl00_txtMon1").get_attribute("value")
        mon2 = br.find_element_by_id("ctl00_MainContent_ctl00_txtMon2").get_attribute("value")
        mon3 = br.find_element_by_id("ctl00_MainContent_ctl00_txtMon3").get_attribute("value")
        tong = float(br.find_element_by_id("ctl00_MainContent_ctl00_txtTongDiem").get_attribute("value"))
        plus = float('0' + br.find_element_by_id("ctl00_MainContent_ctl00_txtDiemThuong").get_attribute("value"))
        tn = br.find_element_by_id("ctl00_MainContent_ctl00_txtNamTN").get_attribute("value")
        xl = br.find_element_by_id("ctl00_MainContent_ctl00_cboXLHT_ob_CbocboXLHTTB").get_attribute("value")

        a.extend([khoi, mon1, mon2, mon3, plus, tong + plus, sbd, tn, xl])
        ###############################
        link_3 = "https://portal4.hcmus.edu.vn/default.aspx?pid=127&t=v&uid=720575940469"+x+"&rel=1"
        br.get(link_3)
        cmnd = br.find_element_by_id("ctl00_MainContent_ctl00_txtCMND").get_attribute("value")
        date = br.find_element_by_id("ctl00_MainContent_ctl00_txtCMNDNgayCap").get_attribute("value")
        sdt = br.find_element_by_id("ctl00_MainContent_ctl00_txtDienThoaiNha").get_attribute("value")
        sdt2 = br.find_element_by_id("ctl00_MainContent_ctl00_txtDienThoaiDD").get_attribute("value")
        addr = br.find_element_by_id("ctl00_MainContent_ctl00_txtFullDiaChiThuongTru").get_attribute("value")
        add = br.find_element_by_id("ctl00_MainContent_ctl00_txtDiaChiTamTru").get_attribute("value")
        emsv = br.find_element_by_id("ctl00_MainContent_ctl00_txtEmail").get_attribute("value")
        emcn = br.find_element_by_id("ctl00_MainContent_ctl00_txtEmailCaNhan").get_attribute("value")
        ll = br.find_element_by_id("ctl00_MainContent_ctl00_txtNguoiLL").get_attribute("value")
        addll = br.find_element_by_id("ctl00_MainContent_ctl00_txtDCNguoiLL").get_attribute("value")
        sdtll = br.find_element_by_id("ctl00_MainContent_ctl00_txtDTNguoiLL").get_attribute("value")
        emll = br.find_element_by_id("ctl00_MainContent_ctl00_txtEmailNguoiLL").get_attribute("value")
        bank = br.find_element_by_id("ctl00_MainContent_ctl00_txtNganHang").get_attribute("value")
        stk = br.find_element_by_id("ctl00_MainContent_ctl00_txtSoThe").get_attribute("value")
        cn = br.find_element_by_id("ctl00_MainContent_ctl00_txtChiNhanh").get_attribute("value")
        a.extend([cmnd, date, sdt, sdt2, addr, add, emsv, emcn, ll, sdtll, addll, emll, bank, stk, cn])
        for i in range(0,n) :
            header =(i>25 and "A" + chr(65+i-26) or chr(65+i)) + str(h)
            f.write(header,a[i])


file_exl.close()
br.close
br.quit
print("\nThe program is completed at :", time.strftime("%H:%M:%S",time.localtime()),'\n')


