from libary import *

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
    print("\nStart Crawling....\n")
    break
links = []
def run(start, end) :
    file_exl = xlsxwriter.Workbook("lin10000.xlsx")
    f = file_exl.add_worksheet("data")
    end = end - 1
    h = 1
    t = int(math.log10(end))
    for i in range (start, end) :
        k = int(math.log10(i))
        s = "0" * (t - k)
        x = s + str(i)
        link ="https://portal4.hcmus.edu.vn/Default.aspx?pid=126&t=v&uid=7205759404" + x + "000&rel=1"
        br.get(link)
        if  br.find_elements_by_xpath("//div[@class='ob_iTC']") != [] :
            print(link)
            links.append(link)
        link ="https://portal4.hcmus.edu.vn/Default.aspx?pid=126&t=v&uid=7205759404"+ x + "500&rel=1"
        br.get(link)
        if  br.find_elements_by_xpath("//div[@class='ob_iTC']") != [] :
            print(link)
            links.append(link)
        print(i)
    for i in links :
         f.write("A"+str(h), i)
         h = h + 1
    file_exl.close()

run(1, 5000)  


br.close
br.quit
os.system("shutdown /s /t 15")