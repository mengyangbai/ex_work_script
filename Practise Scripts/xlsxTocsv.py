import csv
import xlrd
import glob, os
import msvcrt
from tqdm import tqdm
#author bai
firstline = ['C_CONSIGNMENT_ID',	'C_POST_CHARGE_TO_ACCOUNT',	'C_CHARGE_CODE',	'C_MERCHANT_CONSIGNEE_CODE',	'C_CONSIGNEE_NAME',	'C_CONSIGNEE_BUSINESS_NAME',	'C_CONSIGNEE_ADDRESS_1',	'C_CONSIGNEE_ADDRESS_2',	'C_CONSIGNEE_ADDRESS_3',	'C_CONSIGNEE_ADDRESS_4',	'C_CONSIGNEE_SUBURB',	'C_CONSIGNEE_STATE_CODE',	'C_CONSIGNEE_POSTCODE',	'C_CONSIGNEE_COUNTRY_CODE',	'C_CONSIGNEE_PHONE_NUMBER',	'C_PHONE_PRINT_REQUIRED',	'C_CONSIGNEE_FAX_NUMBER',	'C_DELIVERY_INSTRUCTION',	'C_SIGNATURE_REQUIRED',	'C_PART_DELIVERY',	'C_COMMENTS',	'C_ADD_TO_ADDRESS_BOOK',	'C_CTC_AMOUNT',	'C_REF',	'C_REF_PRINT_REQUIRED',	'C_REF2',	'C_REF2_PRINT_REQUIRED',	'C_CHARGEBACK_ACCOUNT',	'C_RECURRING_CONSIGNMENT',	'C_RETURN_NAME',	'C_RETURN_ADDRESS_1',	'C_RETURN_ADDRESS_2',	'C_RETURN_ADDRESS_3',	'C_RETURN_ADDRESS_4',	'C_RETURN_SUBURB',	'C_RETURN_STATE_CODE',	'C_RETURN_POSTCODE',	'C_RETURN_COUNTRY_CODE',	'C_REDIR_COMPANY_NAME',	'C_REDIR_NAME',	'C_REDIR_ADDRESS_1',	'C_REDIR_ADDRESS_2',	'C_REDIR_ADDRESS_3',	'C_REDIR_ADDRESS_4',	'C_REDIR_SUBURB',	'C_REDIR_STATE_CODE',	'C_REDIR_POSTCODE',	'C_REDIR_COUNTRY_CODE',	'C_MANIFEST_ID',	'C_CONSIGNEE_EMAIL',	'C_EMAIL_NOTIFICATION',	'C_APCN',	'C_SURVEY',	'C_DELIVERY_SUBSCRIPTION',	'C_EMBARGO_DATE',	'C_SPECIFIED_DATE',	'C_DELIVER_DAY',	'C_DO_NOT_DELIVER_DAY',	'C_DELIVERY_WINDOW',	'C_CDP_LOCATION',	'C_IMPORTERREFNBR',	'C_SENDER_NAME',	'C_SENDER_CUSTOMS_REFERENCE',	'C_SENDER_BUSINESS_NAME',	'C_SENDER_ADDRESS_LINE1',	'C_SENDER_ADDRESS_LINE2',	'C_SENDER_ADDRESS_LINE3',	'C_SENDER_SUBURB_CITY',	'C_SENDER_STATE_CODE',	'C_SENDER_POSTCODE',	'C_SENDER_COUNTRY_CODE',	'C_SENDER_PHONE_NUMBER',	'C_SENDER_EMAIL',	'C_RTN_LABEL',	'A_ACTUAL_CUBIC_WEIGHT',	'A_LENGTH',	'A_WIDTH',	'A_HEIGHT',	'A_NUMBER_IDENTICAL_ARTS',	'A_CONSIGNMENT_ARTICLE_TYPE_DESCRIPTION',	'A_IS_DANGEROUS_GOODS',	'A_IS_TRANSIT_COVER_REQUIRED',	'A_TRANSIT_COVER_AMOUNT',	'A_CUSTOMS_DECLARED_VALUE',	'A_CLASSIFICATION_EXPLANATION',	'A_EXPORT_CLEARANCE_NUMBER',	'A_IS_RETURN_SURFACE',	'A_IS_RETURN_AIR',	'A_IS_ABANDON',	'A_IS_REDIRECT_SURFACE',	'A_IS_REDIRECT_AIR',	'A_PROD_CLASSIFICATION',	'A_IS_COMMERCIAL_VALUE',	'G_ORIGIN_COUNTRY_CODE',	'G_HS_TARIFF',	'G_DESCRIPTION',	'G_PRODUCT_TYPE',	'G_PRODUCT_CLASSIFICATION',	'G_QUANTITY',	'G_WEIGHT',	'G_UNIT_VALUE',	'G_TOTAL_VALUE']

secondline =['IGNORED',	'OPTIONAL',	'MANDATORY',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'OPTIONAL',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'IGNORED',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'MANDATORY',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'OPTIONAL',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'OPTIONAL',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY',	'OPTIONAL',	'OPTIONAL',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE',	'MANDATORY/OPTIONAL REFER TO GUIDE']

dictone =['Customer Ref',	'Warehouse Code',	'Consignee Name',	'Consignee Phone',	'Consignee Address 1',	'Consignee Address 2',	'Consignee Address 3',	'Consignee Address 4',	'Consignee City',	'Consignee Province',	'Consignee Country',	'Consignee Zip',	'Billing Name',	'Billing Address 1',	'Billing Address 2',	'Billing Address 3',	'Billing Address 4',	'Billing City',	'Billing Province',	'Billing Country',	'Billing Zip',	'Carrier ID',	'Market Place',	'Total Order Value',	'Currency',	'Owner',	'SKU Code',	'SKU Description',	'QTY',	'Unit Price',	'PO',	'Instruction',	'Email',	'Shipping Method',	'Weight']

countryCode ={"AFGHANISTAN":"AF", "ALBANIA":"AL", "ALGERIA":"DZ", "ANDORIA AD ANGOLA":"AO", "ANGUILLA":"AI", "ANTIGUA AND BARBUDA":"AG", "ARGENTINA":"AR", "ARMENIA":"AM", "ARUBA":"AW", "ASCENSION AND ST HELENA":"SH", "AUSTRALIA":"AU", "AUSTRIA":"AT", "AZERBAIJAN":"AZ", "BAHAMAS":"BS", "BAHRAIN":"BH", "BANGLADESH":"BD", "BARBADOS":"BB", "BELARUS":"BY", "BELGIUM":"BE", "BELIZE":"BZ", "BENIN":"BJ", "BERMUDA":"BM", "BHUTAN":"BT", "BOLIVIA":"BO", "BOSNIA-HERZEGOVINA":"BA", "BOTSWANA":"BW", "BRAZIL":"BR", "BRITISH INDIAN OCEAN TERRITORY":"IO", "BRUNEI DARUSSALAM":"BN", "BULGARIA":"BG", "BURKINA FASO":"BF", "BURUNDI":"BI", "CAMBODIA":"KH", "CAMEROON":"CM", "CANADA":"CA", "CAPE VERDE":"CV", "CAYMAN ISLANDS":"KY", "CENTRAL AFRICAN REPUBLIC":"CF", "CHAD":"TD", "CHILE":"CL", "CHINA, PEOPLE'S REPUBLIC":"CN", "COLOMBIA":"CO", "COMOROS":"KM", "CONGO":"CG", "CONGO, DEMOCRATIC REPUBLIC OF":"CD", "COOK ISLANDS":"CK", "COSTA RICA":"CR", "CÔTE D'IVOIRE":"CI", "CROATIA":"HR", "CUBA":"CU", "CYPRUS":"CY", "CZECH REPUBLIC":"CZ", "DENMARK":"DK", "DJIBOUTI":"DJ", "DOMINICA":"DM", "DOMINICAN REPUBLIC":"DO", "EAST TIMOR (TIMOR-LESTE) TP ECUADOR":"EC", "EGYPT":"EG", "EL SALVADOR":"SV", "EQUATORIAL GUINEA":"GQ", "ERITREA":"ER", "ESTONIA":"EE", "ETHIOPIA":"ET", "FALKLAND ISLANDS":"FK", "FAROE ISLANDS FO FIJI":"FJ", "FINLAND":"FI", "FRANCE":"FR", "FRENCH GUIANA":"GF", "FRENCH POLYNESIA":"PF", "GABON":"GA", "GAMBIA":"GM", "GEORGIA":"GE", "GERMANY":"DE", "GHANA":"GH", "GIBRALTAR":"GI", "GREECE":"GR", "GREENLAND":"GL", "GRENADA":"GD", "GUADELOUPE":"GP", "GUAM":"GU", "GUATEMALA":"GT", "GUINEA":"GN", "GUINEA-BISSAU":"GW", "GUYANA":"GY", "HAITI HT HOLY SEE (VATICAN CITY STATE) VA HONDURAS":"HN", "HONG KONG":"HK", "HUNGARY":"HU", "ICELAND":"IS", "INDIA":"IN", "INDONESIA":"ID", "IRAN, ISLAMIC REPUBLIC":"IR", "IRAQ":"IQ", "IRELAND":"IE", "ISRAEL":"IL", "ITALY":"IT", "JAMAICA":"JM", "JAPAN":"JP", "JORDAN":"JO", "KAZAKHSTAN":"KZ", "KENYA":"KE", "KIRIBATI":"KI", "KOREA, DEMOCRATIC PEOPLE'S REPUBLIC":"KP", "KOREA, REPUBLIC":"KR", "KUWAIT":"KW", "KYRGYZSTAN":"KG", "LAO, PEOPLE'S DEMOCRATIC REPUBLIC":"LA", "LATVIA":"LV", "LEBANON":"LB", "LESOTHO":"LS", "LIBERIA":"LR", "LIECHTENSTEIN":"LI", "LITHUANIA":"LT", "LUXEMBOURG":"LU", "MACAO":"MO", "MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF":"MK", "MADAGASCAR":"MG", "MALAWI":"MW", "MALAYSIA":"MY", "MALDIVES":"MV", "MALI":"ML", "LIBYAN ARAB JAMAHIRIYA":"LY", "MALTA":"MT", "MARIANA ISLANDS":"MP", "MARSHALL ISLANDS":"MH", "MARTINIQUE":"MQ", "MAURITANIA":"MR", "MAURITIUS":"MU", "MAYOTTE YT MEXICO":"MX", "MICRONESIA, FEDERATED STATES OF":"FM", "MOLDOVA":"MD", "MONACO MC MONGOLIA":"MN", "MONTENEGRO ME MONTSERRAT":"MS", "MOROCCO":"MA", "MOZAMBIQUE":"MZ", "MYANMAR":"MM", "NAMIBIA":"NA", "NAURU":"NR", "NEPAL":"NP", "NETHERLANDS":"NL", "NETHERLANDS ANTILLES AND ARUBA":"AN", "NEW CALEDONIA":"NC", "NEW ZEALAND":"NZ", "NICARAGUA":"NI", "NIGER":"NE", "NIGERIA":"NG", "NIUE ISLAND":"NU", "NORWAY":"NO", "OMAN":"OM", "PAKISTAN":"PK", "PALAU, REPUBLIC OF":"PW", "PANAMA, REPUBLIC OF":"PA", "PAPUA NEW GUINEA":"PG", "PARAGUAY":"PY", "PERU":"PE", "PHILIPPINES":"PH", "PITCAIRN ISLANDS":"PN", "POLAND, REPUBLIC OF":"PL", "PORTUGAL":"PT", "PUERTO RICO":"PR", "QATAR":"QA", "REUNION":"RE", "ROMANIA":"RO", "RUSSIAN FEDERATION":"RU", "RWANDA":"RW", "SAINT CHRISTOPHER (ST KITTS) AND NEVIS":"KN", "SAMOA, AMERICAN":"AS", "SAMOA, WESTERN":"WS", "SAN MARINO SM SAO TOME AND PRINCIPE":"ST", "SAUDI ARABIA, KINGDOM OF":"SA", "SENEGAL":"SN", "SERBIA":"RS", "SEYCHELLES":"SC", "SIERRA LEONE":"SL", "SINGAPORE":"SG", "SLOVAKIA":"SK", "SLOVENIA":"SI", "SOLOMON ISLANDS":"SB", "SOMALIA":"SO", "SOUTH AFRICA":"ZA", "SPAIN":"ES", "SRI LANKA":"LK", "ST LUCIA":"LC", "ST PIERRE AND MIQUELON":"PM", "ST VINCENT AND THE GRENADINES":"VC", "SUDAN":"SD", "SURINAME":"SR", "SWAZILAND":"SZ", "SWEDEN":"SE", "SWITZERLAND":"CH", "SYRIAN ARAB REPUBLIC":"SY", "TAIWAN":"TW", "TAJIKISTAN":"TJ", "TANZANIA":"TZ", "THAILAND":"TH", "TOGO":"TG", "TOKELAU":"TK", "TONGA":"TO", "TRINIDAD AND TOBAGO":"TT", "TRISTAN DA CUNHA TA TUNISIA":"TN", "TURKEY":"TR", "TURKMENISTAN":"TM", "TURKS AND CAICOS ISLANDS":"TC", "TUVALU":"TV", "UGANDA":"UG", "UKRAINE":"UA", "UNITED ARAB EMIRATES":"AE", "UNITED KINGDOM":"GB", "UNITED STATES":"US", "UNITED STATES MINOR OUTLYING ISLANDS UM URUGUAY":"UY", "UZBEKISTAN":"UZ", "VANUATU":"VU", "VENEZUELA":"VE", "VIETNAM":"VN", "VIRGIN ISLANDS, BRITISH":"VG", "VIRGIN ISLANDS, USA":"VI", "WALLIS AND FUTUNA ISLANDS":"WF", "YEMEN":"YE", "ZAMBIA":"ZM", "ZIMBABWE":"ZW"}

def getstr(rx,str,sh):
    return sh.cell_value(rx,dictone.index(str))

def xlsxToCsv(filename,dir):
    output = filename.replace(" ", "").rstrip(filename[-5:])
    filename = dir + "\\" + filename    
    print("开始转换 "+filename)
    book = xlrd.open_workbook(filename)
    sh = book.sheet_by_index(0)
    rows = []
    refNodict={}    
    for rx in tqdm(range(sh.nrows)):
        if rx != 0:
            #先检查有没有skucode
            refNo = getstr(rx,"Customer Ref",sh)       
            if refNo not in refNodict:
                #没有就新建
                skuCode = getstr(rx,"SKU Code",sh)
                tmpstr = "{:.0f}".format(skuCode) 
                refNodict[refNo] = tmpstr
                oneline = []
                oneline.append("")
                oneline.append("")
                oneline.append("ECM8")
                oneline.append("")
                oneline.append(getstr(rx,"Consignee Name",sh))
                oneline.append("")
                oneline.append(getstr(rx,"Consignee Address 1",sh))
                oneline.append(getstr(rx,"Consignee Address 2",sh))
                oneline.append(getstr(rx,"Consignee Address 3",sh))
                oneline.append(getstr(rx,"Consignee Address 4",sh))
                oneline.append(getstr(rx,"Consignee City",sh))
                oneline.append(getstr(rx,"Consignee Province",sh))
                oneline.append(getstr(rx,"Consignee Zip",sh))
                country = getstr(rx,"Consignee Country",sh)
                try:
                    if country is not '':
                        oneline.append(countryCode[country.upper()])
                except KeyError:
                    oneline.append(getstr(rx,"Consignee Country",sh))        
                oneline.append(getstr(rx,"Consignee Phone",sh))
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append(getstr(rx,"Customer Ref",sh))
                oneline.append("Y")
                oneline.append(tmpstr)
                oneline.append("Y")
                oneline.append("")
                oneline.append("")
                oneline.append("EWE")
                oneline.append("2/21 Worth Street")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("Chullora")
                oneline.append("NSW")
                oneline.append("2190")
                oneline.append("AU")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append(getstr(rx,"Email",sh))
                oneline.append("DESPATCH")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("EWE")
                oneline.append("")
                oneline.append("")
                oneline.append("2/21 Worth Street")
                oneline.append("")
                oneline.append("")
                oneline.append("Chullora")
                oneline.append("NSW")
                oneline.append("2190")
                oneline.append("AU")
                oneline.append("0061 2 9644 2648")
                oneline.append("")
                oneline.append("")
                try:
                    oneline.append(getstr(rx,"Weight",sh))
                except IndexError:
                    oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("GIFT")
                oneline.append("")
                oneline.append("AU")
                oneline.append("")
                oneline.append(getstr(rx,"SKU Description",sh))
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                oneline.append("")
                rows.append(oneline)
            else:
                #有就修改
                skuCode = getstr(rx,"SKU Code",sh)
                tmpstr = "{:.0f}".format(skuCode) 
                refNodict[refNo] ="%s;%s"%(refNodict[refNo],tmpstr)
                for alterLine in rows:
                    if  alterLine[23] == refNo:
                        alterLine[25] = refNodict[refNo]
    outputfile = dir + "\\"+output+".csv"
    with open(outputfile,'w',newline='') as f:
        f_csv = csv.writer(f)
        f_csv.writerow(firstline)
        f_csv.writerow(secondline)
        f_csv.writerows(rows)
    
    print(outputfile+" 输出完成")

if __name__ == '__main__':
    print("Xlsx转换csv开始！")
    dir = 'xlsxTocsv'
    try:
        files = os.listdir(dir)
    except FileNotFoundError:
        print("请把待转换的xlsx文件放到程序的xlsxTocsv\目录下！")
        ord(msvcrt.getch())
        quit()
    n=0
    for file in files:
        if file.endswith(".xlsx"):
            xlsxToCsv(file,dir)
            n+=1
    if n==0:
        print("未发现扩展名为xlsx的文件，按D键退出")
    else:
        print("共计"+str(n)+"个文件，按D键退出")
    while True:
        if ord(msvcrt.getch()) in [68, 100]:
            break
