{
    "CapitalName" : "하나캐피탈",
    "ExcelFile" : "D:\\Rent\\Excel\\HANA\\HANA_20180310.xlsx",
    "Worksheet" : "(국산차)견적",
    "CarInfo" : {
        "Company" : {"Row" : 8, "Col" : 59 },
        "Model"   : {"Row" : 8, "Col" : 76 },
        "Trim"    : {"Row" : 14, "Col" : 59 }
    },
    "Price" : {
        "TotalPrice"  : {"Row" : 16, "Col" : 35 },
        "BasePrice"   : {"Row" : 14, "Col" : 59 },
        "OptionPrice" : {"Row" : 14, "Col" : 76 },
        "OptionInfo"  : {"Row" : 16, "Col" : 76 }
    },
    "Payment" : {
        "Deposit"       : {"Row" : 20, "Col" : 8 },    // 비율 : 0 ~ 100% / 문자열 형식으로만 입력
        "PrePayment"    : {"Row" : 34, "Col" : 59 },   // 금액 (비율 X)
        "Duration"      : {"Row" : 19, "Col" : 8 },    // 36, 48, 60 / 문자열 형식으로만 입력
        "MonthlyFee"    : {"Row" : 30, "Col" : 8 },
        "ResidualValue" : {"Row" : 26, "Col" : 11 },
        "ResidualRate"  : {"Row" : 26, "Col" : 8 }
    },
    
    "Extra" : [
        {"Row" : 24, "Col" : 8 },    // "1만km", "2만km", "3만km", "4만km", "5만km", "제한없음"
        {"Row" : 48, "Col" : 49 },
        {"Row" : 48, "Col" : 62 },
        {"Row" : 23, "Col" : 8 }     // "Self service", "Semi service", "Standard service", "Special service"
    ]
}
