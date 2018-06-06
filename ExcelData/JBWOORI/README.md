{
    "CapitalName" : "JB우리캐피탈",
    "ExcelFile" : "D:\\Rent\\Excel\\JBWOORI\\JBWOORI_20180421-UnLock.xlsm",
    "Worksheet" : "국산",
    "CarInfo" : {
        "Company" : {"Row" : 12, "Col" : 50 },
        "Model"   : {"Row" : 13, "Col" : 50 },
        "Trim"    : {"Row" : 14, "Col" : 50 }
    },
    "Price" : {
        "TotalPrice"  : {"Row" : 17, "Col" : 33 },    // 자동완성
        "BasePrice"   : {"Row" : 16, "Col" : 50 },    // 입력금액
        "OptionPrice" : {"Row" : 18, "Col" : 50 },    // 입력금액
        "OptionInfo"  : {"Row" : 17, "Col" : 50 }     // 문자열정보
    },
    "Payment" : {
        "Deposit"       : {"Row" : 22, "Col" : 9 },
        "PrePayment"    : {"Row" : 23, "Col" : 9 },
        "Duration"      : {"Row" : 21, "Col" : 9 },   // 36개월, 48개월, 60개월
        "MonthlyFee"    : {"Row" : 27, "Col" : 9 },
        "ResidualValue" : {"Row" : 24, "Col" : 12 },
        "ResidualRate"  : {"Row" : 24, "Col" : 9 }
    },
    "Extra" : [
        {"Row" : 25, "Col" : 9 },   // Index 0 / 약정거리 : 1만2천km , 2만km , 3만km , 무제한
        {"Row" : 9,  "Col" : 50 },  // Index 1 / CM 수수료
        {"Row" : 10, "Col" : 50 },  // Index 2 / AG 수수료
        {"Row" : 19, "Col" : 9 },   // Index 3 / 정비 : 순회정비, 입고정비, 정비제외
        {"Row" : 20, "Col" : 9 },   // Index 4 / 순회정비 - 프리미엄(순회 I), 디럭스+(순회 II),
                                                         디럭스(순회 III), 스탠다드(순회 IV)
                                    //           입고정비 - 디럭스(입고 I), 스탠다드(입고 II)
                                    //           정비제외 - 셀프
    ]
}
