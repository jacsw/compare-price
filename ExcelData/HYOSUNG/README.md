<사전작업>
1. SheetID 1에 대해서 보호 시트 해제 필요
2. 숨겨진 Sheet 군들에 대한 정보 확인 필요 : 
   * "잔가 / 잔가보장사_잔가" : 차량/트림별 잔가에 대한 정보를 포함하고 있음 

3. 잔가율 : CD21
   1> 원래식 : =VLOOKUP($CD$13,잔가보장사_잔가!$C$4:$F$165,2,0)
   2> 차량/트림별 잔가율에 대한 비율은 다름

4. Json Config
{
    "CapitalName" : "효성캐피탈",
    "ExcelFile" : "D:\\Rent\\Excel\\HYOSUNG\\HYOSUNG_20180418-UnLock.xlsx",
    "Worksheet" : "견적서및입력시트",
    "CarInfo" : {
        "Company" : {"Row" : 6, "Col" : 100 }, // 숫자 정보 입력 필요 : 1, 2, 4, 3, 5
        "Model"   : {"Row" : 13, "Col" : 82 }, // 차종 정보 입력
        "Trim"    : {"Row" : 15, "Col" : 82 }  // Trim (의미 없음)
    },
    "Price" : {
        "TotalPrice"  : {"Row" : 18, "Col" : 11 },  // 읽기 전용
        "BasePrice"   : {"Row" : 17, "Col" : 65 },
        "OptionPrice" : {"Row" : 17, "Col" : 82 },
        "OptionInfo"  : {"Row" : 19, "Col" : 82 }   // 입력 안하는게 좋음
    },
    "Payment" : {
        "Deposit"       : {"Row" : 38, "Col" : 65 },  // 단위 : %
        "PrePayment"    : {"Row" : 39, "Col" : 65 },  // 단위 : %
        "Duration"      : {"Row" : 36, "Col" : 65 },  // 기간 : 24, 36, 48, 60
        "MonthlyFee"    : {"Row" : 18, "Col" : 46 },
        "ResidualValue" : {"Row" : 18, "Col" : 25 },  // 단위 : 원
        "ResidualRate"  : {"Row" : 19, "Col" : 27 }   // 단위 : %
    },
    "Extra" : [
        {"Row" : 37, "Col" : 65 },   // 약정거리 : 20000, 30000, 40000, 50000
        {"Row" : 41, "Col" : 65 },   // CM 수수료 : 단위 - %
        {"Row" : 40, "Col" : 65 },   // AG 수수료 : 단위 - %
        {"Row" : 80, "Col" : 101 },  // 정비 : 1(VIP), 2(Premium),
                                             3(Standard), 4(Basic) - 기본값
        {"Row" : 82, "Col" : 102 }   // 정비 : 스노우 타이어 제공여부
                                     //      1(제공), 2(미제공)
    ]
}
