import pandas as pd
import datetime

# a = pd.read_excel("E02-2023.xlsx", sheet_name="ECART T2")
def handler(startTimeCheck, endTimeCheck, checkMonth) -> (bool, str):
    dataFinal = pd.DataFrame()
    try:
        for loopRangeDate in range(int(startTimeCheck), int(endTimeCheck)+1):
            timeCheck = str(loopRangeDate).zfill(2) + "-" + str(checkMonth).zfill(2)
            dataRaw = pd.read_excel("ECART 02-2023.xlsx", sheet_name=timeCheck)
            dtf = pd.DataFrame()
            for lineToJoinDataFrame in dataRaw.values[2:]:
                dtf = pd.concat([dtf, pd.DataFrame.from_records(
                    [{'MSTN': lineToJoinDataFrame[0], timeCheck: lineToJoinDataFrame[1]}])], ignore_index=True)

            dtf = dtf.dropna()
            dtf = dtf.astype({'MSTN': str, timeCheck: float})

            if dataFinal.empty:
                dataFinal = dtf
            else:
                dataFinal = dataFinal.merge(dtf, how='outer', on="MSTN")
        return True, "Success"
    except Exception:
        return False, f"Sai định dạng thời gian: {startTimeCheck}-{checkMonth}; {endTimeCheck}-{checkMonth}"

if __name__ == "__main__":
    startTimeCheck = input("Nhap ngay bat dau (type: DD):")
    endTimeCheck = input("Nhap ngay ket thuc (type: DD):")
    checkMonth = datetime.datetime.now().month - 1
    status, resultData = handler(startTimeCheck, endTimeCheck, checkMonth)
    print(status)
    print(resultData)
