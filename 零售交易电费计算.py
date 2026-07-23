import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')

#中长期电价分摊参数（元/MWh）
AUX_SERVICE_FEE = 0 #辅助服务费用分摊费
OPERATION_FEE = 0 #运行成本类市场运行费用分摊费
CAP_PRICE_FEE = 0 #封顶价格上浮数值及固定价格

TARGET_MONTH = "x月" #计算月份
TOTAL_DAYS = 0 #当月总天数

#文件路径配置（从平台导出的应严格按原格式）
USAGE_FILE_DIR = r""      #小时级用电量文件（当日用电量详情 (1).xlsx...）
SPOT_PRICE_FILE = r""     #现货价格文件（实时+日前表）
RETAIL_PRICE_FILE = r""   #零售市场参考价文件
MONTH_USAGE_FILE = r""    #企业月度总用电量文件
SAVE_RESULT_DIR = r""     #结果保存目录（自动创建）

def a_read_excel(file_path, sheet_name=None):
    if sheet_name is None:
        xl = pd.ExcelFile(file_path)
        return pd.read_excel(file_path, sheet_name=xl.sheet_names[0])
    return pd.read_excel(file_path, sheet_name=sheet_name)

def get_unified_date(day):
    return f"x月{day:02d}日"

def read_usage():
    rows = []
    ent = "未知企业"
    for day in range(1, *):  # *为当月天数+1
        f = os.path.join(USAGE_FILE_DIR, f"当日用电量详情 ({day}).xlsx")
        df = pd.read_excel(f)
        if day == 1:
            ent = df.iloc[0]["用户名称"].strip()
        for h in range(24):
            dt = get_unified_date(day)
            tm = f"{h+1:02d}:00"
            val = float(df.iloc[0][f"{h+1:02d}:00"])
            rows.append([ent, dt, tm, val])
    df = pd.DataFrame(rows, columns=["企业名称","日期","时段","Q_每小时总_t"])
    Q_month = df["Q_每小时总_t"].sum()
    return df, ent, Q_month

#现货价格
def read_spot():
    real = pd.read_excel(SPOT_PRICE_FILE, sheet_name="实时")
    daya = pd.read_excel(SPOT_PRICE_FILE, sheet_name="日前")
    rows = []
    for day in range(1,*):
        for h in range(24):
            dt = get_unified_date(day)
            tm = f"{h+1:02d}:00"
            p_r = float(real.iloc[day-1, h+1])
            p_d = float(daya.iloc[day-1, h+1])
            rows.append([dt, tm, p_r, p_d])
    return pd.DataFrame(rows, columns=["日期","时段","P_实时统一_t","P_日前统一_t"])

#零售价格
def read_retail():
    df = pd.read_excel(RETAIL_PRICE_FILE)
    df["时刻"] = df["时刻"].astype(str).str.replace(" 24:00", " 00:00")
    df["时刻"] = pd.to_datetime(df["时刻"])
    df["日期"] = df["时刻"].dt.day.apply(get_unified_date)
    df["时段"] = df["时刻"].dt.strftime("%H:%M")
    df = df.rename(columns={"价格(元/MWh)":"P_零售市场参考价_t"})
    return df[["日期","时段","P_零售市场参考价_t"]]

#月用电量
def get_Q_month(ent):
    df = pd.read_excel(MONTH_USAGE_FILE)
    row = df[(df["用户名称"].str.strip() == ent) & (df["类型"] == "电量(MWh)")]
    return row[[f"{i}日" for i in range(1,*)]].sum(axis=1).iloc[0]

#合并
def calc():
    usage_df, ent, Q_month = read_usage()
    spot_df = read_spot()
    retail_df = read_retail()
    
    df = usage_df.merge(spot_df, on=["日期","时段"])
    df = df.merge(retail_df, on=["日期","时段"])
    df["Q_实时_t"] = df["Q_每小时总_t"]
    df["Q_日前_t"] = df["Q_每小时总_t"].shift(-1)
    df["Q_中长期_t"] = df["Q_每小时总_t"]
    df["P_中长期_t"] = df["P_零售市场参考价_t"] + AUX_SERVICE_FEE + OPERATION_FEE + CAP_PRICE_FEE
    df["C实时_t"] = df["Q_实时_t"] * df["P_实时统一_t"]
    df["C日前_t"] = df["Q_日前_t"] * (df["P_日前统一_t"].shift(-1) - df["P_实时统一_t"])
    df["C中长期_t"] = df["Q_中长期_t"] * (df["P_中长期_t"] - df["P_实时统一_t"])
    df["C总_t"] = df["C实时_t"] + df["C日前_t"] + df["C中长期_t"]

    C实时合计 = df["C实时_t"].sum()
    C日前合计 = df["C日前_t"].sum()
    C中长期合计 = df["C中长期_t"].sum()
    C总合计 = df["C总_t"].sum()

    平均电价 = C总合计 / Q_month if Q_month !=0 else 0
    
    汇总 = {
        "企业名称": ent,
        "月度总用电量": Q_month,
        "实时电费合计": C实时合计,
        "日前电费合计": C日前合计,
        "中长期电费合计": C中长期合计,
        "总电费": C总合计,
        "平均电价": 平均电价,
    }
    return df, 汇总

#导出Excel
def export(df, 汇总):
    os.makedirs(SAVE_RESULT_DIR, exist_ok=True)
    fname = f"{汇总['企业名称']}_x月电费.xlsx"
    path = os.path.join(SAVE_RESULT_DIR, fname)

    汇总表 = pd.DataFrame(pd.Series(汇总), columns=["数值"])

    with pd.ExcelWriter(path, engine="openpyxl") as f:
        df.to_excel(f, sheet_name="分时电费明细", index=False)
        汇总表.to_excel(f, sheet_name="月度汇总")
    return path

#run
if __name__ == "__main__":
    df, 汇总 = calc()
    path = export(df, 汇总)

    print("计算完成")
    print(f"企业：{汇总['企业名称']}")
    print(f"总用电量：{汇总['月度总用电量']:.2f} MWh")
    print(f"总电费：{汇总['总电费']:.2f} 元")
    print(f"平均电价：{汇总['平均电价']:.2f} 元/MWh")
    print(f"文件已保存：{path}")