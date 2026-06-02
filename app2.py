import streamlit as st
import pandas as pd
import io

# تنظیمات هدر صفحه
st.set_page_config(page_title="سیستم هوشمند پردازش و اصلاح گزارشات تردد", layout="wide")

st.markdown("""
    <div style="text-align: right; direction: rtl;">
        <h2>سامانه پردازش و تطبیق داده‌های تردد کارمندان (نسخه اصلاح‌شده) 🚀</h2>
        <p>این برنامه با اصلاح هوشمند ردیف‌های ادغام‌شده، تمام ترددهای ناقص پرسنل را در روزهای مختلف به طور کامل استخراج می‌کند.</p>
    </div>
""", unsafe_allow_html=True)

st.divider()

# ایجاد دو ستون برای آپلود همزمان فایل‌ها
col1, col2 = st.columns(2)

with col1:
    st.markdown("<div style='text-align: right;'><b>۱. آپلود فایل پرسنلی مرجع (اکسل استاندارد)</b></div>", unsafe_allow_html=True)
    uploaded_master = st.file_uploader("فایل آمار نهایی را انتخاب کنید (xlsx)", type=["xlsx", "csv"], key="master")

with col2:
    st.markdown("<div style='text-align: right;'><b>۲. آپلود گزارش تردد روزانه (خروجی سیستم)</b></div>", unsafe_allow_html=True)
    uploaded_attendance = st.file_uploader("فایل گزارش تردد روزانه را انتخاب کنید (xlsx)", type=["xlsx", "csv"], key="attendance")

if uploaded_master and uploaded_attendance:
    
    # --- ۱. خواندن فایل مرجع پرسنلی ---
    try:
        if uploaded_master.name.endswith('.csv'):
            df_master = pd.read_csv(uploaded_master)
        else:
            df_master = pd.read_excel(uploaded_master)
    except Exception as e:
        st.error(f"خطا در خواندن فایل مرجع پرسنلی: {e}")
        st.stop()
        
    # --- ۲. خواندن و پاک‌سازی پیشرفته فایل تردد روزانه ---
    try:
        if uploaded_attendance.name.endswith('.csv'):
            df_att_raw = pd.read_csv(uploaded_attendance, header=None)
        else:
            df_att_raw = pd.read_excel(uploaded_attendance, header=None)
            
        # یافتن ردیف هدر اصلی ستون‌ها در فایل خروجی سیستم
        header_row_index = 12 
        for idx, row in df_att_raw.iterrows():
            if row.astype(str).str.contains('کدپرسنلی|کد_پرسنلی').any():
                header_row_index = idx
                break
        
        # بارگذاری مجدد داده‌ها از ردیف هدر به بعد
        if uploaded_attendance.name.endswith('.csv'):
            df_attendance = pd.read_csv(uploaded_attendance, skiprows=header_row_index)
        else:
            df_attendance = pd.read_excel(uploaded_attendance, skiprows=header_row_index)
            
        # تمیزکاری نام ستون‌ها
        df_attendance.columns = df_attendance.columns.str.strip()
        df_master.columns = df_master.columns.str.strip()
        
        # --- حل مشکل ردیف‌های خالی روزهای بعدی (اصلاح ساختار ادغامی راهکاران) ---
        # ابتدا مقادیر خالی یا فضاهای خالی متنی را به NaN واقعی تبدیل می‌کنیم
        df_attendance['کدپرسنلی'] = df_attendance['کدپرسنلی'].astype(str).str.strip()
        df_attendance['کدپرسنلی'] = df_attendance['کدپرسنلی'].replace(['nan', 'NaN', ''], pd.NA)
        
        if 'نام و نام خانوادگی' in df_attendance.columns:
            df_attendance['نام و نام خانوادگی'] = df_attendance['نام و نام خانوادگی'].astype(str).str.strip()
            df_attendance['نام و نام خانوادگی'] = df_attendance['نام و نام خانوادگی'].replace(['nan', 'NaN', ''], pd.NA)
        
        # با استفاده از Forward Fill، کد پرسنلی و نام را برای روزهای بعدی همان فرد کپی می‌کنیم
        df_attendance['کدپرسنلی'] = df_attendance['کدپرسنلی'].ffill()
        if 'نام و نام خانوادگی' in df_attendance.columns:
            df_attendance['نام و نام خانوادگی'] = df_attendance['نام و نام خانوادگی'].ffill()

    except Exception as e:
        st.error(f"خطا در پردازش ساختار فایل تردد: {e}")
        st.stop()

    st.success("هر دو فایل بارگذاری شدند و ساختار سلول‌های خالیِ متوالی اصلاح شد.")
    st.divider()
    
    # --- ۳. تنظیمات نگاشت داده‌ها ---
    st.markdown("<div style='text-align: right;'><h3>تنظیمات نگاشت و تطبیق داده‌ها</h3></div>", unsafe_allow_html=True)
    
    common_col_master = st.selectbox(
        "ستون کلید مشترک در فایل اول (پرسنلی مرجع) را انتخاب کنید:",
        options=df_master.columns,
        index=list(df_master.columns).index('کد_پرسنلی') if 'کد_پرسنلی' in df_master.columns else 0
    )
    
    common_col_attendance = st.selectbox(
        "ستون کلید مشترک در فایل دوم (گزارش تردد) را انتخاب کنید:",
        options=df_attendance.columns,
        index=list(df_attendance.columns).index('کدپرسنلی') if 'کدپرسنلی' in df_attendance.columns else 0
    )
    
    if st.button("پردازش و استخراج گزارش نهایی ⚙️"):
        
        # یکسان‌سازی فرمت کد پرسنلی جهت انطباق ۱۰۰٪
        df_master[common_col_master] = df_master[common_col_master].astype(str).str.strip().str.replace('.0', '', regex=False)
        df_attendance[common_col_attendance] = df_attendance[common_col_attendance].astype(str).str.strip().str.replace('.0', '', regex=False)
        
        # انتخاب ستون‌های درخواستی از فایل مرجع
        target_cols = [common_col_master, 'واحد_سازمانی', 'انتقالی_به', 'زیر_مجموعه']
        available_targets = [col for col in target_cols if col in df_master.columns]
        
        df_master_sliced = df_master[available_targets].drop_duplicates(subset=[common_col_master])
        
        # عملیات ادغام (حالا تمام روزها کد پرسنلی معتبر دارند و نگاشت می‌شوند)
        df_final = pd.merge(
            df_attendance,
            df_master_sliced,
            left_on=common_col_attendance,
            right_on=common_col_master,
            how='left'
        )
        
        # ستون‌های نهایی برای خروجی
        desired_columns = ['کدپرسنلی', 'نام و نام خانوادگی', 'تاریخ', 'روز', 'ورود', 'خروج', 'واحد_سازمانی', 'انتقالی_به', 'زیر_مجموعه']
        final_cols_to_display = [c for c in desired_columns if c in df_final.columns]
        
        df_output = df_final[final_cols_to_display]
        
        # حذف ردیف‌های زائد سیستمی (مثل سطر مجموع که اصالت داده پرسنلی ندارند)
        # سطر مجموع معمولا در ستون تاریخ یا روز مقداری ندارد یا کلمه مجموع دارد، پس بر اساس تاریخ معتبر فیلتر می‌کنیم
        if 'تاریخ' in df_output.columns:
            df_output = df_output[df_output['تاریخ'].notna() & (df_output['تاریخ'].astype(str).str.strip() != 'nan')]
        
        df_output = df_output[df_output[common_col_attendance].notna() & (df_output[common_col_attendance] != 'nan')]
        
        st.markdown("<div style='text-align: right;'><h4>پیش‌نمایش گزارش اصلاح‌شده نهایی (شامل تمام روزهای تردد ناقص کارمندان)</h4></div>", unsafe_allow_html=True)
        st.dataframe(df_output, use_container_width=True)
        
        # تبدیل به بافر اکسل
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_output.to_excel(writer, index=False, sheet_name='گزارش تردد جامع')
            
            # راست‌چین کردن خودکار اکسل خروجی
            workbook  = writer.book
            worksheet = writer.sheets['گزارش تردد جامع']
            worksheet.right_to_left()
            
        buffer.seek(0)
        
        st.divider()
        st.markdown("<div style='text-align: right;'><b>📥 دریافت فایل خروجی:</b></div>", unsafe_allow_html=True)
        st.download_button(
            label="دانلود فایل اکسل گزارش جامع پرسنل",
            data=buffer,
            file_name="گزارش_جامع_تردد_اصلاح_شده.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("لطفاً هر دو فایل اکسل را آپلود کنید تا سیستم پردازش را آغاز کند.")
