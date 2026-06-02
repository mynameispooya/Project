import streamlit as st
import pandas as pd
import io

# تنظیمات هدر صفحه
st.set_page_config(page_title="سیستم هوشمند پردازش و اصلاح گزارشات تردد", layout="wide")

st.markdown("""
    <div style="text-align: right; direction: rtl;">
        <h2>سامانه پردازش و تطبیق داده‌های تردد کارمندان 🚀</h2>
        <p>این اپلیکیشن فایل خروجی سیستمی تردد (ناقص) را دریافت کرده و پس از تطبیق با فایل جامع آمار پرسنلی، اطلاعات واحد سازمانی، انتقالی و زیرمجموعه را به آن اضافه می‌کند.</p>
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
    
    # --- ۱. خواندن و پیش‌پردازش فایل مرجع پرسنلی ---
    try:
        if uploaded_master.name.endswith('.csv'):
            df_master = pd.read_csv(uploaded_master)
        else:
            df_master = pd.read_excel(uploaded_master)
    except Exception as e:
        st.error(f"خطا در خواندن فایل مرجع پرسنلی: {e}")
        st.stop()
        
    # --- ۲. خواندن و پیش‌پردازش فایل تردد روزانه (سیستمی) ---
    try:
        # با توجه به اینکه فایل سیستمی حدود ۱۲ ردیف هدر متفرقه دارد، ردیف‌های ابتدایی را رد می‌کنیم
        if uploaded_attendance.name.endswith('.csv'):
            # خواندن موقت برای پیدا کردن ردیف هدر اصلی
            df_att_raw = pd.read_csv(uploaded_attendance, header=None)
        else:
            df_att_raw = pd.read_excel(uploaded_attendance, header=None)
            
        # یافتن ردیفی که حاوی اطلاعات اصلی ستون‌ها است (مثلا کلمه کدپرسنلی یا اطلاعات تردد)
        header_row_index = 12 # پیش‌فرض ردیف ۱۳ اکسل (ایندکس ۱۲)
        for idx, row in df_att_raw.iterrows():
            if row.astype(str).str.contains('کدپرسنلی|کد_پرسنلی').any():
                header_row_index = idx
                break
        
        # بارگذاری مجدد فایل با هدر اصلاح شده
        if uploaded_attendance.name.endswith('.csv'):
            df_attendance = pd.read_csv(uploaded_attendance, skiprows=header_row_index)
        else:
            df_attendance = pd.read_excel(uploaded_attendance, skiprows=header_row_index)
            
        # تمیزکاری نام ستون‌ها (حذف فضاهای خالی احتمالی)
        df_attendance.columns = df_attendance.columns.str.strip()
        df_master.columns = df_master.columns.str.strip()
        
    except Exception as e:
        st.error(f"خطا در پردازش ساختار فایل تردد: {e}")
        st.stop()

    st.success("هر دو فایل با موفقیت بارگذاری و ساختار سنجی شدند.")
    st.divider()
    
    # --- ۳. انتخاب ستون مشترک توسط کاربر ---
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
    
    # دکمه اجرای عملیات ادغام
    if st.button("پردازش و استخراج گزارش نهایی ⚙️"):
        
        # یکسان‌سازی نوع داده به رشته جهت جلوگیری از خطای تطبیق اعداد
        df_master[common_col_master] = df_master[common_col_master].astype(str).str.strip().str.replace('.0', '', regex=False)
        df_attendance[common_col_attendance] = df_attendance[common_col_attendance].astype(str).str.strip().str.replace('.0', '', regex=False)
        
        # ستون‌های مورد نیاز از فایل پرسنلی مرجع
        target_cols = [common_col_master, 'واحد_سازمانی', 'انتقالی_به', 'زیر_مجموعه']
        # بررسی وجود ستون‌ها در فایل مرجع
        available_targets = [col for col in target_cols if col in df_master.columns]
        
        df_master_sliced = df_master[available_targets].drop_duplicates(subset=[common_col_master])
        
        # عملیات ادغام دو جدول (Left Join) برای حفظ تمام ردیف‌های گزارش تردد
        df_final = pd.merge(
            df_attendance,
            df_master_sliced,
            left_on=common_col_attendance,
            right_on=common_col_master,
            how='left'
        )
        
        # ستون‌های درخواستی شما برای خروجی نهایی
        desired_columns = ['کدپرسنلی', 'نام و نام خانوادگی', 'تاریخ', 'روز', 'ورود', 'خروج', 'واحد_سازمانی', 'انتقالی_به', 'زیر_مجموعه']
        
        # بررسی ستون‌های موجود و فیلتر کردن نهایی بر اساس فیلدهای موجود
        final_cols_to_display = [c for c in desired_columns if c in df_final.columns]
        
        df_output = df_final[final_cols_to_display]
        
        # حذف ردیف‌هایی که کاملا خالی هستند یا ردیف‌های مجموع سیستمی که فاقد کد پرسنلی معتبر هستند
        df_output = df_output[df_output[common_col_attendance].notna() & (df_output[common_col_attendance] != 'nan')]
        
        st.markdown("<div style='text-align: right;'><h4>پیش‌نمایش گزارش اصلاح‌شده نهایی (۱۰ ردیف اول)</h4></div>", unsafe_allow_html=True)
        st.dataframe(df_output.head(10), use_container_width=True)
        
        # تبدیل دیتافریم به بافر اکسل برای دانلود ارگونومیک
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_output.to_excel(writer, index=False, sheet_name='گزارش تردد جامع')
            
            # راست‌چین کردن خودکار جدول در اکسل خروجی برای زیباتر شدن فرمت فارسی
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
