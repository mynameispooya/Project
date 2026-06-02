import streamlit as st
import pandas as pd
import io
import zipfile

# تنظیمات صفحه و استایل راست‌چین
st.set_page_config(page_title="سیستم جامع و هوشمند پردازش تردد", layout="wide")

st.markdown("""
    <div style="text-align: right; direction: rtl;">
        <h2>سامانه جامع پردازش، تطبیق و تفکیک گزارشات تردد کارمندان 🚀</h2>
        <p>این برنامه تمام ستون‌های گزارش تردد (تاریخ، روز، ورود، خروج، گیت) را استخراج کرده و پس از نگاشت با فایل مرجع، خروجی‌های درخواستی را به همراه واحدهای سازمانی تولید می‌کند.</p>
    </div>
""", unsafe_allow_html=True)

st.divider()

# ایجاد دو ستون برای آپلود فایل‌ها
col1, col2 = st.columns(2)

with col1:
    st.markdown("<div style='text-align: right;'><b>۱. آپلود فایل پرسنلی مرجع (اکسل استاندارد آمار نهایی)</b></div>", unsafe_allow_html=True)
    uploaded_master = st.file_uploader("فایل آمار نهایی پرسنل را انتخاب کنید (xlsx)", type=["xlsx", "csv"], key="master")

with col2:
    st.markdown("<div style='text-align: right;'><b>۲. آپلود گزارش تردد روزانه (خروجی سیستم راهکاران)</b></div>", unsafe_allow_html=True)
    uploaded_attendance = st.file_uploader("فایل گزارش تردد روزانه را انتخاب کنید (xlsx)", type=["xlsx", "csv"], key="attendance")

if uploaded_master and uploaded_attendance:
    
    # --- ۱. خواندن فایل مرجع پرسنلی ---
    try:
        if uploaded_master.name.endswith('.csv'):
            df_master = pd.read_csv(uploaded_master)
        else:
            df_master = pd.read_excel(uploaded_master)
        df_master.columns = df_master.columns.str.strip()
    except Exception as e:
        st.error(f"خطا در خواندن فایل مرجع پرسنلی: {e}")
        st.stop()
        
    # --- ۲. خواندن و بازسازی هدرهای چندسطحی فایل تردد روزانه ---
    try:
        if uploaded_attendance.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_attendance, header=None)
        else:
            df_raw = pd.read_excel(uploaded_attendance, header=None)
            
        # یافتن ردیف اصلی اطلاعات تردد (ردیف ۱۲ یا همان ردیف ۱۳ اکسل)
        header_idx = 12
        for idx, row in df_raw.iterrows():
            if row.astype(str).str.contains('کدپرسنلی|کد_پرسنلی').any():
                header_idx = idx
                break
                
        # استخراج ردیف اول هدر و ردیف دوم هدر برای ترکیب لایه‌ها
        row1 = df_raw.iloc[header_idx].astype(str).str.strip().tolist()
        row2 = df_raw.iloc[header_idx + 1].astype(str).str.strip().tolist()
        
        # ساخت نام ستون‌های ترکیبی و یکتا
        combined_headers = []
        for r1, r2 in zip(row1, row2):
            r1_clean = "" if r1 in ['nan', 'None'] else r1
            r2_clean = "" if r2 in ['nan', 'None'] else r2
            
            if r1_clean and r2_clean:
                combined_headers.append(f"{r1_clean}_{r2_clean}")
            elif r1_clean:
                combined_headers.append(r1_clean)
            elif r2_clean:
                combined_headers.append(r2_clean)
            else:
                combined_headers.append("نامشخص")
                
        # بارگذاری داتا از ردیف بعد از هدرها
        if uploaded_attendance.name.endswith('.csv'):
            df_attendance = pd.read_csv(uploaded_attendance, skiprows=header_idx + 2, header=None)
        else:
            df_attendance = pd.read_excel(uploaded_attendance, skiprows=header_idx + 2, header=None)
            
        # تخصیص هدرهای اصلاح شده به دیتافریم
        df_attendance.columns = combined_headers[:len(df_attendance.columns)]
        
        # یکسان‌سازی دقیق ستون‌های کلیدی و اطلاعات تردد
        for col in df_attendance.columns:
            if 'کدپرسنلی' in col or 'کد_پرسنلی' in col:
                df_attendance.rename(columns={col: 'کدپرسنلی'}, inplace=True)
            if 'نام و نام خانوادگی' in col:
                df_attendance.rename(columns={col: 'نام و نام خانوادگی'}, inplace=True)
            if 'تاریخ' in col:
                df_attendance.rename(columns={col: 'تاریخ'}, inplace=True)
            if 'روز' in col:
                df_attendance.rename(columns={col: 'روز'}, inplace=True)
            if 'ورود' in col:
                df_attendance.rename(columns={col: 'ورود'}, inplace=True)
            if 'خروج' in col:
                df_attendance.rename(columns={col: 'خروج'}, inplace=True)
            if 'گیت' in col:
                df_attendance.rename(columns={col: 'گیت'}, inplace=True)

        # اعمال متد Forward Fill برای پر کردن ردیف‌های خالی روزهای بعدی پرسنل
        df_attendance['کدپرسنلی'] = df_attendance['کدپرسنلی'].astype(str).str.strip().replace(['nan', 'NaN', ''], pd.NA)
        if 'نام و نام خانوادگی' in df_attendance.columns:
            df_attendance['نام و نام خانوادگی'] = df_attendance['نام و نام خانوادگی'].astype(str).str.strip().replace(['nan', 'NaN', ''], pd.NA)
            
        df_attendance['کدپرسنلی'] = df_attendance['کدپرسنلی'].ffill()
        if 'نام و نام خانوادگی' in df_attendance.columns:
            df_attendance['نام و نام خانوادگی'] = df_attendance['نام و نام خانوادگی'].ffill()

    except Exception as e:
        st.error(f"خطا در مدل‌سازی ستون‌ها و هدرهای ترکیبی: {e}")
        st.stop()

    st.success("ساختار ستون‌های فایل تردد و فایل مرجع با موفقیت همترازی شد.")
    st.divider()
    
    # --- ۳. تنظیمات نگاشت و کلیدهای مشترک ---
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
        
        # یکسان‌سازی فرمت کد پرسنلی برای انطباق دقیق
        df_master[common_col_master] = df_master[common_col_master].astype(str).str.strip().str.replace('.0', '', regex=False)
        df_attendance[common_col_attendance] = df_attendance[common_col_attendance].astype(str).str.strip().str.replace('.0', '', regex=False)
        
        # انتخاب فیلدهای هدف از فایل مرجع
        target_cols = [common_col_master, 'واحد_سازمانی', 'انتقالی_به', 'زیر_مجموعه']
        available_targets = [col for col in target_cols if col in df_master.columns]
        
        df_master_sliced = df_master[available_targets].drop_duplicates(subset=[common_col_master])
        
        # ادغام جداول (Left Join)
        df_final = pd.merge(
            df_attendance,
            df_master_sliced,
            left_on=common_col_attendance,
            right_on=common_col_master,
            how='left'
        )
        
        # لیست دقیق ستون‌های خروجی با تضمین وجود اطلاعات تردد روزانه و پرسنلی مرجع
        desired_columns = ['کدپرسنلی', 'نام و نام خانوادگی', 'تاریخ', 'روز', 'ورود', 'خروج', 'گیت', 'واحد_سازمانی', 'انتقالی_به', 'زیر_مجموعه']
        final_cols_to_display = [c for c in desired_columns if c in df_final.columns]
        
        df_output = df_final[final_cols_to_display]
        
        # تمیزکاری نهایی ردیف‌های کمکی و زائد سیستم
        if 'تاریخ' in df_output.columns:
            df_output = df_output[df_output['تاریخ'].notna() & (df_output['تاریخ'].astype(str).str.strip() != 'nan')]
        df_output = df_output[df_output[common_col_attendance].notna() & (df_output[common_col_attendance] != 'nan')]
        
        st.markdown("<div style='text-align: right;'><h4>پیش‌نمایش گزارش جامع (شامل ستون‌های تاریخ، روز، ورود، خروج، گیت و واحدها)</h4></div>", unsafe_allow_html=True)
        st.dataframe(df_output, use_container_width=True)
        
        # --- خروجی اول: دانلود یکجای کل فایل ---
        buffer_all = io.BytesIO()
        with pd.ExcelWriter(buffer_all, engine='xlsxwriter') as writer:
            df_output.to_excel(writer, index=False, sheet_name='گزارش تردد جامع')
            writer.sheets['گزارش تردد جامع'].right_to_left()
        buffer_all.seek(0)
        
        st.divider()
        st.markdown("<div style='text-align: right;'><h3>📥 بخش دانلود خروجی‌ها</h3></div>", unsafe_allow_html=True)
        
        c_down1, c_down2 = st.columns(2)
        
        with c_down1:
            st.download_button(
                label="دانلود فایل اکسل یکجای گزارش جامع",
                data=buffer_all,
                file_name="گزارش_جامع_تردد_کارمندان.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        # --- خروجی دوم: تفکیک هوشمند بر اساس واحد سازمانی و ساخت ZIP ---
        with c_down2:
            if 'واحد_سازمانی' in df_output.columns:
                df_output['واحد_سازمانی'] = df_output['واحد_سازمانی'].fillna('نامشخص')
                unique_departments = df_output['واحد_سازمانی'].unique()
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for dept in unique_departments:
                        df_dept = df_output[df_output['واحد_سازمانی'] == dept]
                        
                        excel_buffer = io.BytesIO()
                        sheet_name_clean = str(dept)[:31]
                        
                        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                            df_dept.to_excel(writer, index=False, sheet_name=sheet_name_clean)
                            writer.sheets[sheet_name_clean].right_to_left()
                        excel_buffer.seek(0)
                        
                        clean_dept_name = str(dept).replace('/', '_').replace('\\', '_')
                        zip_file.writestr(f"گزارش_تردد_واحد_{clean_dept_name}.xlsx", excel_buffer.getvalue())
                
                zip_buffer.seek(0)
                
                st.download_button(
                    label="⚡ دانلود فایل‌های اکسل به تفکیک واحد سازمانی (فایل ZIP)",
                    data=zip_buffer,
                    file_name="گزارشات_تردد_به_تفکیک_واحد.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                st.warning("ستون 'واحد_سازمانی' برای تفکیک دیتا یافت نشد.")

else:
    st.info("لطفاً هر دو فایل اکسل را آپلود کنید تا سیستم پردازش را آغاز کند.")
