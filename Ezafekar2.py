import streamlit as st
import pandas as pd
import re

# تنظیمات اولیه صفحه
st.set_page_config(page_title="محاسبه‌گر تردد و اضافه کار", layout="centered")

st.markdown("""
    <style>
    .reportview-container { direction: rtl; text-align: right; }
    .metric-container {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-top: 20px;
    }
    .metric-value { font-size: 28px; font-weight: bold; color: #2E7D32; }
    </style>
""", unsafe_allow_html=True)

def fa_to_en_numbers(text):
    """تبدیل اعداد فارسی و عربی به انگلیسی برای پردازش صحیح زمان"""
    if not isinstance(text, str):
        return text
    fa_digits = "۰۱۲۳۴۵۶۷۸۹"
    ar_digits = "٠١٢٣٤٥٦٧٨٩"
    en_digits = "0123456789"
    
    # نقشه تبدیل
    translation_table = str.maketrans(fa_digits + ar_digits, en_digits * 2)
    return text.translate(translation_table)

def time_to_seconds(time_str):
    """تبدیل رشته زمانی به ثانیه با پشتیبانی از فرمت‌های مختلف خروجی سیستم‌های تردد"""
    if pd.isna(time_str):
        return 0
    
    # تبدیل به رشته و پاکسازی
    time_str = fa_to_en_numbers(str(time_str)).strip()
    
    if not time_str or time_str in ["-", "nan", "null"]:
        return 0
        
    # الگوی بررسی فرمت HH:MM:SS یا H:MM:SS
    match = re.match(r'^(\d+):([0-5]?\d):([0-5]?\d)$', time_str)
    if match:
        hours, minutes, seconds = map(int, match.groups())
        return hours * 3600 + minutes * 60 + seconds
        
    # الگوی دوم در صورتی که ثانیه ثبت نشده باشد HH:MM
    match_short = re.match(r'^(\d+):([0-5]?\d)$', time_str)
    if match_short:
        hours, minutes = map(int, match_short.groups())
        return hours * 3600 + minutes * 60
        
    return 0

def seconds_to_extended_time(total_seconds):
    """تبدیل ثانیه به رشته زمانی [h]:mm:ss"""
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

st.title("📊 سیستم محاسبات تردد راهکاران")
st.write("فایل خروجی تردد خود را در کادر زیر بارگذاری کنید.")

uploaded_file = st.file_uploader("انتخاب فایل اکسل یا CSV", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # بارگذاری فایل
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success("فایل با موفقیت بارگذاری شد.")
        
        # پیدا کردن ستون داینامیک (بررسی وجود کلمه اضافه کار عادی)
        target_column = None
        for col in df.columns:
            if "اضافه کار عادی" in str(col):
                target_column = col
                break
                
        if target_column:
            # کپی از دیتاست اصلی جهت پاکسازی ردیف‌های توتال خود اکسل
            # معمولاً ردیف‌های نهایی در ستون‌های تاریخ یا کد پرسنلی خالی یا حاوی متن "مجموع" هستند
            # برای امنیت بیشتر، ردیف‌هایی که در این ستون زمان معتبر ندارند یا مقدار متنی متفرقه دارند فیلتر می‌شوند
            
            df['clean_time_seconds'] = df[target_column].apply(time_to_seconds)
            
            # محاسبه مجموع کل ثانیه‌ها
            total_seconds = df['clean_time_seconds'].sum()
            
            # تبدیل به فرمت درخواستی
            formatted_time = seconds_to_extended_time(total_seconds)
            
            # رندر کردن جدول داده‌ها برای اطمینان کاربر
            with st.expander("👁️ مشاهده جدول داده‌های پردازش شده"):
                st.dataframe(df[[target_column, 'clean_time_seconds']])
            
            # نمایش خروجی نهایی
            st.markdown("---")
            st.markdown(f"""
                <div class="metric-container">
                    <p style="margin-bottom: 5px; color: #333; font-size: 16px;">
                        مجموع نهایی ستون <strong>«{target_column}»</strong>
                    </p>
                    <div class="metric-value" dir="ltr">[{formatted_time}]</div>
                </div>
            """, unsafe_allow_html=True)
            
        else:
            st.error("خطا: ستونی که شامل عبارت «اضافه کار عادی» باشد در فایل شما پیدا نشد.")
            st.info(f"ستون‌های یافت شده در فایل: {', '.join(df.columns)}")
            
    except Exception as e:
        st.error(f"خطا در پردازش فایل: {str(e)}")
