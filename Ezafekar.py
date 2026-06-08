import streamlit as st
import pandas as pd
import re

# تنظیمات اولیه صفحه با تم ملایم و راست‌چین برای متن‌ها
st.set_page_config(page_title="محاسبه‌گر تردد و اضافه کار", layout="centered")

# تزریق CSS برای بهبود ظاهر و هماهنگی استایل
st.markdown("""
    <style>
    .reportview-container {
        direction: rtl;
        text-align: right;
    }
    .metric-container {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .metric-value {
        font-size: 24px;
        font-weight: bold;
        color: #1E88E5;
    }
    </style>
""", unsafe_allow_html=True)

def time_to_seconds(time_str):
    """تبدیل رشته زمانی با فرمت HH:MM:SS یا H:MM:SS به ثانیه"""
    if pd.isna(time_str) or not isinstance(time_str, str):
        return 0
    
    time_str = time_str.strip()
    if not time_str or time_str == "-":
        return 0
        
    # بررسی تطابق با فرمت زمان
    match = re.match(r'^(\d+):([0-5]?\d):([0-5]?\d)$', time_str)
    if match:
        hours, minutes, seconds = map(int, match.groups())
        return hours * 3600 + minutes * 60 + seconds
    return 0

def seconds_to_extended_time(total_seconds):
    """تبدیل ثانیه به رشته زمانی تعمیم یافته [h]:mm:ss"""
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

# عنوان اپلیکیشن
st.title("📊 سیستم محاسبه مجموع اضافه کار نهایی")
st.write("فایل اکسل یا CSV تردد خود را در بخش زیر آپلود کنید تا مجموع اضافه کار عادی نهایی محاسبه شود.")

# المان آپلود فایل
uploaded_file = st.file_uploader("انتخاب فایل تردد (Excel یا CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # تشخیص نوع فایل و خواندن آن
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success("فایل با موفقیت بارگذاری شد!")
        
        # نمایش پیش‌نمایش داده‌ها در یک باکس جمع‌وجور
        with st.expander("👁️ مشاهده پیش‌نمایش داده‌های فایل"):
            st.dataframe(df.head(10))
            
        # نام ستون هدف بر اساس ساختار فایل شما
        target_column = 'اضافه کار عادی نهایی'
        
        if target_column in df.columns:
            # اعمال تابع تبدیل به ثانیه روی تک‌تک ردیف‌های ستون هدف
            df['Overtime_Seconds'] = df[target_column].apply(time_to_seconds)
            
            # محاسبه مجموع ثانیه‌ها
            total_seconds = df['Overtime_Seconds'].sum()
            
            # تبدیل مجموع به فرمت [h]:mm:ss
            formatted_total_time = seconds_to_extended_time(total_seconds)
            
            # نمایش خروجی نهایی به صورت کاملاً متمایز
            st.markdown("---")
            st.subheader("⏱️ نتیجه محاسبه نهایی:")
            
            st.markdown(f"""
                <div class="metric-container">
                    <p style="margin-bottom: 5px; color: #555;">مجموع ستون <strong>{target_column}</strong></p>
                    <div class="metric-value" dir="ltr">[{formatted_total_time}]</div>
                </div>
            """, unsafe_allow_html=True)
            
            # نمایش ردیف‌های دارای اضافه کار برای اطمینان کاربر
            active_overtimes = df[df['Overtime_Seconds'] > 0][[target_column]]
            if not active_overtimes.empty:
                with st.expander("📋 مشاهده روزهای دارای اضافه کار"):
                    st.dataframe(active_overtimes)
                    
        else:
            st.error(f"خطا: ستونی با نام 정확 '{target_column}' در فایل یافت نشد. لطفاً نام ستون‌های فایل خود را بررسی کنید.")
            st.info(f"ستون‌های موجود در فایل شما: {', '.join(df.columns)}")
            
    except Exception as e:
        st.error(f"خطایی در پردازش فایل رخ داد: {str(e)}")
