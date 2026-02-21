import pandas as pd
import logging


def save_to_excel(data, filename):
    """حفظ البيانات بصيغة Excel مع تنسيق احترافي (اليمين لليسار وتوسيع الأعمدة)"""
    if not data:
        print(" لا توجد بيانات لحفظها.")
        return

    # التأكد أن الامتداد هو .xlsx وليس .csv
    if not filename.endswith('.xlsx'):
        filename = filename.replace('.csv', '') + '.xlsx'

    try:
        df = pd.DataFrame(data)

        # استخدام ExcelWriter لضبط التنسيق
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='البيانات المستخرجة')

            # الوصول لورقة العمل لضبط الخصائص
            worksheet = writer.sheets['البيانات المستخرجة']

            # 1. ضبط اتجاه الصفحة من اليمين لليسار (RTL)
            worksheet.sheet_view.rightToLeft = True

            # 2. توسيع الأعمدة تلقائياً (Auto-fit) لمنع "قطش" البيانات
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"✅ تم حفظ وتنسيق الملف بنجاح: {filename}")

    except Exception as e:
        logging.error(f" خطأ أثناء حفظ ملف Excel: {e}")
        print(f" فشل الحفظ بصيغة Excel، تأكد من إغلاق الملف إذا كان مفتوحاً.")

