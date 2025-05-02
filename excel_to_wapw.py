
import pandas as pd
import xml.etree.ElementTree as ET

# تحميل بيانات Excel
excel_path = "payroll_template.xlsx"  # قم بتحديث المسار حسب الحاجة
df = pd.read_excel(excel_path)

# إنشاء جذر الملف XML
root = ET.Element("WaPWCPR")

# المشروع
projectIntent = ET.SubElement(root, "projectIntent")
ET.SubElement(projectIntent, "intentId").text = "123456"  # يمكن تعديله حسب الحاجة

# أسبوع الرواتب
payroll = ET.SubElement(root, "payroll")
payrollWeek = ET.SubElement(payroll, "payrollWeek")
ET.SubElement(payrollWeek, "endOfWeekDate").text = df['endOfWeekDate'].iloc[0].strftime('%Y-%m-%d')
ET.SubElement(payrollWeek, "noWorkPerformFlag").text = "false"
employees = ET.SubElement(payrollWeek, "employees")

# لكل موظف
for _, row in df.iterrows():
    employee = ET.SubElement(employees, "employee")
    for tag in ['firstName', 'midName', 'lastName', 'ssn', 'ethnicity', 'gender', 'veteranStatus',
                'address1', 'address2', 'city', 'state', 'zip', 'grossPay', 'fica', 'taxWitholding']:
        ET.SubElement(employee, tag).text = str(row[tag]) if pd.notnull(row[tag]) else ""

    # الخصومات
    otherDeductions = ET.SubElement(employee, "otherDeductions")
    for i in [1, 2]:
        name = row.get(f'deduction{i}Name')
        amt = row.get(f'deduction{i}Amt')
        if pd.notnull(name) and pd.notnull(amt):
            deduction = ET.SubElement(otherDeductions, "otherDeduction")
            ET.SubElement(deduction, "deductionName").text = str(name)
            ET.SubElement(deduction, "deductionHourlyAmt").text = f"{amt:.2f}"

    # الأجور والساعات
    tradeHoursWages = ET.SubElement(employee, "tradeHoursWages")
    wage = ET.SubElement(tradeHoursWages, "tradeHoursWage")
    for tag in ['trade', 'jobClass', 'tradeNotes', 'county', 'regularHourRateAmt', 'overtimeHourRateAmt',
                'doubletimeHourRateAmt', 'hourlyPensionRateAmt', 'hourlyMedicalAmt',
                'hourlyVacationAmt', 'hourlyHolidayAmt', 'apprenticeBenefitAmt', 'apprenticeFlg',
                'apprenticeId', 'apprenticeState', 'apprenticeOccpnName', 'apprenticeStepName',
                'apprenticeStepBeginHours', 'apprenticeStepEndHours']:
        ET.SubElement(wage, tag).text = str(row[tag]) if pd.notnull(row[tag]) else ""

    for i in range(1, 8):
        ET.SubElement(wage, f'regularDay{i}Hours').text = str(row.get(f'regularDay{i}Hours', 0.0))
        ET.SubElement(wage, f'overtimeDay{i}Hours').text = str(row.get(f'overtimeDay{i}Hours', 0.0))
        ET.SubElement(wage, f'doubletimeDay{i}Hours').text = str(row.get(f'doubletimeDay{i}Hours', 0.0))

    # الفوائد
    tradeBenefits = ET.SubElement(wage, "tradeBenefits")
    for i in [1, 2]:
        name = row.get(f'benefit{i}Name')
        amt = row.get(f'benefit{i}Amt')
        if pd.notnull(name) and pd.notnull(amt):
            benefit = ET.SubElement(tradeBenefits, "tradeBenefit")
            ET.SubElement(benefit, "benefitHourlyName").text = str(name)
            ET.SubElement(benefit, "benefitHourlyAmt").text = f"{amt:.2f}"

# حفظ الملف XML
tree = ET.ElementTree(root)
tree.write("output_payroll.xml", encoding="utf-8", xml_declaration=True)
print("✅ XML file created: output_payroll.xml")
