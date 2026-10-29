import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment







df = pd.read_csv("hotel_bookings_practice_fixed.csv")

#print(df.head())

#print(df.shape)

#print(df.isnull().sum())
print("Original:", len(df))

before = len(df)

duplicates = df[df.duplicated()].copy()

df = df.drop_duplicates()

removed = df[
    df[["Check_In", "Check_Out", "Full_Name"]].isna().all(axis=1)
].copy()

df = df.dropna(
    subset=["Check_In", "Check_Out", "Full_Name"],
    how='all'
)


print("remaining",len(df))




df['Check_In'] = pd.to_datetime(
    df["Check_In"],
    format='mixed',
    errors='coerce'
)

df['Check_Out'] = pd.to_datetime(
    df["Check_Out"],
    format='mixed',
    errors='coerce'
)

#print(df[['Check_In','Check_Out']].head())

#print(df[['Check_In','Check_Out']].isnull().sum())

#print(df[df['Check_In'].isna()])



missing = df[df['Check_In'].isna()].copy()

missing['Price_USD'] = pd.to_numeric(
    missing['Price_USD'],
    errors='coerce'
)


missing['Nights'] = (missing['Price_USD'] / 40).astype(int)

missing["Check_Out"] = pd.to_datetime(missing["Check_Out"])

missing["Estimated_ci"] = (
    missing["Check_Out"] -
    pd.to_timedelta(missing["Nights"], unit="D")
)


df['Nights'] = (df['Check_Out'] - df['Check_In']).dt.days



print("missing both dates removed: ", before - len(df))





invalid = df[df['Nights'] <= 0].copy()

valid = df[df['Nights'] > 0].copy()


valid["Expected_price"] = valid['Nights'] * 40


valid['Price_USD'] = pd.to_numeric(valid['Price_USD'], errors='coerce')

valid['Price_was_missing'] = valid['Price_USD'].isna()
valid['Price_USD']= valid['Price_USD'].fillna(valid["Expected_price"])


wrong_values = valid[
    valid['Expected_price'] != valid['Price_USD']
]


wrong_values = wrong_values.copy()

wrong_values['Price_Difference'] = (
    wrong_values['Price_USD'] - wrong_values['Expected_price']
)


long_stays = valid[valid['Nights'] > 30].copy()

print("total bookings: " , len(df))

print("valid bookings: ", len(valid))
print("invalid bookings: ", len(invalid))
print("missing bookings: ", len(missing))

total_checked  = len(valid) + len(invalid) + len(missing)

if total_checked == len(df):
    print("all bookings are accounted for!")

else:
    print("warning: some bookings are not accounted for!")    

invalid.to_excel("invalid_bookings.xlsx",index=False)

long_stays.to_excel("long_stays.xlsx",index=False)


summary = pd.DataFrame({
    "Caregory" : [
        "Original bookings",
        "Duplicates",
        "Removed missing Fields",
        "Valid bookings",
        "Invalid Dates",
        "Missing Check IN",
        "Long Stays",
        "Wrong Prices",
        "Estimated Prices"
        
    ],

    "Count" : [
        before,
        len(duplicates),
        len(removed),
        len(valid),
        len(invalid),
        len(missing),
        len(long_stays),
        len(wrong_values),
        valid["Price_was_missing"].sum()
    ]
})

print(summary)


original_revenue = valid.loc[
    ~valid["Price_was_missing"], "Price_USD"
].sum()


estimated_revenue = valid.loc[
    valid["Price_was_missing"], "Price_USD"
].sum()

total_revenue = original_revenue + estimated_revenue


print("Original Prices Total:", original_revenue)
print("Estimated Prices Total:", estimated_revenue)
print("Combined Total:", total_revenue)



missing_revenue = missing["Price_USD"].sum()
all_revenue = total_revenue + missing_revenue

print("Valid Bookings Revenue:", total_revenue)
print("Missing Check-In Revenue:", missing_revenue)
print("All Bookings Revenue:", all_revenue)


financial_summary = pd.DataFrame({
    "Category": [
        "Original Prices Total",
        "Estimated Prices Total",
        "Valid Bookings Revenue",
        "Missing Check-In Revenue",
        "All Bookings Revenue"
    ],

    "Amount_USD": [
        original_revenue,
        estimated_revenue,
        total_revenue,
        missing_revenue,
        all_revenue
    ]
})














with pd.ExcelWriter(
    "Hotel_Report_v2.xlsx",
    engine="openpyxl",
    date_format="YYYY-MM-DD",
    datetime_format="YYYY-MM-DD"
) as writer:

    reports = {
    "Summary": summary,
    "Financial_Summary": financial_summary,
    "Valid_Bookings": valid,
    "Invalid_Dates": invalid,
    "Wrong_Prices": wrong_values,
    "Missing_CheckIn": missing,
    "Long_Stays": long_stays,
    "Duplicates": duplicates,
    "Removed_Rows": removed
}

    for sheet_name, data in reports.items():

        data.to_excel(writer, sheet_name=sheet_name, index=False)

        worksheet = writer.sheets[sheet_name]

        for cell in worksheet[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                start_color="1F4E78",
                end_color="1F4E78",
                fill_type="solid"
            )
            cell.alignment = Alignment(
                horizontal="center"
            )
        for column in worksheet.columns:
            letter = column[0].column_letter
            worksheet.column_dimensions[letter].width = 22

            if column[0].value in ["Check_In", "Check_Out", "Estimated_ci"]:
                for cell in column[1:]:
                    cell.number_format = "yyyy-mm-dd"

        if sheet_name == "Valid_Bookings":
            price_column = list(data.columns).index("Price_USD") + 1

            for row_number , was_missing in enumerate(data["Price_was_missing"],start=2):

                if was_missing:
                    cell = worksheet.cell(row=row_number,column=price_column)
                    cell.number_format = '0.00" Estemated"'
























