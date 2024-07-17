import streamlit as st
import pandas as pd
import openpyxl 
import os

# Load data from Excel file
file_path = "D:\\Interlining Data.xlsx"

@st.cache_data
def load_data():
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()  # Trim spaces from column names
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            "Indent Number", "Stage", "Customer", "Style", "Wash",
            "Content", "GSM", "Structure", "Count_Cons", "Type of construction",
            "Collar Skin", "Collar Patch", "Inner Collar", "Inner NB", "NB Patch",
            "Outer NB", "CF T P", "CF D P", "Top Cuff", "In cuff", "Top SP",
            "Inner SP", "Label Patch", "Moon Patch", "Welt", "Flap"
        ])
        df.to_excel(file_path, index=False)
    return df  # Return the DataFrame, not just "Data Loaded"

# Function to save new data entry to the Excel file
@st.cache_data
def save_data(new_data):
    try:
        df = load_data()  # Load existing data
        df.columns = df.columns.str.strip()  # Trim spaces from column names
        # Concatenate the new data with the existing DataFrame
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
        df.to_excel(file_path, index=False)  # Save to Excel
        st.success("Data saved successfully!")
    except PermissionError:
        st.error("Permission denied: Ensure the file is not open.")
    except Exception as e:
        st.error(f"Error saving data: {e}")

# Create Streamlit app
st.title("Data Entry and Retrieval Dashboard")

# Display the data entry form
st.title("Data Entry")

indent_number = st.number_input("Indent Number", min_value=9999)
stage = st.selectbox("Stage", [" ", "Design", "Development", "FIT", "GFE", "GPT", "GPT,PP", "Mock", "Offer", "Photoshoot", "Pre-Production", "Proto",
                 "Quotation", "Sealer", "Size Set", "SMS"])
customer = st.text_input("Customer")
style = st.text_input("Style")
wash = st.text_input("Wash")
content = st.text_input("Content")
gsm = st.number_input("GSM", min_value=9)
structure = st.selectbox("Structure", [" ", "Corduroy", "Dobby", "Denim", "French Terry", "Herringbone", "Interlock (Knit)", "Jersey",
                                                  "Jacquard", "Knit", "Matt", "Miss Jersey Knit", "Oxford", "Oxford Twill",
                                                  "Pique", "Plain", "Poplin", "Satin", "Seersucker", "Single Jersey", "Twill", "Twill Knit"])
count_cons = st.text_input("Count_Cons")
type_of_construction = st.selectbox("Type of construction", [" ", "Woven", "Knit"])
collar_skin = st.text_input("Collar Skin")
collar_patch = st.text_input("Collar Patch")
inner_collar = st.text_input("Inner Collar")
inner_nb = st.text_input("Inner NB")
nb_patch = st.text_input("NB Patch")
outer_nb = st.text_input("Outer NB")
cf_t_p = st.text_input("CF T P")
cf_d_p = st.text_input("CF D P")
top_cuff = st.text_input("Top Cuff")
in_cuff = st.text_input("In cuff")
top_sp = st.text_input("Top SP")
inner_sp = st.text_input("Inner SP")
label_patch = st.text_input("Label Patch")
moon_patch = st.text_input("Moon patch")
welt = st.text_input("Welt")
flap = st.text_input("Flap")

if st.button("Save Data"):
    new_data = {
        'Indent Number': indent_number,
        'Stage': stage,
        'Customer': customer,
        'Style': style,
        'Wash': wash,
        'Content': content,
        'GSM': gsm,
        'Structure': structure,
        'Count_Cons': count_cons,
        'Type of construction': type_of_construction,
        'Collar Skin': collar_skin,
        'Collar Patch': collar_patch,
        'Inner Collar': inner_collar,
        'Inner NB': inner_nb,
        'NB Patch': nb_patch,
        'Outer NB': outer_nb,
        'CF T P': cf_t_p,
        'CF D P': cf_d_p,
        'Top Cuff': top_cuff,
        'In cuff': in_cuff,
        'Top SP': top_sp,
        'Inner SP': inner_sp,
        'Label Patch': label_patch,
        'Moon patch': moon_patch,
        'Welt': welt,
        'Flap': flap
    }
    save_data(new_data)

# Data Retrieval section
st.header("Data Retrieval")
with st.form("data_retrieval"):
    indent_number_retrieve = st.text_input("Indent Number")
    stage_retrieve = st.selectbox("Stage", [" ", "Design", "Development", "FIT", "GFE", "GPT", "GPT,PP", "Mock", "Offer", "Photoshoot", "Pre-Production", "Proto",
                 "Quotation", "Sealer", "Size Set", "SMS"])
    customer_retrieve = st.text_input("Customer")
    style_retrieve = st.text_input("Style")
    wash_retrieve = st.text_input("Wash")
    content_retrieve = st.text_input("Content")
    gsm_retrieve = st.text_input("GSM")
    structure_retrieve = st.selectbox("Structure", [" ", "Corduroy", "Dobby", "Denim", "French Terry", "Herringbone", "Interlock (Knit)", "Jersey",
                                                  "Jacquard", "Knit", "Matt", "Miss Jersey Knit", "Oxford", "Oxford Twill",
                                                  "Pique", "Plain", "Poplin", "Satin", "Seersucker", "Single Jersey", "Twill", "Twill Knit"])
    type_of_construction_retrieve = st.selectbox("Type of construction", [" ", "Woven", "Knit"])

    submitted = st.form_submit_button("Retrieve")

    if submitted:
        filters = {}
        if indent_number_retrieve:
            filters["Indent Number"] = indent_number_retrieve
        if customer_retrieve:
            filters["Customer"] = customer_retrieve
        if style_retrieve:
            filters["Style"] = style_retrieve
        if wash_retrieve:
            filters["Wash"] = wash_retrieve
        if content_retrieve:
            filters["Content"] = content_retrieve
        if gsm_retrieve:
            filters["GSM"] = gsm_retrieve
        if structure_retrieve:
            filters["Structure"] = structure_retrieve
        if type_of_construction_retrieve:
            filters["Type of construction"] = type_of_construction_retrieve

        df = load_data()  # Load data after filtering

        filtered_df = df
        for key, value in filters.items():
            if value:
                try:
                    if key in ['Indent Number', 'GSM']:  # Check if key is in the list
                        filtered_df = filtered_df.loc[filtered_df[key] == int(value)]
                    else:
                        filtered_df = filtered_df.loc[filtered_df[key] == value]
                except ValueError as ve:
                    st.error(f"ValueError: {ve}")
                except KeyError as ke:
                    st.error(f"KeyError: {ke}")
                except Exception as e:
                    st.error(f"Unexpected error: {e}")

            if filtered_df.empty:
                st.error("No matching records found.")
        else:
            st.write(filtered_df)

if __name__ == "__main__":
    pass
