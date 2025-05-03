import streamlit as st
import os
import commission_reconciliation
import bcbs_commission_reconciliation

# Set page configuration
st.set_page_config(page_title="CSV Reconciliation Tool", page_icon="📊")

# App title and description
st.title("Statement Reconciliation Tool")
st.write(
    "Upload client list and CSV statements files to create a master reconciliation."
)

# Company selector
selected_company = st.selectbox("Select the company", options=list(["UHC", "BCBS"]))

# File uploader for client list
client_list_uploaded = st.file_uploader(
    "Upload your XLSX client list", type=["xlsx"], accept_multiple_files=False
)

# File uploader for multiple CSV files
file_types = {"UHC": "CSV", "BCBS": "XLSX"}
file_type = file_types[selected_company]
statements_uploaded = st.file_uploader(
    f"Upload your {file_type} statements", type=[file_type], accept_multiple_files=True
)

# Month selector for last statement month
selected_month = st.selectbox(
    "Select the month of the last statement", options=list(range(1, 13))
)


# Process files when uploaded
if statements_uploaded and client_list_uploaded:
    st.write(f"Selected {len(statements_uploaded)} files")

    # Display the files
    for file in statements_uploaded:
        st.success(f"Ready to process: {file.name}")

    # Button to run reconciliation
    if st.button("Run Reconciliation"):
        try:
            st.write("### Running reconciliation process...")

            # Get the paths of the uploaded files
            file_paths = [file.name for file in statements_uploaded]

            # Create the master conciliation excel
            print(type(selected_month))
            if selected_company == "UHC":
                output_path = commission_reconciliation.main(
                    file_paths,
                    client_list_uploaded.name,
                    selected_month,
                    selected_company,
                )
            elif selected_company == "BCBS":
                output_path = bcbs_commission_reconciliation.main(
                    file_paths,
                    client_list_uploaded.name,
                    selected_month,
                    selected_company,
                )
            # For demonstration purposes
            st.write("Creating master conciliation file:")
            for path in file_paths:
                st.code(path)

            # Show success message
            st.success(f"Reconciliation completed successfully: {output_path}")

            # download button here
            if os.path.exists(output_path):
                with open(output_path, "rb") as file:
                    btn = st.download_button(
                        label=f"Download {output_path}",
                        data=file,
                        file_name=os.path.basename(output_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            else:
                st.warning(
                    "Output file not found. Please check the reconciliation process."
                )

        except Exception as e:
            st.error(f"Error during reconciliation: {e}")

# Instructions section
with st.expander("Instructions"):
    st.markdown(
        """
    1. Upload client list and your CSV statement files using the file uploader above
    2. Select the month of the last statement
    2. Click the 'Run Reconciliation' button
    3. The reconciliation process will run 
    4. Download the results when processing is complete
    """
    )

