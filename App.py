import streamlit as st
from DatabaseManager import excelManager

em = excelManager("dataExcel.xlsx")
options = ["Choose Action", "Insert", "Edit", "Delete"]
choice = st.selectbox("Choose an action:", options)
saveChange = st.checkbox("SaveChanges", value=False)

if choice in ("Edit", "Delete"):
    nim = st.text_input("Enter targeted NIM:", key="targetNim")
    if choice == "Delete":
        if any(str(i).isalpha() for i in nim):
            st.error("Input NIM harus angka semua")
        elif st.button("Delete"):
            if not em.getData("NIM", nim):
                st.error("NIM not found")
            else:
                em.deleteData(nim, saveChange)
                if not em.getData("NIM", nim):
                    st.success("Deleted")

if choice in ("Insert", "Edit"):
    newNim = st.text_input("Enter New NIM:", key="newNim")
    newName = st.text_input("Enter New Name:", key="newName")
    newGrade = st.text_input("Enter New Grade :", key="newGrade")

    if choice == "Edit":
        if st.button("Edit"):
            if any(str(i).isalpha() for i in newNim):
                st.error("Input NIM harus angka semua")
            elif any(str(i).isdigit() for i in newName):
                st.error("Input nama harus alphabet semua")
            elif any(str(i).isalpha() for i in newGrade):
                st.error("Input nilai harus angka semua")
            else:
                if not em.getData("NIM", nim):
                    st.error("NIM not found")
                else:
                    result = em.editData(
                        str(nim),
                        {"NIM": str(newNim).strip(), "Nama": str(newName).strip(), "Nilai": int(newGrade.strip())},
                        saveChange
                    )
                    if em.getData("NIM", newNim):
                        st.success("Edited")
                    else:
                        st.error("Edit failed")

    if choice == "Insert":
        if st.button("Insert"):
            if any(str(i).isalpha() for i in newNim):
                st.error("Input NIM harus angka semua")
            elif any(str(i).isdigit() for i in newName):
                st.error("Input nama harus alphabet semua")
            elif any(str(i).isalpha() for i in newGrade):
                st.error("Input nilai harus angka semua")
            else:
                if em.getData("NIM", newNim):
                    st.error("NIM already exists")
                else:
                    em.insertData(
                        {"NIM": str(newNim).strip(), "Nama": str(newName).strip(), "Nilai": int(newGrade.strip())},
                        saveChange
                    )
                    if em.getData("NIM", newNim):
                        st.success("Inserted")
                    else:
                        st.error("Insert failed")

option = ["None", ">", "<", "=", "<=", ">="]
filterSelectBox = st.selectbox("Opsi Filter: ", option)

if filterSelectBox == "None":
    st.table(em.getDataFrame())
else:
    targetFilterColumn = st.selectbox("Target Column", ["NIM", "Nilai"])
    filter = st.text_input("Filter Nilai")

    if filter:
        try:
            filter_val = int(filter)
            df = em.getDataFrame()
            if filterSelectBox == ">":
                st.table(df[df[targetFilterColumn] > filter_val])
            elif filterSelectBox == "<":
                st.table(df[df[targetFilterColumn] < filter_val])
            elif filterSelectBox == "=":
                st.table(df[df[targetFilterColumn] == filter_val])
            elif filterSelectBox == "<=":
                st.table(df[df[targetFilterColumn] <= filter_val])
            elif filterSelectBox == ">=":
                st.table(df[df[targetFilterColumn] >= filter_val])
        except ValueError:
            st.error("Filter harus berupa angka")