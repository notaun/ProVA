# test_xlwings_save.py
import xlwings as xw
app = xw.App(visible=True)
wb = app.books.add()
wb.sheets[0]["A1"].value = "Hello"
try:
    wb.save(r"C:\Users\Pc\PyCharmMiscProject\test_save.xlsx")
    print("Saved ok")
except Exception as e:
    print("Save failed:", e)
finally:
    app.quit()
