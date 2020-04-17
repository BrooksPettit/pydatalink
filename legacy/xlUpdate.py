def xlUpdate(full_path,quit_xl=False):
    import win32com.client as win32
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    workbook = excel.Workbooks.Open(full_path)
    # this must be the absolute path (r'C:/abc/def/ghi')
    workbook.Save()
    workbook.Close()
    if quit_xl:
        excel.Quit()
